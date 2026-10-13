# -*- coding: utf-8 -*-
"""一次性数据操作脚本：把指定编号范围内的项目合同「所属部门」由 A 科室调整为 B 科室。

本次用途（2026-10-13）：
    编号 266301~266314 的项目合同：湘潭市机关事务管理局后勤中心-物业组 → 后勤中心-工程组

用法（默认预演，不写库）：
    python tests/fix_dept_range.py --db "data/合同台账.db"                         # 预演
    python tests/fix_dept_range.py --db "data/合同台账.db" --run                   # 执行
    python tests/fix_dept_range.py --db "data/合同台账.db" --run --i-know-this-is-prod

安全设计（与 D-15 清洗脚本一致）：
  ① --db 必填；② 目标若为 data/合同台账.db 需 --i-know-this-is-prod；
  ③ --run 前自动整库备份；④ 仅改「当前部门 == --from」的记录，其余跳过并告警；
  ⑤ 目标部门必须在 departments 中存在（否则外键/触发器会失败，脚本提前拒绝）；
  ⑥ 单事务提交，失败回滚；⑦ 写 field_change_logs + operation_logs 审计留痕。
"""
import argparse
import datetime
import os
import shutil
import sqlite3
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROD_DB = os.path.abspath(os.path.join(ROOT, "data", "合同台账.db"))
BASE = "湘潭市机关事务管理局后勤中心"


def main():
    ap = argparse.ArgumentParser(description="一次性数据操作：调整项目合同所属部门（默认 dry-run）")
    ap.add_argument("--db", required=True, help="目标 SQLite 库路径")
    ap.add_argument("--no-min", default="266301", help="编号范围下界（含）")
    ap.add_argument("--no-max", default="266314", help="编号范围上界（含）")
    ap.add_argument("--from", dest="src", default=BASE + "-物业组", help="原所属部门（完整字符串）")
    ap.add_argument("--to", dest="dst", default=BASE + "-工程组", help="新所属部门（完整字符串）")
    ap.add_argument("--run", action="store_true", help="真正执行（缺省仅预演）")
    ap.add_argument("--i-know-this-is-prod", action="store_true", help="允许对生产库执行（危险）")
    args = ap.parse_args()

    db = os.path.abspath(args.db)
    if not os.path.exists(db):
        print("[ERROR] 目标库不存在：%s" % db)
        return 2
    if db == PROD_DB and not args.i_know_this_is_prod:
        print("[REFUSED] 目标指向生产库，请先复制到 test_data/ 演练；确需直连请加 --i-know-this-is-prod")
        return 3

    con = sqlite3.connect(db, timeout=20)
    try:
        con.execute("PRAGMA foreign_keys=ON")
        # 目标部门必须存在（departments 为 (单位名称, 科室名称) 复合主键）
        unit, _, dept = args.dst.partition("-")
        ok = con.execute('SELECT COUNT(*) FROM departments WHERE "单位名称"=? AND "科室名称"=?',
                         (unit, dept)).fetchone()[0]
        print("目标库：%s；模式：%s" % (db, "【执行】" if args.run else "【预演 dry-run】"))
        print("调整：%s  →  %s　（编号 %s ~ %s）" % (args.src, args.dst, args.no_min, args.no_max))
        if not ok:
            print("[REFUSED] 目标部门不存在于 departments：%s" % args.dst)
            return 3
        print("目标部门存在性校验：通过（%s / %s）" % (unit, dept))
        print("-" * 100)

        rows = con.execute(
            'SELECT id, "编号", "项目名称", "所属部门", "所属单位", "所属科室", "合同金额（万元）", "是否归档" '
            'FROM project_contracts WHERE "编号">=? AND "编号"<=? ORDER BY "编号"',
            (args.no_min, args.no_max)).fetchall()
        hit = [r for r in rows if str(r[3]).strip() == args.src]
        skip = [r for r in rows if str(r[3]).strip() != args.src]
        print("范围内共 %d 条；命中「%s」%d 条；跳过 %d 条" % (len(rows), args.src, len(hit), len(skip)))
        print("| 编号 | 项目名称 | 当前部门 | 所属单位/科室 | 金额(万元) | 归档 |")
        print("|---|---|---|---|---|---|")
        for r in rows:
            mark = "" if str(r[3]).strip() == args.src else "（跳过：部门不符）"
            print("| %s | %s | %s %s | %s/%s | %s | %s |" % (
                r[1], str(r[2])[:24], r[3], mark, r[4], r[5],
                ("%.2f" % float(r[6])) if r[6] not in (None, "") else "-", r[7] or "-"))
        money = sum(float(r[6]) for r in hit if r[6] not in (None, ""))
        print("-" * 100)
        print("将调整 %d 条；合同金额合计 %.2f 万元（该金额将从「%s」转移到「%s」的部门统计）" % (
            len(hit), money, args.src.split("-")[-1], args.dst.split("-")[-1]))

        if not args.run:
            print("\n[DRY-RUN] 未做任何修改。确认无误后加 --run 执行。")
            return 0
        if not hit:
            print("\n[SKIP] 没有需要调整的记录。")
            return 0

        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        bak = os.path.join(os.path.dirname(db), "所属部门调整前_%s.db" % ts)
        shutil.copyfile(db, bak)
        print("\n[BACKUP] 已备份：%s" % bak)

        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        con.execute("BEGIN")
        try:
            n = 0
            for r in hit:
                con.execute('UPDATE project_contracts SET "所属部门"=? WHERE id=?', (args.dst, r[0]))
                n += 1
                con.execute(
                    'INSERT INTO field_change_logs ("时间","用户名","姓名","角色","IP","模块","对象编号",'
                    '"对象名称","字段","原值","新值") VALUES (?,?,?,?,?,?,?,?,?,?,?)',
                    (now, "data_fix", "数据调整脚本", "系统", "本机", "合同台账", r[1], r[2], "所属部门",
                     str(r[3]), args.dst))
            con.execute(
                'INSERT INTO operation_logs ("时间","用户名","姓名","角色","IP","模块","操作类型","对象","详情","结果") '
                'VALUES (?,?,?,?,?,?,?,?,?,?)',
                (now, "data_fix", "数据调整脚本", "系统", "本机", "合同台账", "数据调整", args.no_min + "-" + args.no_max,
                 "所属部门由「%s」调整为「%s」，共 %d 条" % (args.src, args.dst, n), "成功"))
            bad = con.execute("PRAGMA foreign_key_check").fetchall()
            if bad:
                raise RuntimeError("外键校验失败：%s" % (bad[:3],))
            con.execute("COMMIT")
        except Exception as e:  # noqa: BLE001
            con.execute("ROLLBACK")
            print("[ERROR] 执行失败已回滚：%s" % e)
            return 4

        print("\n[DONE] 已调整 %d 条。" % n)
        after = con.execute('SELECT COUNT(*) FROM project_contracts WHERE "所属部门"=?', (args.dst,)).fetchone()[0]
        left = con.execute('SELECT COUNT(*) FROM project_contracts WHERE "所属部门"=?', (args.src,)).fetchone()[0]
        print("调整后：%s 共 %d 条；%s 共 %d 条" % (args.dst, after, args.src, left))
        print("触发器同步检查（前 3 条）：")
        for r in con.execute('SELECT "编号","所属部门","所属单位","所属科室" FROM project_contracts '
                             'WHERE "编号">=? AND "编号"<=? ORDER BY "编号" LIMIT 3', (args.no_min, args.no_max)):
            print("  %s → %s | %s/%s" % (r[0], r[1], r[2], r[3]))
        print("回滚方式：用 %s 覆盖回 %s" % (bak, db))
        return 0
    finally:
        con.close()


if __name__ == "__main__":
    sys.exit(main())
