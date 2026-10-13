# -*- coding: utf-8 -*-
"""D-15 数据清洗脚本：修正「决策的招标控制价」的单位错位（误按“元”填写，应为“万元”）。

范围：仅针对已核对确认的 10 条项目合同（见 FIXES），其他数据一律不动。

用法（默认预演，不做任何修改）：
    python tests/clean_d15_control_price.py --db test_data/清洗演练.db
    python tests/clean_d15_control_price.py --db test_data/清洗演练.db --run
可选参数：
    --no-log                 不写审计日志（默认会写 field_change_logs + operation_logs）
    --i-know-this-is-prod    允许对 data/合同台账.db 直接执行（默认拒绝，防误操作）

安全设计：
  ① 目标库必须显式给出（--db），不会默认指向任何库；
  ② 生产库（data/合同台账.db）需显式开关才允许写入；
  ③ --run 前自动整库备份为 <同目录>/清洗前备份_<时间戳>.db；
  ④ 逐条校验“当前值 == 预期错误值”，不符则跳过并告警（防错改）；
  ⑤ 幂等：已是修正后数值的记录自动跳过，可安全重复执行；
  ⑥ 全部修改在单个事务内提交，失败自动回滚。
"""
import argparse
import datetime
import os
import shutil
import sqlite3
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROD_DB = os.path.abspath(os.path.join(ROOT, "data", "合同台账.db"))
COL = "决策的招标控制价"
AMT = "合同金额（万元）"

# (编号, 预期当前值, 修正后值, 项目名称)  —— 修正后值 = 预期当前值 / 10000
FIXES = [
    ("266303", 663600.00, 66.36, "市民之家中央空调运行维护服务"),
    ("266314", 130500.00, 13.05, "市“三馆”外围梁柱及市博物馆观光电梯井钢结构"),
    ("266313", 115600.00, 11.56, "市委二院办公楼二楼办公用房维修项目"),
    ("266312", 101772.00, 10.18, "市政府原机关一院2、3、4栋住宅区公共部位"),
    ("266302", 80000.00, 8.00, "市行政中心部分绿化带改造为人行通道"),
    ("266309", 62600.00, 6.26, "市政府原机关一、二院屋顶漏水等安全隐患维修"),
    ("266308", 36300.00, 3.63, "市政府B栋1楼西侧6间办公室修缮"),
    ("266306", 25200.00, 2.52, "市博物馆无水消火栓维修项目"),
    ("266301", 22300.00, 2.23, "“四大家”和市民之家中央空调机房监控消防报警"),
    ("266102", 3000.00, 0.30, "2026年2月-3月市本级公车平台易耗品物资"),
]


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def survey(con):
    """返回 (控制价合计, 合同额合计, 异常行数, {编号: 行})。异常 = 控制价/合同额 >= 100。"""
    rows = con.execute('SELECT "编号","项目名称","%s","%s" FROM project_contracts' % (COL, AMT)).fetchall()
    tot = 0.0
    amt_tot = 0.0
    bad = 0
    idx = {}
    for no, name, bid, amt in rows:
        b, a = num(bid), num(amt)
        tot += b or 0
        amt_tot += a or 0
        if b is not None and a is not None and a > 0 and b / a >= 100:
            bad += 1
            idx[str(no)] = {"编号": str(no), "项目名称": name, "控制价": b, "合同额": a}
    return tot, amt_tot, bad, idx


def main():
    ap = argparse.ArgumentParser(description="D-15 招标控制价单位错位清洗（默认 dry-run）")
    ap.add_argument("--db", required=True, help="目标 SQLite 库路径（建议先用 test_data/ 下的副本）")
    ap.add_argument("--run", action="store_true", help="真正执行修改（缺省仅预演）")
    ap.add_argument("--no-log", action="store_true", help="不写审计日志")
    ap.add_argument("--i-know-this-is-prod", action="store_true", help="允许对生产库执行（危险）")
    args = ap.parse_args()

    db = os.path.abspath(args.db)
    if not os.path.exists(db):
        print("[ERROR] 目标库不存在：%s" % db)
        return 2
    if db == PROD_DB and not args.i_know_this_is_prod:
        print("[REFUSED] 目标指向生产库 data/合同台账.db。请先复制到 test_data/ 演练；"
              "确需直连请显式加 --i-know-this-is-prod")
        return 3

    con = sqlite3.connect(db, timeout=20)
    try:
        con.execute("PRAGMA foreign_keys=ON")
        tot0, amt0, bad0, idx = survey(con)
        print("=" * 100)
        print("D-15 招标控制价单位错位清洗　目标库：%s" % db)
        print("模式：%s" % ("【执行】" if args.run else "【预演 dry-run，不修改数据】"))
        print("-" * 100)
        print("修正前：控制价合计 %.2f 万元；合同额合计 %.2f 万元；单位异常(≥100倍) %d 条" % (tot0, amt0, bad0))
        print("-" * 100)

        plan, skip = [], []
        for no, old, new, name in FIXES:
            row = idx.get(no)
            cur = num(con.execute('SELECT "%s" FROM project_contracts WHERE "编号"=?' % COL, (no,)).fetchone()[0]) \
                if con.execute('SELECT COUNT(*) FROM project_contracts WHERE "编号"=?', (no,)).fetchone()[0] else None
            if cur is None:
                skip.append((no, "库中不存在该编号"))
                continue
            if abs(cur - new) < 0.005:
                skip.append((no, "已是修正后数值 %.2f，跳过（幂等）" % new))
                continue
            if abs(cur - old) > 0.005:
                skip.append((no, "当前值 %.2f 与预期错误值 %.2f 不符（人工确认后再处理）" % (cur, old)))
                continue
            plan.append((no, cur, new, name, row["合同额"] if row else None))

        print("| 编号 | 项目名称 | 当前值 | → 修正为 | 合同额(万元) | 比值(修正后) |")
        print("|---|---|---|---|---|---|")
        for no, cur, new, name, amt in plan:
            ratio = ("%.2f" % (new / amt)) if amt else "-"
            print("| %s | %s | %.2f | **%.2f** | %s | %s |" % (
                no, str(name)[:24], cur, new, ("%.2f" % amt) if amt else "-", ratio))
        for no, why in skip:
            print("| %s | （跳过） | - | - | - | %s |" % (no, why))
        print("-" * 100)

        delta = sum(new - cur for _, cur, new, _, _ in plan)
        print("计划修正 %d 条，跳过 %d 条；控制价合计将变化 %+.2f 万元（%.2f → %.2f）" % (
            len(plan), len(skip), delta, tot0, tot0 + delta))
        print("预估修正后异常(≥100倍)条数：%d" % max(0, bad0 - sum(1 for no, *_ in plan if no in idx)))

        if not args.run:
            print("\n[DRY-RUN] 未做任何修改。确认无误后加 --run 执行。")
            return 0

        if not plan:
            print("\n[SKIP] 没有需要修正的记录。")
            return 0

        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        bak = os.path.join(os.path.dirname(db), "清洗前备份_%s.db" % ts)
        shutil.copyfile(db, bak)
        print("\n[BACKUP] 已备份：%s" % bak)

        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        con.execute("BEGIN")
        try:
            for no, cur, new, name, _amt in plan:
                con.execute('UPDATE project_contracts SET "%s"=? WHERE "编号"=?' % COL, (new, no))
                if not args.no_log:
                    con.execute(
                        'INSERT INTO field_change_logs ("时间","用户名","姓名","角色","IP","模块","对象编号",'
                        '"对象名称","字段","原值","新值") VALUES (?,?,?,?,?,?,?,?,?,?,?)',
                        (now, "data_cleanup", "数据清洗脚本", "系统", "本机", "合同台账", no, name, COL,
                         str(cur), str(new)))
            if not args.no_log:
                con.execute(
                    'INSERT INTO operation_logs ("时间","用户名","姓名","角色","IP","模块","操作类型","对象","详情","结果") '
                    'VALUES (?,?,?,?,?,?,?,?,?,?)',
                    (now, "data_cleanup", "数据清洗脚本", "系统", "本机", "合同台账", "数据清洗", "-",
                     "D-15 招标控制价单位错位修正 %d 条（元→万元）" % len(plan), "成功"))
            con.execute("COMMIT")
        except Exception as e:  # noqa: BLE001
            con.execute("ROLLBACK")
            print("[ERROR] 执行失败已回滚：%s" % e)
            return 4

        tot1, amt1, bad1, _ = survey(con)
        print("[DONE] 已修正 %d 条。" % len(plan))
        print("修正后：控制价合计 %.2f 万元；合同额合计 %.2f 万元；单位异常(≥100倍) %d 条" % (tot1, amt1, bad1))
        print("回滚方式：用 %s 覆盖回 %s" % (bak, db))
        return 0
    finally:
        con.close()


if __name__ == "__main__":
    sys.exit(main())
