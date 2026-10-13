# -*- coding: utf-8 -*-
"""M14 日志审计测试（L6）。
覆盖：登录/登出留痕、写操作埋点覆盖（增删改/归档/附件/导入导出/用户/部门/配置）、字段级变更日志内容、
日志与字段日志的权限过滤、筛选参数、只追加（无删除接口）、导出与导出留痕、保留天数清理、分页边界。
说明：保留天数清理依赖启动时执行（_maybe_cleanup_logs 每进程每天一次），故本模块用 copy_db=False 自管副本，
服务启动前预置「200 天前」日志与 log_keep_days=90。
"""
import os
import re
import sys
import shutil
import sqlite3
import datetime
from io import BytesIO

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _common as C  # noqa: E402
import openpyxl  # noqa: E402


def main():
    C.setup_stdout()
    C.ensure_dirs()
    C.prepare_runtime()
    C.prepare_base()
    db = C.work_db("m14_logs.db")
    prod_before = C.prod_fingerprint()
    r = C.R("M14", "M14 日志审计测试报告（L6）")
    srv = C.Server(db, copy_db=False)
    shutil.copyfile(db, C.RUNTIME_DB)
    L = C.RUNTIME_DB
    stamp = datetime.datetime.now().strftime("%H%M%S")
    old_ts = (datetime.datetime.now() - datetime.timedelta(days=200)).strftime("%Y-%m-%d %H:%M:%S")
    recent_ts = (datetime.datetime.now() - datetime.timedelta(days=3)).strftime("%Y-%m-%d %H:%M:%S")
    con = sqlite3.connect(L)
    try:
        con.execute('INSERT OR REPLACE INTO settings ("配置项","配置值") VALUES (?,?)', ("log_keep_days", "90"))
        for ts, tag in ((old_ts, "200天前"), (recent_ts, "3天前")):
            con.execute('INSERT INTO operation_logs ("时间","用户名","姓名","角色","IP","模块","操作类型","对象","详情","结果") '
                        'VALUES (?,?,?,?,?,?,?,?,?,?)', (ts, "admin", "系统管理员", "管理员", "127.0.0.1",
                                                         "保留天数测试", "保留天数测试", "-", tag, "成功"))
            con.execute('INSERT INTO field_change_logs ("时间","用户名","姓名","角色","IP","模块","对象编号","对象名称","字段","原值","新值") '
                        'VALUES (?,?,?,?,?,?,?,?,?,?,?)', (ts, "admin", "系统管理员", "管理员", "127.0.0.1",
                                                           "保留天数测试", "-", "-", "保留天数测试", tag, tag))
        con.commit()
    finally:
        con.close()
    srv.start()
    B = srv.base
    try:
        P = C.Server(B.replace("http://127.0.0.1:", ""))
        A, ia = C.login("admin", base=B)
        assert A, "admin 登录失败"
        p_un = C.query(L, "SELECT 用户名 FROM users WHERE 角色='普通用户' ORDER BY 用户名 LIMIT 1")[0]["用户名"]
        P, ip = C.login(p_un, base=B)

        def logs(sess=None, **params):
            return (sess or A).get(B + "/api/logs", params=params, timeout=60)

        def flogs(sess=None, **params):
            return (sess or A).get(B + "/api/field_logs", params=params, timeout=60)

        def db_count(table, where="", params=()):
            return C.query(L, "SELECT COUNT(*) c FROM %s %s" % (table, where), params)[0]["c"]

        def do_writes():
            body = {"category": "服务类", "name": "M14日志-" + stamp, "procurement_method": "自行采购",
                    "amount": 2, "main_content": "日志测试", "handler": ia["name"], "fund_source": "其他"}
            no = A.post(B + "/api/contracts", json=body, timeout=60).json()["number"]
            cid = "row_%s" % C.query(L, "SELECT id FROM project_contracts WHERE 编号=?", (no,))[0]["id"]
            A.put(B + "/api/contracts/" + cid, json={"remark": "改备注"}, timeout=60)
            A.post(B + "/api/contracts/%s/attachments" % cid, files={"file": ("a.png", b"\x89PNG\r\n\x1a\n123")},
                   data={"type": "decision"}, timeout=60)
            A.put(B + "/api/contracts/%s/archive" % cid, json={}, timeout=30)
            no2 = "M14DEL" + stamp[:2]
            b2 = dict(body)
            b2["name"] = "M14删除-" + stamp
            b2["number"] = no2
            A.post(B + "/api/contracts", json=b2, timeout=60)
            rid = C.query(L, "SELECT id FROM project_contracts WHERE 编号=?", (no2,))[0]["id"]
            A.delete(B + "/api/contracts/row_%s" % rid, json={"reason": "日志测试"}, timeout=30)
            A.post(B + "/api/users", json={"username": "zz_m14_%s" % stamp, "name": "M14用户",
                                           "department": ia["department"], "role": "普通用户"}, timeout=30)
            A.post(B + "/api/departments", json={"unit": "M14单位" + stamp, "dept": "M14科室"}, timeout=30)
            A.put(B + "/api/config", json={"default_remind_days": 61}, timeout=30)
            A.put(B + "/api/config", json={"default_remind_days": 60}, timeout=30)
            A.get(B + "/api/contracts/export", timeout=180)
            return no, no2, cid

        def t01():
            s1, i1 = C.login("admin", base=B)
            bad = C.login("admin", password="wrong-pw", base=B)
            assert s1 and bad[0] is None, "登录样本异常"
            rows = C.query(L, "SELECT 结果 r, COUNT(*) c FROM operation_logs WHERE 模块='用户登录' AND 操作类型='登录' GROUP BY 结果")
            got = {x["r"]: x["c"] for x in rows}
            assert got.get("成功", 0) > 0, "未记录成功登录"
            assert got.get("失败", 0) > 0, "未记录失败登录"
            s1.get(B + "/api/logout", timeout=30)
            s2, i2 = C.login("admin", base=B)
            assert s2, "登出后重新登录失败"
            out = C.query(L, "SELECT COUNT(*) c FROM operation_logs WHERE 模块='用户登录' AND 操作类型='登出'")[0]["c"]
            assert out > 0, "未记录登出"
            return "登录留痕：成功 %d 条、失败 %d 条；登出 %d 条" % (got.get("成功", 0), got.get("失败", 0), out)

        r.case("TC-M14-01", "登录/登出留痕", "成功与失败均记录", t01)

        def t02():
            no, no2, cid = do_writes()
            required = [("用户登录", "登录"), ("合同台账", "新增"), ("合同台账", "修改"), ("合同台账", "归档"),
                        ("合同台账", "删除"), ("合同台账", "上传附件"), ("合同台账", "导出"),
                        ("用户管理", "新增用户"), ("部门管理", "新增部门"), ("系统配置", "修改配置")]
            rows = set((x["模块"], x["操作类型"]) for x in
                       C.query(L, "SELECT DISTINCT 模块, 操作类型 FROM operation_logs"))
            miss = [x for x in required if x not in rows]
            assert not miss, "埋点缺失：%s" % miss
            return "10 类写操作埋点齐全（新增/修改/归档/删除/上传附件/导出/新增用户/新增部门/改配置/登录）"

        r.case("TC-M14-02", "写操作埋点覆盖", "关键动作均留痕", t02)

        def t03():
            no = "M14FL" + stamp[:2]
            body = {"number": no, "category": "服务类", "name": "M14字段日志-" + stamp, "procurement_method": "自行采购",
                    "amount": 1, "main_content": "x", "handler": ia["name"], "fund_source": "其他",
                    "winner": "原单位A"}
            A.post(B + "/api/contracts", json=body, timeout=60)
            rid = C.query(L, "SELECT id FROM project_contracts WHERE 编号=?", (no,))[0]["id"]
            A.put(B + "/api/contracts/row_%s" % rid, json={"winner": "新单位B"}, timeout=60)
            rows = C.query(L, "SELECT * FROM field_change_logs WHERE 对象编号=? AND 字段='中标单位'", (no,))
            assert rows, "未记录字段级变更"
            x = rows[-1]
            for k in ("时间", "用户名", "姓名", "模块", "对象编号", "对象名称", "字段", "原值", "新值"):
                assert str(x.get(k) or "").strip(), "字段日志缺 %s" % k
            assert x["原值"] == "原单位A" and x["新值"] == "新单位B", "前后值不符：%s→%s" % (x["原值"], x["新值"])
            return "字段日志完整：%s·%s %s→%s（操作人 %s）" % (x["模块"], x["字段"], x["原值"], x["新值"], x["姓名"])

        r.case("TC-M14-03", "字段级变更日志内容", "含前后值/操作人/时间", t03)

        def t04():
            total_db = db_count("operation_logs")
            ja = logs().json()
            jp = logs(P).json()
            assert ja["total"] == total_db, "管理员日志 total %s ≠ 库内 %s" % (ja["total"], total_db)
            others = set(x["用户名"] for x in jp["items"])
            assert others <= {p_un}, "普通用户看到他人日志：%s" % others
            mine = db_count("operation_logs", 'WHERE "用户名"=?', (p_un,))
            assert jp["total"] == mine, "普通用户 total %s ≠ 本人 %s" % (jp["total"], mine)
            fa = flogs().json()
            fp = flogs(P).json()
            assert fa["total"] == db_count("field_change_logs"), "字段日志管理员 total 不符"
            assert set(x["用户名"] for x in fp["items"]) <= {p_un}, "普通用户看到他人字段日志"
            return "管理员日志 %d 条/字段日志 %d 条；普通用户仅见本人（%d/%d）" % (
                ja["total"], fa["total"], jp["total"], fp["total"])

        r.case("TC-M14-04", "日志可见性（管理员全量/其他本人）", "口径与库内一致", t04)

        def t05():
            today = datetime.date.today().strftime("%Y-%m-%d")
            j1 = logs(start_date=today).json()
            exp1 = db_count("operation_logs", 'WHERE "时间">=?', (today + " 00:00:00",))
            assert j1["total"] == exp1, "start_date 筛选 %s≠%s" % (j1["total"], exp1)
            j2 = logs(module="用户登录").json()
            exp2 = db_count("operation_logs", 'WHERE "模块"=?', ("用户登录",))
            assert j2["total"] == exp2, "module 筛选 %s≠%s" % (j2["total"], exp2)
            j3 = logs(action="新增").json()
            exp3 = db_count("operation_logs", 'WHERE "操作类型"=?', ("新增",))
            assert j3["total"] == exp3, "action 筛选 %s≠%s" % (j3["total"], exp3)
            old = (datetime.date.today() - datetime.timedelta(days=1)).strftime("%Y-%m-%d")
            j4 = logs(start_date=old, end_date=old).json()
            exp4 = db_count("operation_logs", 'WHERE "时间">=? AND "时间"<=?', (old + " 00:00:00", old + " 23:59:59"))
            assert j4["total"] == exp4, "日期区间筛选 %s≠%s" % (j4["total"], exp4)
            return "start_date/module/action/日期区间 筛选均与库内一致（%s/%s/%s/%s）" % (
                j1["total"], j2["total"], j3["total"], j4["total"])

        r.case("TC-M14-05", "日志筛选参数", "与 SQL 期望一致", t05)

        def t06():
            j = flogs(field="中标单位").json()
            exp = db_count("field_change_logs", 'WHERE "字段" LIKE ?', ("%中标单位%",))
            assert j["total"] == exp, "字段筛选 %s≠%s" % (j["total"], exp)
            j2 = flogs(module="合同台账").json()
            assert j2["total"] == db_count("field_change_logs", 'WHERE "模块"=?', ("合同台账",)), "模块筛选不符"
            j3 = flogs(obj="M14FL" + stamp[:2]).json()
            assert j3["total"] >= 1, "按对象编号筛选无结果"
            assert isinstance(j["fields"], list) and j["fields"], "未返回字段列表 fields"
            assert "中标单位" in j["fields"], "fields 未包含「中标单位」"
            return "字段/模块/对象筛选正确；fields 共 %d 个字段（含「中标单位」）" % len(j["fields"])

        r.case("TC-M14-06", "字段日志筛选与字段清单", "筛选正确 + fields 列表", t06, priority="P1")

        def t07():
            resp = A.get(B + "/api/logs/export", timeout=180)
            assert resp.status_code == 200, "导出 HTTP %s" % resp.status_code
            ws = openpyxl.load_workbook(BytesIO(resp.content)).active
            hdr = [str(c.value or "") for c in ws[1]]
            for k in ("时间", "用户名", "姓名", "模块", "操作类型", "对象", "详情", "结果"):
                assert k in hdr, "导出缺少列 %s" % k
            n = ws.max_row - 1
            total = logs(page_size=1).json()["total"]
            assert n >= total - 2, "导出行数 %d 明显少于日志总数 %d" % (n, total)
            return "日志导出 200：%d 行、%d 列（含时间/用户/模块/操作/对象/详情/结果）" % (n, len(hdr))

        r.case("TC-M14-07", "日志导出（管理员）", "200 + xlsx 内容完整", t07)

        def t08():
            src = open(os.path.join(C.PROJECT_ROOT, "app.py"), encoding="utf-8").read()
            log_routes = re.findall(r'@app\.route\("(/api/(?:logs|field_logs)[^"]*)"(?:,\s*methods=\[([^\]]*)\])?', src)
            dels = [x for x in log_routes if x[1] and "DELETE" in x[1]]
            assert not dels, "存在删除日志的接口：%s" % dels
            methods = ["%s:%s" % (p, m or "GET") for p, m in log_routes]
            return "日志相关路由 %s，均无 DELETE 方法（只追加，不可改删）" % methods

        r.case("TC-M14-08", "日志只追加（无删除/修改接口）", "无 DELETE 路由", t08)

        def t09():
            old = C.query(L, "SELECT COUNT(*) c FROM operation_logs WHERE 模块='保留天数测试' AND 详情='200天前'")[0]["c"]
            recent = C.query(L, "SELECT COUNT(*) c FROM operation_logs WHERE 模块='保留天数测试' AND 详情='3天前'")[0]["c"]
            old_f = C.query(L, "SELECT COUNT(*) c FROM field_change_logs WHERE 模块='保留天数测试' AND 原值='200天前'")[0]["c"]
            recent_f = C.query(L, "SELECT COUNT(*) c FROM field_change_logs WHERE 模块='保留天数测试' AND 原值='3天前'")[0]["c"]
            keep = C.query(L, "SELECT 配置值 v FROM settings WHERE 配置项='log_keep_days'")[0]["v"]
            assert old == 0, "保留 90 天下，200 天前的操作日志未被清理"
            assert old_f == 0, "200 天前的字段日志未被清理"
            assert recent == 1 and recent_f == 1, "3 天前日志被误清理（%s/%s）" % (recent, recent_f)
            return "log_keep_days=%s：200 天前日志（操作/字段）已清理，3 天前保留" % keep

        r.case("TC-M14-09", "日志保留天数清理", "超期清理、期内保留", t09, priority="P1")

        def t10():
            before = db_count("operation_logs", 'WHERE "操作类型"=?', ("导出",))
            A.get(B + "/api/logs/export", timeout=180)
            after = db_count("operation_logs", 'WHERE "操作类型"=?', ("导出",))
            assert after > before, "日志导出未留痕（%d→%d）" % (before, after)
            row = C.query(L, 'SELECT * FROM operation_logs WHERE "操作类型"=? ORDER BY id DESC LIMIT 1', ("导出",))[0]
            return "导出行为留痕：%s · %s · %s" % (row["模块"], row["操作类型"], str(row["详情"])[:40])

        r.case("TC-M14-10", "日志导出自身留痕", "导出后新增审计记录", t10, priority="P1")

        def t11():
            outs = []
            for params in ({"page": 0, "page_size": 0}, {"page": -1, "page_size": 9999},
                           {"page": 999999, "page_size": 20}, {"page": "abc", "page_size": "xyz"}):
                resp = logs(**params)
                assert resp.status_code == 200, "分页边界 %s 返回 %s" % (params, resp.status_code)
                j = resp.json()
                outs.append("p=%s/ps=%s→%s,%s" % (params.get("page"), params.get("page_size"),
                                                  j["page"], j["page_size"]))
                assert 1 <= j["page_size"] <= 200, "page_size 越界：%s" % j["page_size"]
                assert j["page"] >= 1, "page 未归一到 ≥1：%s" % j["page"]
            return "分页边界均 200 且归一化（%s）" % ";".join(outs)

        r.case("TC-M14-11", "日志分页边界", "无 500、参数归一化", t11, priority="P2")

        srv.sync_db()
    finally:
        srv.stop()
    prod_after = C.prod_fingerprint()
    meta = {
        "服务地址": B,
        "工作库": "test_data/m14_logs.db",
        "保留天数": "预置 log_keep_days=90（原值见 settings）",
        "生产库未改动": "是" if prod_before == prod_after else "否（指纹变化！）",
    }
    r.report("M14_日志审计测试.md", meta=meta)
    n, p, f, e = r.stat()
    print("\n[SUMMARY] M14 共 %d 例：通过 %d，失败 %d，错误 %d" % (n, p, f, e))
    sys.stdout.flush()
    os._exit(0 if (f == 0 and e == 0) else 1)


if __name__ == "__main__":
    main()
