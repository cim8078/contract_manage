# -*- coding: utf-8 -*-
"""M06 资产租金收缴（整表替换）测试（L2）。
覆盖：整表替换新增/修改/删除、stats 统计、confirm_clear 防误删、逐行校验返回行号与字段、
金额非负、缴款方式枚举、日期格式、归档 403、查询用户 403、资产合同ID 由路径强制覆盖、
200 条上限、totals 计算、字段级变更日志「租金明细[期次]·列名」。
隔离：运行时副本 + test_data 工作库。
"""
import os
import sys
import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _common as C  # noqa: E402


def main():
    C.setup_stdout()
    C.ensure_dirs()
    C.prepare_runtime()
    C.prepare_base()
    db = C.work_db("m06_rent.db")
    prod_before = C.prod_fingerprint()
    r = C.R("M06", "M06 资产租金收缴（整表替换）测试报告（L2）")
    srv = C.Server(db)
    srv.start()
    B = srv.base
    L = C.RUNTIME_DB
    stamp = datetime.datetime.now().strftime("%H%M%S")
    try:
        A, ia = C.login("admin", base=B)
        assert A, "admin 登录失败"
        q_un = C.query(L, "SELECT 用户名 FROM users WHERE 角色='查询用户' ORDER BY 用户名 LIMIT 1")[0]["用户名"]
        Q, iq = C.login(q_un, base=B)
        assert Q, "查询用户登录失败"

        seq = [0]

        def new_asset():
            seq[0] += 1
            no = "990%03d" % seq[0]
            body = {"number": no, "place": "租金测试位置-" + stamp, "asset_name": "租金测试资产-" + stamp,
                    "tenant": "租金承租人", "dept": ia["department"], "annual_rent": 1200}
            resp = A.post(B + "/api/asset_contracts", json=body, timeout=60)
            assert resp.status_code == 201, "创建资产失败 %s %s" % (resp.status_code, resp.text[:160])
            return no

        def rent_rows(aid):
            return C.query(L, "SELECT * FROM asset_rent_collections WHERE CAST(资产合同ID AS TEXT)=? "
                              "ORDER BY id", (str(aid),))

        def put_rent(aid, items, sess=None, **extra):
            body = {"items": items}
            body.update(extra)
            return (sess or A).put(B + "/api/asset_contracts/arow_%s/rent_collections" % aid, json=body, timeout=60)

        def t01():
            aid = new_asset()
            resp = put_rent(aid, [
                {"period_label": "2026年第1季度", "period_start": "2026-01-01", "period_end": "2026-03-31",
                 "due_amount": 30000, "paid_amount": 30000, "pay_date": "2026-01-15", "pay_method": "直缴非税", "payee": "张三"},
                {"period_label": "2026年第2季度", "due_amount": 30000, "paid_amount": 0},
                {"period_label": "2026年第3季度", "due_amount": 30000, "paid_amount": 0}])
            assert resp.status_code == 200, "HTTP %s %s" % (resp.status_code, resp.text[:200])
            j = resp.json()
            rows = rent_rows(aid)
            assert len(rows) == 3, "落库 %d 条（期望 3）" % len(rows)
            assert j["stats"]["inserted"] == 3, "stats.inserted=%s" % j["stats"]
            assert j["totals"]["count"] == 3 and abs(j["totals"]["due"] - 90000) < 0.001, "totals=%s" % j["totals"]
            return "PUT 200；落库 3 条；stats=%s；totals=%s" % (j["stats"], {k: j["totals"][k] for k in ("count", "due", "paid")})

        r.case("TC-M06-01", "整表替换新增 3 条明细", "200 + 落库 3 + stats.inserted=3", t01)

        def t02():
            aid = new_asset()
            put_rent(aid, [
                {"period_label": "2026年第1季度", "due_amount": 30000, "paid_amount": 30000},
                {"period_label": "2026年第2季度", "due_amount": 30000, "paid_amount": 0},
                {"period_label": "2026年第3季度", "due_amount": 30000, "paid_amount": 0}])
            rows = rent_rows(aid)
            by_label = {x["收款期次"]: x for x in rows}
            q1 = by_label["2026年第1季度"]
            keep = {"id": q1["id"], "period_label": "2026年第1季度", "due_amount": 30000, "paid_amount": 31000}
            keep2 = {"id": by_label["2026年第2季度"]["id"], "period_label": "2026年第2季度",
                     "due_amount": 30000, "paid_amount": 5000}
            resp = put_rent(aid, [keep, keep2])
            assert resp.status_code == 200, "HTTP %s %s" % (resp.status_code, resp.text[:200])
            j = resp.json()
            rows2 = rent_rows(aid)
            assert len(rows2) == 2, "剩余 %d 条（期望 2）" % len(rows2)
            labels = sorted(x["收款期次"] for x in rows2)
            assert labels == ["2026年第1季度", "2026年第2季度"], "剩余期次=%s" % labels
            v = [x for x in rows2 if x["收款期次"] == "2026年第1季度"][0]["实缴金额（元）"]
            assert abs(float(v) - 31000) < 0.001, "Q1 实缴未更新=%s" % v
            assert j["stats"]["deleted"] == 1, "stats.deleted=%s" % j["stats"]
            return "修改 1 条+删除 1 条；stats=%s；Q1 实缴=31000" % j["stats"]

        r.case("TC-M06-02", "整表替换：修改与删除", "剩余 2 条、stats.deleted=1", t02)

        def t03():
            aid = new_asset()
            put_rent(aid, [{"period_label": "2026年第1季度", "due_amount": 100}])
            resp = put_rent(aid, [])
            assert resp.status_code == 400, "空 items 未确认返回 %s（期望 400）" % resp.status_code
            assert len(rent_rows(aid)) == 1, "未确认却清空了明细"
            return "空 items 且未确认 → 400；明细保留 %d 条" % len(rent_rows(aid))

        r.case("TC-M06-03", "清空明细需 confirm_clear（防误删）", "400 且不清空", t03)

        def t04():
            aid = new_asset()
            put_rent(aid, [{"period_label": "2026年第1季度", "due_amount": 100},
                           {"period_label": "2026年第2季度", "due_amount": 100}])
            resp = put_rent(aid, [], confirm_clear=True)
            assert resp.status_code == 200, "确认清空 HTTP %s %s" % (resp.status_code, resp.text[:200])
            assert len(rent_rows(aid)) == 0, "确认后仍有 %d 条" % len(rent_rows(aid))
            return "confirm_clear=true → 200，明细清空"

        r.case("TC-M06-04", "确认清空明细", "200 + 明细清空", t04)

        def t05():
            aid = new_asset()
            cases = [
                ("缺收款期次", [{"due_amount": 1}]),
                ("应缴非数字", [{"period_label": "P1", "due_amount": "abc"}]),
                ("实缴为负", [{"period_label": "P1", "due_amount": 1, "paid_amount": -5}]),
                ("日期非法", [{"period_label": "P1", "due_amount": 1, "pay_date": "2026-13-40"}]),
                ("缴款方式非法", [{"period_label": "P1", "due_amount": 1, "pay_method": "比特币"}]),
                ("第二行错误", [{"period_label": "P1", "due_amount": 1},
                                {"period_label": "P2", "due_amount": "x"}]),
            ]
            out = []
            for name, items in cases:
                resp = put_rent(aid, items)
                msg = ""
                try:
                    msg = (resp.json() or {}).get("error", "")
                except Exception:
                    msg = resp.text[:80]
                out.append("%s=%s%s" % (name, resp.status_code, ("(" + msg[:28] + ")") if msg else ""))
                assert resp.status_code == 400, "%s 返回 %s（期望 400）" % (name, resp.status_code)
                assert msg, "%s 未返回错误说明" % name
            assert len(rent_rows(aid)) == 0, "校验失败却写入了明细"
            return "；".join(out)

        r.case("TC-M06-05", "逐行校验返回原因", "400 + 中文原因，且不落库", t05)

        def t06():
            aid = new_asset()
            put_rent(aid, [{"period_label": "2026年第1季度", "due_amount": 100}])
            ar = A.put(B + "/api/asset_contracts/arow_%s/archive" % aid, json={}, timeout=30)
            assert ar.status_code == 200, "归档失败 %s" % ar.status_code
            resp = put_rent(aid, [{"period_label": "2026年第1季度", "due_amount": 999}])
            assert resp.status_code == 403, "归档后改租金返回 %s（期望 403）" % resp.status_code
            v = rent_rows(aid)[0]["应缴金额（元）"]
            assert abs(float(v) - 100) < 0.001, "归档后明细被改动"
            return "归档后 PUT 明细 → 403；原明细未被改动"

        r.case("TC-M06-06", "归档后租金明细只读", "403", t06)

        def t07():
            aid = new_asset()
            put_rent(aid, [{"period_label": "2026年第1季度", "due_amount": 30000, "paid_amount": 30000}])
            row = rent_rows(aid)[0]
            put_rent(aid, [{"id": row["id"], "period_label": "2026年第1季度",
                            "due_amount": 30000, "paid_amount": 31000}])
            fl = C.query(L, "SELECT * FROM field_change_logs WHERE 模块='资产合同台账' AND 对象编号=? "
                            "AND 字段 LIKE '租金明细%'", (aid,))
            assert fl, "未写租金明细字段级日志"
            hit = [x for x in fl if "2026年第1季度" in x["字段"] and "实缴金额" in x["字段"]]
            assert hit, "未命中「租金明细[期次]·实缴金额（元）」格式：%s" % [x["字段"] for x in fl]
            f0 = hit[-1]
            assert abs(float(f0["原值"]) - 30000) < 0.001 and abs(float(f0["新值"]) - 31000) < 0.001, \
                "日志原值/新值不符：%s→%s" % (f0["原值"], f0["新值"])
            return "命中字段「%s」原值 %s→新值 %s" % (f0["字段"], f0["原值"], f0["新值"])

        r.case("TC-M06-07", "租金明细字段级变更日志", "字段=租金明细[期次]·列名，原/新值正确", t07)

        def t08():
            x = new_asset()
            y = new_asset()
            resp = put_rent(x, [{"asset_id": y, "period_label": "2026年第1季度", "due_amount": 100}])
            assert resp.status_code == 200, "HTTP %s %s" % (resp.status_code, resp.text[:200])
            assert len(rent_rows(x)) == 1, "路径合同的明细未落库"
            assert len(rent_rows(y)) == 0, "body 中的 asset_id 被误当作目标（越权写入 %s）" % y
            return "body 内 asset_id=%s 被忽略；明细落在路径合同 %s" % (y, x)

        r.case("TC-M06-08", "资产合同ID 以路径为准（防越权）", "只写入路径合同", t08)

        def t09():
            aid = new_asset()
            put_rent(aid, [{"period_label": "Q1", "due_amount": 100, "paid_amount": 40},
                           {"period_label": "Q2", "due_amount": 200, "paid_amount": 0}])
            j = A.get(B + "/api/asset_contracts/arow_%s/rent_collections" % aid, timeout=30).json()
            t = j["totals"]
            assert t["count"] == 2, "totals.count=%s" % t["count"]
            assert abs(t["due"] - 300) < 0.001 and abs(t["paid"] - 40) < 0.001, "totals 金额不符：%s" % t
            assert abs(t.get("diff", (t["due"] - t["paid"])) - 260) < 0.001, "diff 不符：%s" % t
            return "totals=%s" % {k: t[k] for k in t}

        r.case("TC-M06-09", "明细汇总 totals", "count/due/paid/diff 正确", t09, priority="P1")

        def t10():
            aid = new_asset()
            resp = put_rent(aid, [{"period_label": "Q1", "due_amount": -1}])
            assert resp.status_code == 400, "负应缴返回 %s（期望 400）" % resp.status_code
            return "应缴 -1 → 400 %s" % (resp.json() or {}).get("error", "")

        r.case("TC-M06-10", "金额非负校验（Q14）", "400", t10, priority="P1")

        def t11():
            aid = new_asset()
            resp = put_rent(aid, [{"period_label": "Q1", "due_amount": 1}], sess=Q)
            assert resp.status_code == 403, "查询用户改明细返回 %s（期望 403）" % resp.status_code
            g = Q.get(B + "/api/asset_contracts/arow_%s/rent_collections" % aid, timeout=30)
            assert g.status_code in (200, 403, 404), "查询用户读明细异常 %s" % g.status_code
            return "查询用户 PUT 403；GET 明细返回 %s" % g.status_code

        r.case("TC-M06-11", "查询用户不可改明细", "403", t11)

        def t12():
            aid = new_asset()
            items = [{"period_label": "P%03d" % i, "due_amount": 1} for i in range(201)]
            resp = put_rent(aid, items)
            assert resp.status_code == 400, "201 条返回 %s（期望 400 超上限）" % resp.status_code
            msg = (resp.json() or {}).get("error", "")
            assert "200" in msg, "错误信息未提示上限：%s" % msg
            assert len(rent_rows(aid)) == 0, "超限请求仍写入数据"
            ok = put_rent(aid, [{"period_label": "P%03d" % i, "due_amount": 1} for i in range(200)])
            assert ok.status_code == 200, "恰好 200 条返回 %s（期望 200）" % ok.status_code
            return "201 条 → 400（%s）；200 条 → 200，实际落库 %d 条" % (msg[:30], len(rent_rows(aid)))

        r.case("TC-M06-12", "单合同 200 条上限（ASSET_RENT_MAX_ROWS）", "201→400 / 200→200", t12)

        srv.sync_db()
    finally:
        srv.stop()

    prod_after = C.prod_fingerprint()
    meta = {
        "服务地址": B,
        "工作库": "test_data/m06_rent.db",
        "生产库未改动": "是" if prod_before == prod_after else "否（指纹变化！）",
    }
    r.report("M06_租金整表替换测试.md", meta=meta)
    n, p, f, e = r.stat()
    print("\n[SUMMARY] M06 共 %d 例：通过 %d，失败 %d，错误 %d" % (n, p, f, e))
    sys.stdout.flush()
    os._exit(0 if (f == 0 and e == 0) else 1)


if __name__ == "__main__":
    main()
