# -*- coding: utf-8 -*-
"""M05 资产合同 CRUD/归档 与数据质量测试（L2）。
覆盖：新增（必填 place/asset_name/tenant、金额非负、日期 yyyy-mm-dd、6 位编号）、
列表筛选（位置精确、名称/承租人/管理部门/责任人模糊、归档、排序）、计算列（paid_total/due_total/rent_count）、
详情（含 attachments/rent_collections）、修改留痕、归档与归档后只读、删除（reason、有租金 409、delete_rent 级联）、
归档权限（管理员/责任人）、清空提醒、不存在资源 404。
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
    db = C.work_db("m05_asset.db")
    prod_before = C.prod_fingerprint()
    r = C.R("M05", "M05 资产合同 CRUD/归档与数据质量测试报告（L2）")
    srv = C.Server(db)
    srv.start()
    B = srv.base
    L = C.RUNTIME_DB
    stamp = datetime.datetime.now().strftime("%H%M%S")
    try:
        A, ia = C.login("admin", base=B)
        assert A, "admin 登录失败：%s" % ia
        p_un = C.query(L, "SELECT 用户名 FROM users WHERE 角色='普通用户' ORDER BY 用户名 LIMIT 1")[0]["用户名"]
        P, ip = C.login(p_un, base=B)
        assert P, "普通用户登录失败"

        def asset_payload(**kw):
            base = {"place": "测试位置-" + stamp, "asset_name": "测试资产-" + stamp,
                    "tenant": "测试承租人", "dept": ia["department"], "person": ip["name"],
                    "area": 12.5, "lease_start": "2026-01-01", "lease_end": "2026-12-31",
                    "deposit": 100, "annual_rent": 1200, "quarter_rent": 300}
            base.update(kw)
            return base

        seq = [0]

        def new_asset(num=None, **kw):
            seq[0] += 1
            body = asset_payload(**kw)
            body["number"] = num or ("999%03d" % seq[0])
            return A.post(B + "/api/asset_contracts", json=body, timeout=60)

        def arow(no):
            rows = C.query(L, "SELECT * FROM asset_contracts WHERE id=?", (no,))
            return rows[0] if rows else None

        def t01():
            resp = new_asset()
            assert resp.status_code == 201, "HTTP %s %s" % (resp.status_code, resp.text[:200])
            no = resp.json()["number"]
            assert len(no) == 6 and no.isdigit(), "编号不是 6 位：%s" % no
            row = arow(no)
            assert row, "库内未找到资产合同 %s" % no
            lg = C.query(L, "SELECT * FROM operation_logs WHERE 模块='资产合同台账' AND 操作类型='新增'")
            assert lg, "未写操作日志"
            return "新增 %s（arow_%s）库内 1 行、操作日志 %d 条" % (no, row["id"], len(lg))

        r.case("TC-M05-01", "新增资产合同（含 6 位编号）", "201 + 库内 1 行 + 日志", t01)

        def t02():
            hits = []
            for miss in ("place", "asset_name", "tenant"):
                body = asset_payload()
                body.pop(miss)
                body["number"] = "998" + str(100 + len(hits))
                rr = A.post(B + "/api/asset_contracts", json=body, timeout=60)
                hits.append("%s=%s" % (miss, rr.status_code))
            bad = [h for h in hits if not h.endswith("400")]
            assert not bad, "必填缺失未被拒绝：%s" % bad
            return "缺 place/asset_name/tenant 均 400（%s）" % ",".join(hits)

        r.case("TC-M05-02", "必填校验 place/asset_name/tenant（Q13）", "400", t02)

        def t03():
            hits = []
            vals = (("deposit", -1), ("annual_rent", -1), ("quarter_rent", -5), ("area", -3), ("total_amount", -9))
            for i, (k, v) in enumerate(vals):
                body = asset_payload(**{k: v})
                body["number"] = "997%03d" % i
                rr = A.post(B + "/api/asset_contracts", json=body, timeout=60)
                hits.append("%s=%s" % (k, rr.status_code))
            bad = [h for h in hits if not h.endswith("400")]
            assert not bad, "负值金额未被拒绝：%s" % bad
            return "押金/年租金/季租/面积/总合同额 负值均 400（%s）" % ",".join(hits)

        r.case("TC-M05-03", "金额非负校验（Q14）", "400", t03)

        def t04():
            # 宽松解析：2026/1/1 应被规范化为 2026-01-01 后接受（属设计行为）
            body = asset_payload(lease_start="2026/1/1")
            body["number"] = "996000"
            r1 = A.post(B + "/api/asset_contracts", json=body, timeout=60)
            assert r1.status_code == 201, "宽松日期 2026/1/1 返回 %s（期望规范化后接受）" % r1.status_code
            got = arow("996000")["租期开始日"]
            assert got == "2026-01-01", "日期未规范化为 yyyy-mm-dd：%r" % got
            hits = []
            vals = (("lease_end", "2026-13-40"), ("next_lease_date", "abc"))
            for i, (k, v) in enumerate(vals):
                b2 = asset_payload(**{k: v})
                b2["number"] = "996%03d" % (i + 1)
                rr = A.post(B + "/api/asset_contracts", json=b2, timeout=60)
                hits.append("%s=%s" % (k, rr.status_code))
            bad = [h for h in hits if not h.endswith("400")]
            assert not bad, "非法日期未被拒绝：%s" % bad
            return "2026/1/1→规范化 %s；非法 %s 均 400" % (got, ",".join(hits))

        r.case("TC-M05-04", "日期格式校验（Q15）", "400", t04)

        def t05():
            body = asset_payload(annual_rent="约一万")
            body["number"] = "995001"
            rr = A.post(B + "/api/asset_contracts", json=body, timeout=60)
            assert rr.status_code == 400, "金额非数字返回 %s（期望 400；对比：项目合同接口会接受文本）" % rr.status_code
            return "年租金「约一万」→ 400（资产接口金额校验正确）"

        r.case("TC-M05-05", "金额数字校验", "400", t05, priority="P1")

        def t06():
            out = []
            rows = C.query(L, "SELECT * FROM asset_contracts ORDER BY id")
            sample = rows[0]
            jj = A.get(B + "/api/asset_contracts", params={"place": sample["位置"], "page_size": 200}, timeout=60).json()
            exp = C.query(L, "SELECT COUNT(*) c FROM asset_contracts WHERE 位置=?", (sample["位置"],))[0]["c"]
            assert jj["total"] == exp, "place 精确筛选 total=%s，期望 %s" % (jj["total"], exp)
            out.append("place=%d" % exp)
            j2 = A.get(B + "/api/asset_contracts", params={"archived": "否", "page_size": 200}, timeout=60).json()
            exp2 = C.query(L, "SELECT COUNT(*) c FROM asset_contracts WHERE COALESCE(是否归档,'')<>'是'")[0]["c"]
            assert j2["total"] == exp2, "archived=否 total=%s，期望 %s" % (j2["total"], exp2)
            out.append("archived=否:%d" % exp2)
            sub = str(sample["资产名称"])[:4]
            j3 = A.get(B + "/api/asset_contracts", params={"asset_name": sub, "page_size": 200}, timeout=60).json()
            exp3 = C.query(L, "SELECT COUNT(*) c FROM asset_contracts WHERE 资产名称 LIKE ?", ("%" + sub + "%",))[0]["c"]
            assert j3["total"] == exp3, "asset_name 模糊 total=%s，期望 %s" % (j3["total"], exp3)
            out.append("asset_name~%s:%d" % (sub, exp3))
            ids = [int(x["_row"]) for x in j3["items"] if str(x["_row"]).isdigit()]
            assert ids == sorted(ids, reverse=True), "未按 id 降序：%s" % ids
            out.append("排序=id 降序")
            return "；".join(out)

        r.case("TC-M05-06", "资产列表筛选与排序", "与 SQL 期望一致", t06)

        def t07():
            jj = A.get(B + "/api/asset_contracts", params={"page_size": 200}, timeout=60).json()
            bad = []
            for it in jj["items"]:
                rid = it["_row"]
                agg = C.query(L, "SELECT COUNT(*) c, COALESCE(SUM(应缴金额（元）),0) d, "
                                 "COALESCE(SUM(实缴金额（元）),0) p FROM asset_rent_collections WHERE 资产合同ID=?", (rid,))[0]
                if int(it["rent_count"]) != agg["c"] or abs(float(it["due_total"]) - float(agg["d"])) > 0.001 \
                        or abs(float(it["paid_total"]) - float(agg["p"])) > 0.001:
                    bad.append("%s: (%s,%s,%s) vs %s" % (rid, it["rent_count"], it["due_total"], it["paid_total"], agg))
            assert not bad, "计算列与子表聚合不一致：%s" % bad[:3]
            return "%d 条资产的 rent_count/due_total/paid_total 与子表聚合一致" % len(jj["items"])

        r.case("TC-M05-07", "计算列与租金子表一致", "完全一致", t07, priority="P1")

        def t08():
            resp = new_asset()
            no = resp.json()["number"]
            d = A.get(B + "/api/asset_contracts/arow_%s" % no, timeout=30)
            assert d.status_code == 200, "详情 HTTP %s %s" % (d.status_code, d.text[:160])
            j = d.json()
            for k in ("id", "place", "asset_name", "tenant", "attachments", "rent_collections"):
                assert k in j, "详情缺字段 %s" % k
            rc = j["rent_collections"]
            n_rent = len(rc.get("items", [])) if isinstance(rc, dict) else len(rc)
            return "详情 200，含 attachments(%d) 与 rent_collections(%d)" % (len(j["attachments"]), n_rent)

        r.case("TC-M05-08", "资产详情含附件与租金明细", "200 + 两子结构", t08, priority="P1")

        def t09():
            resp = new_asset(person="原责任人")
            no = resp.json()["number"]
            up = A.put(B + "/api/asset_contracts/arow_%s" % no, json={"person": "新责任人"}, timeout=60)
            assert up.status_code == 200, "修改 HTTP %s %s" % (up.status_code, up.text[:160])
            assert arow(no)["责任人"] == "新责任人", "库内未更新"
            fl = C.query(L, "SELECT * FROM field_change_logs WHERE 模块='资产合同台账' AND 对象编号=? AND 字段='责任人'", (no,))
            assert fl, "未写字段级变更日志"
            return "PUT 200；field_change_logs 责任人 %s→%s" % (fl[-1]["原值"], fl[-1]["新值"])

        r.case("TC-M05-09", "修改资产合同并留痕", "200 + 字段日志", t09)

        def t10():
            resp = new_asset()
            no = resp.json()["number"]
            ar = A.put(B + "/api/asset_contracts/arow_%s/archive" % no, json={}, timeout=30)
            assert ar.status_code == 200, "归档 HTTP %s" % ar.status_code
            up = A.put(B + "/api/asset_contracts/arow_%s" % no, json={"remark": "x"}, timeout=30)
            dl = A.delete(B + "/api/asset_contracts/arow_%s" % no, json={"reason": "x"}, timeout=30)
            rn = A.put(B + "/api/asset_contracts/arow_%s/rent_collections" % no,
                       json={"items": [{"period_label": "2027年第1季度", "due_amount": 1}]}, timeout=30)
            assert up.status_code == 403, "归档后 PUT 返回 %s" % up.status_code
            assert dl.status_code == 403, "归档后 DELETE 返回 %s" % dl.status_code
            assert rn.status_code == 403, "归档后改租金返回 %s" % rn.status_code
            return "归档 200（归档人=%s）；归档后 PUT/DELETE/租金 均 403" % arow(no)["归档人"]

        r.case("TC-M05-10", "归档后只读（含租金明细）", "403", t10)

        def t11():
            resp = new_asset()
            no = resp.json()["number"]
            dl = A.delete(B + "/api/asset_contracts/arow_%s" % no, json={}, timeout=30)
            assert dl.status_code == 400, "无 reason 删除返回 %s（期望 400）" % dl.status_code
            assert arow(no), "资产被误删"
            return "无 reason → 400，未删除"

        r.case("TC-M05-11", "删除必须填写原因", "400", t11)

        def t12():
            with_rent = C.query(L, "SELECT 资产合同ID a, COUNT(*) c FROM asset_rent_collections "
                                   "GROUP BY 资产合同ID HAVING c>0 ORDER BY c DESC LIMIT 1")
            assert with_rent, "库内无带租金记录的资产合同样本"
            aid = str(with_rent[0]["a"])
            dl = A.delete(B + "/api/asset_contracts/arow_%s" % aid, json={"reason": "自动化测试"}, timeout=30)
            assert dl.status_code == 409, "有租金记录删除返回 %s（期望 409）" % dl.status_code
            body = dl.json() if dl.content else {}
            assert str(body.get("rent_count", "")) == str(with_rent[0]["c"]), \
                "409 未返回 rent_count（实际 %s，期望 %s）" % (body.get("rent_count"), with_rent[0]["c"])
            assert C.query(L, "SELECT COUNT(*) c FROM asset_rent_collections WHERE 资产合同ID=?", (aid,))[0]["c"] > 0, \
                "409 后租金记录被清理"
            return "有 %d 条租金记录且未传 delete_rent → 409 rent_count=%s" % (
                with_rent[0]["c"], body.get("rent_count"))

        r.case("TC-M05-12", "有租金记录删除 → 409 确认", "409 + rent_count", t12)

        def t13():
            resp = new_asset()
            no = resp.json()["number"]
            rent = A.put(B + "/api/asset_contracts/arow_%s/rent_collections" % no, json={"items": [
                {"period_label": "2026年第1季度", "due_amount": 100, "paid_amount": 100},
                {"period_label": "2026年第2季度", "due_amount": 100, "paid_amount": 0}]}, timeout=30)
            assert rent.status_code == 200, "写入租金失败 %s %s" % (rent.status_code, rent.text[:160])
            assert C.query(L, "SELECT COUNT(*) c FROM asset_rent_collections WHERE 资产合同ID=?", (no,))[0]["c"] == 2
            dl = A.delete(B + "/api/asset_contracts/arow_%s" % no,
                          json={"reason": "级联清理", "delete_rent": True}, timeout=30)
            assert dl.status_code == 200, "带 delete_rent 删除返回 %s %s" % (dl.status_code, dl.text[:160])
            assert arow(no) is None, "主记录未删除"
            assert C.query(L, "SELECT COUNT(*) c FROM asset_rent_collections WHERE 资产合同ID=?", (no,))[0]["c"] == 0, \
                "租金明细未级联清理"
            return "delete_rent=true → 200，主记录与 2 条明细一并清理"

        r.case("TC-M05-13", "带 delete_rent 删除并级联", "200 + 明细清理", t13)

        def t14():
            no1 = new_asset(person="别人").json()["number"]
            ok = A.put(B + "/api/asset_contracts/arow_%s/archive" % no1, json={}, timeout=30)
            assert ok.status_code == 200, "管理员归档失败 %s" % ok.status_code
            no2 = new_asset(person="别人").json()["number"]
            bad = P.put(B + "/api/asset_contracts/arow_%s/archive" % no2, json={}, timeout=30)
            assert bad.status_code == 403, "非责任人归档返回 %s（期望 403）" % bad.status_code
            no3 = new_asset(person=ip["name"]).json()["number"]
            ok2 = P.put(B + "/api/asset_contracts/arow_%s/archive" % no3, json={}, timeout=30)
            assert ok2.status_code == 200, "责任人归档返回 %s（期望 200）%s" % (ok2.status_code, ok2.text[:160])
            return "管理员 200；非责任人 403；责任人(%s) 200" % ip["name"]

        r.case("TC-M05-14", "归档限管理员或责任人", "200/403", t14)

        def t15():
            no = new_asset(next_lease_date="2027-01-31").json()["number"]
            assert str(arow(no)["下次出租时间"]).strip(), "样本未写入下次出租时间"
            rc = A.put(B + "/api/asset_contracts/arow_%s/clear_remind" % no, json={}, timeout=30)
            assert rc.status_code == 200, "clear_remind HTTP %s %s" % (rc.status_code, rc.text[:160])
            assert not str(arow(no)["下次出租时间"] or "").strip(), "下次出租时间未清空"
            fl = C.query(L, "SELECT * FROM field_change_logs WHERE 模块='资产合同台账' AND 对象编号=? "
                            "AND 字段='下次出租时间'", (no,))
            assert fl, "清空提醒未写字段级日志"
            return "清空提醒 200，置空并留痕 %d 条" % len(fl)

        r.case("TC-M05-15", "清空资产提醒", "200 + 置空 + 留痕", t15, priority="P1")

        def t16():
            g = A.get(B + "/api/asset_contracts/arow_999999", timeout=30)
            u = A.put(B + "/api/asset_contracts/arow_999999", json={"remark": "x"}, timeout=30)
            d = A.delete(B + "/api/asset_contracts/arow_999999", json={"reason": "x"}, timeout=30)
            b = A.get(B + "/api/asset_contracts/xyz", timeout=30)
            assert g.status_code == 404, "不存在详情 %s" % g.status_code
            assert u.status_code == 404, "不存在修改 %s" % u.status_code
            assert d.status_code == 404, "不存在删除 %s" % d.status_code
            assert b.status_code == 400, "非法 id %s" % b.status_code
            return "不存在资源 404（GET/PUT/DELETE）；非法 id 格式 400"

        r.case("TC-M05-16", "不存在/非法资产 id", "404 / 400", t16, priority="P1")

        srv.sync_db()
    finally:
        srv.stop()

    prod_after = C.prod_fingerprint()
    meta = {
        "服务地址": B,
        "工作库": "test_data/m05_asset.db",
        "生产库未改动": "是" if prod_before == prod_after else "否（指纹变化！）",
    }
    r.report("M05_资产合同CRUD测试.md", meta=meta)
    n, p, f, e = r.stat()
    print("\n[SUMMARY] M05 共 %d 例：通过 %d，失败 %d，错误 %d" % (n, p, f, e))
    sys.stdout.flush()
    os._exit(0 if (f == 0 and e == 0) else 1)


if __name__ == "__main__":
    main()
