# -*- coding: utf-8 -*-
"""M09 统计维度测试（L3）。
覆盖：项目侧多维度（部门/分类/资金来源/经办人/时间/未支付/预算执行）可用性与数值校验、
筛选（单位/部门/经办人/年度/日期区间）生效、按数据权限过滤、非法 type 不 500；
资产侧 by_dept/by_person/by_time 数值与筛选与权限。
期望值由 tests/_oracle.py 独立实现（Python 侧聚合，与页面口径一致）。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _common as C  # noqa: E402
import _oracle as O  # noqa: E402

PROJ_TYPES = ["by_dept", "by_category", "by_fund_source", "by_person", "by_time",
              "unpaid", "budget", "pay_dept", "pay_project", "budget_dept", "budget_project"]
ASSET_TYPES = ["by_dept", "by_person", "by_time"]


def main():
    C.setup_stdout()
    C.ensure_dirs()
    C.prepare_runtime()
    C.prepare_base()
    db = C.work_db("m09_stat.db")
    prod_before = C.prod_fingerprint()
    r = C.R("M09", "M09 统计维度测试报告（L3）")
    srv = C.Server(db)
    srv.start()
    B = srv.base
    L = C.RUNTIME_DB
    try:
        ctx = {}
        for role, un in (("管理员", "admin"), ("普通用户", "章伟"), ("查询用户", "谢玮")):
            s, info = C.login(un, base=B)
            if not s:
                rows = C.query(L, "SELECT 用户名 FROM users WHERE 角色=? ORDER BY 用户名 LIMIT 1", (role,))
                s, info = C.login(rows[0]["用户名"], base=B)
            assert s, "%s 登录失败" % role
            ctx[role] = {"s": s, "un": info["username"]}

        def proj_rows(un=None):
            rows = C.query(L, "SELECT * FROM project_contracts")
            if un is None:
                return rows
            vis = C.visible_depts_of(L, un)
            return [c for c in rows if O.in_scope(c.get("所属部门"), vis)]

        def asset_rows(un=None):
            rows = C.query(L, "SELECT * FROM asset_contracts")
            if un is None:
                return rows
            vis = C.visible_depts_of(L, un)
            return [a for a in rows if O.in_scope(a.get("管理部门"), vis)]

        def stat(role, **params):
            resp = ctx[role]["s"].get(B + "/api/statistics", params=params, timeout=60)
            return resp

        def rent_totals():
            return O.rent_totals(C.query(L, "SELECT * FROM asset_rent_collections"))

        def astat(role, **params):
            return ctx[role]["s"].get(B + "/api/asset_statistics", params=params, timeout=60)

        def t01():
            ok, empty, bad = [], [], []
            for t in PROJ_TYPES:
                resp = stat("管理员", type=t)
                if resp.status_code != 200:
                    bad.append("%s=%s" % (t, resp.status_code))
                    continue
                j = resp.json()
                if not isinstance(j.get("data"), list):
                    bad.append("%s:无 data" % t)
                elif j["data"]:
                    ok.append(t)
                else:
                    empty.append(t)
            assert not bad, "维度异常：%s" % bad
            return "11 个维度全部 200 且返回 data：非空 %d 个 %s；空 %s" % (len(ok), ok, empty or "无")

        r.case("TC-M09-01", "项目统计各维度可用", "全部 200 且 data 为列表", t01)

        def t02():
            rows = proj_rows("admin")
            exp = {}
            for c in rows:
                k = c.get("所属部门") or "未知"
                g = exp.setdefault(k, [0, 0.0])
                g[0] += 1
                g[1] += (O.num(c.get("合同金额（万元）")) or 0)
            j = stat("管理员", type="by_dept").json()
            got = {x["部门"]: (x["合同数"], x["合同金额"]) for x in j["data"]}
            assert set(got) == set(exp), "部门集合不一致：%s vs %s" % (sorted(got)[:3], sorted(exp)[:3])
            for k, (n, amt) in exp.items():
                assert got[k][0] == n, "部门 %s 合同数 %s≠%s" % (k, got[k][0], n)
                assert abs(got[k][1] - round(amt, 2)) < 0.05, "部门 %s 金额 %s≠%s" % (k, got[k][1], round(amt, 2))
            return "%d 个部门的合同数与金额均与独立聚合一致（合计 %d 条）" % (len(exp), len(rows))

        r.case("TC-M09-02", "by_dept 数值=独立聚合", "逐部门一致", t02)

        def t03():
            exp = {}
            for c in proj_rows("admin"):
                k = c.get("项目分类") or "未知"
                g = exp.setdefault(k, [0, 0.0])
                g[0] += 1
                g[1] += (O.num(c.get("合同金额（万元）")) or 0)
            j = stat("管理员", type="by_category").json()
            got = {x["分类"]: (x["合同数"], x["合同金额"]) for x in j["data"]}
            assert set(got) == set(exp), "分类集合不一致"
            for k in exp:
                assert got[k][0] == exp[k][0] and abs(got[k][1] - round(exp[k][1], 2)) < 0.05, "分类 %s 数值不符" % k
            return "%d 个项目分类的合同数与金额一致" % len(exp)

        r.case("TC-M09-03", "by_category 数值=独立聚合", "逐分类一致", t03)

        def t04():
            row = C.query(L, "SELECT 所属部门 d, COUNT(*) c FROM project_contracts "
                             "WHERE 所属部门<>'' GROUP BY 所属部门 ORDER BY c DESC LIMIT 1")[0]
            sub = str(row["d"]).split("-")[-1]
            j = stat("管理员", type="by_dept", dept=sub).json()
            exp_n = sum(x["合同数"] for x in j["data"])
            sql_n = len([c for c in proj_rows("admin") if sub in str(c.get("所属部门", ""))])
            assert exp_n == sql_n, "dept 筛选后合同数 %s≠%s" % (exp_n, sql_n)
            person = C.query(L, "SELECT 经办人 p, COUNT(*) c FROM project_contracts "
                                "WHERE 经办人<>'' GROUP BY 经办人 ORDER BY c DESC LIMIT 1")[0]["p"]
            j2 = stat("管理员", type="by_dept", person=person).json()
            exp2 = sum(x["合同数"] for x in j2["data"])
            sql2 = len([c for c in proj_rows("admin") if person in str(c.get("经办人", ""))])
            assert exp2 == sql2, "person 筛选后合同数 %s≠%s" % (exp2, sql2)
            return "部门筛选(%s)=%d、经办人筛选(%s)=%d 均与独立计算一致" % (sub, sql_n, person, sql2)

        r.case("TC-M09-04", "筛选参数（dept/person）生效", "与独立计算一致", t04)

        def t05():
            years = {}
            for c in proj_rows("admin"):
                d = O.parse_date(c.get("采购时间"))
                if d:
                    years[d.year] = years.get(d.year, 0) + 1
            assert years, "库内无带采购时间的合同"
            y = str(max(years, key=lambda k: years[k]))
            j = stat("管理员", type="by_dept", year=y).json()
            got = sum(x["合同数"] for x in j["data"])
            assert got == years[int(y)], "year=%s 筛选结果 %s≠%s" % (y, got, years[int(y)])
            sd, ed = "%s-01-01" % y, "%s-12-31" % y
            j2 = stat("管理员", type="by_dept", start_date=sd, end_date=ed).json()
            got2 = sum(x["合同数"] for x in j2["data"])
            assert got2 == years[int(y)], "日期区间筛选 %s≠%s" % (got2, years[int(y)])
            return "year=%s 与 start/end_date 区间筛选结果均为 %d 条，与独立计算一致" % (y, got)

        r.case("TC-M09-05", "筛选参数（year/日期区间）生效", "与独立计算一致", t05, priority="P1")

        def t06():
            out = []
            for role in ctx:
                j = stat(role, type="by_dept").json()
                got = sum(x["合同数"] for x in j["data"])
                exp = len(proj_rows(ctx[role]["un"]))
                assert got == exp, "%s 统计合同数 %s≠可见范围 %s" % (role, got, exp)
                out.append("%s:%d" % (role, got))
            return "；".join(out)

        r.case("TC-M09-06", "项目统计按数据权限过滤", "与各账号可见范围一致", t06)

        def t07():
            rows = proj_rows("admin")
            exp = {}
            for c in rows:
                k = c.get("所属部门") or "未知"
                exp[k] = exp.get(k, 0.0) + O.pending_unpaid_project(c)
            j = stat("管理员", type="pay_dept").json()
            got = {}
            for x in j["data"]:
                got[x["部门"]] = x["未支付"]
            assert set(got) == set(exp), "未支付统计部门集合不一致"
            diff = [k for k in exp if abs(got[k] - round(exp[k], 2)) > 0.05]
            assert not diff, "未支付金额不符：%s" % diff[:3]
            total_got = sum(got.values())
            total_exp = sum(O.pending_unpaid_project(c) for c in rows)
            assert abs(total_got - total_exp) < 0.1, "未支付合计 %s≠%s" % (total_got, total_exp)
            return "按部门的未支付合计 %.2f 万元与独立计算一致（%d 个部门）" % (total_got, len(exp))

        r.case("TC-M09-07", "未支付（支付情况）维度数值", "逐部门一致", t07, priority="P1")

        def t08():
            j = stat("管理员", type="budget_dept").json()
            assert j["data"], "预算执行维度无数据"
            x = j["data"][0]
            for k in ("招标控制价", "合同金额", "差额", "偏差率", "合同数"):
                assert k in x, "预算执行缺少字段 %s" % k
            if x["招标控制价"]:
                exp_rate = round((x["合同金额"] - x["招标控制价"]) / x["招标控制价"] * 100, 2)
                assert abs(x["偏差率"] - exp_rate) < 0.05, "偏差率 %s≠%s" % (x["偏差率"], exp_rate)
            assert abs(x["差额"] - round(x["合同金额"] - x["招标控制价"], 2)) < 0.05, "差额计算不符"
            return "%d 个部门；首行 招标控制价=%.2f 合同金额=%.2f 差额=%.2f 偏差率=%s%%（口径正确）" % (
                len(j["data"]), x["招标控制价"], x["合同金额"], x["差额"], x["偏差率"])

        r.case("TC-M09-08", "预算执行维度（偏差率/差额）", "口径正确", t08, priority="P1")

        def t09():
            resp = stat("管理员", type="unknown_type_xyz")
            assert resp.status_code != 500, "非法 type 返回 500（应 200 空数据或 4xx）"
            body = ""
            try:
                j = resp.json()
                body = "data=%s" % (len(j.get("data", [])) if isinstance(j.get("data"), list) else j.get("data"))
            except Exception:
                body = resp.text[:60]
            return "type=unknown_type_xyz → HTTP %s（%s）" % (resp.status_code, body)

        r.case("TC-M09-09", "非法 type 不返回 500", "200 空数据 或 4xx", t09, priority="P2")

        def t10():
            rows = asset_rows("admin")
            exp = {}
            for a in rows:
                k = a.get("管理部门") or "未指定"
                g = exp.setdefault(k, [0, 0.0, 0.0, 0.0])
                g[0] += 1
                g[1] += (O.num(a.get("面积（㎡）")) or 0)
                g[2] += (O.num(a.get("年租金（元）")) or 0)
                t = rent_totals().get(O.asset_key(a)) or {"paid": 0.0}
                g[3] += float(t["paid"] or 0)
            j = astat("管理员", type="by_dept").json()
            assert j.get("data") is not None, "asset_statistics 无 data"
            got = {}
            for x in j["data"]:
                got[x["管理部门"]] = x
            assert set(got) == set(exp), "资产统计部门集合不一致：%s vs %s" % (sorted(got), sorted(exp))
            for k, (n, area, rent, paid) in exp.items():
                x = got[k]
                assert x["资产数"] == n, "%s 资产数 %s≠%s" % (k, x["资产数"], n)
                assert abs(x["面积合计(㎡)"] - round(area, 2)) < 0.05, "%s 面积不符" % k
                assert abs(x["年租金合计(元)"] - round(rent, 2)) < 0.05, "%s 年租金不符" % k
                assert abs(x["已缴租金合计(元)"] - round(paid, 2)) < 0.05, "%s 已缴租金不符" % k
            return "%d 个管理部门的 资产数/面积/年租金/已缴租金 均与独立聚合一致" % len(exp)

        r.case("TC-M09-10", "资产统计 by_dept 数值", "逐部门一致（含已缴租金）", t10)

        def t11():
            out = []
            for t in ("by_person", "by_time"):
                resp = astat("管理员", type=t)
                assert resp.status_code == 200, "asset_statistics type=%s HTTP %s" % (t, resp.status_code)
                j = resp.json()
                assert isinstance(j.get("data"), list), "type=%s 无 data" % t
                total = sum(x.get("资产数", 0) for x in j["data"])
                assert total == len(asset_rows("admin")), "type=%s 资产数合计 %s≠%s" % (t, total, len(asset_rows("admin")))
                out.append("%s:%d行/合计%d" % (t, len(j["data"]), total))
            dep = asset_rows("admin")[0]["管理部门"]
            j2 = astat("管理员", type="by_dept", dept=str(dep).split("-")[-1]).json()
            exp_n = len([a for a in asset_rows("admin") if str(dep).split("-")[-1] in str(a.get("管理部门", ""))])
            assert sum(x["资产数"] for x in j2["data"]) == exp_n, "资产统计 dept 筛选不符"
            out.append("dept筛选=%d" % exp_n)
            return "；".join(out)

        r.case("TC-M09-11", "资产统计 by_person/by_time 与筛选", "合计一致、筛选生效", t11, priority="P1")

        def t12():
            out = []
            for role in ctx:
                j = astat(role, type="by_dept").json()
                got = sum(x["资产数"] for x in j["data"])
                exp = len(asset_rows(ctx[role]["un"]))
                assert got == exp, "%s 资产统计 %s≠可见 %s" % (role, got, exp)
                out.append("%s:%d" % (role, got))
            return "；".join(out)

        r.case("TC-M09-12", "资产统计按数据权限过滤", "与各账号可见范围一致", t12)

        srv.sync_db()
    finally:
        srv.stop()

    prod_after = C.prod_fingerprint()
    meta = {
        "服务地址": B,
        "工作库": "test_data/m09_stat.db",
        "生产库未改动": "是" if prod_before == prod_after else "否（指纹变化！）",
    }
    r.report("M09_统计维度测试.md", meta=meta)
    n, p, f, e = r.stat()
    print("\n[SUMMARY] M09 共 %d 例：通过 %d，失败 %d，错误 %d" % (n, p, f, e))
    sys.stdout.flush()
    os._exit(0 if (f == 0 and e == 0) else 1)


if __name__ == "__main__":
    main()
