# -*- coding: utf-8 -*-
"""M07 首页卡片口径测试（L3）。
覆盖：卡片字段齐全；总合同数=项目+资产；总合同额=Σ项目金额+Σ资产总合同额/10000；
到期计数；待支口径（未归档 且 待支付>0；未支付为空时用 合同金额−已支付）；
待收口径（Σ应缴−Σ实缴>0 的资产合同数 / 差额合计万元）；合计卡片=待收+待支；悬停构成分量；权限过滤；归档参与提醒。
期望值由 tests/_oracle.py 独立实现，不照抄被测代码。
"""
import os
import sys
import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _common as C  # noqa: E402
import _oracle as O  # noqa: E402


def main():
    C.setup_stdout()
    C.ensure_dirs()
    C.prepare_runtime()
    C.prepare_base()
    db = C.work_db("m07_dashboard.db")
    prod_before = C.prod_fingerprint()
    r = C.R("M07", "M07 首页卡片口径测试报告（L3）")
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
            assert s, "%s 账号登录失败" % role
            ctx[role] = {"s": s, "un": info["username"]}
        today = datetime.date.today()
        dd = int(float(C.query(L, "SELECT 配置值 v FROM settings WHERE 配置项='default_remind_days'")[0]["v"]))
        projects = C.query(L, "SELECT * FROM project_contracts")
        assets = C.query(L, "SELECT * FROM asset_contracts")
        rent_totals = O.rent_totals(C.query(L, "SELECT * FROM asset_rent_collections"))

        def expect(un):
            vis = C.visible_depts_of(L, un)
            projects = C.query(L, "SELECT * FROM project_contracts")
            assets = C.query(L, "SELECT * FROM asset_contracts")
            rt = O.rent_totals(C.query(L, "SELECT * FROM asset_rent_collections"))
            ps = [c for c in projects if O.in_scope(c.get("所属部门"), vis)]
            as_ = [a for a in assets if O.in_scope(a.get("管理部门"), vis)]
            proj_amt = sum((O.num(c.get("合同金额（万元）")) or 0) for c in ps)
            asset_amt = sum((O.num(a.get("总合同额")) or 0) for a in as_) / 10000.0
            exp_p = sum(1 for c in ps if O.is_due(c, "下次招采时间", dd, today) is not None)
            exp_a = sum(1 for a in as_ if O.is_due(a, "下次出租时间", dd, today) is not None)
            unpaid = sum(O.pending_unpaid_project(c) for c in ps)
            pend = sum(1 for c in ps if not O.archived_project(c) and O.pending_unpaid_project(c) > 0)
            diffs = [O.receivable_diff(a, rt) for a in as_]
            recv_c = sum(1 for d in diffs if d > 0)
            recv_amt = round(sum(diffs) / 10000.0, 2)
            return {
                "project_count": len(ps), "asset_count": len(as_), "total_count": len(ps) + len(as_),
                "project_amount": round(proj_amt, 2), "asset_amount": round(asset_amt, 2),
                "total_amount": round(proj_amt + asset_amt, 2),
                "exp_project": exp_p, "exp_asset": exp_a, "exp_total": exp_p + exp_a,
                "pending_count": pend, "unpaid_total": round(unpaid, 2),
                "recv_count": recv_c, "recv_amount": recv_amt,
                "pending_total_count": recv_c + pend,
                "pending_total_amount": round(recv_amt + unpaid, 2),
                "vis": vis, "ps": ps, "as": as_,
            }

        def dash(role):
            resp = ctx[role]["s"].get(B + "/api/dashboard", timeout=60)
            assert resp.status_code == 200, "%s dashboard HTTP %s" % (role, resp.status_code)
            return resp.json()

        def t01():
            j = dash("管理员")
            need = ["total_contracts", "project_contract_count", "asset_contract_count", "project_amount",
                    "asset_amount", "total_amount", "expiring_contracts", "expiring_project_count",
                    "expiring_asset_count", "pending_project_count", "unpaid_amount_total",
                    "pending_receivable_count", "pending_receivable_amount",
                    "pending_total_count", "pending_total_amount"]
            miss = [k for k in need if k not in j]
            assert not miss, "缺少卡片字段：%s" % miss
            return "15 个卡片字段齐全（含待收/待支拆分与合并值）"

        r.case("TC-M07-01", "卡片字段齐全", "文档列明的字段全部返回", t01)

        def t02():
            out = []
            for role in ctx:
                j, e = dash(role), expect(ctx[role]["un"])
                assert j["project_contract_count"] == e["project_count"], "%s 项目数 %s≠%s" % (role, j["project_contract_count"], e["project_count"])
                assert j["asset_contract_count"] == e["asset_count"], "%s 资产数 %s≠%s" % (role, j["asset_contract_count"], e["asset_count"])
                assert j["total_contracts"] == e["total_count"], "%s 总数 %s≠%s" % (role, j["total_contracts"], e["total_count"])
                out.append("%s:%d=%d+%d" % (role, j["total_contracts"], j["project_contract_count"], j["asset_contract_count"]))
            return "；".join(out)

        r.case("TC-M07-02", "总合同数=项目+资产（按可见范围）", "与独立计算一致", t02)

        def t03():
            out = []
            for role in ctx:
                j, e = dash(role), expect(ctx[role]["un"])
                assert abs(j["project_amount"] - e["project_amount"]) < 0.05, "%s 项目额 %s≠%s" % (role, j["project_amount"], e["project_amount"])
                assert abs(j["asset_amount"] - e["asset_amount"]) < 0.05, "%s 资产额(万元) %s≠%s" % (role, j["asset_amount"], e["asset_amount"])
                assert abs(j["total_amount"] - e["total_amount"]) < 0.05, "%s 总额 %s≠%s" % (role, j["total_amount"], e["total_amount"])
                out.append("%s:%.2f=%.2f+%.2f" % (role, j["total_amount"], j["project_amount"], j["asset_amount"]))
            return "；".join(out)

        r.case("TC-M07-03", "总合同额=Σ项目金额+Σ资产总合同额/10000", "与独立计算一致", t03)

        def t04():
            out = []
            for role in ctx:
                j, e = dash(role), expect(ctx[role]["un"])
                assert j["expiring_project_count"] == e["exp_project"], "%s 项目到期 %s≠%s" % (role, j["expiring_project_count"], e["exp_project"])
                assert j["expiring_asset_count"] == e["exp_asset"], "%s 资产到期 %s≠%s" % (role, j["expiring_asset_count"], e["exp_asset"])
                assert j["expiring_contracts"] == e["exp_total"], "%s 到期合计 %s≠%s" % (role, j["expiring_contracts"], e["exp_total"])
                out.append("%s:%d(项%d+资%d)" % (role, j["expiring_contracts"], j["expiring_project_count"], j["expiring_asset_count"]))
            return "；".join(out)

        r.case("TC-M07-04", "到期提醒计数", "项目+资产，与独立计算一致", t04, priority="P1")

        def t05():
            out = []
            for role in ctx:
                j, e = dash(role), expect(ctx[role]["un"])
                assert j["pending_project_count"] == e["pending_count"], "%s 待支合同数 %s≠%s" % (role, j["pending_project_count"], e["pending_count"])
                assert abs(j["unpaid_amount_total"] - e["unpaid_total"]) < 0.05, "%s 待支金额 %s≠%s" % (role, j["unpaid_amount_total"], e["unpaid_total"])
                out.append("%s:%d个/%.2f万" % (role, j["pending_project_count"], j["unpaid_amount_total"]))
            return "；".join(out)

        r.case("TC-M07-05", "待支口径（未归档 且 待支付>0）", "与独立计算一致", t05, priority="P1")

        def t06():
            out = []
            for role in ctx:
                j, e = dash(role), expect(ctx[role]["un"])
                assert j["pending_receivable_count"] == e["recv_count"], "%s 待收合同数 %s≠%s" % (role, j["pending_receivable_count"], e["recv_count"])
                assert abs(j["pending_receivable_amount"] - e["recv_amount"]) < 0.05, "%s 待收金额(万元) %s≠%s" % (role, j["pending_receivable_amount"], e["recv_amount"])
                out.append("%s:%d个/%.2f万" % (role, j["pending_receivable_count"], j["pending_receivable_amount"]))
            return "；".join(out)

        r.case("TC-M07-06", "待收口径（Σ应缴−Σ实缴>0）", "与独立计算一致", t06, priority="P1")

        def t07():
            out = []
            for role in ctx:
                j = dash(role)
                assert j["pending_total_count"] == j["pending_receivable_count"] + j["pending_project_count"], \
                    "%s 计数合计不符：%s ≠ %s+%s" % (role, j["pending_total_count"], j["pending_receivable_count"], j["pending_project_count"])
                assert abs(j["pending_total_amount"] - (j["pending_receivable_amount"] + j["unpaid_amount_total"])) < 0.05, \
                    "%s 金额合计不符：%s" % (role, j["pending_total_amount"])
                out.append("%s:%d/%.2f万" % (role, j["pending_total_count"], j["pending_total_amount"]))
            return "；".join(out)

        r.case("TC-M07-07", "待收/支合计卡片=待收+待支", "计数与金额均成立", t07, priority="P1")

        def t08():
            j = dash("管理员")
            assert j["project_contract_count"] + j["asset_contract_count"] == j["total_contracts"], "悬停构成（项目+资产）≠总数"
            assert j["expiring_project_count"] + j["expiring_asset_count"] == j["expiring_contracts"], "悬停构成（到期拆分）≠到期总数"
            return "构成分量之和=总数（合同数 %d、到期 %d）" % (j["total_contracts"], j["expiring_contracts"])

        r.case("TC-M07-08", "悬停气泡构成分量之和=总数", "误差 0", t08, priority="P1")

        def t09():
            nums = {role: dash(role)["total_contracts"] for role in ctx}
            assert nums["管理员"] >= nums["普通用户"], "普通用户可见数多于管理员"
            assert nums["管理员"] >= nums["查询用户"], "查询用户可见数多于管理员"
            for role in ctx:
                j, e = dash(role), expect(ctx[role]["un"])
                assert j["total_amount"] == e["total_amount"] or abs(j["total_amount"] - e["total_amount"]) < 0.05
            return "管理员 %d / 普通用户 %d / 查询用户 %d（均与各自可见范围一致）" % (
                nums["管理员"], nums["普通用户"], nums["查询用户"])

        r.case("TC-M07-09", "卡片数据随权限过滤", "各角色数值=其可见范围", t09)

        def t10():
            # 判别性构造：新建一条「今天到期 + 未支付 100 万」的合同，再归档，观察两处口径变化
            def fv(col):
                return C.query(L, "SELECT %s v FROM project_contracts WHERE TRIM(COALESCE(%s,''))<>'' LIMIT 1"
                               % (col, col))[0]["v"]

            admin_name = C.query(L, "SELECT 姓名 n FROM users WHERE 用户名='admin'")[0]["n"]
            stamp = datetime.datetime.now().strftime("%H%M%S")
            body = {"category": fv("项目分类"), "name": "M07-10归档口径-" + stamp,
                    "procurement_method": fv("采购方式"), "amount": 1, "main_content": "口径测试",
                    "handler": admin_name, "fund_source": fv("资金来源"),
                    "next_bid_date": today.strftime("%Y-%m-%d"), "unpaid_amount": 100}
            before = dash("管理员")
            resp = ctx["管理员"]["s"].post(B + "/api/contracts", json=body, timeout=60)
            assert resp.status_code == 201, "构造样本失败 %s" % resp.status_code
            row = C.query(L, "SELECT id FROM project_contracts WHERE 编号=?", (resp.json()["number"],))[0]
            mid = dash("管理员")
            assert mid["expiring_project_count"] == before["expiring_project_count"] + 1, \
                "新建到期合同后到期计数未 +1：%s→%s" % (before["expiring_project_count"], mid["expiring_project_count"])
            assert mid["pending_project_count"] == before["pending_project_count"] + 1, \
                "新建待支合同后待支计数未 +1"
            ar = ctx["管理员"]["s"].put(B + "/api/contracts/row_%s/archive" % row["id"], json={}, timeout=30)
            assert ar.status_code == 200, "归档失败 %s" % ar.status_code
            after = dash("管理员")
            assert after["expiring_project_count"] == mid["expiring_project_count"], \
                "归档后到期计数变化（归档应仍参与提醒）：%s→%s" % (mid["expiring_project_count"], after["expiring_project_count"])
            assert after["pending_project_count"] == before["pending_project_count"], \
                "归档后待支计数未回落（归档不应参与待支）：%s→%s" % (mid["pending_project_count"], after["pending_project_count"])
            # 与 oracle 复核
            e = expect("admin")
            assert after["expiring_project_count"] == e["exp_project"], "归档后到期计数与独立计算不符"
            assert after["pending_project_count"] == e["pending_count"], "归档后待支计数与独立计算不符"
            return "构造「今天到期+未支付100万」样本：到期 %d→%d 归档后仍 %d（参与提醒）；待支 %d→%d 归档后回落 %d（不参与）" % (
                before["expiring_project_count"], mid["expiring_project_count"], after["expiring_project_count"],
                before["pending_project_count"], mid["pending_project_count"], after["pending_project_count"])

        r.case("TC-M07-10", "归档合同参与到期提醒、不参与待支", "构造样本验证口径", t10, priority="P1")

        srv.sync_db()
    finally:
        srv.stop()

    prod_after = C.prod_fingerprint()
    meta = {
        "服务地址": B,
        "工作库": "test_data/m07_dashboard.db",
        "账号": "admin / %s / %s" % (ctx["普通用户"]["un"], ctx["查询用户"]["un"]),
        "默认提醒天数": dd,
        "生产库未改动": "是" if prod_before == prod_after else "否（指纹变化！）",
    }
    r.report("M07_首页卡片口径测试.md", meta=meta)
    n, p, f, e = r.stat()
    print("\n[SUMMARY] M07 共 %d 例：通过 %d，失败 %d，错误 %d" % (n, p, f, e))
    sys.stdout.flush()
    os._exit(0 if (f == 0 and e == 0) else 1)


if __name__ == "__main__":
    main()
