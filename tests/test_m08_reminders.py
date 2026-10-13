# -*- coding: utf-8 -*-
"""M08 到期提醒 + 待收/支提醒测试（L3）。
覆盖：提醒合并与升序、kind/next_label/剩余天数、提前天数（合同级优先、空取全局默认）、今天到期边界、
清空提醒后消失、权限过滤；待收在前待支在后、待收字段、计数一致、差额=0 与实缴>应缴不计入、权限过滤。
期望值由 tests/_oracle.py 独立实现。
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
    db = C.work_db("m08_reminders.db")
    prod_before = C.prod_fingerprint()
    r = C.R("M08", "M08 到期提醒与待收/支提醒测试报告（L3）")
    srv = C.Server(db)
    srv.start()
    B = srv.base
    L = C.RUNTIME_DB
    stamp = datetime.datetime.now().strftime("%H%M%S")
    try:
        ctx = {}
        for role, un in (("管理员", "admin"), ("普通用户", "章伟"), ("查询用户", "谢玮")):
            s, info = C.login(un, base=B)
            if not s:
                rows = C.query(L, "SELECT 用户名 FROM users WHERE 角色=? ORDER BY 用户名 LIMIT 1", (role,))
                s, info = C.login(rows[0]["用户名"], base=B)
            assert s, "%s 登录失败" % role
            ctx[role] = {"s": s, "un": info["username"], "name": info["name"]}
        today = datetime.date.today()
        dd = int(float(C.query(L, "SELECT 配置值 v FROM settings WHERE 配置项='default_remind_days'")[0]["v"]))

        def proj_rows():
            return C.query(L, "SELECT * FROM project_contracts")

        def asset_rows():
            return C.query(L, "SELECT * FROM asset_contracts")

        def rent_totals():
            return O.rent_totals(C.query(L, "SELECT * FROM asset_rent_collections"))

        def exp_reminders(un, use_row_days=True):
            vis = C.visible_depts_of(L, un)
            out = []
            for c in proj_rows():
                if not O.in_scope(c.get("所属部门"), vis):
                    continue
                d = O.parse_date(c.get("下次招采时间"))
                if not d:
                    continue
                rd = dd
                if use_row_days and O.num(c.get("提醒提前天数")) is not None:
                    rd = int(O.num(c.get("提醒提前天数")))
                if today < d - datetime.timedelta(days=rd):
                    continue
                out.append(("project", str(c.get("编号")), (d - today).days))
            for a in asset_rows():
                if not O.in_scope(a.get("管理部门"), vis):
                    continue
                d = O.parse_date(a.get("下次出租时间"))
                if not d:
                    continue
                rd = dd
                if use_row_days and O.num(a.get("提醒提前天数")) is not None:
                    rd = int(O.num(a.get("提醒提前天数")))
                if today < d - datetime.timedelta(days=rd):
                    continue
                out.append(("asset", str(a.get("id")), (d - today).days))
            return out

        def rem(role):
            resp = ctx[role]["s"].get(B + "/api/reminders", timeout=60)
            assert resp.status_code == 200, "%s reminders HTTP %s" % (role, resp.status_code)
            return resp.json()

        def t01():
            j = rem("管理员")
            items = j["items"]
            assert j["count"] == len(items), "count=%s 与 items=%d 不符" % (j["count"], len(items))
            kinds = set(x.get("kind") for x in items)
            assert kinds <= {"project", "asset"}, "kind 取值异常：%s" % kinds
            for x in items:
                assert "next_label" in x and "remaining_days" in x, "缺少 next_label/remaining_days：%s" % list(x)[:6]
            days = [x["remaining_days"] for x in items]
            assert days == sorted(days), "未按剩余天数升序：%s" % days[:10]
            exp = exp_reminders("admin")
            assert j["count"] == len(exp), "提醒条数 %s≠期望 %s" % (j["count"], len(exp))
            return "共 %d 条（项目/资产合并），剩余天数升序，含 kind/next_label；与独立计算一致" % j["count"]

        r.case("TC-M08-01", "提醒合并、字段与升序", "与独立计算一致且升序", t01)

        def t02():
            # 判别性构造：下次招采 = 今天+30，提醒提前天数 = 0 → 用「合同级 0」判定不命中；
            # 若错误地使用全局默认（dd>30）则会命中。以此证明采用了合同级取值。
            cand = C.query(L, "SELECT id,编号 FROM project_contracts "
                              "WHERE COALESCE(是否归档,'')<>'是' ORDER BY id LIMIT 1")[0]
            no = str(cand["编号"])
            target = (today + datetime.timedelta(days=30)).strftime("%Y-%m-%d")
            up = ctx["管理员"]["s"].put(B + "/api/contracts/row_%s" % cand["id"],
                                        json={"next_bid_date": target, "remind_days": 0}, timeout=60)
            assert up.status_code == 200, "构造样本失败 %s" % up.status_code
            items = rem("管理员")["items"]
            hit = [x for x in items if x.get("number") == no]
            assert not hit, "提前天数=0 仍被提醒（疑似错误使用全局默认 %d）：%s" % (dd, no)
            # 反证：把提前天数设为 31（>30）→ 应命中并剩余 30 天
            up2 = ctx["管理员"]["s"].put(B + "/api/contracts/row_%s" % cand["id"],
                                         json={"remind_days": 31}, timeout=60)
            assert up2.status_code == 200, "二次构造失败 %s" % up2.status_code
            hit2 = [x for x in rem("管理员")["items"] if x.get("number") == no]
            assert hit2 and hit2[0]["remaining_days"] == 30, "提前天数=31 应命中且剩余 30 天：%s" % hit2
            return "合同级提醒天数生效：置 0 不命中、置 31 命中（剩余 30 天），未错用全局默认 %d" % dd

        r.case("TC-M08-02", "提前天数：合同级优先", "采用合同级提醒提前天数", t02, priority="P1")

        def t03():
            # 通过接口构造边界样本（同时验证接口写入会正确失效只读缓存）：
            # 下次招采 = 今天+30；提前天数置空 → 走全局默认 dd（60）→ 应命中（剩余 30 天）
            cand = C.query(L, "SELECT id,编号 FROM project_contracts "
                              "WHERE COALESCE(是否归档,'')<>'是' ORDER BY id LIMIT 1")[0]
            no = str(cand["编号"])
            target = (today + datetime.timedelta(days=30)).strftime("%Y-%m-%d")
            up = ctx["管理员"]["s"].put(B + "/api/contracts/row_%s" % cand["id"],
                                        json={"next_bid_date": target, "remind_days": ""}, timeout=60)
            assert up.status_code == 200, "构造样本失败 %s %s" % (up.status_code, up.text[:120])
            assert O.num(C.query(L, "SELECT 提醒提前天数 v FROM project_contracts WHERE id=?", (cand["id"],))[0]["v"]) is None, \
                "提醒提前天数未被置空"
            hit = [x for x in rem("管理员")["items"] if x.get("number") == no]
            assert hit, "空提醒天数未按全局默认 %d 天判定（应命中）" % dd
            assert hit[0]["remaining_days"] == 30, "剩余天数=%s（应为 30）" % hit[0]["remaining_days"]
            # 反证：提前天数=0 → 今天 < 下次招采 − 0 → 不应命中
            up2 = ctx["管理员"]["s"].put(B + "/api/contracts/row_%s" % cand["id"],
                                         json={"remind_days": 0}, timeout=60)
            assert up2.status_code == 200, "置 0 失败 %s" % up2.status_code
            assert all(x.get("number") != no for x in rem("管理员")["items"]), "提前天数=0 时不应命中"
            return "空提前天数按全局默认 %d 天命中（剩余 30 天）；置 0 后不命中 → 默认值回退正确" % dd

        r.case("TC-M08-03", "提前天数：空值取全局默认", "与全局默认口径一致", t03, priority="P1")

        def t04():
            items = rem("管理员")["items"]
            zero = [x for x in items if x["remaining_days"] == 0]
            neg = [x for x in items if x["remaining_days"] < 0]
            exp_zero = [x for x in exp_reminders("admin") if x[2] == 0]
            assert len(zero) == len(exp_zero), "今天到期条目数 %d≠期望 %d" % (len(zero), len(exp_zero))
            return "剩余 0 天条目 %d 条（与期望一致）；已过期仍在列 %d 条" % (len(zero), len(neg))

        r.case("TC-M08-04", "边界：今天到期（剩余 0 天）", "出现在提醒中", t04, priority="P1")

        def t05():
            items = rem("管理员")["items"]
            assert items, "当前无提醒样本"
            target = None
            for x in items:
                if x.get("kind") == "project":
                    target = x
                    break
            if target is None:
                return "当前仅有资产提醒，跳过（项目清提醒由 M04-18 覆盖）"
            no = target["number"]
            row = C.query(L, "SELECT id FROM project_contracts WHERE 编号=?", (no,))[0]
            rc = ctx["管理员"]["s"].put(B + "/api/contracts/row_%s/clear_remind" % row["id"], json={}, timeout=30)
            assert rc.status_code == 200, "clear_remind HTTP %s" % rc.status_code
            after = rem("管理员")["items"]
            assert all(x.get("number") != no for x in after), "清空提醒后仍出现在列表中：%s" % no
            return "对编号 %s 执行清空提醒后，该条从提醒列表消失（%d → %d）" % (no, len(items), len(after))

        r.case("TC-M08-05", "清空提醒后不再提醒", "从列表消失", t05, priority="P1")

        def t06():
            out = []
            for role in ctx:
                j = rem(role)
                vis = C.visible_depts_of(L, ctx[role]["un"])
                exp = exp_reminders(ctx[role]["un"])
                got_set = set((x.get("kind"), str(x.get("number"))) for x in j["items"])
                exp_set = set((k, str(n)) for (k, n, _d) in exp)
                assert got_set == exp_set, "%s 提醒集合不一致：仅接口=%s；仅期望=%s（接口 %d 条 / 期望 %d 条）" % (
                    role, sorted(got_set - exp_set)[:3], sorted(exp_set - got_set)[:3], len(got_set), len(exp_set))
                if vis is not None:
                    bad = [x["number"] for x in j["items"] if str(x.get("department", "")).strip() not in vis]
                    assert not bad, "%s 返回越权提醒：%s" % (role, bad[:3])
                out.append("%s:%d" % (role, j["count"]))
            return "；".join(out)

        r.case("TC-M08-06", "提醒按数据权限过滤", "与各账号可见范围一致", t06)

        def pend(role):
            resp = ctx[role]["s"].get(B + "/api/pending_payments", timeout=60)
            assert resp.status_code == 200, "%s pending HTTP %s" % (role, resp.status_code)
            return resp.json()

        def t07():
            j = pend("管理员")
            kinds = [x.get("kind") for x in j["items"]]
            seen_payable = False
            for k in kinds:
                if k == "payable":
                    seen_payable = True
                elif seen_payable and k == "receivable":
                    raise AssertionError("待收项出现在待支项之后（排序不符）")
            return "顺序校验通过：receivable×%d 在前、payable×%d 在后" % (
                j["receivable_count"], j["payable_count"])

        r.case("TC-M08-07", "待收在前、待支在后", "receivable 全部先于 payable", t07)

        def t08():
            j = pend("管理员")
            rec = [x for x in j["items"] if x.get("kind") == "receivable"]
            assert rec, "无可校验的待收项"
            miss = [k for k in ("due_amount", "received_amount", "next_lease_date", "pending_amount")
                    if k not in rec[0]]
            assert not miss, "待收项缺字段：%s" % miss
            for x in rec:
                assert abs(x["pending_amount"] - (x["due_amount"] - x["received_amount"])) < 0.02, \
                    "待收金额≠应缴−实缴（万元）：%s" % x
            return "待收项 %d 条，含 due/received/next_lease_date/pending_amount（万元），且 pending=due−received" % len(rec)

        r.case("TC-M08-08", "待收项字段与万元换算", "字段齐全且等式成立", t08, priority="P1")

        def t09():
            j = pend("管理员")
            rc = sum(1 for x in j["items"] if x.get("kind") == "receivable")
            pc = sum(1 for x in j["items"] if x.get("kind") == "payable")
            assert j["receivable_count"] == rc and j["payable_count"] == pc, \
                "计数与明细不符：%s/%s vs %s/%s" % (j["receivable_count"], j["payable_count"], rc, pc)
            assert j["count"] == rc + pc, "总计数不符：%s≠%s" % (j["count"], rc + pc)
            return "count=%d = receivable %d + payable %d" % (j["count"], rc, pc)

        r.case("TC-M08-09", "待收/待支计数与明细一致", "计数=明细条数", t09, priority="P1")

        def t10():
            j = pend("管理员")
            bad = [x for x in j["items"] if x.get("kind") == "receivable" and float(x["pending_amount"]) <= 0]
            assert not bad, "存在差额≤0 的待收项：%s" % bad[:2]
            vis = None
            exp = []
            for a in asset_rows():
                if O.receivable_diff(a, rent_totals()) > 0:
                    exp.append(str(a.get("id")))
            got = sorted(str(x["number"]) for x in j["items"] if x.get("kind") == "receivable")
            assert got == sorted(exp), "待收集合与独立计算不符：%s vs %s" % (got[:5], sorted(exp)[:5])
            zero = [a for a in asset_rows() if O.receivable_diff(a, rent_totals()) == 0]
            return "差额>0 的资产合同 %d 条全部计入；差额=0 的 %d 条未计入" % (len(exp), len(zero))

        r.case("TC-M08-10", "差额=0 不计入待收", "集合与独立计算一致", t10, priority="P1")

        def t11():
            over = [a for a in asset_rows() if O.receivable_diff(a, rent_totals()) < 0]
            j = pend("管理员")
            ids = set(str(x["number"]) for x in j["items"] if x.get("kind") == "receivable")
            bad = [str(a.get("id")) for a in over if str(a.get("id")) in ids]
            assert not bad, "多缴（实缴>应缴）合同被计入待收：%s" % bad
            return "多缴（差额<0）合同 %d 条，均未计入待收" % len(over)

        r.case("TC-M08-11", "实缴>应缴不计入待收", "差额为负不计入", t11, priority="P2")

        def t12():
            out = []
            for role in ctx:
                j = pend(role)
                vis = C.visible_depts_of(L, ctx[role]["un"])
                if vis is not None:
                    bad = [x.get("department") for x in j["items"] if str(x.get("department", "")).strip() not in vis]
                    assert not bad, "%s 返回越权待收/支项：%s" % (role, bad[:3])
                out.append("%s:%d(收%d/支%d)" % (role, j["count"], j["receivable_count"], j["payable_count"]))
            return "；".join(out)

        r.case("TC-M08-12", "待收/支按数据权限过滤", "无越权项", t12)

        srv.sync_db()
    finally:
        srv.stop()

    prod_after = C.prod_fingerprint()
    meta = {
        "服务地址": B,
        "工作库": "test_data/m08_reminders.db",
        "默认提醒天数": dd,
        "生产库未改动": "是" if prod_before == prod_after else "否（指纹变化！）",
    }
    r.report("M08_提醒与待收待支测试.md", meta=meta)
    n, p, f, e = r.stat()
    print("\n[SUMMARY] M08 共 %d 例：通过 %d，失败 %d，错误 %d" % (n, p, f, e))
    sys.stdout.flush()
    os._exit(0 if (f == 0 and e == 0) else 1)


if __name__ == "__main__":
    main()
