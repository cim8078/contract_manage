# -*- coding: utf-8 -*-
"""M04 项目合同 CRUD 与数据质量测试（L2 / L4 交叉）。
覆盖：新增（必填/编号唯一/编号自动生成/所属部门注入/默认提醒天数）、详情、修改（留痕、编号唯一）、
删除（reason）、归档与归档后只读、清空提醒、分页与筛选、附件随编号变更同步；
数据质量规则 Q1~Q7（编号唯一、枚举、日期、金额、部门）。
隔离：运行时副本 + test_data 工作库；只改测试库。
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
    db = C.work_db("m04_crud.db")
    prod_before = C.prod_fingerprint()
    r = C.R("M04", "M04 项目合同 CRUD 与数据质量测试报告（L2/L4）")
    srv = C.Server(db)
    srv.start()
    B = srv.base
    L = C.RUNTIME_DB
    stamp = datetime.datetime.now().strftime("%H%M%S")
    try:
        sess_admin, info_admin = C.login("admin", base=B)
        assert sess_admin, "admin 登录失败：%s" % info_admin
        sess_pu, info_pu = C.login("章伟", base=B)
        if not sess_pu:
            pu = C.query(L, "SELECT 用户名 FROM users WHERE 角色='普通用户' LIMIT 1")
            sess_pu, info_pu = C.login(pu[0]["用户名"], base=B)
        assert sess_pu, "普通用户登录失败"
        PU_NAME = info_pu["name"]

        def first_val(col):
            rows = C.query(L, "SELECT %s v FROM project_contracts WHERE %s IS NOT NULL AND TRIM(%s)<>'' "
                              "GROUP BY %s ORDER BY COUNT(*) DESC LIMIT 1" % (col, col, col, col))
            return rows[0]["v"] if rows else ""

        ENUM_CAT = first_val("项目分类")
        ENUM_METHOD = first_val("采购方式")
        FUND = first_val("资金来源")

        def payload(**kw):
            base = {"category": ENUM_CAT, "name": "自动化测试合同-" + stamp,
                    "procurement_method": ENUM_METHOD, "amount": 12.34,
                    "main_content": "测试内容", "handler": PU_NAME, "fund_source": FUND}
            base.update(kw)
            return base

        def row_of(number):
            rows = C.query(L, "SELECT * FROM project_contracts WHERE 编号=?", (number,))
            return rows[0] if rows else None

        def new_contract(**kw):
            return sess_admin.post(B + "/api/contracts", json=payload(**kw), timeout=60)

        def t01():
            resp = new_contract(name="M04-01新增-" + stamp)
            assert resp.status_code == 201, "HTTP %s %s" % (resp.status_code, resp.text[:200])
            no = resp.json()["number"]
            row = row_of(no)
            assert row, "库内未找到新增合同 %s" % no
            assert row["所属部门"] == info_admin["department"], "所属部门=%s" % row["所属部门"]
            logs = C.query(L, "SELECT * FROM operation_logs WHERE 模块='合同台账' AND 操作类型='新增' AND 对象=?", (no,))
            assert logs, "未写入操作日志"
            return "新增 %s（id=row_%s）库内 1 行、操作日志 1 条" % (no, row["id"])

        r.case("TC-M04-01", "新增项目合同成功", "201 + 库内 1 行 + 操作日志", t01)

        def t02():
            depts = [d["d"] for d in C.query(L, "SELECT 单位名称||'-'||科室名称 d FROM departments ORDER BY 1")]
            ruled = None
            norule = None
            for d in depts:
                jj = sess_admin.get(B + "/api/contract_number/next",
                                    params={"kind": "project", "dept": d}, timeout=30).json()
                if jj.get("number") and ruled is None:
                    ruled = (d, jj["number"])
                if (not jj.get("number")) and jj.get("error") and norule is None:
                    norule = (d, jj["error"])
            assert norule, "未找到「未配置编号规则」的部门样本"
            out = ["未配置规则部门 %s → number 空、error=%s" % (norule[0], norule[1][:40])]
            if ruled:
                un = "zz_m04_%s" % stamp
                cu = sess_admin.post(B + "/api/users", json={
                    "username": un, "name": "测试员M04", "department": ruled[0], "role": "普通用户"}, timeout=30)
                assert cu.status_code == 201, "创建测试用户失败 %s %s" % (cu.status_code, cu.text[:200])
                s2, i2 = C.login(un, base=B)
                assert s2, "测试用户登录失败"
                resp = s2.post(B + "/api/contracts", json={
                    "category": ENUM_CAT, "name": "M04-02自动编号-" + stamp, "procurement_method": ENUM_METHOD,
                    "amount": 1, "main_content": "x", "handler": "测试员M04", "fund_source": FUND}, timeout=60)
                assert resp.status_code == 201, "自动编号新增 HTTP %s %s" % (resp.status_code, resp.text[:200])
                no = resp.json()["number"]
                assert len(no) == 6 and no.isdigit(), "自动编号不是 6 位数字：%s" % no
                assert row_of(no), "自动编号合同未落库"
                out.append("有规则部门 %s → 自动编号 %s（预览 %s）" % (ruled[0], no, ruled[1]))
            else:
                out.append("未找到有编号规则的部门（跳过正例）")
            return "；".join(out)

        r.case("TC-M04-02", "编号自动生成 / 未配置规则科室报错", "6 位编号；无规则时给出 error", t02)

        def t03():
            no = "TSTDUP01"
            a = sess_admin.post(B + "/api/contracts", json=payload(number=no), timeout=60)
            assert a.status_code == 201, "首条新增失败 %s" % a.status_code
            b = sess_admin.post(B + "/api/contracts", json=payload(number=no), timeout=60)
            assert b.status_code == 400, "重复编号返回 %s（期望 400）" % b.status_code
            assert no in b.text, "错误信息未指明编号：%s" % b.text[:120]
            return "重复编号 %s → 400 %s" % (no, b.json().get("error", ""))

        r.case("TC-M04-03", "编号唯一校验（Q1）", "重复编号 400", t03)

        def t04():
            d = payload()
            d.pop("name")
            resp = sess_admin.post(B + "/api/contracts", json=d, timeout=60)
            assert resp.status_code == 400, "缺 name 返回 %s" % resp.status_code
            hits = []
            for miss in ("category", "amount", "main_content", "handler", "fund_source", "procurement_method"):
                dd = payload()
                dd.pop(miss)
                rr = sess_admin.post(B + "/api/contracts", json=dd, timeout=60)
                hits.append("%s=%s" % (miss, rr.status_code))
            bad = [h for h in hits if not h.endswith("400")]
            assert not bad, "部分必填缺失未被拒绝：%s" % bad
            return "缺 name 400；其余 6 项必填均 400（%s）" % ",".join(hits)

        r.case("TC-M04-04", "必填字段校验（9 必填）", "缺项 400", t04)

        def t05():
            resp = new_contract(name="M04-05日期-" + stamp, procurement_date="2026/13/40")
            assert resp.status_code == 400, "非法日期返回 HTTP %s（文档 Q5 要求后端校验并 400；响应体：%s）" % (
                resp.status_code, resp.text[:160])
            return "非法采购时间 → 400"

        r.case("TC-M04-05", "日期格式校验（Q5）", "400", t05, priority="P1")

        def t06():
            resp = new_contract(name="M04-06金额-" + stamp, amount="约三万")
            stored = ""
            if resp.status_code == 201:
                row = row_of(resp.json()["number"])
                stored = "；实际落库值=%r（列类型 REAL，文本被原样接受）" % row["合同金额（万元）"]
            assert resp.status_code == 400, "金额「约三万」返回 HTTP %s（文档 Q6 要求后端转数值、非法报错）%s" % (
                resp.status_code, stored)
            return "金额非数字 → 400"

        r.case("TC-M04-06", "金额数字校验（Q6）", "400", t06, priority="P1")

        def t07():
            resp = new_contract(name="M04-07枚举-" + stamp, category="不存在的分类XYZ",
                                procurement_method="不存在的方式XYZ", is_settled="半结算")
            stored = ""
            if resp.status_code == 201:
                row = row_of(resp.json()["number"])
                stored = "；实际落库 项目分类=%r 采购方式=%r 是否结算=%r" % (
                    row["项目分类"], row["采购方式"], row["是否结算"])
            assert resp.status_code == 400, "非法枚举返回 HTTP %s（文档 Q2/Q3/Q4 要求后端校验）%s" % (
                resp.status_code, stored)
            return "非法枚举 → 400"

        r.case("TC-M04-07", "枚举取值校验（Q2/Q3/Q4）", "400", t07, priority="P1")

        def t08():
            other = C.query(L, "SELECT 单位名称||'-'||科室名称 d FROM departments WHERE "
                               "单位名称||'-'||科室名称 <> ? LIMIT 1", (info_pu["department"],))[0]["d"]
            resp = sess_pu.post(B + "/api/contracts", json=payload(
                number="TSTM08" + stamp[:2], name="M04-08部门-" + stamp,
                department=other, dept=other, owner_dept=other), timeout=60)
            assert resp.status_code == 201, "HTTP %s %s" % (resp.status_code, resp.text[:160])
            row = row_of(resp.json()["number"])
            assert row["所属部门"] == info_pu["department"], \
                "所属部门被写成 %s（应强制为本人部门 %s）" % (row["所属部门"], info_pu["department"])
            return "普通用户 body 指定他部门被忽略，落库所属部门=%s" % row["所属部门"]

        r.case("TC-M04-08", "所属部门自动注入/防越权指定（Q7）", "落库为本部门", t08)

        def t09():
            dfault = C.query(L, "SELECT 配置值 v FROM settings WHERE 配置项='default_remind_days'")
            exp = dfault[0]["v"] if dfault else "60"
            resp = new_contract(name="M04-09提醒-" + stamp)
            row = row_of(resp.json()["number"])
            got = row["提醒提前天数"]
            assert got is not None and str(got).strip() != "", "提醒提前天数未写默认值（=%r）" % got
            assert int(float(got)) == int(float(exp)), "默认提醒天数 %s ≠ 全局 %s" % (got, exp)
            return "空提醒天数 → 落库 %s（全局默认 %s）" % (got, exp)

        r.case("TC-M04-09", "提醒提前天数默认值", "取全局默认值", t09, priority="P1")

        detail_holder = {}

        def t10():
            resp = new_contract(name="M04-10详情-" + stamp, next_bid_date="2026-12-31", remark="详情样例")
            no = resp.json()["number"]
            row = row_of(no)
            detail_holder["id"] = row["id"]
            d = sess_admin.get(B + "/api/contracts/row_%s" % row["id"], timeout=30)
            assert d.status_code == 200, "详情 HTTP %s" % d.status_code
            j = d.json()
            for k in ("id", "number", "name", "category", "amount", "department", "attachments"):
                assert k in j, "详情缺字段 %s" % k
            assert j["number"] == no and j["remark"] == "详情样例", "详情字段值不符"
            return "详情 200，含 %d 个字段（含 attachments）" % len(j)

        r.case("TC-M04-10", "合同详情", "200 + 字段齐全", t10)

        def t11():
            d = sess_admin.get(B + "/api/contracts/row_99999999", timeout=30)
            assert d.status_code == 404, "不存在详情返回 %s" % d.status_code
            bad = sess_admin.get(B + "/api/contracts/abc", timeout=30)
            assert bad.status_code == 400, "非法 id 返回 %s（期望 400）" % bad.status_code
            return "不存在 → 404；非法 id 格式 → 400"

        r.case("TC-M04-11", "详情异常路径", "404 / 400", t11, priority="P1")

        def t12():
            resp = new_contract(name="M04-12修改-" + stamp, winner="原单位A")
            no = resp.json()["number"]
            row = row_of(no)
            up = sess_admin.put(B + "/api/contracts/row_%s" % row["id"], json={"winner": "新单位B"}, timeout=60)
            assert up.status_code == 200, "修改 HTTP %s %s" % (up.status_code, up.text[:160])
            assert row_of(no)["中标单位"] == "新单位B", "库内未更新"
            fl = C.query(L, "SELECT * FROM field_change_logs WHERE 对象编号=? AND 字段='中标单位'", (no,))
            assert fl, "未写字段级变更日志"
            f0 = fl[-1]
            assert str(f0["原值"]) == "原单位A" and str(f0["新值"]) == "新单位B", \
                "字段日志原值/新值不符：%s→%s" % (f0["原值"], f0["新值"])
            return "PUT 200；field_change_logs 中标单位 原值A→新值B"

        r.case("TC-M04-12", "修改成功并写字段级变更日志（Q17）", "200 + 前后值留痕", t12)

        def t13():
            a = new_contract(number="TSTM13A", name="M04-13a-" + stamp)
            b = new_contract(number="TSTM13B", name="M04-13b-" + stamp)
            assert a.status_code == 201 and b.status_code == 201, "样本创建失败"
            rb = row_of(b.json()["number"])
            up = sess_admin.put(B + "/api/contracts/row_%s" % rb["id"],
                                json={"number": a.json()["number"]}, timeout=60)
            assert up.status_code == 400, "改重编号返回 %s（期望 400）" % up.status_code
            assert "不允许修改" in up.json().get("error", ""), "提示不符：%s" % up.text[:120]
            assert row_of(b.json()["number"]), "原编号的记录被错误改动"
            return "改重编号 → 400「%s」（编号不可修改规则先于唯一性校验）" % up.json().get("error", "")

        r.case("TC-M04-13", "修改编号被拒（含改重场景）", "400 且原编号不变", t13)

        def t14():
            resp = new_contract(name="M04-14删除-" + stamp)
            no = resp.json()["number"]
            row = row_of(no)
            dl = sess_admin.delete(B + "/api/contracts/row_%s" % row["id"], json={}, timeout=30)
            assert dl.status_code == 400, "无 reason 删除返回 %s（期望 400）" % dl.status_code
            dl2 = sess_admin.delete(B + "/api/contracts/row_%s" % row["id"], json={"reason": "  "}, timeout=30)
            assert dl2.status_code == 400, "空白 reason 返回 %s" % dl2.status_code
            assert row_of(no), "合同被误删"
            return "无/空白 reason → 400，合同未删除"

        r.case("TC-M04-14", "删除必须填写原因", "400", t14)

        def t15():
            resp = new_contract(name="M04-15删除-" + stamp)
            no = resp.json()["number"]
            row = row_of(no)
            dl = sess_admin.delete(B + "/api/contracts/row_%s" % row["id"],
                                   json={"reason": "自动化测试清理"}, timeout=30)
            assert dl.status_code == 200, "删除 HTTP %s %s" % (dl.status_code, dl.text[:160])
            assert row_of(no) is None, "库内仍存在"
            lg = C.query(L, "SELECT * FROM operation_logs WHERE 模块='合同台账' AND 操作类型='删除' AND 对象=?", (no,))
            assert lg and "自动化测试清理" in str(lg[-1]["详情"]), "删除日志缺失原因（日志：%s）" % (lg[-1] if lg else None)
            return "带 reason 删除 200，行已删除且日志含原因"

        r.case("TC-M04-15", "删除成功并留痕", "200 + 行删除 + 日志", t15)

        archived = {}

        def t16():
            resp = new_contract(name="M04-16归档-" + stamp, next_bid_date="2026-11-30")
            no = resp.json()["number"]
            row = row_of(no)
            archived["id"] = row["id"]
            archived["no"] = no
            ar = sess_admin.put(B + "/api/contracts/row_%s/archive" % row["id"], json={}, timeout=30)
            assert ar.status_code == 200, "归档 HTTP %s %s" % (ar.status_code, ar.text[:160])
            row = row_of(no)
            assert str(row["是否归档"]) == "是", "是否归档未置是"
            assert str(row["归档人"]).strip(), "归档人为空"
            assert str(row["归档时间"]).strip(), "归档时间为空"
            up = sess_admin.put(B + "/api/contracts/row_%s" % row["id"], json={"remark": "x"}, timeout=30)
            dl = sess_admin.delete(B + "/api/contracts/row_%s" % row["id"], json={"reason": "x"}, timeout=30)
            assert up.status_code == 403, "归档后 PUT 返回 %s" % up.status_code
            assert dl.status_code == 403, "归档后 DELETE 返回 %s" % dl.status_code
            return "归档 200（归档人=%s），归档后 PUT/DELETE 均 403" % row["归档人"]

        r.case("TC-M04-16", "归档写入与归档后只读", "200 + 三字段落库 + 只读 403", t16)

        def t17():
            assert archived, "缺少归档样本"
            fl = C.query(L, "SELECT * FROM field_change_logs WHERE 对象编号=? AND 字段='是否归档'", (archived["no"],))
            lg = C.query(L, "SELECT * FROM operation_logs WHERE 模块='合同台账' AND 操作类型='归档' AND 对象=?",
                         (archived["no"],))
            assert fl and lg, "归档留痕缺失（字段日志=%d，操作日志=%d）" % (len(fl), len(lg))
            return "归档字段日志 %d 条、操作日志 %d 条" % (len(fl), len(lg))

        r.case("TC-M04-17", "归档留痕", "字段日志 + 操作日志", t17, priority="P1")

        def t18():
            resp = new_contract(name="M04-18清提醒-" + stamp, next_bid_date="2026-12-01")
            no = resp.json()["number"]
            row = row_of(no)
            assert str(row["下次招采时间"]).strip(), "样本未写入下次招采时间"
            rc = sess_admin.put(B + "/api/contracts/row_%s/clear_remind" % row["id"], json={}, timeout=30)
            assert rc.status_code == 200, "clear_remind HTTP %s %s" % (rc.status_code, rc.text[:160])
            row = row_of(no)
            assert not str(row["下次招采时间"] or "").strip(), "下次招采时间未清空（=%r）" % row["下次招采时间"]
            fl = C.query(L, "SELECT * FROM field_change_logs WHERE 对象编号=? AND 字段='下次招采时间'", (no,))
            assert fl, "清空提醒未写字段级日志"
            return "清空提醒 200，下次招采时间置空，字段日志 %d 条" % len(fl)

        r.case("TC-M04-18", "清空提醒（永久关闭）", "200 + 置空 + 留痕", t18, priority="P1")

        def t19():
            base = sess_admin.get(B + "/api/contracts", params={"page": 1, "page_size": 5}, timeout=60).json()
            assert len(base["items"]) <= 5, "page_size 未生效"
            nums = []
            for it in base["items"]:
                try:
                    nums.append(float(it["number"]))
                except Exception:
                    pass
            assert nums == sorted(nums, reverse=True), "首页未按编号降序：%s" % nums
            p2 = sess_admin.get(B + "/api/contracts", params={"page": 2, "page_size": 5}, timeout=60).json()
            assert p2["page"] == 2 and p2["total"] == base["total"], "分页 total 不稳定"
            rep = sess_admin.get(B + "/api/contracts", params={"category": ENUM_CAT, "page_size": 200}, timeout=60).json()
            exp = C.query(L, "SELECT COUNT(*) c FROM project_contracts WHERE 项目分类=?", (ENUM_CAT,))[0]["c"]
            assert rep["total"] == exp, "分类筛选 total=%s，期望 %s" % (rep["total"], exp)
            assert all(it["category"] == ENUM_CAT for it in rep["items"]), "分类筛选结果含非目标分类"
            dept_sub = str(info_admin["department"]).split("-")[-1]
            rep2 = sess_admin.get(B + "/api/contracts", params={"dept": dept_sub, "page_size": 1}, timeout=60).json()
            exp2 = C.query(L, "SELECT COUNT(*) c FROM project_contracts WHERE 所属部门 LIKE ?",
                           ("%" + dept_sub + "%",))[0]["c"]
            assert rep2["total"] == exp2, "部门筛选 total=%s，期望 %s" % (rep2["total"], exp2)
            return "分页/降序/分类筛选(%s=%d)/部门筛选(%s=%d) 均与 SQL 一致" % (ENUM_CAT, exp, dept_sub, exp2)

        r.case("TC-M04-19", "分页、默认排序与筛选", "与 SQL 期望一致", t19)

        def t20():
            # 1) 有附件的合同改编号 → 400 拒绝，编号与附件均不变
            att = C.query(L, "SELECT a.附件ID i, a.合同编号 n FROM project_attachments a "
                             "WHERE a.合同编号 IS NOT NULL AND TRIM(a.合同编号)<>'' "
                             "AND EXISTS (SELECT 1 FROM project_contracts c WHERE c.编号=a.合同编号 "
                             "AND COALESCE(c.是否归档,'')<>'是') LIMIT 1")
            assert att, "库内无「附件+对应合同」样本"
            old_no = str(att[0]["n"])
            row = C.query(L, "SELECT id FROM project_contracts WHERE 编号=?", (old_no,))[0]
            r1 = sess_admin.put(B + "/api/contracts/row_%s" % row["id"],
                                json={"number": old_no + "Z"}, timeout=60)
            assert r1.status_code == 400, "有附件合同改号返回 %s（期望 400）" % r1.status_code
            msg = r1.json().get("error", "")
            assert "不允许修改" in msg, "错误提示不符：%s" % msg
            assert row_of(old_no), "原编号记录丢失"
            a1 = C.query(L, "SELECT 合同编号 FROM project_attachments WHERE 附件ID=?", (att[0]["i"],))[0]
            assert a1["合同编号"] == old_no, "附件合同编号被改动：%s" % a1["合同编号"]
            # 2) 无附件的合同改编号 → 同样 400
            resp = new_contract(name="M04-20无附件-" + stamp)
            no2 = resp.json()["number"]
            r2 = sess_admin.put(B + "/api/contracts/row_%s" % row_of(no2)["id"],
                                json={"number": no2 + "Y"}, timeout=60)
            assert r2.status_code == 400, "无附件合同改号返回 %s（期望 400）" % r2.status_code
            # 3) 提交与现编号相同 → 200，且不产生编号变更日志
            before = C.query(L, "SELECT COUNT(*) c FROM field_change_logs WHERE 对象编号=? AND 字段='编号'",
                             (no2,))[0]["c"]
            r3 = sess_admin.put(B + "/api/contracts/row_%s" % row_of(no2)["id"],
                                json={"number": no2, "remark": "同编号提交"}, timeout=60)
            after = C.query(L, "SELECT COUNT(*) c FROM field_change_logs WHERE 对象编号=? AND 字段='编号'",
                            (no2,))[0]["c"]
            assert r3.status_code == 200, "同编号提交返回 %s" % r3.status_code
            assert after == before, "同编号提交产生了编号变更日志"
            return "改号（有/无附件）均 400「%s」；编号与附件不变；同编号提交 200 无留痕" % msg

        r.case("TC-M04-20", "合同编号不可修改（业务规则）", "改号 400、编号与附件不变", t20)

        srv.sync_db()
    finally:
        srv.stop()

    prod_after = C.prod_fingerprint()
    meta = {
        "服务地址": B,
        "工作库": "test_data/m04_crud.db",
        "枚举样本": "项目分类=%s；采购方式=%s；资金来源=%s" % (ENUM_CAT, ENUM_METHOD, FUND),
        "生产库未改动": "是" if prod_before == prod_after else "否（指纹变化！）",
    }
    r.report("M04_项目合同CRUD测试.md", meta=meta)
    n, p, f, e = r.stat()
    print("\n[SUMMARY] M04 共 %d 例：通过 %d，失败 %d，错误 %d" % (n, p, f, e))
    sys.stdout.flush()
    os._exit(0 if (f == 0 and e == 0) else 1)


if __name__ == "__main__":
    main()
