# -*- coding: utf-8 -*-
"""M17 数据质量与约束复核（L4）。
覆盖《数据字典表》Q1~Q17 的可执行部分 + 三处业务待确认项的行为：
编号唯一/非空、枚举（含「部分结算」）、日期格式（含宽松分隔符规范化）、金额数字、
部门外键、经办人白名单（提示但不强制）、勾稽 Q9/Q10、提醒源列数据质量、Q12 是/否枚举、
以及导入侧三级判定（ok/warn/error）与历史表头归一化。
"""
import os
import sys
import datetime
from io import BytesIO

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _common as C  # noqa: E402
import _oracle as O  # noqa: E402
import openpyxl  # noqa: E402


def main():
    C.setup_stdout()
    C.ensure_dirs()
    C.prepare_runtime()
    C.prepare_base()
    db = C.work_db("m17_dq.db")
    prod_before = C.prod_fingerprint()
    r = C.R("M17", "M17 数据质量与约束复核报告（L4）")
    srv = C.Server(db)
    srv.start()
    B = srv.base
    L = C.RUNTIME_DB
    stamp = datetime.datetime.now().strftime("%H%M%S")
    try:
        A, ia = C.login("admin", base=B)
        assert A, "admin 登录失败"
        dept = ia["department"]
        base_body = {"category": "服务类", "procurement_method": "自行采购", "fund_source": "其他",
                     "amount": 1, "main_content": "数据质量测试", "handler": ia["name"]}

        def post(**kw):
            body = dict(base_body)
            body["name"] = "M17测试-" + stamp
            body.update(kw)
            return A.post(B + "/api/contracts", json=body, timeout=60)

        def t01():
            r1 = post(number="TSTQ1" + stamp[:2])
            assert r1.status_code == 201, "正常新增失败 %s" % r1.status_code
            no = r1.json()["number"]
            r2 = post(number=no)
            assert r2.status_code == 400, "重复编号未被拒绝：%s" % r2.status_code
            r3 = post(number="")
            assert r3.status_code in (201, 400), "空编号行为异常 %s" % r3.status_code
            idx = C.query(L, "SELECT sql FROM sqlite_master WHERE type='index' AND tbl_name='project_contracts' "
                             "AND sql LIKE '%UNIQUE%'")
            assert idx, "缺少编号唯一索引"
            return "重复编号 400；空编号 HTTP %s（自动生成）；唯一索引存在" % r3.status_code

        r.case("TC-M17-01", "Q1 编号非空且唯一", "重复拒绝 + 唯一索引兜底", t01)

        def t02():
            bad = []
            for field, val in (("category", "不存在的分类"), ("procurement_method", "不存在的采购方式"),
                               ("fund_source", "不存在的资金来源")):
                rr = post(**{field: val})
                if rr.status_code != 400:
                    bad.append("%s=%s" % (field, rr.status_code))
            semi = post(is_settled="部分结算")
            other = post(is_settled="半结算")
            assert not bad, "非法枚举未被拒绝：%s" % bad
            assert semi.status_code == 201, "「部分结算」被拒绝（Q4 允许）：%s" % semi.status_code
            assert other.status_code == 400, "非法结算值未被拒绝：%s" % other.status_code
            return "分类/采购方式/资金来源 非法值均 400；是否结算：部分结算 201、半结算 400"

        r.case("TC-M17-02", "Q2/Q3/Q4 枚举取值", "非法拒绝、部分结算放行", t02)

        def t03():
            bad = post(procurement_date="2026-13-40")
            assert bad.status_code == 400, "非法日期 %s" % bad.status_code
            loose = post(procurement_date="2026/1/9")
            assert loose.status_code == 201, "宽松日期 %s" % loose.status_code
            row = C.query(L, "SELECT 采购时间 v FROM project_contracts WHERE 编号=?", (loose.json()["number"],))[0]
            assert row["v"] == "2026-01-09", "未规范化为 yyyy-mm-dd：%r" % row["v"]
            empty = post(procurement_date="")
            assert empty.status_code == 201, "空日期应允许：%s" % empty.status_code
            return "2026-13-40 → 400；2026/1/9 → 201 且落库 2026-01-09；空值放行"

        r.case("TC-M17-03", "Q5 日期格式（校验+规范化）", "非法 400、宽松值规范化", t03)

        def t04():
            bad = post(amount="约三万")
            assert bad.status_code == 400, "文本金额未被拒绝：%s" % bad.status_code
            neg = post(amount=-1)
            assert neg.status_code == 400, "负金额未被拒绝：%s" % neg.status_code
            n_text, n_neg = 0, 0
            for c in C.query(L, "SELECT 合同金额（万元） v, 结算金额 s, 已支付 p, 未支付 u FROM project_contracts"):
                for k in ("v", "s", "p", "u"):
                    v = c.get(k)
                    if v in (None, ""):
                        continue
                    try:
                        f = float(v)
                        if f < 0:
                            n_neg += 1
                    except (TypeError, ValueError):
                        n_text += 1
            return "新写入文本/负金额均 400；全库现存「非数字金额」字段值 %d 处、负值 %d 处（待业务确认清单）" % (n_text, n_neg)

        r.case("TC-M17-04", "Q6 金额字段为数字（含存量盘点）", "接口拒绝文本/负值", t04)

        def t05():
            r1 = post()
            assert r1.status_code == 201, "新增失败 %s" % r1.status_code
            row = C.query(L, "SELECT 所属部门 d, 所属单位 u, 所属科室 k FROM project_contracts WHERE 编号=?",
                          (r1.json()["number"],))[0]
            assert row["d"] == dept, "所属部门未注入：%s" % row["d"]
            assert row["u"] and row["k"], "所属单位/科室未由触发器派生：%s/%s" % (row["u"], row["k"])
            depts = [x["d"] for x in C.query(L, "SELECT 单位名称||'-'||科室名称 d FROM departments")]
            assert row["d"] in depts, "所属部门不在部门列表"
            return "所属部门自动注入=%s，触发器派生 所属单位/科室=%s/%s" % (row["d"], row["u"], row["k"])

        r.case("TC-M17-05", "Q7 所属部门∈部门列表（外键+触发器）", "注入且派生正确", t05)

        def t06():
            ghost = "不存在的人XYZ"
            rr = post(handler=ghost)
            users = [x["n"] for x in C.query(L, "SELECT DISTINCT 姓名 n FROM users")]
            n_ghost = len([c for c in C.query(L, "SELECT 经办人 h FROM project_contracts")
                           if str(c["h"] or "").strip() and str(c["h"]).strip() not in users])
            body = ""
            if rr.status_code == 201:
                body = str(rr.json())
            return "经办人「%s」写入 → HTTP %s（文档 Q8：中级别，提示但不强制）；返回体提示=%s；全库经办人不在用户表的合同 %d 条" % (
                ghost, rr.status_code, ("无 warning 字段" if "warn" not in body.lower() else "含提示"), n_ghost)

        r.case("TC-M17-06", "Q8 经办人∈用户主数据（提示不强制）", "记录实际行为", t06, priority="P1")

        def t07():
            rows = C.query(L, "SELECT 编号 n, 合同金额（万元） a, 结算金额 s, 已支付 p, 未支付 u FROM project_contracts")
            q9 = q10 = 0
            for c in rows:
                a, s, p, u = O.num(c["a"]), O.num(c["s"]), O.num(c["p"]), O.num(c["u"])
                if s is not None and p is not None and u is not None and abs(p + u - s) > 0.01:
                    q9 += 1
                if s is not None and a is not None and s > a * 1.01:
                    q10 += 1
            return "Q9（已支付+未支付≠结算金额，容差0.01）违背 %d 条；Q10（结算金额>合同金额，容差1%%）违背 %d 条" % (q9, q10)

        r.case("TC-M17-07", "Q9/Q10 金额勾稽（月审计提示项）", "输出违背条数", t07, priority="P1")

        def t08():
            bad_date = empty_rd = 0
            for c in C.query(L, "SELECT 下次招采时间 d, 提醒提前天数 r FROM project_contracts"):
                if str(c["d"] or "").strip() and O.parse_date(c["d"]) is None:
                    bad_date += 1
                if not str(c["r"] or "").strip():
                    empty_rd += 1
            assets = C.query(L, "SELECT 下次出租时间 d FROM asset_contracts")
            bad_asset = len([a for a in assets if str(a["d"] or "").strip() and O.parse_date(a["d"]) is None])
            return "下次招采时间非空但无法解析（永不提醒）：项目 %d 条 / 资产 %d 条；提醒提前天数为空的项目 %d 条（回退全局默认）" % (
                bad_date, bad_asset, empty_rd)

        r.case("TC-M17-08", "Q11/Q16 提醒源列的数据质量", "输出不可解析条数", t08, priority="P2")

        def t09():
            bad = post(has_change="不确定")
            ok = post(has_change="否")
            assert bad.status_code == 400, "非法「是否有签证或变更」未被拒绝：%s" % bad.status_code
            assert ok.status_code == 201, "合法值被拒绝：%s" % ok.status_code
            n_bad = len([x for x in C.query(L, "SELECT 是否有签证或变更 v FROM project_contracts")
                         if str(x["v"] or "").strip() not in ("", "是", "否")])
            return "「不确定」→400、「否」→201；全库该列非 是/否/空 的行数=%d" % n_bad

        r.case("TC-M17-09", "Q12 是否有签证或变更∈{是,否}", "非法 400", t09, priority="P1")

        def t10():
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "合同台账"
            headers = ["编号", "项目分类", "项目名称", "采购方式", "合同金额（万元）", "资金来源",
                       "合同主要内容", "经办人/分管领导", "所属部门", "端工验收时间", "甲方现场负责人",
                       "是否有签证或变更", "是否结算", "提醒提前天数"]
            ws.append(headers)

            def row(no, name, amount, settled, accept="2026-05-05", handler="张三"):
                ws.append([no, "服务类", name, "自行采购", amount, "其他", "x", handler, dept,
                           accept, "李四", "否", settled, 60])

            row("IMPTEST01", "导入-正常行", 1.5, "部分结算")
            row("IMPTEST02", "导入-文本金额", "约三万", "否")
            row("IMPTEST03", "导入-非法日期", 2, "否", accept="2026-13-40")
            row("IMPTEST04", "导入-缺经办人", 2, "否", handler="")
            bio = BytesIO()
            wb.save(bio)
            files = {"file": ("dq_test.xlsx", bio.getvalue(),
                              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
            resp = A.post(B + "/api/contracts/import/preview", files=files, timeout=60)
            assert resp.status_code == 200, "预览 HTTP %s %s" % (resp.status_code, resp.text[:160])
            j = resp.json()
            items = {str(x["number"]): x for x in j["items"]}
            assert set(items) == {"IMPTEST01", "IMPTEST02", "IMPTEST03", "IMPTEST04"}, "行解析异常：%s" % list(items)
            r1 = items["IMPTEST01"]
            assert r1["status"] in ("ok", "warn"), "正常行状态=%s" % r1["status"]
            assert not any("经办人为空" in x for x in r1["reasons"]), \
                "历史表头「经办人/分管领导」未归一化到经办人：%s" % r1["reasons"]
            assert any("是否结算" in x for x in r1["reasons"]), "「部分结算」应给出警告：%s" % r1["reasons"]
            assert items["IMPTEST02"]["status"] == "error" and \
                any("合同金额（万元）" in x and "数字" in x for x in items["IMPTEST02"]["reasons"]), \
                "文本金额未判为错误：%s" % items["IMPTEST02"]
            assert items["IMPTEST03"]["status"] == "error" and \
                any("竣工验收时间" in x for x in items["IMPTEST03"]["reasons"]), \
                "历史表头「端工验收时间」未归一化或未校验：%s" % items["IMPTEST03"]
            assert items["IMPTEST04"]["status"] == "error" and \
                any("经办人" in x for x in items["IMPTEST04"]["reasons"]), \
                "缺经办人未判错：%s" % items["IMPTEST04"]
            return "preview：总计 %d（ok/warn/error=%d/%d/%d）；表头归一化（端工验收时间/经办人-分管领导）生效；" \
                   "文本金额=error、「部分结算」=warn、缺经办人=error" % (
                       j["total"], j["ok_count"], j["warn_count"], j["error_count"])

        r.case("TC-M17-10", "导入三级判定与历史表头归一化", "ok/warn/error 正确、归一化生效", t10, priority="P1")

        def t11():
            # 历史脏数据保护：找一条「合同金额为非数字文本」的存量合同，
            # 仅改其它字段必须成功（否则新校验会锁死 182 处历史数据）
            target = None
            for c in C.query(L, "SELECT id,编号,合同金额（万元） v FROM project_contracts"):
                v = c["v"]
                if v in (None, ""):
                    continue
                try:
                    float(v)
                except (TypeError, ValueError):
                    target = c
                    break
            assert target, "库内无非数字金额样本（跳过）"
            rid = "row_%s" % target["id"]
            r1 = A.put(B + "/api/contracts/" + rid, json={"remark": "仅改备注"}, timeout=60)
            assert r1.status_code == 200, "仅改其它字段被拦截：HTTP %s %s" % (r1.status_code, r1.text[:120])
            r2 = A.put(B + "/api/contracts/" + rid, json={"amount": target["v"]}, timeout=60)
            assert r2.status_code == 200, "原值回填被拦截（未变更字段应跳过校验）：%s" % r2.status_code
            r3 = A.put(B + "/api/contracts/" + rid, json={"amount": "另一个文本金额"}, timeout=60)
            assert r3.status_code == 400, "改成新的文本金额未被拒绝：%s" % r3.status_code
            return "样本 %s（金额=%r）：仅改备注 200；原值回填 200；改成新文本 400 → 未锁死历史数据" % (
                target["编号"], target["v"])

        r.case("TC-M17-11", "历史脏数据不被新校验锁死", "未变更字段跳过校验", t11, priority="P1")

        srv.sync_db()
    finally:
        srv.stop()

    prod_after = C.prod_fingerprint()
    meta = {
        "服务地址": B,
        "工作库": "test_data/m17_dq.db",
        "生产库未改动": "是" if prod_before == prod_after else "否（指纹变化！）",
    }
    r.report("M17_数据质量与约束复核.md", meta=meta)
    n, p, f, e = r.stat()
    print("\n[SUMMARY] M17 共 %d 例：通过 %d，失败 %d，错误 %d" % (n, p, f, e))
    sys.stdout.flush()
    os._exit(0 if (f == 0 and e == 0) else 1)


if __name__ == "__main__":
    main()
