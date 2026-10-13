# -*- coding: utf-8 -*-
"""M12 打印为 PDF 测试（L5）。
打印由前端在新标签页生成（window.print），无服务端 PDF 渲染，故本轮验证：
主页/静态资源可达、打印所依赖的数据接口字段齐全、打印视图含打印人/时间水印、资产打印附租金明细、版本号一致性。
"""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _common as C  # noqa: E402

INDEX = os.path.join(C.PROJECT_ROOT, "templates", "index.html")
DEV_DOC = os.path.join(C.PROJECT_ROOT, "docs", "系统开发文档.md")
PRINT_FIELDS = ["number", "name", "category", "procurement_method", "amount", "fund_source",
                "main_content", "procurement_date", "start_date", "end_date", "acceptance_date",
                "site_manager", "handler", "department", "is_settled", "settled_amount",
                "paid_amount", "unpaid_amount", "term", "remark"]


def main():
    C.setup_stdout()
    C.ensure_dirs()
    C.prepare_runtime()
    C.prepare_base()
    db = C.work_db("m12_print.db")
    prod_before = C.prod_fingerprint()
    r = C.R("M12", "M12 打印为 PDF 测试报告（L5）")
    srv = C.Server(db)
    srv.start()
    B = srv.base
    L = C.RUNTIME_DB
    try:
        A, ia = C.login("admin", base=B)
        assert A, "admin 登录失败"
        html = open(INDEX, encoding="utf-8").read()

        def t01():
            resp = A.get(B + "/", timeout=30)
            assert resp.status_code == 200, "主页 HTTP %s" % resp.status_code
            body = resp.text
            assert "潭合同" in body, "主页标题缺失"
            assert ("vue.global" in body) or ("createApp" in body) or ("id=\"app\"" in body), "未找到 Vue 挂载点"
            return "GET / → 200，含 Vue 挂载与本地化前端资源引用（%d 字节）" % len(body)

        r.case("TC-M12-01", "主页可达", "200 + 前端挂载点", t01)

        def t02():
            assets = ["/static/vendor/js/vue.global.prod.js", "/static/vendor/js/xlsx.full.min.js",
                      "/static/vendor/js/chart.umd.min.js", "/static/vendor/css/bootstrap.min.css",
                      "/static/logo.jpg"]
            bad = []
            for a in assets:
                rr = A.get(B + a, timeout=30)
                if rr.status_code != 200:
                    bad.append("%s=%s" % (a, rr.status_code))
            assert not bad, "静态资源不可达：%s" % bad
            return "5 个本地化静态资源（含打印页引用的 bootstrap/chart）均可访问"

        r.case("TC-M12-02", "前端静态资源可达（离线可用）", "全部 200", t02)

        def t03():
            c = C.query(L, "SELECT id FROM project_contracts ORDER BY id LIMIT 1")[0]
            d = A.get(B + "/api/contracts/row_%s" % c["id"], timeout=30)
            assert d.status_code == 200, "项目详情 HTTP %s" % d.status_code
            j = d.json()
            miss = [k for k in PRINT_FIELDS if k not in j]
            assert not miss, "打印所需字段缺失：%s" % miss
            a = C.query(L, "SELECT id FROM asset_contracts ORDER BY id LIMIT 1")[0]
            da = A.get(B + "/api/asset_contracts/arow_%s" % a["id"], timeout=30)
            assert da.status_code == 200, "资产详情 HTTP %s" % da.status_code
            ja = da.json()
            for k in ("attachments", "rent_collections"):
                assert k in ja, "资产打印所需 %s 缺失" % k
            return "项目详情含打印所需 %d 个全要素字段；资产详情含 attachments 与 rent_collections" % len(PRINT_FIELDS)

        r.case("TC-M12-03", "打印数据源字段齐全", "项目全要素 + 资产附件/租金", t03)

        def t04():
            m = re.search(r"function printContractPdf\(\)(.*?)function printAssetPdf", html, re.S)
            assert m, "未找到 printContractPdf 实现"
            body = m.group(1)
            assert "window.print()" in body, "打印视图未调用 window.print()"
            assert "打印人：" in body, "缺少打印人水印"
            assert "打印时间：" in body, "缺少打印时间水印"
            return "printContractPdf：含 window.print()、水印「打印人：」「打印时间：」"

        r.case("TC-M12-04", "打印视图调用与打印人/时间水印", "水印存在", t04)

        def t05():
            m = re.search(r"function printAssetPdf\(\)(.*?)\n\s*function ", html, re.S)
            assert m, "未找到 printAssetPdf 实现"
            body = m.group(1)
            assert "租金" in body, "资产打印视图未包含租金明细区域"
            assert "打印人：" in body and "打印时间：" in body, "资产打印缺少水印"
            return "printAssetPdf：含租金收取明细区域与打印人/时间水印"

        r.case("TC-M12-05", "资产打印附租金明细表", "含租金区域 + 水印", t05, priority="P1")

        def t06():
            m = re.search(r"系统版本\s*V(\d+(?:\.\d+)+)", html)
            page_ver = m.group(1) if m else "(未找到)"
            doc = open(DEV_DOC, encoding="utf-8").read()
            md = re.search(r"\*\*V(\d+\.\d+(?:\.\d+)?)（当前）\*\*", doc)
            doc_ver = md.group(1) if md else "(未找到)"
            assert page_ver == doc_ver, \
                "登录页版本 V%s 与《系统开发文档》当前版本 V%s 不一致（D-1）" % (page_ver, doc_ver)
            return "登录页版本 V%s = 文档当前版本 V%s" % (page_ver, doc_ver)

        r.case("TC-M12-06", "版本号一致性（登录页 vs 文档）", "两者一致", t06, priority="P2")

    finally:
        srv.stop()

    prod_after = C.prod_fingerprint()
    meta = {
        "服务地址": B,
        "工作库": "test_data/m12_print.db",
        "生产方式": "前端新标签页 + window.print()（无服务端 PDF 渲染）",
        "生产库未改动": "是" if prod_before == prod_after else "否（指纹变化！）",
    }
    r.report("M12_打印PDF测试.md", meta=meta)
    n, p, f, e = r.stat()
    print("\n[SUMMARY] M12 共 %d 例：通过 %d，失败 %d，错误 %d" % (n, p, f, e))
    sys.stdout.flush()
    os._exit(0 if (f == 0 and e == 0) else 1)


if __name__ == "__main__":
    main()
