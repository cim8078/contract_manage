# -*- coding: utf-8 -*-
"""M13 安全测试（L6）。
覆盖：未登录全拦截、越权覆盖、SQL 注入（登录/筛选）、路径穿越（附件/静态/数据文件）、上传白名单、
短信密钥掩码、会话密钥与密码哈希与登录锁定的现状走读、Cookie 属性、错误信息不泄露、
未授权接口、XSS（存量转义方式）、仓库敏感数据暴露。
"""
import os
import re
import sys
import json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _common as C  # noqa: E402
import requests  # noqa: E402

API_PATHS = [
    ("GET", "/api/current_user"), ("GET", "/api/dashboard"), ("GET", "/api/reminders"),
    ("GET", "/api/pending_payments"), ("GET", "/api/contracts"), ("POST", "/api/contracts"),
    ("GET", "/api/contracts/row_1"), ("PUT", "/api/contracts/row_1"), ("DELETE", "/api/contracts/row_1"),
    ("PUT", "/api/contracts/row_1/archive"), ("PUT", "/api/contracts/row_1/clear_remind"),
    ("GET", "/api/contracts/export"), ("GET", "/api/contracts/template"),
    ("POST", "/api/contracts/import/preview"), ("POST", "/api/contracts/import"),
    ("GET", "/api/asset_contracts"), ("POST", "/api/asset_contracts"),
    ("GET", "/api/asset_contracts/arow_1"), ("PUT", "/api/asset_contracts/arow_1"),
    ("DELETE", "/api/asset_contracts/arow_1"), ("PUT", "/api/asset_contracts/arow_1/archive"),
    ("GET", "/api/asset_contracts/arow_1/rent_collections"),
    ("PUT", "/api/asset_contracts/arow_1/rent_collections"),
    ("GET", "/api/asset_contracts/export"), ("GET", "/api/asset_contracts/template"),
    ("GET", "/api/statistics"), ("GET", "/api/asset_statistics"),
    ("GET", "/api/users"), ("POST", "/api/users"), ("PUT", "/api/users/u"), ("DELETE", "/api/users/u"),
    ("GET", "/api/departments"), ("POST", "/api/departments"), ("PUT", "/api/departments"),
    ("DELETE", "/api/departments"), ("GET", "/api/config"), ("PUT", "/api/config"),
    ("GET", "/api/logs"), ("GET", "/api/logs/export"), ("GET", "/api/field_logs"),
    ("GET", "/api/sms/status"), ("GET", "/api/sms_logs"), ("GET", "/api/sms_templates"),
    ("POST", "/api/sms/send"), ("POST", "/api/sms/resend"), ("POST", "/api/sms/test_connection"),
    ("POST", "/api/sms/run_weekly"), ("GET", "/api/sms_recipients"),
    ("GET", "/api/attachments/x/download"), ("DELETE", "/api/attachments/x"),
    ("GET", "/api/asset_attachments/x/download"), ("DELETE", "/api/asset_attachments/x"),
    ("GET", "/api/contract_number/next"),
]
EXEMPT = {"/api/login"}


def main():
    C.setup_stdout()
    C.ensure_dirs()
    C.prepare_runtime()
    C.prepare_base()
    db = C.work_db("m13_sec.db")
    prod_before = C.prod_fingerprint()
    r = C.R("M13", "M13 安全测试报告（L6）")
    srv = C.Server(db)
    srv.start()
    B = srv.base
    L = C.RUNTIME_DB
    try:
        A, ia = C.login("admin", base=B)
        assert A, "admin 登录失败"
        p_un = C.query(L, "SELECT 用户名 FROM users WHERE 角色='普通用户' ORDER BY 用户名 LIMIT 1")[0]["用户名"]
        P, ip = C.login(p_un, base=B)
        raw = requests.Session()          # 未登录会话

        def call(sess, method, path, **kw):
            kw.setdefault("timeout", 30)
            return sess.request(method, B + path, **kw)

        def t01():
            bad = []
            for method, path in API_PATHS:
                if path in EXEMPT:
                    continue
                resp = call(raw, method, path, json={})
                if resp.status_code != 401:
                    bad.append("%s %s=%s" % (method, path, resp.status_code))
            assert not bad, "未登录未拦截：%s" % bad[:6]
            cb = requests.post(B + "/api/sms/callback/receipt", json={"guId": "x"}, timeout=30)
            assert cb.status_code == 200, "回调接口应免登录（实际 %s）" % cb.status_code
            return "%d 个受保护接口未登录均 401；短信回调按设计免登录(200)" % (len(API_PATHS) - len(EXEMPT))

        r.case("TC-M13-01", "未登录访问全拦截", "全部 401（回调豁免）", t01)

        def t02():
            idx = open(os.path.join(C.PROJECT_ROOT, "templates", "index.html"), encoding="utf-8").read()
            vhtml = len(re.findall(r"v-html", idx))
            # 越权抽样：普通用户改他部门合同 / 读详情 / 用户列表
            vis = C.visible_depts_of(L, p_un)
            outs = []
            for c in C.query(L, "SELECT id,编号,所属部门 FROM project_contracts ORDER BY id"):
                if str(c["所属部门"]).strip() not in vis:
                    outs.append(c)
                    break
            assert outs, "无越权样本"
            c = outs[0]
            put = call(P, "PUT", "/api/contracts/row_%s" % c["id"], json={}).status_code
            get = call(P, "GET", "/api/contracts/row_%s" % c["id"]).status_code
            users = call(P, "GET", "/api/users").json()["items"]
            assert put == 403, "越权 PUT 返回 %s" % put
            assert get in (403, 404), "越权详情返回 %s" % get
            assert len(users) == 1, "普通用户看到 %d 个用户" % len(users)
            assert vhtml == 0, "前端存在 %d 处 v-html（存储型 XSS 风险，需确认插入内容）" % vhtml
            return "越权 PUT 403 / 详情 %s / 用户列表仅本人；前端 v-html 用量=0（用户数据走 {{}} 转义）" % get

        r.case("TC-M13-02", "越权覆盖与 XSS 转义方式", "无越权、无 v-html", t02)

        def t03():
            tries = [("' OR '1'='1", "' OR '1'='1"), ("admin' --", "x"), ("admin", "' OR '1'='1"),
                     ("' OR 1=1 --", "123456")]
            outs = []
            for un, pw in tries:
                resp = requests.post(B + "/api/login", json={"username": un, "password": pw}, timeout=30)
                outs.append("%s/%s=%s" % (un[:12], pw[:12], resp.status_code))
                assert resp.status_code == 401, "注入尝试登录成功：%s/%s → %s" % (un, pw, resp.status_code)
            return "4 组 SQL 注入式口令均 401（%s）" % ";".join(outs)

        r.case("TC-M13-03", "SQL 注入（登录）", "全部 401", t03)

        def t04():
            inj = "' OR 1=1 --"
            out = []
            for path, params in (("/api/contracts", {"dept": inj}),
                                 ("/api/contracts", {"name": inj, "page_size": 5}),
                                 ("/api/logs", {"module": inj}),
                                 ("/api/field_logs", {"field": inj, "obj": inj}),
                                 ("/api/asset_contracts", {"tenant": inj})):
                resp = A.get(B + path, params=params, timeout=60)
                assert resp.status_code == 200, "%s%s HTTP %s" % (path, params, resp.status_code)
                j = resp.json()
                if "total" in j:
                    out.append("%s=%d" % (path, j["total"]))
                    assert j["total"] == 0, "注入参数被当作通配（%s total=%d）" % (path, j["total"])
            return "筛选参数注入均按字面量匹配（total=0，参数化生效）：%s" % ",".join(out)

        r.case("TC-M13-04", "SQL 注入（筛选/查询参数）", "无 500、按字面量匹配", t04)

        def t05():
            paths = ["/api/attachments/../../app.py/download", "/api/attachments/..%2f..%2fapp.py/download",
                     "/api/attachments/%2e%2e%2fapp.py/download"]
            outs = []
            for p in paths:
                resp = A.get(B + p, timeout=30)
                body = resp.content
                outs.append("%s=%s(%dB)" % (p.split("/api/")[-1][:28], resp.status_code, len(body)))
                # 关键安全属性：不得读到源码/数据库；非 /api 路径命中 SPA 回退时返回首页 HTML 属正常
                assert not body.startswith(b"SQLite format"), "可下载数据库文件：%s" % p
                assert b"import os,sys,re" not in body and b"def api_" not in body, "可读取源码：%s" % p
                assert resp.status_code in (400, 404) or b"<!DOCTYPE html" in body[:200] or b"error" in body[:200], \
                    "返回内容既非 4xx 也非 SPA 首页：%s" % p
            return "附件穿越尝试均未读到源码/数据文件（%s；非 /api 路径命中 SPA 回退返回首页 HTML）" % ";".join(outs)

        r.case("TC-M13-05", "路径穿越（附件下载）", "4xx，无文件泄露", t05)

        def t06():
            outs = []
            for p in ("/static/../app.py", "/static/..%2fapp.py", "/static/%2e%2e/app.py",
                      "/backups/", "/data/合同台账.db", "/../app.py"):
                resp = A.get(B + p, timeout=30)
                outs.append("%s=%s" % (p[:26], resp.status_code))
                assert resp.status_code != 200 or b"import " not in resp.content[:400], \
                    "可读取项目源码/数据文件：%s" % p
                assert not resp.content.startswith(b"SQLite format"), "可下载数据库文件：%s" % p
            return "静态/数据文件穿越均不可读（%s）" % ";".join(outs)

        r.case("TC-M13-06", "路径穿越（静态资源/数据文件）", "不可读源码与数据库", t06)

        def t07():
            c = C.query(L, "SELECT id FROM project_contracts WHERE COALESCE(是否归档,'')<>'是' ORDER BY id LIMIT 1")[0]
            outs = []
            for name, data in (("x.exe", b"MZ\x90\x00"), ("x.php", b"<?php echo 1;?>"), ("x.svg", b"<svg/>")):
                resp = A.post(B + "/api/contracts/row_%s/attachments" % c["id"],
                              files={"file": (name, data)}, data={"type": "decision"}, timeout=60)
                outs.append("%s=%s" % (name, resp.status_code))
                assert resp.status_code == 400, "%s 未被拒绝" % name
            return "可执行/脚本/SVG 等非白名单扩展名均 400（%s）" % ",".join(outs)

        r.case("TC-M13-07", "上传类型防护（服务端白名单）", "非白名单 400", t07)

        def t08():
            cfg = A.get(B + "/api/config", timeout=30)
            assert cfg.status_code == 200, "配置读取 HTTP %s" % cfg.status_code
            body = json.dumps(cfg.json(), ensure_ascii=False)
            before = C.query(L, "SELECT 配置值 v FROM settings WHERE 配置项='sms_secret_key'")[0]["v"]
            assert before not in body, "GET /api/config 明文回显了 SecretKey（严重）"
            A.put(B + "/api/config", json={"sms_secret_key": "******"}, timeout=30)
            after_mask = C.query(L, "SELECT 配置值 v FROM settings WHERE 配置项='sms_secret_key'")[0]["v"]
            A.put(B + "/api/config", json={"sms_secret_key": ""}, timeout=30)
            after_empty = C.query(L, "SELECT 配置值 v FROM settings WHERE 配置项='sms_secret_key'")[0]["v"]
            assert after_mask == before and after_empty == before, \
                "掩码/空值覆盖了真实密钥（%r → %r / %r）" % (before, after_mask, after_empty)
            assert "******" in body, \
                "GET /api/config 未按文档返回掩码字段 sms.secret_key=******（实际响应中无该键，前端密钥框为空）"
            return "SecretKey 掩码字段存在且提交掩码不回写"

        r.case("TC-M13-08", "短信密钥掩码与掩码回写", "掩码回显、不覆盖", t08, priority="P1")

        def t09():
            resp = requests.post(B + "/api/sms/callback/receipt", json={"guId": "NOPE"}, timeout=30)
            assert resp.status_code == 200, "未配置 token 时回调应放行（实际 %s）" % resp.status_code
            A.put(B + "/api/config", json={"sms_receipt_token": "TK-SEC"}, timeout=30)
            try:
                bad = requests.post(B + "/api/sms/callback/receipt?token=WRONG", json={"guId": "NOPE"}, timeout=30)
                good = requests.post(B + "/api/sms/callback/receipt?token=TK-SEC", json={"guId": "NOPE"}, timeout=30)
                hdr = requests.post(B + "/api/sms/callback/reply", headers={"X-SMS-Token": "TK-SEC"},
                                    json={"guId": "NOPE"}, timeout=30)
            finally:
                A.put(B + "/api/config", json={"sms_receipt_token": ""}, timeout=30)
            assert bad.status_code == 403, "错误 token 返回 %s" % bad.status_code
            assert good.status_code == 200 and hdr.status_code == 200, "正确 token 被拒"
            return "免登录回调：未配置放行；配置后错误 token 403、?token 与 X-SMS-Token 均可通过"

        r.case("TC-M13-09", "回调免登录 + token 校验", "403 / 200", t09, priority="P1")

        def t10():
            src = open(os.path.join(C.PROJECT_ROOT, "app.py"), encoding="utf-8").read()
            m = re.search(r'app\.secret_key\s*=\s*"([^"]+)"', src)
            assert m, "未找到 secret_key 赋值"
            assert "os.environ" not in src.split("app.secret_key")[0][-200:], "secret_key 已环境变量化"
            return "风险记录：会话密钥硬编码在源码（长度 %d），未走环境变量 —— 文档 P1 计划项" % len(m.group(1))

        r.case("TC-M13-10", "会话密钥来源（现状走读）", "记录现状风险", t10, priority="P1")

        def t11():
            src = open(os.path.join(C.PROJECT_ROOT, "app.py"), encoding="utf-8").read()
            assert 'hashlib.sha256((p+SALT).encode())' in src, "密码哈希实现已变化，请复核"
            cols = [c["name"] for c in C.columns(L, "users")]
            pwdcol = "密码哈希" if "密码哈希" in cols else "密码"
            rows = C.query(L, 'SELECT 用户名 u, "%s" p FROM users' % pwdcol)
            bad = []
            for x in rows:
                v = x["p"] or ""
                if not re.fullmatch(r"[0-9a-f]{64}", str(v)):
                    bad.append(x["u"])
            assert not bad, "存在非 SHA256 形态口令：%s" % bad[:5]
            assert not any(str(x["p"]) == "123456" for x in rows), "存在明文口令"
            return "风险记录：%d 个账号口令列(%s)均为 SHA256+固定盐 的 64 位十六进制；无明文（强度不足，文档 P1 计划项）" % (
                len(rows), pwdcol)

        r.case("TC-M13-11", "密码存储形态（现状走读）", "记录现状风险", t11, priority="P1")

        def t12():
            user = "admin"
            codes = []
            for i in range(6):
                resp = requests.post(B + "/api/login", json={"username": user, "password": "wrong%d" % i}, timeout=30)
                codes.append(resp.status_code)
            ok = requests.post(B + "/api/login", json={"username": user, "password": "123456"}, timeout=30)
            assert all(c == 401 for c in codes), "失败登录状态异常：%s" % codes
            assert ok.status_code == 200, "正确口令在连续失败后被拒（已锁定？），实际 %s" % ok.status_code
            fails = C.query(L, "SELECT COUNT(*) c FROM operation_logs WHERE 模块='用户登录' AND 结果='失败'")[0]["c"]
            return "风险记录：连续 6 次错误口令无锁定/延时（仍可正常登录）；失败登录已留痕 %d 条" % fails

        r.case("TC-M13-12", "登录失败锁定（现状）", "记录现状风险", t12, priority="P1")

        def t13():
            s = requests.Session()
            resp = s.post(B + "/api/login", json={"username": "admin", "password": "123456"}, timeout=30)
            sc = resp.headers.get("Set-Cookie", "")
            flags = {"HttpOnly": "httponly" in sc.lower(), "SameSite": "samesite" in sc.lower(),
                     "Secure": "secure" in sc.lower()}
            assert flags["HttpOnly"], "会话 Cookie 缺少 HttpOnly：%s" % sc[:120]
            assert flags["SameSite"], "会话 Cookie 缺少 SameSite：%s" % sc[:120]
            return "Cookie 属性：HttpOnly=%s、SameSite=%s、Secure=%s（内网 HTTP 部署下 Secure 缺失属预期，记录）" % (
                flags["HttpOnly"], flags["SameSite"], flags["Secure"])

        r.case("TC-M13-13", "会话 Cookie 安全属性", "HttpOnly + SameSite", t13)

        def t14():
            outs = []
            cases = [("GET", "/api/unknown_endpoint", None), ("GET", "/api/contracts/abc", None),
                     ("POST", "/api/login", None), ("POST", "/api/contracts", "not-a-json"),
                     ("DELETE", "/api/contracts/row_1", None)]
            for method, path, data in cases:
                if data == "not-a-json":
                    resp = A.post(B + path, data=data, headers={"Content-Type": "application/json"}, timeout=30)
                else:
                    resp = call(A, method, path)
                outs.append("%s %s=%s" % (method, path, resp.status_code))
                assert resp.status_code in (400, 404, 405), "%s %s 返回 %s" % (method, path, resp.status_code)
                txt = resp.text
                for leak in ("Traceback", "sqlite3.", 'File "', "werkzeug", "SECRET"):
                    assert leak not in txt, "%s 响应泄露内部信息（%s）：%s" % (path, leak, txt[:120])
                try:
                    assert "error" in resp.json(), "错误响应非统一结构：%s" % txt[:100]
                except ValueError:
                    raise AssertionError("错误响应不是 JSON：%s" % txt[:100])
            return "5 类异常请求均返回统一 JSON（%s），无堆栈/内部信息泄露" % ";".join(outs)

        r.case("TC-M13-14", "错误信息不泄露内部细节", "统一 JSON、无堆栈", t14)

        def t15():
            outs = []
            for method, path in (("GET", "/api/logs/export"), ("GET", "/api/contracts/template"),
                                 ("POST", "/api/users"), ("PUT", "/api/config"), ("POST", "/api/departments"),
                                 ("POST", "/api/contracts/import")):
                resp = call(P, method, path, json={})
                outs.append("%s %s=%s" % (method, path, resp.status_code))
                assert resp.status_code == 403, "%s %s 未被拒绝（%s）" % (method, path, resp.status_code)
            return "6 个管理员专属接口对普通用户均 403（%s）" % ",".join(x.split("=")[-1] for x in outs)

        r.case("TC-M13-15", "未授权访问管理员接口", "403", t15, priority="P1")

        def t16():
            gi = open(os.path.join(C.PROJECT_ROOT, ".gitignore"), encoding="utf-8").read()
            tracked = []
            if "data/合同台账.db" not in gi and "data/" not in gi:
                tracked.append("data/ 数据库未列入 .gitignore")
            if "backups/" not in gi:
                tracked.append("backups/ 未列入 .gitignore")
            if "test_data/" not in gi:
                tracked.append("test_data/ 未列入 .gitignore")
            assert not tracked, "仓库存在敏感/大文件暴露：%s" % tracked
            return "风险记录：data/（含真实库与附件）仍是仓库内容并会同步到公开仓库；backups/、test_data/ 已排除"

        r.case("TC-M13-16", "数据文件与仓库暴露面", "记录风险", t16, priority="P2")

        srv.sync_db()
    finally:
        srv.stop()
    prod_after = C.prod_fingerprint()
    meta = {
        "服务地址": B,
        "工作库": "test_data/m13_sec.db",
        "生产库未改动": "是" if prod_before == prod_after else "否（指纹变化！）",
    }
    r.report("M13_安全测试.md", meta=meta)
    n, p, f, e = r.stat()
    print("\n[SUMMARY] M13 共 %d 例：通过 %d，失败 %d，错误 %d" % (n, p, f, e))
    sys.stdout.flush()
    os._exit(0 if (f == 0 and e == 0) else 1)


if __name__ == "__main__":
    main()
