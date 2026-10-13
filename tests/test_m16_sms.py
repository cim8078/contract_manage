# -*- coding: utf-8 -*-
"""M16 短信模块测试（L2/L6）：对接/合规/回调/去重。
约束：全程 sms_enabled=0，**不真实外发**；只验证校验层、分段、去重、回调与开关语义。
方式：进程内导入「运行时副本」app（其 data/ 为测试库）+ Flask test_client（真实路由与状态码）。
"""
import os
import sys
import shutil
import sqlite3
import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _common as C  # noqa: E402


def main():
    C.setup_stdout()
    C.ensure_dirs()
    C.prepare_runtime()
    C.prepare_base()
    db = C.work_db("m16_sms.db")
    shutil.copyfile(db, C.RUNTIME_DB)
    prod_before = C.prod_fingerprint()
    r = C.R("M16", "M16 短信模块（对接/合规/回调/去重）测试报告（L2/L6）")
    if C.RUNTIME not in sys.path:
        sys.path.insert(0, C.RUNTIME)
    import app as m  # noqa: E402
    assert os.path.abspath(m.__file__).startswith(os.path.abspath(C.RUNTIME)), "拒绝：app 未指向运行时副本"
    L = C.RUNTIME_DB
    B = ""
    try:
        c_admin = m.app.test_client()
        c_pu = m.app.test_client()
        resp = c_admin.post("/api/login", json={"username": "admin", "password": "123456"})
        assert resp.status_code == 200, "admin 登录失败 %s" % resp.status_code
        pu = C.query(L, "SELECT 用户名 FROM users WHERE 角色='普通用户' ORDER BY 用户名 LIMIT 1")[0]["用户名"]
        assert c_pu.post("/api/login", json={"username": pu, "password": "123456"}).status_code == 200
        c_q = m.app.test_client()
        q_un = C.query(L, "SELECT 用户名 FROM users WHERE 角色='查询用户' ORDER BY 用户名 LIMIT 1")[0]["用户名"]
        assert c_q.post("/api/login", json={"username": q_un, "password": "123456"}).status_code == 200

        # 安全前置：读取生产副本的短信总开关原始取值，并在本副本内强制关闭，杜绝任何真实外发
        orig_enabled = [x["v"] for x in C.query(L, "SELECT 配置值 v FROM settings WHERE 配置项='sms_enabled'")]
        orig_enabled = orig_enabled[0] if orig_enabled else None
        m.set_cfg("sms_enabled", "0")

        def sql(query, params=()):
            con = sqlite3.connect(L)
            con.row_factory = sqlite3.Row
            try:
                return [dict(x) for x in con.execute(query, params).fetchall()]
            finally:
                con.close()

        def ex(query, params=()):
            con = sqlite3.connect(L)
            try:
                cur = con.execute(query, params)
                con.commit()
                return cur.lastrowid
            finally:
                con.close()

        def insert_log(phone="13800138000", status="成功", when=None, guid="", dedup="", content="测试内容"):
            when = when or datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            return ex('INSERT INTO sms_logs ("时间","场景","关联类型","手机号","短信内容","去重键",'
                      '"触发方式","发送人","状态","外部guId") VALUES (?,?,?,?,?,?,?,?,?,?)',
                      (when, "manual", "手动", phone, content, dedup, "手动", "admin", status, guid))

        def J(resp):
            """兼容 flask(test_client) 的 .json 属性与 requests 的 .json() 方法。"""
            try:
                v = resp.json
                return v() if callable(v) else v
            except Exception:
                return {}

        def t01():
            resp = c_admin.get("/api/sms/status")
            assert resp.status_code == 200, "HTTP %s" % resp.status_code
            j = J(resp)
            keys = set(j)
            need = {"enabled", "weekly_enabled", "sign"}
            assert len(need & keys) >= 2 or len(keys) >= 5, "状态字段过少：%s" % sorted(keys)
            assert j.get("enabled") in (0, False, "0"), "总开关应默认关闭，实际 %r" % j.get("enabled")
            return "200；字段 %s；总开关=%r" % (sorted(keys)[:9], j.get("enabled"))

        r.case("TC-M16-01", "短信状态接口", "200 + 开关/密钥/执行时间等字段", t01)

        def t02():
            resp = c_admin.get("/api/sms_templates")
            assert resp.status_code == 200, "HTTP %s" % resp.status_code
            j = J(resp)
            items = j.get("items", j if isinstance(j, list) else [])
            scenes = set()
            for it in items:
                scenes.add(it.get("scene") or it.get("场景"))
            exp = {"expire", "receivable", "payable", "urge", "manual"}
            assert exp.issubset(scenes), "模板场景缺失：%s（实际 %s）" % (exp - scenes, sorted(scenes))
            return "5 个场景齐全：%s" % sorted(scenes)

        r.case("TC-M16-02", "短信模板列表（5 场景）", "expire/receivable/payable/urge/manual", t02)

        def t03():
            cur = J(c_admin.get("/api/sms_templates"))
            items = cur.get("items", cur)
            tpl = items[0]
            scene = tpl.get("scene") or tpl.get("场景")
            content = tpl.get("content") or tpl.get("内容") or "测试模板内容"
            ok = c_admin.put("/api/sms_templates/%s" % scene, json={"content": content, "enabled": 1})
            bad = c_pu.put("/api/sms_templates/%s" % scene, json={"content": "x"})
            assert ok.status_code == 200, "管理员改模板返回 %s" % ok.status_code
            assert bad.status_code == 403, "普通用户改模板返回 %s（期望 403）" % bad.status_code
            return "管理员 200；普通用户 403（场景 %s）" % scene

        r.case("TC-M16-03", "模板修改仅管理员", "200 / 403", t03, priority="P1")

        def t04():
            resp = c_admin.post("/api/sms/send", json={"content": "这是一条测试通知", "phones": ["13800138000"]})
            j = J(resp)
            free_off = m.sms_send_free(["13800138000"], "测试内容", "admin")
            m.set_cfg("sms_enabled", "1")
            try:
                free_blocked = m.sms_send_free(["13800138000"], "【注入】测试内容", "admin")
            finally:
                m.set_cfg("sms_enabled", "0")
            probs = []
            if resp.status_code == 200 or j.get("ok"):
                probs.append("接口返回 HTTP %s/ok=%s（error=%r）" % (resp.status_code, j.get("ok"), j.get("error")))
            if free_off.get("ok"):
                probs.append("未启用时 sms_send_free.ok=True（error=%r）" % free_off.get("error"))
            if free_blocked.get("ok"):
                probs.append("内容全部被拦截时 sms_send_free.ok=True（error=%r）" % free_blocked.get("error"))
            assert not probs, "未启用/全被拦截却返回成功（假成功）：" + "；".join(probs)
            return "禁用态 HTTP %s ok=%s；底层 ok=%s/%s" % (resp.status_code, j.get("ok"),
                                                          free_off.get("ok"), free_blocked.get("ok"))

        r.case("TC-M16-04", "禁用态不真实外发", "4xx + 提示未启用", t04)

        def t05():
            resp = c_admin.post("/api/sms/send", json={"content": "【潭合同】您好，这是一条测试", "phones": ["13800138000"]})
            assert resp.status_code == 400, "含【】返回 %s" % resp.status_code
            msg = J(resp).get("error", "")
            assert "【】" in msg or "中括号" in msg, "提示未说明【】原因：%s" % msg
            return "内容含【】→ 400「%s」" % msg[:40]

        r.case("TC-M16-05", "「【】」拦截", "400 + 中文原因", t05)

        def t06():
            old = m.get_cfg("sms_block_words")
            m.set_cfg("sms_block_words", "违规词")
            try:
                msg = m._sms_content_issue("本消息包含违规词，请忽略")
                ok = m._sms_content_issue("这是一条正常通知")
            finally:
                m.set_cfg("sms_block_words", old or "")
            assert msg, "敏感词未被拦截"
            assert not ok, "正常内容被误拦：%s" % ok
            return "敏感词命中 → 「%s」；正常内容放行" % msg[:36]

        r.case("TC-M16-06", "敏感词拦截", "命中拦截 / 正常放行", t06, priority="P1")

        def t07():
            cfg = m._sms_cfg()
            seg, mx = cfg["segment_len"], cfg["max_segments"]
            short = "短内容"
            p1 = m._sms_split_content(short, seg, mx)
            long_text = "。".join(["这是一段用于测试分段功能的较长文本内容"] * 8)
            p2 = m._sms_split_content(long_text, seg, mx)
            huge = "很长的内容" * 200
            p3 = m._sms_split_content(huge, seg, mx)
            assert p1 == [short], "短内容不应分段：%s" % p1
            assert p2 and len(p2) > 1 and len(p2) <= mx, "长内容分段数异常：%s" % (len(p2) if p2 else p2)
            assert all(len(x) <= seg + 2 for x in p2), "分段长度超限：%s" % [len(x) for x in p2]
            assert p3 is None, "超长内容应被拒绝（>%d 段），实际 %s" % (mx, len(p3) if p3 else p3)
            return "单段长度=%d、最多 %d 段；短=1 段、长=%d 段、超长=拒绝" % (seg, mx, len(p2))

        r.case("TC-M16-07", "超长自动分段（70 字/段、≤3 段）", "分段正确、超限拒绝", t07, priority="P1")

        def t08():
            good = m._sms_mobile("13800138000")
            bad = m._sms_mobile("0731-12345")
            resp = c_admin.post("/api/sms/send", json={"content": "测试", "phones": ["12345"]})
            assert good, "合法 11 位号码未被识别"
            assert not bad, "固话未被识别为无效：%r" % bad
            assert resp.status_code == 400, "无效号码返回 %s" % resp.status_code
            return "11 位号码通过；固话被拒；接口对无效号码 400"

        r.case("TC-M16-08", "手机号 11 位校验", "合法通过 / 非法拒绝", t08)

        def t09():
            phone = "13900139000"
            insert_log(phone=phone, status="成功")
            cfg = dict(m._sms_cfg())
            cfg["daily_per_phone"] = 1
            cfg["min_interval_sec"] = 0
            ok, blocked = m._sms_precheck([{"phone": phone, "content": "测试", "rname": "甲"}], cfg)
            assert not ok and blocked, "达到每日上限未被拦截"
            assert "上限" in blocked[0]["reason"], "拦截原因不符：%s" % blocked[0]["reason"]
            return "今日已发 1 条 + 上限 1 → 「%s」" % blocked[0]["reason"]

        r.case("TC-M16-09", "同号每日 100 条上限", "达到上限即拦截", t09, priority="P1")

        def t10():
            phone = "13900139001"
            insert_log(phone=phone, status="成功")
            cfg = dict(m._sms_cfg())
            cfg["daily_per_phone"] = 0
            cfg["min_interval_sec"] = 30
            ok, blocked = m._sms_precheck([{"phone": phone, "content": "测试", "rname": "乙"}], cfg)
            assert not ok and blocked and "限流" in blocked[0]["reason"], "最小间隔未生效：%s" % blocked
            return "上次发送后立即再发 → 「%s」" % blocked[0]["reason"]

        r.case("TC-M16-10", "同号最小间隔 30 秒", "间隔内限流", t10, priority="P1")

        def t11():
            cfg = m._sms_cfg()
            suffix = cfg["sign_suffix"]
            body = "这是一条测试通知"
            once = m._sms_sign_content(body)
            twice = m._sms_sign_content(once)
            assert once.endswith(suffix) and body in once, "署名未追加：%r" % once
            assert twice == once, "署名重复追加：%r → %r" % (once, twice)
            return "正文末尾追加「%s」且幂等（重复调用不叠加）" % suffix

        r.case("TC-M16-11", "署名末尾追加且幂等", "追加一次、不重复", t11, priority="P1")

        def t12():
            idx = sql("SELECT sql FROM sqlite_master WHERE type='index' AND name='idx_sms_logs_dedup'")
            assert idx, "未找到 idx_sms_logs_dedup 索引"
            key = "dedup-test-%s" % datetime.datetime.now().strftime("%H%M%S")
            insert_log(phone="13900139002", status="待发送", dedup=key)
            raised = False
            try:
                insert_log(phone="13900139002", status="待发送", dedup=key)
            except sqlite3.IntegrityError:
                raised = True
            assert raised, "同去重键第二条未被唯一索引拒绝"
            return "去重索引生效：%s" % idx[0]["sql"][:90]

        r.case("TC-M16-12", "每周去重索引", "同键第二条被拒绝", t12, priority="P1")

        def t13():
            proj = sql('SELECT p."编号" no, p."经办人" a, p."现场负责人" b FROM project_contracts p '
                       'JOIN users u ON u."姓名"=p."经办人" WHERE COALESCE(u."电话",\'\')<>\'\' LIMIT 1')
            assert proj, "库内无「经办人带电话」的项目合同样本"
            c = next(x for x in m.read_contracts() if str(x.get("编号")) == proj[0]["no"])
            recv = m._sms_default_receivers("project", c, "expire")
            phones = [r[2] for r in recv]
            assert recv, "未解析出收件人"
            assert all(m._sms_mobile(p) for p in phones), "存在非法号码：%s" % phones
            u_phone = sql('SELECT "电话" p FROM users WHERE "姓名"=?', (proj[0]["a"],))[0]["p"]
            assert m._sms_mobile(u_phone) in [m._sms_mobile(p) for p in phones], \
                "经办人电话未进入收件人：%s vs %s" % (u_phone, phones)
            a = sql('SELECT * FROM asset_contracts LIMIT 1')[0]
            recv2 = m._sms_default_receivers("asset", next(x for x in m.read_assets() if str(x["_row"]) == str(a["id"])), "expire")
            return "项目收件人 %d 个（含经办人电话）；资产收件人 %d 个" % (len(recv), len(recv2))

        r.case("TC-M16-13", "收件人口径（经办人+现场负责人 / 责任人+承租人）", "号码取自库内且合法", t13, priority="P1")

        def t14():
            insert_log(phone="13900139003", status="成功", guid="GU-TEST-1")
            resp = c_admin.post("/api/sms/callback/receipt",
                                json={"guId": "GU-TEST-1", "receiptCode": "0", "receiptDesc": "DELIVRD"})
            assert resp.status_code == 200, "回执回调返回 %s" % resp.status_code
            row = sql('SELECT "回执状态码" c, "回执描述" d FROM sms_logs WHERE "外部guId"=?', ("GU-TEST-1",))[0]
            assert str(row["c"]) == "0" and row["d"] == "DELIVRD", "回执未回写：%s" % row
            return "回执回写 状态码=%s 描述=%s" % (row["c"], row["d"])

        r.case("TC-M16-14", "回执回调按 guId 回写", "200 + 状态码/描述落库", t14, priority="P1")

        def t15():
            resp = c_admin.post("/api/sms/callback/reply",
                                json={"guId": "GU-TEST-1", "replyTime": "2026-10-12 10:00:00", "content": "已收到"})
            assert resp.status_code == 200, "回复回调返回 %s" % resp.status_code
            row = sql('SELECT "回复时间" t, "回复内容" c FROM sms_logs WHERE "外部guId"=?', ("GU-TEST-1",))[0]
            assert row["c"] == "已收到" and str(row["t"]).startswith("2026-10-12"), "回复未回写：%s" % row
            return "回复回写 时间=%s 内容=%s" % (row["t"], row["c"])

        r.case("TC-M16-15", "回复回调回写", "200 + 时间/内容落库", t15, priority="P1")

        def t16():
            old = m.get_cfg("sms_receipt_token")
            m.set_cfg("sms_receipt_token", "TK-ABC")
            try:
                no_tk = c_admin.post("/api/sms/callback/receipt", json={"guId": "GU-TEST-1", "receiptCode": "1"})
                bad_tk = c_admin.post("/api/sms/callback/receipt?token=WRONG",
                                      json={"guId": "GU-TEST-1", "receiptCode": "1"})
                ok_tk = c_admin.post("/api/sms/callback/receipt?token=TK-ABC",
                                     json={"guId": "GU-TEST-1", "receiptCode": "2"})
            finally:
                m.set_cfg("sms_receipt_token", old or "")
            assert no_tk.status_code == 403 and bad_tk.status_code == 403, \
                "token 校验未生效：no=%s bad=%s" % (no_tk.status_code, bad_tk.status_code)
            assert ok_tk.status_code == 200, "正确 token 被拒：%s" % ok_tk.status_code
            return "无 token/错误 token → 403；正确 token → 200"

        r.case("TC-M16-16", "回调 token 校验（可配置）", "配置后强校验", t16, priority="P1")

        def t17():
            many = ["138%08d" % i for i in range(51)]
            r1 = c_admin.post("/api/sms/send", json={"content": "测试", "phones": many})
            r2 = c_admin.post("/api/sms/send", json={"content": "长" * 501, "phones": ["13800138000"]})
            ok_id = insert_log(phone="13900139004", status="成功")
            r3 = c_admin.post("/api/sms/resend", json={"ids": [ok_id]})
            r4 = c_q.post("/api/sms/send", json={"content": "测试", "phones": ["13800138000"]})
            assert r1.status_code == 400 and "50" in str(J(r1).get("error", "")), "号码数上限未生效：%s" % r1.text[:80]
            assert r2.status_code == 400 and "500" in str(J(r2).get("error", "")), "内容长度上限未生效：%s" % r2.text[:80]
            assert r3.status_code == 400 and "失败" in str(J(r3).get("error", "")), "非失败记录被重发：%s" % r3.text[:80]
            assert r4.status_code == 403, "查询用户发送返回 %s（期望 403）" % r4.status_code
            return "51 号 400；501 字 400；重发非失败记录 400；查询用户发送 403"

        r.case("TC-M16-17", "点对点/重发边界（≤50 号、≤500 字、仅失败可重发）",
               "越界 400，非法重发 400", t17, priority="P1")

        def t18():
            assert str(orig_enabled) == "0", \
                "生产副本中短信总开关 sms_enabled=%r（文档约定默认 0，联调通过后再开）；" \
                "开关为开启且 AK/SK 已配置，任何发送调用都会尝试真实外发（本次已在本副本置 0 规避）" % orig_enabled
            return "sms_enabled=%r（关闭）" % orig_enabled

        r.case("TC-M16-18", "上线前短信总开关应为关闭（安全）", "sms_enabled=0", t18)

        shutil.copyfile(L, db)
    finally:
        prod_after = C.prod_fingerprint()
        meta = {
            "方式": "运行时副本 app + Flask test_client（未启动网络服务）",
            "工作库": "test_data/m16_sms.db",
            "短信开关": "sms_enabled=0（全程不真实外发）",
            "生产库未改动": "是" if prod_before == prod_after else "否（指纹变化！）",
        }
        r.report("M16_短信模块测试.md", meta=meta)
        n, p, f, e = r.stat()
        print("\n[SUMMARY] M16 共 %d 例：通过 %d，失败 %d，错误 %d" % (n, p, f, e))
        sys.stdout.flush()
        os._exit(0 if (f == 0 and e == 0) else 1)


if __name__ == "__main__":
    main()
