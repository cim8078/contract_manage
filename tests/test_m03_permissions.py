# -*- coding: utf-8 -*-
"""M03 角色与数据权限测试（L2 权限矩阵）。
覆盖：三角色登录、项目/资产可见范围（所属部门 / 管理部门 ∪ 分管部门）、查询用户只读、
越权编辑/详情/归档/附件、导出范围一致、用户列表可见性、仅管理员接口、部门写权限、自身信息、admin 保护。
隔离：运行时副本 + test_data 工作库；用例本身不产生有效写操作（均为被拒绝路径）。
"""
import os
import sys
from io import BytesIO

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _common as C  # noqa: E402
import openpyxl  # noqa: E402

NAMED = ["章伟", "谢玮", "曾明峰", "熊玉婷"]


def main():
    C.setup_stdout()
    C.ensure_dirs()
    C.prepare_runtime()
    C.prepare_base()
    db = C.work_db("m03_perm.db")
    prod_before = C.prod_fingerprint()
    r = C.R("M03", "M03 角色与数据权限测试报告（L2）")
    srv = C.Server(db)
    srv.start()
    B = srv.base
    try:
        users = C.query(db, "SELECT 用户名,姓名,角色,所属部门,分管部门 FROM users")

        def try_login(un):
            return C.login(un, base=B)

        logged = {}
        for role in ("管理员", "普通用户", "查询用户"):
            cands = [u for u in users if u["角色"] == role]
            cands.sort(key=lambda u: (0 if (u["用户名"] in NAMED or u["姓名"] in NAMED) else 1, u["用户名"]))
            for u in cands:
                s, info = try_login(u["用户名"])
                if s:
                    logged[role] = {"u": u, "s": s, "info": info}
                    break
        assert set(logged) == {"管理员", "普通用户", "查询用户"}, "三角色账号登录失败：%s" % sorted(logged)

        PU = logged["普通用户"]
        QU = logged["查询用户"]
        AU = logged["管理员"]

        def proj_id_out_of_scope(un):
            vis = C.visible_depts_of(db, un)
            for c in C.query(db, "SELECT id,编号,所属部门,是否归档,经办人 FROM project_contracts ORDER BY id"):
                if vis is not None and str(c["所属部门"] or "").strip() not in vis:
                    return c
            return None

        def proj_id_in_scope_not_handler(un, name):
            vis = C.visible_depts_of(db, un)
            for c in C.query(db, "SELECT id,编号,所属部门,是否归档,经办人 FROM project_contracts ORDER BY id"):
                if str(c["是否归档"] or "") == "是":
                    continue
                if vis is None or str(c["所属部门"] or "").strip() in vis:
                    if name not in str(c["经办人"] or ""):
                        return c
            return None

        # ── TC-M03-01 三角色登录覆盖 ──
        def t01():
            out = []
            for role in ("管理员", "普通用户", "查询用户"):
                l = logged[role]
                assert l["info"]["role"] == l["u"]["角色"], "%s 返回 role=%s，库内=%s" % (
                    l["u"]["用户名"], l["info"]["role"], l["u"]["角色"])
                out.append("%s→%s" % (l["u"]["用户名"], l["info"]["role"]))
            for name in NAMED:
                hit = [u for u in users if u["用户名"] == name or u["姓名"] == name]
                if not hit:
                    out.append("%s(不在库)" % name)
                    continue
                u0 = hit[0]
                s, info = try_login(u0["用户名"])
                assert s, "指定账号 %s 默认口令登录失败：%s" % (name, info)
                assert info["role"] == u0["角色"], "%s 角色不符" % name
                out.append("%s(%s)→%s" % (name, u0["用户名"], info["role"]))
            return "；".join(out)

        r.case("TC-M03-01", "三角色与指定账号登录", "200 且角色与库内一致", t01)

        # ── TC-M03-02 项目合同可见范围 ──
        def t02():
            out = []
            for role in ("管理员", "普通用户", "查询用户"):
                l = logged[role]
                un = l["u"]["用户名"]
                vis = C.visible_depts_of(db, un)
                if vis is None:
                    exp = C.query(db, "SELECT COUNT(*) c FROM project_contracts")[0]["c"]
                elif not vis:
                    exp = 0
                else:
                    ph = ",".join("?" * len(vis))
                    exp = C.query(db, "SELECT COUNT(*) c FROM project_contracts WHERE 所属部门 IN (%s)" % ph,
                                  tuple(vis))[0]["c"]
                resp = l["s"].get(B + "/api/contracts", params={"page": 1, "page_size": 200}, timeout=60)
                assert resp.status_code == 200, "%s 列表 HTTP %s" % (un, resp.status_code)
                j = resp.json()
                assert j["total"] == exp, "%s 可见项目合同 total=%s，期望 %s" % (un, j["total"], exp)
                if vis is not None:
                    bad = [it["number"] for it in j["items"] if str(it.get("department", "")).strip() not in vis]
                    assert not bad, "%s 返回越权合同：%s" % (un, bad[:3])
                out.append("%s:%d(期望%d)" % (un, j["total"], exp))
            return "；".join(out)

        r.case("TC-M03-02", "项目合同可见范围=所属部门∪分管部门", "total 与 SQL 一致、无越权", t02)

        # ── TC-M03-03 资产合同可见范围（管理部门） ──
        def t03():
            out = []
            for role in ("管理员", "普通用户", "查询用户"):
                l = logged[role]
                un = l["u"]["用户名"]
                vis = C.visible_depts_of(db, un)
                if vis is None:
                    exp = C.query(db, "SELECT COUNT(*) c FROM asset_contracts")[0]["c"]
                elif not vis:
                    exp = 0
                else:
                    ph = ",".join("?" * len(vis))
                    exp = C.query(db, "SELECT COUNT(*) c FROM asset_contracts WHERE 管理部门 IN (%s)" % ph,
                                  tuple(vis))[0]["c"]
                resp = l["s"].get(B + "/api/asset_contracts", params={"page": 1, "page_size": 200}, timeout=60)
                assert resp.status_code == 200, "%s 资产列表 HTTP %s" % (un, resp.status_code)
                j = resp.json()
                assert j["total"] == exp, "%s 可见资产合同 total=%s，期望 %s" % (un, j["total"], exp)
                if vis is not None:
                    bad = [it.get("id") for it in j["items"]
                           if str(it.get("dept", it.get("department", ""))).strip() not in vis]
                    assert not bad, "%s 返回越权资产：%s" % (un, bad[:3])
                out.append("%s:%d(期望%d)" % (un, j["total"], exp))
            return "；".join(out)

        r.case("TC-M03-03", "资产合同可见范围=管理部门∪分管部门", "total 与 SQL 一致、无越权", t03)

        # ── TC-M03-04 查询用户只读 ──
        def t04():
            s = QU["s"]
            cid = "row_1"
            aid = "arow_%s" % (C.query(db, "SELECT id FROM asset_contracts ORDER BY id LIMIT 1")[0]["id"])
            checks = [
                ("POST /api/contracts", s.post(B + "/api/contracts", json={"name": "x"}, timeout=30)),
                ("PUT /api/contracts", s.put(B + "/api/contracts/" + cid, json={}, timeout=30)),
                ("DELETE /api/contracts", s.delete(B + "/api/contracts/" + cid, json={"reason": "t"}, timeout=30)),
                ("PUT archive", s.put(B + "/api/contracts/" + cid + "/archive", json={}, timeout=30)),
                ("POST asset", s.post(B + "/api/asset_contracts", json={"place": "x"}, timeout=30)),
                ("PUT asset", s.put(B + "/api/asset_contracts/" + aid, json={}, timeout=30)),
                ("DELETE asset", s.delete(B + "/api/asset_contracts/" + aid, json={"reason": "t"}, timeout=30)),
                ("PUT rent", s.put(B + "/api/asset_contracts/" + aid + "/rent_collections",
                                   json={"items": []}, timeout=30)),
                ("POST asset attachment", s.post(B + "/api/asset_contracts/" + aid + "/attachments",
                                                 files={"file": ("a.png", b"x", "image/png")},
                                                 data={"type": "condition"}, timeout=30)),
            ]
            bad = ["%s=%s" % (n, resp.status_code) for n, resp in checks if resp.status_code != 403]
            assert not bad, "查询用户写操作未被拒绝：%s" % bad
            return "查询用户 %s 的 %d 项写操作全部 403" % (QU["u"]["用户名"], len(checks))

        r.case("TC-M03-04", "查询用户只读（全部写操作 403）", "9 项写接口均 403", t04)

        # ── TC-M03-05 越权编辑他部门合同 ──
        def t05():
            un = PU["u"]["用户名"]
            vis = C.visible_depts_of(db, un)
            c = None
            for x in C.query(db, "SELECT id,编号,所属部门,是否归档 FROM project_contracts ORDER BY id"):
                if str(x["是否归档"] or "") == "是":
                    continue
                if vis is not None and str(x["所属部门"] or "").strip() not in vis:
                    c = x
                    break
            assert c, "%s 不存在「未归档且在范围外」的越权样本" % un
            resp = PU["s"].put(B + "/api/contracts/row_%s" % c["id"], json={}, timeout=30)
            assert resp.status_code == 403, "越权 PUT 返回 %s（期望 403）" % resp.status_code
            msg = resp.json().get("error", "")
            assert ("本部门" in msg) or ("分管" in msg), "403 原因非「越权编辑」口径：%s" % msg
            return "PUT 编号 %s（部门 %s）→ 403 %s" % (c["编号"], c["所属部门"], msg)

        r.case("TC-M03-05", "越权编辑他部门项目合同", "403", t05)

        # ── TC-M03-06 越权查看详情 ──
        def t06():
            un = PU["u"]["用户名"]
            c = proj_id_out_of_scope(un)
            assert c, "%s 不存在越权合同样本" % un
            resp = PU["s"].get(B + "/api/contracts/row_%s" % c["id"], timeout=30)
            assert resp.status_code == 404, "越权详情返回 %s（期望 404 隐藏）" % resp.status_code
            return "GET 越权合同 → 404（按设计隐藏为“合同不存在”）"

        r.case("TC-M03-06", "越权查看合同详情", "404（信息隐藏）", t06)

        # ── TC-M03-07 归档权限 ──
        def t07():
            un = PU["u"]["用户名"]
            c = proj_id_in_scope_not_handler(un, PU["u"]["姓名"])
            msgs = []
            if c:
                resp = PU["s"].put(B + "/api/contracts/row_%s/archive" % c["id"], json={}, timeout=30)
                assert resp.status_code == 403, "项目归档越权返回 %s（期望 403）" % resp.status_code
                msgs.append("项目合同 编号%s 非经办人归档→403" % c["编号"])
            else:
                msgs.append("项目无可用样本(跳过)")
            a = None
            for x in C.query(db, "SELECT id,资产名称,责任人,是否归档 FROM asset_contracts ORDER BY id"):
                if str(x["是否归档"] or "") != "是" and PU["u"]["姓名"] not in str(x["责任人"] or ""):
                    a = x
                    break
            if a:
                resp = PU["s"].put(B + "/api/asset_contracts/arow_%s/archive" % a["id"], json={}, timeout=30)
                assert resp.status_code == 403, "资产归档越权返回 %s（期望 403）" % resp.status_code
                msgs.append("资产合同 %s 非责任人归档→403" % a["id"])
            else:
                msgs.append("资产无可用样本(跳过)")
            return "；".join(msgs)

        r.case("TC-M03-07", "归档限管理员或经办人/责任人", "非授权者 403", t07)

        # ── TC-M03-08 附件越权下载 ──
        def t08():
            un = PU["u"]["用户名"]
            vis = C.visible_depts_of(db, un)
            atts = C.query(db, "SELECT 附件ID,合同编号 FROM project_attachments")
            target = None
            for a in atts:
                rows = C.query(db, "SELECT 所属部门 FROM project_contracts WHERE 编号=?", (a["合同编号"],))
                dept = str(rows[0]["所属部门"]).strip() if rows else ""
                if vis is not None and dept not in vis:
                    target = a
                    break
            assert target, "无可越权下载的附件样本"
            resp = PU["s"].get(B + "/api/attachments/%s/download" % target["附件ID"], timeout=30)
            assert resp.status_code == 403, "越权下载附件返回 %s（期望 403）" % resp.status_code
            return "附件 %s（合同 %s）越权下载 → 403" % (target["附件ID"], target["合同编号"])

        r.case("TC-M03-08", "越权下载他人合同附件", "403", t08)

        # ── TC-M03-09 导出范围一致 ──
        def t09():
            un = PU["u"]["用户名"]
            lst = PU["s"].get(B + "/api/contracts", params={"page": 1, "page_size": 1}, timeout=60).json()
            resp = PU["s"].get(B + "/api/contracts/export", timeout=120)
            assert resp.status_code == 200, "导出 HTTP %s" % resp.status_code
            wb = openpyxl.load_workbook(BytesIO(resp.content))
            ws = wb.active
            n = ws.max_row - 1
            assert n == lst["total"], "导出行数 %d 与列表 total %d 不一致" % (n, lst["total"])
            return "%s 导出 %d 行 = 列表 total %d" % (un, n, lst["total"])

        r.case("TC-M03-09", "导出行数与列表口径一致", "行数=total", t09)

        # ── TC-M03-10 用户列表可见性 ──
        def t10():
            total = C.query(db, "SELECT COUNT(*) c FROM users")[0]["c"]
            ja = AU["s"].get(B + "/api/users", timeout=30).json()
            jp = PU["s"].get(B + "/api/users", timeout=30).json()
            assert len(ja["items"]) == total, "管理员看到 %d 个用户，库内 %d" % (len(ja["items"]), total)
            assert len(jp["items"]) == 1 and jp["items"][0]["username"] == PU["u"]["用户名"], \
                "普通用户看到 %d 个用户" % len(jp["items"])
            return "管理员 %d 条；普通用户仅本人 1 条" % len(ja["items"])

        r.case("TC-M03-10", "用户列表可见性", "管理员全量 / 其他仅本人", t10)

        # ── TC-M03-11 仅管理员接口 ──
        def t11():
            s = PU["s"]
            checks = [
                ("GET 项目模板", s.get(B + "/api/contracts/template", timeout=30)),
                ("GET 资产模板", s.get(B + "/api/asset_contracts/template", timeout=30)),
                ("POST 导入预览", s.post(B + "/api/contracts/import/preview", files={"file": ("a.xlsx", b"x")}, timeout=30)),
                ("PUT 配置", s.put(B + "/api/config", json={"default_remind_days": 60}, timeout=30)),
                ("GET 日志导出", s.get(B + "/api/logs/export", timeout=30)),
                ("POST 新增用户", s.post(B + "/api/users", json={"username": "t", "name": "t", "department": "t"}, timeout=30)),
            ]
            bad = ["%s=%s" % (n, resp.status_code) for n, resp in checks if resp.status_code != 403]
            assert not bad, "普通用户访问管理员接口未被拒绝：%s" % bad
            return "6 个管理员专属接口对普通用户均 403"

        r.case("TC-M03-11", "管理员专属接口拦截", "普通用户 403", t11)

        # ── TC-M03-12 部门接口权限 ──
        def t12():
            s = PU["s"]
            checks = [
                ("POST", s.post(B + "/api/departments", json={"unit": "u", "dept": "d"}, timeout=30)),
                ("PUT", s.put(B + "/api/departments", json={"unit": "u", "dept": "d", "new_unit": "u", "new_dept": "d2"}, timeout=30)),
                ("DELETE", s.delete(B + "/api/departments", params={"unit": "u", "dept": "d"}, timeout=30)),
            ]
            bad = ["%s=%s" % (n, resp.status_code) for n, resp in checks if resp.status_code != 403]
            assert not bad, "普通用户部门写操作未被拒绝：%s" % bad
            return "部门 增/改/删 对普通用户均 403"

        r.case("TC-M03-12", "部门管理仅管理员", "普通用户 403", t12)

        # ── TC-M03-13 修改自身信息 ──
        def t13():
            un = PU["u"]["用户名"]
            other = AU["u"]["用户名"]
            r1 = PU["s"].put(B + "/api/users/%s/self" % un, json={}, timeout=30)
            r2 = PU["s"].put(B + "/api/users/%s/self" % other, json={"new_password": "x"}, timeout=30)
            assert r1.status_code != 403, "本人修改自身返回 %s（应可达）" % r1.status_code
            assert r2.status_code == 403, "修改他人返回 %s（期望 403）" % r2.status_code
            return "本人空提交=%s（非 403）；改他人=403" % r1.status_code

        r.case("TC-M03-13", "仅可修改自身信息", "本人可达 / 他人 403", t13)

        # ── TC-M03-14 admin 保护 ──
        def t14():
            resp = AU["s"].delete(B + "/api/users/admin", timeout=30)
            assert resp.status_code in (400, 403), "删除 admin 返回 %s" % resp.status_code
            return "DELETE /api/users/admin → %s %s" % (resp.status_code, resp.json().get("error", ""))

        r.case("TC-M03-14", "默认管理员不可删除", "400/403", t14)

        srv.sync_db()
    finally:
        srv.stop()

    prod_after = C.prod_fingerprint()
    meta = {
        "服务地址": B,
        "工作库": "test_data/m03_perm.db（副本自 %s）" % C.base_info().get("src", ""),
        "账号": "admin / %s / %s" % (PU["u"]["用户名"], QU["u"]["用户名"]),
        "生产库未改动": "是" if prod_before == prod_after else "否（指纹变化！）",
    }
    r.report("M03_权限矩阵测试.md", meta=meta)
    n, p, f, e = r.stat()
    print("\n[SUMMARY] M03 共 %d 例：通过 %d，失败 %d，错误 %d" % (n, p, f, e))
    sys.stdout.flush()
    os._exit(0 if (f == 0 and e == 0) else 1)


if __name__ == "__main__":
    main()
