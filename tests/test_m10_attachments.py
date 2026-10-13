# -*- coding: utf-8 -*-
"""M10 附件上传/下载/删除测试（L5）。
覆盖：合法格式（图片/PDF）上传、非法扩展名、大小上限（19M/20M/21M 等价边界）、无扩展名与双扩展名、
项目/资产附件类型白名单、资产「现状图片」上传 PDF、下载内容哈希一致、越权下载 403、
删除（记录+物理文件）、不存在 404、查询用户 403、归档后 403、合同删除级联清理。
说明：大小边界通过 POST /api/config 把 default_attach_max_mb 调为 1M 后用 0.9M/1.0M/1.1M 文件验证（等价路径），用完恢复。
"""
import os
import sys
import hashlib
import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _common as C  # noqa: E402

RUNTIME_ATTACH = None


def main():
    global RUNTIME_ATTACH
    C.setup_stdout()
    C.ensure_dirs()
    C.prepare_runtime()
    C.prepare_base()
    db = C.work_db("m10_attach.db")
    RUNTIME_ATTACH = os.path.join(C.RUNTIME, "data", "attachments")
    prod_before = C.prod_fingerprint()
    r = C.R("M10", "M10 附件上传下载测试报告（L5）")
    srv = C.Server(db)
    srv.start()
    B = srv.base
    L = C.RUNTIME_DB
    stamp = datetime.datetime.now().strftime("%H%M%S")
    try:
        A, ia = C.login("admin", base=B)
        assert A, "admin 登录失败"
        q_un = C.query(L, "SELECT 用户名 FROM users WHERE 角色='查询用户' ORDER BY 用户名 LIMIT 1")[0]["用户名"]
        Q, iq = C.login(q_un, base=B)
        assert Q, "查询用户登录失败"
        p_un = C.query(L, "SELECT 用户名 FROM users WHERE 角色='普通用户' ORDER BY 用户名 LIMIT 1")[0]["用户名"]
        P, ip = C.login(p_un, base=B)
        assert P, "普通用户登录失败"
        ORIG_MAX = C.query(L, "SELECT 配置值 v FROM settings WHERE 配置项='default_attach_max_mb'")

        def body(size, head=b"\x89PNG\r\n\x1a\n"):
            return head + os.urandom(max(0, size - len(head)))

        def up_project(cid, name, data=None, atype="decision", sess=None):
            return (sess or A).post(B + "/api/contracts/%s/attachments" % cid,
                                    files={"file": (name, data if data is not None else body(1024))},
                                    data={"type": atype}, timeout=120)

        def up_asset(aid, name, data=None, atype="condition", sess=None):
            return (sess or A).post(B + "/api/asset_contracts/%s/attachments" % aid,
                                    files={"file": (name, data if data is not None else body(1024))},
                                    data={"type": atype}, timeout=120)

        def new_project(**kw):
            b = {"category": "服务类", "name": "M10附件-" + stamp, "procurement_method": "自行采购",
                 "amount": 1, "main_content": "附件测试", "handler": ia["name"], "fund_source": "其他"}
            b.update(kw)
            return A.post(B + "/api/contracts", json=b, timeout=60).json()["number"]

        def new_asset(**kw):
            b = {"number": "993%03d" % (len(C.query(L, "SELECT id FROM asset_contracts")) + 1),
                 "place": "附件测试位置-" + stamp, "asset_name": "附件测试资产-" + stamp,
                 "tenant": "承租人", "dept": ia["department"]}
            b.update(kw)
            resp = A.post(B + "/api/asset_contracts", json=b, timeout=60)
            assert resp.status_code == 201, "创建资产失败 %s" % resp.status_code
            return resp.json()["number"]

        def row_of(no):
            rows = C.query(L, "SELECT id FROM project_contracts WHERE 编号=?", (no,))
            return rows[0]["id"] if rows else None

        def attach_rows(number):
            return C.query(L, "SELECT * FROM project_attachments WHERE 合同编号=?", (number,))

        def disk_path(store):
            return os.path.join(RUNTIME_ATTACH, store)

        def t01():
            no = new_project()
            cid = "row_%s" % row_of(no)
            r1 = up_project(cid, "图A.png", atype="decision")
            r2 = up_project(cid, "文B.pdf", atype="contract")
            assert r1.status_code == 201, "png 上传 %s %s" % (r1.status_code, r1.text[:120])
            assert r2.status_code == 201, "pdf 上传 %s %s" % (r2.status_code, r2.text[:120])
            rows = attach_rows(no)
            assert len(rows) == 2, "库内附件 %d 条" % len(rows)
            store = rows[-1]["存储文件名"]
            assert os.path.exists(disk_path(store)), "物理文件未落盘：%s" % store
            return "png/pdf 上传 201；库内 %d 条、物理文件已落盘（%s）" % (len(rows), store)

        r.case("TC-M10-01", "合法格式上传（图片/PDF）", "201 + 库记录 + 物理文件", t01)

        def t02():
            no = new_project()
            cid = "row_%s" % row_of(no)
            hits = []
            for fn in ("bad.exe", "bad.txt", "bad.zip"):
                rr = up_project(cid, fn)
                hits.append("%s=%s" % (fn, rr.status_code))
                assert rr.status_code == 400, "%s 未被拒绝：%s" % (fn, rr.status_code)
                assert "jpg" in rr.text or "pdf" in rr.text, "未回显允许列表：%s" % rr.text[:100]
            assert not attach_rows(no), "非法格式产生了附件记录"
            return "非法扩展名全部 400 且回显允许列表（%s）" % ",".join(hits)

        r.case("TC-M10-02", "非法扩展名白名单", "400 + 回显允许类型", t02)

        def t03():
            orig = ORIG_MAX[0]["v"] if ORIG_MAX else ""
            cfg = A.put(B + "/api/config", json={"default_attach_max_mb": 1}, timeout=30)
            assert cfg.status_code == 200, "设置上限失败 %s" % cfg.status_code
            try:
                no = new_project()
                cid = "row_%s" % row_of(no)
                ok = up_project(cid, "ok.png", data=body(int(0.9 * 1024 * 1024)))
                edge = up_project(cid, "edge.png", data=body(1 * 1024 * 1024))
                big = up_project(cid, "big.png", data=body(int(1.1 * 1024 * 1024)))
                assert ok.status_code == 201, "0.9M 被拒：%s" % ok.status_code
                assert edge.status_code == 201, "1.0M（等于上限）被拒：%s" % edge.status_code
                assert big.status_code == 400, "1.1M 未拒绝：%s" % big.status_code
                msg = (big.json() or {}).get("error", "")
                assert "不能超过" in msg, "超限提示不符：%r" % msg
                rows = attach_rows(no)
                assert len(rows) == 2, "超限文件产生了记录（库内 %d 条）" % len(rows)
                stores = set(x["存储文件名"] for x in rows)
                orphan = [f for f in os.listdir(RUNTIME_ATTACH)
                          if f not in stores and os.path.getsize(os.path.join(RUNTIME_ATTACH, f)) > 1024 * 1024]
                assert not orphan, "超限文件未从磁盘清理：%s" % orphan[:3]
                return "上限设为 1M：0.9M→201、1.0M(等于)→201、1.1M→400；超限文件已从磁盘清理"
            finally:
                A.put(B + "/api/config", json={"default_attach_max_mb": int(float(orig or 20))}, timeout=30)

        r.case("TC-M10-03", "大小上限边界（19M/20M/21M 等价）", "≤上限 201、超限 400 且清理", t03, priority="P1")

        def t04():
            no = new_project()
            cid = "row_%s" % row_of(no)
            noext = up_project(cid, "noextension")
            dbl = up_project(cid, "evil.exe.png")
            assert noext.status_code == 400, "无扩展名未被拒绝：%s" % noext.status_code
            return "无扩展名 → 400；双扩展名 evil.exe.png → HTTP %s（白名单按末位扩展名 .png 判定）" % dbl.status_code

        r.case("TC-M10-04", "无扩展名 / 双扩展名", "无扩展名 400；双扩展名记录实际", t04, priority="P1")

        def t05():
            no = new_project()
            cid = "row_%s" % row_of(no)
            ok = []
            for t in ("decision", "contract", "acceptance"):
                rr = up_project(cid, "%s.png" % t, atype=t)
                assert rr.status_code == 201, "type=%s 被拒 %s" % (t, rr.status_code)
                ok.append(t)
            bad = up_project(cid, "bad.png", atype="unknown")
            assert bad.status_code == 400, "非法 type 未被拒绝：%s" % bad.status_code
            types = set(x["附件类型"] for x in attach_rows(no))
            return "三种合法类型均可上传（%s）；非法 type → 400；落库类型=%s" % (",".join(ok), sorted(types))

        r.case("TC-M10-05", "项目附件类型白名单", "3 类合法 + 非法 400", t05)

        def t06():
            aid = "arow_%s" % new_asset()
            ok = []
            for t in ("condition", "decision", "contract"):
                rr = up_asset(aid, "%s.png" % t, atype=t)
                assert rr.status_code == 201, "type=%s 被拒 %s" % (t, rr.status_code)
                ok.append(t)
            bad = up_asset(aid, "bad.png", atype="nope")
            assert bad.status_code == 400, "非法 type 未被拒绝：%s" % bad.status_code
            assert "condition" in bad.text, "未回显服务端支持的类型：%s" % bad.text[:140]
            return "资产三类类型均可上传（%s）；非法 type 400 且回显支持列表" % ",".join(ok)

        r.case("TC-M10-06", "资产附件类型白名单", "3 类合法 + 非法 400 回显", t06)

        def t07():
            aid = "arow_%s" % new_asset()
            rr = up_asset(aid, "现状.pdf", atype="condition")
            assert rr.status_code == 400, \
                "「资产现状图片」类型上传 PDF 返回 %s（需求口径期望 400：现状图片应仅接受图片）" % rr.status_code
            return "现状图片 + PDF → 400"

        r.case("TC-M10-07", "现状图片类型上传 PDF 应拒绝", "400", t07, priority="P1")

        def t08():
            no = new_project()
            cid = "row_%s" % row_of(no)
            payload = body(4096)
            rr = up_project(cid, "一致性.png", data=payload)
            assert rr.status_code == 201, "上传失败 %s" % rr.status_code
            att = rr.json()["attachment"]
            aid = att.get("id") or att.get("附件ID")
            assert aid, "上传响应未返回附件ID：%s" % list(att)[:6]
            dl = A.get(B + "/api/attachments/%s/download" % aid, timeout=60)
            assert dl.status_code == 200, "下载 HTTP %s" % dl.status_code
            assert hashlib.sha256(dl.content).hexdigest() == hashlib.sha256(payload).hexdigest(), "下载内容与上传不一致"
            return "下载 200，字节哈希与上传一致（%d 字节）" % len(payload)

        r.case("TC-M10-08", "下载内容一致性", "哈希一致", t08)

        def t09():
            atts = C.query(L, "SELECT 附件ID i, 合同编号 n FROM project_attachments")
            vis = C.visible_depts_of(L, p_un)
            target = None
            for a in atts:
                rr = C.query(L, "SELECT 所属部门 d FROM project_contracts WHERE 编号=?", (a["n"],))
                if rr and str(rr[0]["d"]).strip() not in vis:
                    target = a
                    break
            assert target, "无越权附件样本"
            dl = P.get(B + "/api/attachments/%s/download" % target["i"], timeout=30)
            dl2 = P.delete(B + "/api/attachments/%s" % target["i"], timeout=30)
            assert dl.status_code == 403, "越权下载返回 %s" % dl.status_code
            assert dl2.status_code == 403, "越权删除返回 %s" % dl2.status_code
            return "越权下载/删除均 403（样本附件所属合同 %s）" % target["n"]

        r.case("TC-M10-09", "越权下载/删除附件", "403", t09)

        def t10():
            no = new_project()
            cid = "row_%s" % row_of(no)
            rr = up_project(cid, "待删除.png")
            att = rr.json()["attachment"]
            aid = att.get("id") or att.get("附件ID")
            row = [x for x in attach_rows(no) if x["附件ID"] == aid][0]
            store = row["存储文件名"]
            assert os.path.exists(disk_path(store)), "删除前文件不存在"
            dl = A.delete(B + "/api/attachments/%s" % aid, timeout=30)
            assert dl.status_code == 200, "删除 HTTP %s" % dl.status_code
            assert not [x for x in attach_rows(no) if x["附件ID"] == aid], "库内记录未删除"
            assert not os.path.exists(disk_path(store)), "物理文件未删除"
            return "删除 200：库记录与物理文件均已清理（%s）" % store

        r.case("TC-M10-10", "删除附件（记录+文件）", "200 + 双向清理", t10)

        def t11():
            g = A.get(B + "/api/attachments/%s/download" % ("f" * 32), timeout=30)
            d = A.delete(B + "/api/attachments/%s" % ("f" * 32), timeout=30)
            ga = A.get(B + "/api/asset_attachments/%s/download" % ("f" * 32), timeout=30)
            assert g.status_code == 404 and d.status_code == 404, \
                "项目附件不存在未返回 404：%s/%s" % (g.status_code, d.status_code)
            assert ga.status_code == 404, "资产附件不存在未返回 404：%s" % ga.status_code
            return "项目/资产不存在附件 → 404"

        r.case("TC-M10-11", "不存在的附件", "404", t11, priority="P1")

        def t12():
            no = new_project()
            cid = "row_%s" % row_of(no)
            q1 = up_project(cid, "q.png", sess=Q)
            assert q1.status_code == 403, "查询用户上传返回 %s" % q1.status_code
            rr = up_project(cid, "归档前.png")
            att = rr.json()["attachment"]
            aid = att.get("id") or att.get("附件ID")
            ar = A.put(B + "/api/contracts/%s/archive" % cid, json={}, timeout=30)
            assert ar.status_code == 200, "归档失败 %s" % ar.status_code
            arc = up_project(cid, "归档后.png")
            assert arc.status_code == 403, "归档后上传返回 %s" % arc.status_code
            row = [x for x in attach_rows(no) if x["附件ID"] == aid][0]
            store = row["存储文件名"]
            # 归档合同不可删除（403），级联清理另用一条未归档合同验证
            no2 = new_project(name="M10级联-" + stamp)
            cid2 = "row_%s" % row_of(no2)
            rr2 = up_project(cid2, "级联.png")
            att2 = rr2.json()["attachment"]
            aid2 = att2.get("id") or att2.get("附件ID")
            row2 = [x for x in attach_rows(no2) if x["附件ID"] == aid2][0]
            store2 = row2["存储文件名"]
            dl = A.delete(B + "/api/contracts/%s" % cid2, json={"reason": "级联清理测试"}, timeout=30)
            assert dl.status_code == 200, "删除合同 HTTP %s" % dl.status_code
            assert not attach_rows(no2), "合同删除后附件记录未清理"
            assert not os.path.exists(disk_path(store2)), "合同删除后附件文件未清理"
            return "查询用户上传 403、归档后上传 403；删除未归档合同时附件记录与文件级联清理（归档件保留 %s）" % store

        r.case("TC-M10-12", "查询用户/归档后上传 + 删除合同级联", "403；级联清理", t12)

        srv.sync_db()
    finally:
        srv.stop()

    prod_after = C.prod_fingerprint()
    meta = {
        "服务地址": B,
        "工作库": "test_data/m10_attach.db",
        "附件目录": RUNTIME_ATTACH,
        "生产库未改动": "是" if prod_before == prod_after else "否（指纹变化！）",
    }
    r.report("M10_附件上传下载测试.md", meta=meta)
    n, p, f, e = r.stat()
    print("\n[SUMMARY] M10 共 %d 例：通过 %d，失败 %d，错误 %d" % (n, p, f, e))
    sys.stdout.flush()
    os._exit(0 if (f == 0 and e == 0) else 1)


if __name__ == "__main__":
    main()
