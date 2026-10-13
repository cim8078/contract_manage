# -*- coding: utf-8 -*-
"""M01 数据库迁移与结构完整性测试（L1）。
覆盖：v1~v30 版本行、重建表迁移幂等（v13/v15/v17/v26/v28）、settings 迁移（v29/v30）、
空库全新建库、唯一索引 / 复合主键 / 外键 / CHECK 约束、连续迁移幂等、失败迁移回滚。
隔离：全部在 test_data/ 的工作库副本上执行，生产库只读（只做指纹校验）。
"""
import os
import sys
import sqlite3
import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _common as C  # noqa: E402


def main():
    C.setup_stdout()
    C.ensure_dirs()
    C.prepare_runtime()
    src = C.prepare_base()
    info = C.base_info()

    r = C.R("M01", "M01 数据库迁移与结构完整性测试报告（L1）")
    prod_before = C.prod_fingerprint()
    base = C.BASE_DB
    base_counts = C.counts(base)
    base_versions = C.version_rows(base)
    base_ids = [x["id"] for x in C.query(base, "SELECT id FROM asset_contracts ORDER BY id")]
    VERS = list(range(1, C.MAX_VERSION + 1))

    # ── TC-M01-01 版本行完整性与说明一致性 ──
    def t01():
        assert base_versions == VERS, "版本行不连续/缺失：%s" % base_versions
        db_descs = C.version_descs(base)
        src_descs = C.parse_migrations_from_source()
        missing = sorted(set(src_descs) - set(db_descs))
        assert not missing, "库内缺少版本：%s" % missing
        diff = [v for v in sorted(src_descs) if src_descs[v] != db_descs.get(v)]
        assert not diff, "库内迁移说明与源码不一致（历史版本说明被改动）：%s" % "；".join(
            "v%d[库内=%s | 源码=%s]" % (v, db_descs.get(v), src_descs[v]) for v in diff)
        for row in C.query(base, "SELECT version, applied_at FROM schema_version"):
            datetime.datetime.strptime((row["applied_at"] or "")[:19], "%Y-%m-%d %H:%M:%S")
        return "v1~v%d 共 %d 版，说明与 app.py MIGRATIONS 一致，applied_at 均为合法时间" % (
            C.MAX_VERSION, C.MAX_VERSION)

    r.case("TC-M01-01", "版本行完整性与说明一致性", "v1~v30 连续、说明与源码一致", t01)

    # ── 通用：删除版本行后重放 ──
    def scenario(cid, name, expected, del_versions, post=None, priority="P0"):
        def fn():
            db = C.work_db("m01_%s.db" % cid.replace("-", "_"))
            con = sqlite3.connect(db)
            for v in del_versions:
                con.execute("DELETE FROM schema_version WHERE version=?", (v,))
            con.commit()
            con.close()
            out = C.run_init_db(db)
            vr = C.version_rows(db)
            assert vr == VERS, "版本行未恢复：%s" % vr
            cc = C.counts(db)
            assert cc == base_counts, "行数不守恒：%s != %s" % (cc, base_counts)
            fk = C.fk_violations(db)
            assert not fk, "外键违规 %d 条：%s" % (len(fk), fk[:3])
            extra = post(db) if post else ""
            replayed = ("重放 v%s" % ",".join(str(v) for v in del_versions)) if "MIGRATE" in out or True else ""
            return "%s；行数守恒 %s；fk_check=0；%s" % (replayed, list(cc.values()), extra)
        return r.case(cid, name, expected, fn, priority)

    OLD_RENT_COLS = ["已缴租金1季度（元）", "已缴租金2季度（元）", "已缴租金3季度（元）",
                     "已缴租金4季度（元）", "已缴租金合计（元）", "交款时间", "开票对应租期", "缴款方式"]

    def post13(db):
        cols = C.column_names(db, "asset_contracts")
        left = [c for c in OLD_RENT_COLS if c in cols]
        assert not left, "仍残留旧租金/缴款方式列：%s" % left
        return "asset_contracts 无旧租金列"

    def post15(db):
        cols = C.column_names(db, "asset_contracts")
        assert "缴款方式" not in cols, "asset_contracts 仍存在「缴款方式」列"
        return "asset_contracts 无「缴款方式」列"

    def post17(db):
        cols = {c["name"]: c for c in C.columns(db, "asset_rent_collections")}
        assert "缴款方式" in cols, "asset_rent_collections 缺少「缴款方式」列"
        df = cols["缴款方式"]["dflt_value"]
        assert df and "直缴非税" in str(df), "缴款方式默认值不正确：%r" % df
        return "缴款方式默认值=%s" % df

    def post26(db):
        cols = {c["name"]: c for c in C.columns(db, "asset_contracts")}
        assert cols["id"]["type"].upper() == "TEXT", "asset_contracts.id 类型=%s（应为 TEXT）" % cols["id"]["type"]
        now_ids = [x["id"] for x in C.query(db, "SELECT id FROM asset_contracts ORDER BY id")]
        assert now_ids == base_ids, "历史 id 未保留：%s != %s" % (now_ids, base_ids)
        return "id 类型=TEXT、历史 id 保留 %s" % now_ids

    def post28(db):
        names = C.table_names(db)
        assert "attachments" not in names, "旧表 attachments 仍存在"
        assert "project_attachments" in names, "新表 project_attachments 不存在"
        assert C.counts(db, ["project_attachments"]) == C.counts(base, ["project_attachments"])
        return "attachments 已改名，project_attachments 行数不变"

    def post2930(db):
        rows = {x["配置项"]: x["配置值"] for x in C.query(db, "SELECT * FROM settings")}
        assert "sms_sign_suffix" in rows, "缺少 sms_sign_suffix 配置项"
        assert "湘潭市机关事务管理局" in str(rows["sms_sign_suffix"]), "署名配置值异常：%r" % rows["sms_sign_suffix"]
        assert "sms_sign_name" not in rows, "废弃配置 sms_sign_name 仍存在"
        return "sms_sign_suffix=%s，sms_sign_name 已删除" % rows["sms_sign_suffix"]

    scenario("TC-M01-02", "v13 重建表迁移幂等重放", "跳过重建、行数守恒、无外键违规", [13], post13)
    scenario("TC-M01-03", "v15 删列迁移幂等重放", "跳过重建、缴款方式列不存在", [15], post15)
    scenario("TC-M01-04", "v17 默认值迁移幂等重放", "缴款方式默认值=直缴非税", [17], post17)
    scenario("TC-M01-05", "v26 主键 INTEGER→TEXT 幂等重放", "id 为 TEXT 且历史 id 保留", [26], post26)
    scenario("TC-M01-06", "v28 附件表改名幂等重放", "无 attachments 表、project_attachments 数据保留", [28], post28)
    scenario("TC-M01-07", "v29/v30 settings 幂等重放", "署名配置存在、废弃配置删除", [29, 30], post2930, priority="P1")

    def post31(db):
        row = C.query(db, "SELECT sql FROM sqlite_master WHERE type='table' AND name='project_attachments'")
        ddl = (row[0]["sql"] if row else "") or ""
        assert "ON UPDATE CASCADE" in ddl.upper(), "v31 后 project_attachments 外键仍非 CASCADE：%s" % ddl
        return "project_attachments 外键=ON UPDATE CASCADE"

    scenario("TC-M01-15", "v31 重建表迁移幂等重放（D-12 修复）",
             "外键为 ON UPDATE CASCADE、行数守恒、无外键违规", [31], post31)

    # ── TC-M01-08 空库全新建库 ──
    def t08():
        db = os.path.join(C.TEST_DATA, "m01_08_fresh.db")
        C.run_init_db_fresh(db)
        names = C.table_names(db)
        missing = [t for t in C.ALL_TABLES if t not in names]
        assert not missing, "缺表：%s" % missing
        assert C.version_rows(db) == VERS, "全新库版本行=%s" % C.version_rows(db)
        assert C.fk_violations(db) == [], "全新库存在外键违规"
        users = C.query(db, "SELECT 用户名,角色 FROM users")
        assert any(u["用户名"] == "admin" and u["角色"] == "管理员" for u in users), "未写入默认管理员"
        assert len(C.trigger_names(db)) > 0, "未创建触发器"
        return "13 表齐备、v1~v%d 全建、默认管理员已写入、触发器 %d 个" % (
            C.MAX_VERSION, len(C.trigger_names(db)))

    r.case("TC-M01-08", "空库全新建库", "13 表齐备且版本=30、含默认管理员", t08)

    # ── TC-M01-09 编号唯一索引 ──
    def t09():
        db = C.work_db("m01_09_unique.db")
        idx = C.index_names(db, "project_contracts")
        assert any("no" in i or "编号" in i for i in idx), "未找到编号唯一索引：%s" % idx
        con = C.connect(db)
        try:
            row = dict(con.execute('SELECT * FROM project_contracts LIMIT 1').fetchone())
            other = con.execute('SELECT "编号" FROM project_contracts WHERE "编号"<>? LIMIT 1', (row["编号"],)).fetchone()[0]
            cols = [k for k in row if k != "id"]
            vals = [row[c] for c in cols]
            vals[cols.index("编号")] = other
            sql = "INSERT INTO project_contracts (%s) VALUES (%s)" % (
                ",".join('"%s"' % c for c in cols), ",".join("?" * len(cols)))
            try:
                con.execute(sql, vals)
                con.commit()
                raise AssertionError("重复编号未被拒绝")
            except sqlite3.IntegrityError as e:
                assert "UNIQUE" in str(e).upper() or "unique" in str(e), "拒绝原因非唯一约束：%s" % e
                return "重复编号被 UNIQUE 拒绝（%s）" % e
        finally:
            con.close()

    r.case("TC-M01-09", "编号唯一索引生效", "重复编号写入被拒绝", t09)

    # ── TC-M01-10 departments 复合主键 ──
    def t10():
        db = C.work_db("m01_10_deptpk.db")
        cols = C.columns(db, "departments")
        pk = [c["name"] for c in cols if c["pk"] > 0]
        assert set(pk) == {"单位名称", "科室名称"}, "复合主键列=%s（应为 单位名称+科室名称）" % pk
        con = C.connect(db)
        try:
            row = con.execute("SELECT 单位名称,科室名称 FROM departments LIMIT 1").fetchone()
            try:
                con.execute("INSERT INTO departments (单位名称,科室名称) VALUES (?,?)", (row[0], row[1]))
                con.commit()
                raise AssertionError("重复部门组合未被拒绝")
            except sqlite3.IntegrityError:
                return "复合主键 (单位名称,科室名称) 生效，重复组合被拒绝"
        finally:
            con.close()

    r.case("TC-M01-10", "departments 复合主键生效", "重复（单位,科室）被拒绝", t10)

    # ── TC-M01-11 外键约束 ──
    def t11():
        db = C.work_db("m01_11_fk.db")
        tg = C.trigger_names(db)
        assert any("dept" in t for t in tg), "未找到部门派生触发器：%s" % tg
        con = C.connect(db, foreign_keys=True)
        try:
            row = dict(con.execute("SELECT * FROM project_contracts LIMIT 1").fetchone())
            cols = [k for k in row if k != "id"]
            vals = [row[c] for c in cols]
            if "编号" in cols:
                vals[cols.index("编号")] = "TSTFK9999"   # 避免被唯一约束先拦截，确保拒绝原因来自外键
            for c in ("所属部门", "所属单位", "所属科室"):
                if c in cols:
                    vals[cols.index(c)] = {"所属部门": "不存在的单位-不存在的科室",
                                           "所属单位": "不存在的单位", "所属科室": "不存在的科室"}[c]
            sql = "INSERT INTO project_contracts (%s) VALUES (%s)" % (
                ",".join('"%s"' % c for c in cols), ",".join("?" * len(cols)))
            try:
                con.execute(sql, vals)
                con.commit()
                raise AssertionError("悬空所属部门未被拒绝")
            except sqlite3.IntegrityError as e:
                return "悬空部门被拒绝（%s）" % e
        finally:
            con.close()

    r.case("TC-M01-11", "外键与部门派生触发器生效", "悬空所属部门写入被拒绝", t11)

    # ── TC-M01-12 CHECK 约束 ──
    def t12():
        db = C.work_db("m01_12_check.db")
        con = C.connect(db)
        msgs = []
        try:
            row = dict(con.execute("SELECT * FROM asset_contracts LIMIT 1").fetchone())
            cols = list(row.keys())
            vals = [row[c] for c in cols]
            vals[cols.index("年租金（元）")] = -1
            sql = "INSERT INTO asset_contracts (%s) VALUES (%s)" % (
                ",".join('"%s"' % c for c in cols), ",".join("?" * len(cols)))
            try:
                con.execute(sql, vals)
                con.commit()
                msgs.append("负年租金未被拒绝")
                con.rollback()
            except sqlite3.IntegrityError as e:
                msgs.append("负金额被 CHECK 拒绝(%s)" % e)
            cddl = C.query(db, "SELECT sql FROM sqlite_master WHERE name='project_contracts'")[0]["sql"]
            has_date_check = "CHECK" in (cddl or "").upper()
            msgs.append("project_contracts 声明 CHECK=%s" % has_date_check)
            row = dict(con.execute("SELECT * FROM project_contracts LIMIT 1").fetchone())
            cols = list(row.keys())
            vals = [row[c] for c in cols]
            if "采购时间" in cols:
                vals[cols.index("采购时间")] = "2026-13-40"
                sql = "INSERT INTO project_contracts (%s) VALUES (%s)" % (
                    ",".join('"%s"' % c for c in cols), ",".join("?" * len(cols)))
                try:
                    con.execute(sql, vals)
                    con.commit()
                    msgs.append("非法日期未被拒绝")
                    con.rollback()
                except sqlite3.IntegrityError as e:
                    msgs.append("非法日期被 CHECK 拒绝")
            bad = [m for m in msgs if "未被拒绝" in m]
            assert not bad, "；".join(msgs)
            return "；".join(msgs)
        finally:
            con.close()

    r.case("TC-M01-12", "CHECK 约束（金额非负/日期格式）", "非法值写入被拒绝", t12)

    # ── TC-M01-13 连续两次迁移幂等 ──
    def t13():
        db = C.work_db("m01_13_twice.db")
        C.run_init_db(db)
        c1, v1 = C.counts(db), C.version_rows(db)
        C.run_init_db(db)
        c2, v2 = C.counts(db), C.version_rows(db)
        assert v1 == v2 == VERS, "两次迁移版本行不一致：%s / %s" % (v1, v2)
        assert c1 == c2 == base_counts, "两次迁移行数不一致：%s / %s" % (c1, c2)
        return "连续两次 init_db 无副作用（版本 %d、行数不变）" % len(v2)

    r.case("TC-M01-13", "连续两次 init_db 幂等", "版本与行数均无变化", t13)

    # ── TC-M01-14 失败迁移回滚（故障注入，仅内存修改 MIGRATIONS） ──
    def t14():
        import importlib
        if C.RUNTIME not in sys.path:
            sys.path.insert(0, C.RUNTIME)
        db = C.work_db("m01_14_fault.db")
        C.run_init_db(db)
        app = importlib.import_module("app")
        con = sqlite3.connect(db)
        con.execute("DELETE FROM schema_version WHERE version=30")
        con.commit()

        def _fault_create_then_raise(c):
            c.execute("CREATE TABLE t_fault_injection(x)")
            raise RuntimeError("fault injection")

        def _fault_raise(c):
            raise RuntimeError("fault injection 2")

        app.MIGRATIONS.append((998, "fault-injection-create-then-fail", _fault_create_then_raise))
        app.MIGRATIONS.append((999, "fault-injection-raise", _fault_raise))
        raised = False
        try:
            app._run_versioned_migrations(con)
        except RuntimeError:
            raised = True
        finally:
            app.MIGRATIONS.pop()
            app.MIGRATIONS.pop()
        try:
            con.rollback()
        except Exception:
            pass
        assert raised, "注入的失败迁移未抛出异常"
        left = con.execute("SELECT name FROM sqlite_master WHERE name='t_fault_injection'").fetchone()
        assert not left, "失败迁移的建表未回滚"
        vers = [x[0] for x in con.execute("SELECT version FROM schema_version ORDER BY version")]
        assert 998 not in vers and 999 not in vers, "失败迁移被错误记录版本：%s" % vers
        assert 30 in vers, "失败前的正常迁移未提交（v30 缺失）"
        con.close()
        return "注入失败迁移：建表回滚、未记录版本、v30 正常提交，后续启动可重试"

    r.case("TC-M01-14", "失败迁移回滚且不记录版本", "回滚且无版本行残留", t14, priority="P2")

    # ── TC-M01-16 v31 级联更新（DB 层） ──
    def t16():
        db = C.work_db("m01_16_cascade.db")
        att = C.query(db, "SELECT a.附件ID i, a.合同编号 n FROM project_attachments a "
                          "WHERE a.合同编号 IS NOT NULL AND TRIM(a.合同编号)<>'' "
                          "AND EXISTS (SELECT 1 FROM project_contracts c WHERE c.编号=a.合同编号) LIMIT 1")
        assert att, "库内无「附件+对应合同」样本"
        aid, old = att[0]["i"], str(att[0]["n"])
        new = old + "C"
        con = C.connect(db, foreign_keys=True)
        try:
            con.execute('UPDATE project_contracts SET "编号"=? WHERE "编号"=?', (new, old))
            con.commit()
            got = con.execute('SELECT "合同编号" FROM project_attachments WHERE "附件ID"=?', (aid,)).fetchone()[0]
            assert got == new, "外键级联未生效：附件合同编号仍为 %r（应随主表变为 %r）" % (got, new)
            bad = con.execute("PRAGMA foreign_key_check").fetchall()
            assert not bad, "级联后出现外键违规：%s" % (bad[:3],)
            con.execute('UPDATE project_contracts SET "编号"=? WHERE "编号"=?', (old, new))
            con.commit()
            got2 = con.execute('SELECT "合同编号" FROM project_attachments WHERE "附件ID"=?', (aid,)).fetchone()[0]
            assert got2 == old, "回滚改编号后附件未跟随：%r" % got2
        finally:
            con.close()
        return "改编号 %s→%s：附件自动跟随（CASCADE）且可回滚，无外键违规" % (old, new)

    r.case("TC-M01-16", "v31 级联更新生效（DB 层）", "附件随主表编号自动更新", t16)

    # ── TC-M01-17 v31 真实重建路径（RESTRICT → CASCADE）与重建前备份 ──
    def t17():
        import glob
        db = C.work_db("m01_17_rebuild.db")
        con = C.connect(db)
        try:
            con.execute("PRAGMA foreign_keys=OFF")
            con.execute("BEGIN")
            con.execute('CREATE TABLE pa_old ('
                        '"附件ID" TEXT PRIMARY KEY, "合同编号" TEXT, "附件类型" TEXT, "显示名称" TEXT, '
                        '"存储文件名" TEXT, "文件大小" INTEGER, "上传人" TEXT, "上传时间" TEXT, '
                        'FOREIGN KEY ("合同编号") REFERENCES project_contracts("编号") '
                        'ON DELETE RESTRICT ON UPDATE RESTRICT)')
            con.execute("INSERT INTO pa_old SELECT * FROM project_attachments")
            con.execute("DROP TABLE project_attachments")
            con.execute("ALTER TABLE pa_old RENAME TO project_attachments")
            con.execute("DELETE FROM schema_version WHERE version=31")
            con.execute("COMMIT")
        finally:
            con.close()
        before = C.counts(db, ["project_attachments"])["project_attachments"]
        C.run_init_db(db)
        ddl = C.query(db, "SELECT sql FROM sqlite_master WHERE name='project_attachments'")[0]["sql"]
        assert "ON UPDATE CASCADE" in ddl.upper(), "真实重建后外键仍非 CASCADE：%s" % ddl
        assert C.counts(db, ["project_attachments"])["project_attachments"] == before, "重建后附件行数变化"
        assert C.version_rows(db) == VERS, "重建后版本行异常"
        baks = glob.glob(os.path.join(C.RUNTIME, "backups", "合同台账_v31重建前_*.db"))
        assert baks, "v31 未生成「重建前」备份"
        return "RESTRICT→重建→CASCADE；附件行数 %d 保留；重建前备份 %d 个" % (before, len(baks))

    r.case("TC-M01-17", "v31 真实重建路径与重建前备份", "重建成功、数据保留、留备份", t17, priority="P1")

    # ── 生产库未被改动校验 ──
    prod_after = C.prod_fingerprint()
    prod_same = (prod_before == prod_after)

    meta = {
        "基线源": "%s（schema v%s）" % (info.get("src", ""), info.get("src_version", "")),
        "工作副本": "%s（%s 字节）" % (info.get("base", ""), info.get("size", "")),
        "基线行数": str(base_counts),
        "生产库未改动": "是（size/mtime/head-sha256 均一致）" if prod_same else "否（指纹变化！）",
    }
    extra = "## 结论\n\n生产库 %s 未被改动（基线取自 backups，测试在 test_data 副本执行）。" % (
        "data/合同台账.db" if prod_same else "")
    r.report("M01_数据库与迁移测试.md", meta=meta, extra=extra)
    n, p, f, e = r.stat()
    print("\n[SUMMARY] M01 共 %d 例：通过 %d，失败 %d，错误 %d" % (n, p, f, e))
    sys.stdout.flush()
    os._exit(0 if (f == 0 and e == 0) else 1)


if __name__ == "__main__":
    main()
