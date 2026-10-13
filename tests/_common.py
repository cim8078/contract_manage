# -*- coding: utf-8 -*-
"""「潭合同」台账管理系统 测试公共库。
隔离红线：绝不读写生产库 data/合同台账.db（仅只读用于「未被改动」的指纹校验）。
所有写操作只发生在 test_data/ 目录内。
迁移类用例通过「运行时副本」执行真实代码路径：test_data/runtime/ 是 app.py 等文件的只读副本，
其 BASE_DIR 指向副本目录，因此副本的 data/合同台账.db 即测试库，import app 时自动执行 init_db()。
"""
import datetime
import glob
import hashlib
import os
import re
import shutil
import sqlite3
import subprocess
import sys

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TESTS_DIR = os.path.join(PROJECT_ROOT, "tests")
TEST_DATA = os.path.join(PROJECT_ROOT, "test_data")
RUNTIME = os.path.join(TEST_DATA, "runtime")
BACKUP_DIR = os.path.join(PROJECT_ROOT, "backups")
REPORT_DIR = os.path.join(PROJECT_ROOT, "docs", "test_reports")
PROD_DB = os.path.join(PROJECT_ROOT, "data", "合同台账.db")
RUNTIME_DB = os.path.join(RUNTIME, "data", "合同台账.db")
BASE_DB = os.path.join(TEST_DATA, "base_合同台账.db")
PYTHON = sys.executable

CORE_TABLES = [
    "project_contracts", "asset_contracts", "asset_rent_collections", "users",
    "departments", "project_attachments", "asset_attachments",
]
ALL_TABLES = [
    "settings", "users", "departments", "project_contracts", "project_attachments",
    "asset_contracts", "asset_attachments", "asset_rent_collections",
    "field_change_logs", "operation_logs", "sms_templates", "sms_logs", "schema_version",
]
MAX_VERSION = 31

BOOTSTRAP = (
    "import sys, os\n"
    "sys.path.insert(0, os.getcwd())\n"
    "import app\n"
    "sys.stderr.write('INIT_DB_DONE\\n')\n"
    "sys.stderr.flush()\n"
    "os._exit(0)\n"
)


def ensure_dirs():
    for d in (TEST_DATA, RUNTIME, os.path.join(RUNTIME, "data"), REPORT_DIR):
        os.makedirs(d, exist_ok=True)


def sha256(path, limit=None):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        read = 0
        while True:
            b = f.read(1024 * 1024)
            if not b:
                break
            h.update(b)
            read += len(b)
            if limit and read >= limit:
                break
    return h.hexdigest()


def latest_full_backup(min_bytes=1000000):
    """backups/ 中体积 >= min_bytes 的最新完整库（排除 A06 等部分库）。"""
    cands = []
    for p in glob.glob(os.path.join(BACKUP_DIR, "*.db")):
        try:
            if os.path.getsize(p) >= min_bytes:
                cands.append((os.path.getmtime(p), p))
        except OSError:
            pass
    if not cands:
        return None
    return max(cands)[1]


def full_backups(min_bytes=1000000):
    """backups/ 中的完整库，按修改时间倒序。"""
    cands = []
    for p in glob.glob(os.path.join(BACKUP_DIR, "*.db")):
        try:
            if os.path.getsize(p) >= min_bytes:
                cands.append((os.path.getmtime(p), p))
        except OSError:
            pass
    return [p for _, p in sorted(cands, reverse=True)]


def db_max_version(path):
    try:
        con = sqlite3.connect("file:%s?mode=ro" % path.replace("\\", "/"), uri=True)
    except sqlite3.Error:
        return None
    try:
        try:
            row = con.execute("SELECT MAX(version) FROM schema_version").fetchone()
            return row[0] if row else None
        except sqlite3.Error:
            return None
    finally:
        con.close()


_BASE_INFO = {}


def prepare_base(force=False, want_version=MAX_VERSION):
    """建立迁移基线库（只读来源，不触碰生产库）。

    取用规则：优先「schema 版本 == want_version 的最新完整备份」；若备份还没有那么新，
    则取版本最高的最新完整备份，再在本地对该副本执行一次启动迁移（init_db）升级到当前代码版本。
    """
    ensure_dirs()
    prepare_runtime()
    cands = full_backups()
    if not cands:
        raise RuntimeError("backups/ 下未找到完整备份库")
    src = None
    for p in cands:
        if db_max_version(p) == want_version:
            src = p
            break
    if src is None:
        best = max((db_max_version(p) or -1) for p in cands)
        for p in cands:
            if (db_max_version(p) or -1) == best:
                src = p
                break
    src_mtime = os.path.getmtime(src)
    src_version = db_max_version(src)
    if force or not os.path.exists(BASE_DB) or os.path.getmtime(BASE_DB) < src_mtime:
        shutil.copyfile(src, BASE_DB)
        _BASE_INFO.clear()
    base_version = db_max_version(BASE_DB)
    if base_version != want_version:
        run_init_db(BASE_DB)          # 副本内升级到当前代码版本（生产库不受影响）
        base_version = db_max_version(BASE_DB)
    if not _BASE_INFO:
        _BASE_INFO.update({
            "src": os.path.basename(src),
            "src_sha256": sha256(src, limit=4 * 1024 * 1024),
            "src_version": src_version,
            "base": os.path.basename(BASE_DB),
            "base_version": base_version,
            "base_sha256": sha256(BASE_DB, limit=4 * 1024 * 1024),
            "size": os.path.getsize(BASE_DB),
        })
    return src


def base_info():
    return dict(_BASE_INFO)


def prepare_runtime(force=False):
    """把代码只读复制到 test_data/runtime/（副本的 data/ 即测试库）。"""
    ensure_dirs()
    for fn in ("app.py", "contract_no.py"):
        src = os.path.join(PROJECT_ROOT, fn)
        dst = os.path.join(RUNTIME, fn)
        if force or not os.path.exists(dst) or os.path.getmtime(src) > os.path.getmtime(dst):
            shutil.copyfile(src, dst)
    for d in ("templates", "static"):
        src = os.path.join(PROJECT_ROOT, d)
        dst = os.path.join(RUNTIME, d)
        if force and os.path.isdir(dst):
            shutil.rmtree(dst)
        if os.path.isdir(src) and not os.path.isdir(dst):
            shutil.copytree(src, dst)
    os.makedirs(os.path.join(RUNTIME, "data"), exist_ok=True)
    os.makedirs(os.path.join(RUNTIME, "backups"), exist_ok=True)


def work_db(name):
    """由基线库派生一份工作库。"""
    ensure_dirs()
    if not os.path.exists(BASE_DB):
        prepare_base()
    dst = os.path.join(TEST_DATA, name)
    shutil.copyfile(BASE_DB, dst)
    for suffix in ("-wal", "-shm"):
        p = dst + suffix
        if os.path.exists(p):
            os.remove(p)
    return dst


def _bootstrap(dbpath, fresh=False, timeout=300):
    """把 dbpath 放进运行时副本并执行 init_db()（真实代码路径）。"""
    ensure_dirs()
    if fresh and os.path.exists(RUNTIME_DB):
        os.remove(RUNTIME_DB)
    elif not fresh:
        shutil.copyfile(dbpath, RUNTIME_DB)
    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"
    p = subprocess.run([PYTHON, "-c", BOOTSTRAP], cwd=RUNTIME, env=env,
                       capture_output=True, timeout=timeout)
    out = p.stdout.decode("utf-8", "replace")
    err = p.stderr.decode("utf-8", "replace")
    if p.returncode != 0:
        raise RuntimeError("init_db 执行失败 rc=%s\nSTDOUT:\n%s\nSTDERR:\n%s" % (
            p.returncode, out[-4000:], err[-4000:]))
    shutil.copyfile(RUNTIME_DB, dbpath)
    return out + err


def run_init_db(dbpath, timeout=300):
    """对已有库执行一次启动迁移（真实 app.init_db 路径）。"""
    return _bootstrap(dbpath, fresh=False, timeout=timeout)


def run_init_db_fresh(dbpath, timeout=300):
    """删除目标库后执行一次全新建库。"""
    if os.path.exists(dbpath):
        os.remove(dbpath)
    return _bootstrap(dbpath, fresh=True, timeout=timeout)


def connect(dbpath, foreign_keys=False):
    con = sqlite3.connect(dbpath, timeout=15)
    con.row_factory = sqlite3.Row
    if foreign_keys:
        con.execute("PRAGMA foreign_keys=ON")
    return con


def query(dbpath, sql, params=()):
    con = connect(dbpath)
    try:
        return [dict(r) for r in con.execute(sql, params).fetchall()]
    finally:
        con.close()


def counts(dbpath, tables=None):
    tables = tables or CORE_TABLES
    out = {}
    con = connect(dbpath)
    try:
        for t in tables:
            try:
                out[t] = con.execute('SELECT COUNT(*) FROM "%s"' % t).fetchone()[0]
            except sqlite3.Error as e:
                out[t] = "ERR:%s" % e
    finally:
        con.close()
    return out


def table_names(dbpath):
    return [r["name"] for r in query(dbpath, "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")]


def columns(dbpath, table):
    con = connect(dbpath)
    try:
        return [dict(r) for r in con.execute('PRAGMA table_info("%s")' % table).fetchall()]
    finally:
        con.close()


def column_names(dbpath, table):
    return [c["name"] for c in columns(dbpath, table)]


def version_rows(dbpath):
    return [r["version"] for r in query(dbpath, "SELECT version FROM schema_version ORDER BY version")]


def version_descs(dbpath):
    return {r["version"]: r["description"] for r in query(dbpath, "SELECT version, description FROM schema_version")}


def fk_violations(dbpath):
    return [tuple(r.values()) for r in query(dbpath, "PRAGMA foreign_key_check")]


def index_names(dbpath, table):
    return sorted(r["name"] for r in query(
        dbpath, "SELECT name FROM sqlite_master WHERE type='index' AND tbl_name=?", (table,)))


def trigger_names(dbpath):
    return sorted(r["name"] for r in query(dbpath, "SELECT name FROM sqlite_master WHERE type='trigger'"))


def parse_migrations_from_source(app_path=None):
    """从 app.py 源码解析 MIGRATIONS 列表的 (版本, 说明)，用于与库内记录比对。"""
    app_path = app_path or os.path.join(PROJECT_ROOT, "app.py")
    with open(app_path, "r", encoding="utf-8") as f:
        text = f.read()
    m = re.search(r"\nMIGRATIONS\s*=\s*\[(.*?)\n\]", text, re.S)
    if not m:
        raise RuntimeError("未能在 app.py 中定位 MIGRATIONS 列表")
    out = {}
    for line in m.group(1).splitlines():
        mm = re.match(r'\s*\(\s*(\d+)\s*,\s*"((?:[^"\\]|\\.)*)"', line)
        if mm:
            out[int(mm.group(1))] = mm.group(2)
    return out


def prod_fingerprint():
    if not os.path.exists(PROD_DB):
        return {}
    st = os.stat(PROD_DB)
    return {
        "size": st.st_size,
        "mtime": datetime.datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
        "sha256_head": sha256(PROD_DB, limit=4 * 1024 * 1024),
    }


class R:
    """极简用例运行器 + Markdown 报告生成。"""

    def __init__(self, module, title=""):
        self.module = module
        self.title = title or module
        self.rows = []

    def case(self, cid, name, expected, func, priority="P0"):
        try:
            detail = func()
            status = "PASS"
        except AssertionError as e:
            status, detail = "FAIL", str(e)
        except Exception as e:  # noqa: BLE001
            status, detail = "ERROR", "%s: %s" % (type(e).__name__, e)
        self.rows.append({"id": cid, "name": name, "expected": expected,
                          "status": status, "detail": str(detail), "priority": priority})
        print("[%s] %s %s :: %s" % (status, cid, name, detail))
        return status

    def stat(self):
        n = len(self.rows)
        p = sum(1 for r in self.rows if r["status"] == "PASS")
        f = sum(1 for r in self.rows if r["status"] == "FAIL")
        e = sum(1 for r in self.rows if r["status"] == "ERROR")
        return n, p, f, e

    def report(self, filename, meta=None, extra=""):
        ensure_dirs()
        n, p, f, e = self.stat()
        path = os.path.join(REPORT_DIR, filename)
        lines = ["# %s" % self.title, ""]
        lines.append("- **执行时间**：%s" % datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        lines.append("- **模块**：%s" % self.module)
        for k, v in (meta or {}).items():
            lines.append("- **%s**：%s" % (k, v))
        lines.append("- **结果**：共 %d 例，通过 %d，失败 %d，错误 %d（通过率 %.1f%%）" % (
            n, p, f, e, (p * 100.0 / n if n else 0)))
        lines.append("")
        lines.append("## 用例结果")
        lines.append("")
        lines.append("| 编号 | 用例 | 优先级 | 预期 | 结果 | 详情 |")
        lines.append("|---|---|---|---|---|---|")
        for r in self.rows:
            lines.append("| %s | %s | %s | %s | **%s** | %s |" % (
                r["id"], r["name"], r["priority"], r["expected"], r["status"],
                r["detail"].replace("|", "\\|").replace("\n", "<br>")))
        lines.append("")
        lines.append("## 问题清单")
        lines.append("")
        fails = [r for r in self.rows if r["status"] != "PASS"]
        if not fails:
            lines.append("无失败/错误用例。")
        else:
            for r in fails:
                lines.append("- **%s %s（%s）**：预期 %s；实际 %s" % (
                    r["id"], r["name"], r["status"], r["expected"], r["detail"]))
        if extra:
            lines.append("")
            lines.append(extra)
        lines.append("")
        with open(path, "w", encoding="utf-8", newline="\n") as fp:
            fp.write("\n".join(lines))
        print("[REPORT] %s" % path)
        return path


def setup_stdout():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


# ══════════════════════ HTTP 测试基座（M02 起使用）══════════════════════
import time  # noqa: E402
import requests  # noqa: E402

SERVER_PORT = int(os.environ.get("TM_TEST_PORT", "5099"))
BASE_URL = "http://127.0.0.1:%d" % SERVER_PORT


class Server:
    """启动/停止「运行时副本」服务：副本 app.py + 副本 data/ 测试库。"""

    def __init__(self, dbpath=None, port=None, copy_db=True):
        self.port = port or SERVER_PORT
        self.dbpath = dbpath
        self.copy_db = copy_db
        self.proc = None
        self.base = "http://127.0.0.1:%d" % self.port
        self.logpath = os.path.join(TEST_DATA, "server_%d.log" % self.port)
        self._log = None

    def start(self, timeout=90):
        ensure_dirs()
        prepare_runtime()
        if self.copy_db and self.dbpath:
            shutil.copyfile(self.dbpath, RUNTIME_DB)
        env = dict(os.environ)
        env["PYTHONIOENCODING"] = "utf-8"
        env["TM_TEST_PORT"] = str(self.port)
        self._log = open(self.logpath, "wb")
        self.proc = subprocess.Popen(
            [PYTHON, os.path.join(TESTS_DIR, "_serve.py")],
            cwd=PROJECT_ROOT, env=env, stdout=self._log, stderr=subprocess.STDOUT)
        t0 = time.time()
        while time.time() - t0 < timeout:
            if self.proc.poll() is not None:
                raise RuntimeError("测试服务提前退出（rc=%s），日志：%s" % (self.proc.returncode, self.logpath))
            try:
                resp = requests.get(self.base + "/api/current_user", timeout=2)
                if resp.status_code in (200, 401):
                    return self
            except Exception:
                pass
            time.sleep(0.4)
        raise RuntimeError("测试服务启动超时：%s" % self.logpath)

    def sync_db(self):
        """把副本库回写到工作库（服务停止后调用）。"""
        if self.dbpath and os.path.exists(RUNTIME_DB):
            shutil.copyfile(RUNTIME_DB, self.dbpath)

    def stop(self):
        if self.proc and self.proc.poll() is None:
            self.proc.terminate()
            try:
                self.proc.wait(timeout=15)
            except Exception:
                self.proc.kill()
        if self._log:
            try:
                self._log.close()
            except Exception:
                pass

    def __enter__(self):
        return self.start()

    def __exit__(self, *a):
        self.stop()


def login(username, password="123456", base=None):
    """登录并返回 (session, info)；失败返回 (None, {"status":...})。"""
    s = requests.Session()
    try:
        resp = s.post((base or BASE_URL) + "/api/login",
                      json={"username": username, "password": password}, timeout=15)
    except Exception as e:  # noqa: BLE001
        return None, {"status": "ERR", "body": str(e)}
    if resp.status_code == 200:
        return s, resp.json()
    return None, {"status": resp.status_code, "body": resp.text[:200]}


def visible_depts_of(dbpath, username):
    """按 app.py 的 _visible_depts 口径计算某用户可见部门集合；管理员返回 None（全部）。"""
    rows = query(dbpath, "SELECT 用户名,角色,所属部门,分管部门 FROM users WHERE 用户名=?", (username,))
    if not rows:
        return set()
    u = rows[0]
    if u["角色"] == "管理员":
        return None
    depts = [str(u["所属部门"] or "").strip()]
    md = str(u["分管部门"] or "").strip()
    depts += [p.strip() for p in md.replace("、", ",").replace("；", ";").split(",") if p.strip()]
    return set(depts)
