# -*- coding: utf-8 -*-
"""M15 性能与缓存测试（L7）。
判定标准（需求方）：单次请求 < 2s；并发成功率 > 99%。
覆盖：只读缓存一致性、写后缓存失效、列表/详情/dashboard/统计/导出耗时、写前整库备份开销、10 并发读。
"""
import os
import sys
import time
import statistics
import datetime
from concurrent.futures import ThreadPoolExecutor

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _common as C  # noqa: E402

LIMIT_MS = 2000.0


def main():
    C.setup_stdout()
    C.ensure_dirs()
    C.prepare_runtime()
    C.prepare_base()
    db = C.work_db("m15_perf.db")
    prod_before = C.prod_fingerprint()
    r = C.R("M15", "M15 性能与缓存测试报告（L7）")
    srv = C.Server(db)
    srv.start()
    B = srv.base
    L = C.RUNTIME_DB
    RUNTIME_BACKUPS = os.path.join(C.RUNTIME, "backups")
    stamp = datetime.datetime.now().strftime("%H%M%S")
    try:
        A, ia = C.login("admin", base=B)
        assert A, "admin 登录失败"

        def timed(fn, n=1, **kw):
            times = []
            last = None
            for _ in range(n):
                t0 = time.perf_counter()
                last = fn(**kw)
                times.append((time.perf_counter() - t0) * 1000)
            return last, times

        def stat(times):
            return {"avg": statistics.mean(times), "min": min(times), "max": max(times), "n": len(times)}

        def fmt(s):
            return "avg=%.0fms min=%.0fms max=%.0fms(n=%d)" % (s["avg"], s["min"], s["max"], s["n"])

        def t01():
            src = open(os.path.join(C.PROJECT_ROOT, "app.py"), encoding="utf-8").read()
            assert "READ_CACHE_TTL = 2.0" in src, "未找到 2s 只读缓存常量"
            assert "_invalidate_caches()" in src and src.count("_invalidate_caches()") >= 4, "写后失效调用点不足"
            # 同一时间窗内连续调用三组首页接口，结果必须完全一致（命中同一份只读快照）
            snaps = []
            for _ in range(3):
                d = A.get(B + "/api/dashboard", timeout=60).json()
                rm = A.get(B + "/api/reminders", timeout=60).json()
                pp = A.get(B + "/api/pending_payments", timeout=60).json()
                snaps.append((d["total_contracts"], d["total_amount"], rm["count"], pp["count"]))
            assert len(set(snaps)) == 1, "2s 窗口内多次读取结果不一致：%s" % snaps
            warm, wtimes = timed(lambda: A.get(B + "/api/dashboard", timeout=60), n=10)
            st = stat(wtimes)
            assert st["max"] < LIMIT_MS, "缓存态 dashboard 最慢 %.0fms 超 2s" % st["max"]
            return "只读缓存生效（TTL=2s、写后失效）：窗口内 3 次快照一致 %s；热读取 %s" % (snaps[0], fmt(st))

        r.case("TC-M15-01", "只读缓存命中与一致性", "窗口内快照一致 + 热读取 < 2s", t01, priority="P1")

        def t02():
            body = {"category": "服务类", "name": "M15缓存-" + stamp, "procurement_method": "自行采购",
                    "amount": 1, "main_content": "缓存失效测试", "handler": ia["name"],
                    "fund_source": "其他", "remark": "初始备注"}
            no = A.post(B + "/api/contracts", json=body, timeout=60).json()["number"]
            rid = C.query(L, "SELECT id FROM project_contracts WHERE 编号=?", (no,))[0]["id"]
            first = A.get(B + "/api/contracts/row_%s" % rid, timeout=30).json()["remark"]
            A.get(B + "/api/contracts", params={"page_size": 1}, timeout=30)     # 预热列表缓存
            newval = "缓存失效-" + stamp
            t0 = time.perf_counter()
            A.put(B + "/api/contracts/row_%s" % rid, json={"remark": newval}, timeout=60)
            ms = (time.perf_counter() - t0) * 1000
            after = A.get(B + "/api/contracts/row_%s" % rid, timeout=30).json()["remark"]
            lst = A.get(B + "/api/contracts", params={"no": no, "page_size": 5}, timeout=30).json()
            hit = [x for x in lst["items"] if x["number"] == no]
            assert after == newval, "写后立即读取仍为旧值（缓存未失效）：%r" % after
            assert hit and hit[0]["remark"] == newval, "列表缓存未失效：%s" % (hit[0]["remark"] if hit else "无记录")
            assert ms < LIMIT_MS, "单次写操作 %.0fms 超 2s" % ms
            return "写后立即读（无 sleep）取得新值：详情=%s、列表=%s；写耗时 %.0fms（预热值 %r）" % (
                after, hit[0]["remark"], ms, first)

        r.case("TC-M15-02", "写操作后缓存立即失效", "立即读到新值", t02, priority="P0")

        def t03():
            out = []
            for ps in (20, 200):
                _, ts = timed(lambda: A.get(B + "/api/contracts", params={"page": 1, "page_size": ps}, timeout=60), n=5)
                s = stat(ts)
                out.append("page_size=%d %s" % (ps, fmt(s)))
                assert s["max"] < LIMIT_MS, "列表(page_size=%d) 最慢 %.0fms 超 2s" % (ps, s["max"])
            return "项目合同列表（673 行）%s" % "；".join(out)

        r.case("TC-M15-03", "列表分页响应时间", "单次 < 2s", t03, priority="P1")

        def t04():
            rid = C.query(L, "SELECT id FROM project_contracts ORDER BY id LIMIT 1")[0]["id"]
            _, ts = timed(lambda: A.get(B + "/api/contracts/row_%s" % rid, timeout=30), n=10)
            s = stat(ts)
            assert s["max"] < LIMIT_MS, "详情最慢 %.0fms 超 2s" % s["max"]
            return "合同详情 %s" % fmt(s)

        r.case("TC-M15-04", "详情响应时间", "单次 < 2s", t04, priority="P2")

        def t05():
            _, ts = timed(lambda: A.get(B + "/api/dashboard", timeout=60), n=10)
            s = stat(ts)
            assert s["max"] < LIMIT_MS, "dashboard 最慢 %.0fms 超 2s" % s["max"]
            return "首页 dashboard %s" % fmt(s)

        r.case("TC-M15-05", "首页 dashboard 响应时间", "单次 < 2s", t05, priority="P1")

        def t06():
            out = []
            for path in ("/api/statistics", "/api/asset_statistics"):
                _, ts = timed(lambda p=path: A.get(B + p, params={"type": "by_dept"}, timeout=60), n=5)
                s = stat(ts)
                out.append("%s %s" % (path, fmt(s)))
                assert s["max"] < LIMIT_MS, "%s 最慢 %.0fms 超 2s" % (path, s["max"])
            return "；".join(out)

        r.case("TC-M15-06", "统计接口响应时间", "单次 < 2s", t06, priority="P2")

        def t07():
            out = []
            for path, name in (("/api/contracts/export", "项目导出"), ("/api/asset_contracts/export", "资产导出")):
                t0 = time.perf_counter()
                resp = A.get(B + path, timeout=300)
                ms = (time.perf_counter() - t0) * 1000
                assert resp.status_code == 200, "%s HTTP %s" % (name, resp.status_code)
                assert ms < LIMIT_MS * 4, "%s 耗时 %.0fms（超 8s）" % (name, ms)
                out.append("%s %.0fms/%.2fMB" % (name, ms, len(resp.content) / 1048576.0))
            return "导出耗时与体积：%s（判定阈值放宽为 8s，属离线导出）" % "；".join(out)

        r.case("TC-M15-07", "导出耗时与体积", "项目导出 < 8s", t07, priority="P1")

        def t08():
            before_files = len(os.listdir(RUNTIME_BACKUPS))
            before_size = sum(os.path.getsize(os.path.join(RUNTIME_BACKUPS, f)) for f in os.listdir(RUNTIME_BACKUPS))
            body = {"category": "服务类", "name": "M15备份-" + stamp, "procurement_method": "自行采购",
                    "amount": 1, "main_content": "备份开销测试", "handler": ia["name"], "fund_source": "其他"}
            t0 = time.perf_counter()
            resp = A.post(B + "/api/contracts", json=body, timeout=60)
            ms = (time.perf_counter() - t0) * 1000
            assert resp.status_code == 201, "新增失败 %s" % resp.status_code
            after_files = len(os.listdir(RUNTIME_BACKUPS))
            after_size = sum(os.path.getsize(os.path.join(RUNTIME_BACKUPS, f)) for f in os.listdir(RUNTIME_BACKUPS))
            read_ms = statistics.mean(timed(lambda: A.get(B + "/api/dashboard", timeout=60), n=5)[1])
            db_mb = os.path.getsize(L) / 1048576.0
            assert ms < LIMIT_MS * 3, "单次写（含整库备份）耗时 %.0fms 过长" % ms
            return "单次写操作 %.0fms（含整库备份，库 %.2fMB）；对照热读 %.0fms；backups 文件 %d→%d、体积 %+.2fMB" % (
                ms, db_mb, read_ms, before_files, after_files, (after_size - before_size) / 1048576.0)

        r.case("TC-M15-08", "写前整库备份开销", "单次写 < 6s，量化磁盘增长", t08, priority="P1")

        def t09():
            N = 10

            def one(_):
                try:
                    resp = A.get(B + "/api/contracts", params={"page": 1, "page_size": 50}, timeout=60)
                    return resp.status_code
                except Exception as e:  # noqa: BLE001
                    return "ERR:%s" % e

            t0 = time.perf_counter()
            with ThreadPoolExecutor(max_workers=N) as ex:
                codes = list(ex.map(one, range(N)))
            total_ms = (time.perf_counter() - t0) * 1000
            ok = sum(1 for c in codes if c == 200)
            rate = ok * 100.0 / N
            assert rate > 99, "并发成功率 %.1f%% ≤ 99%%（%s）" % (rate, codes)
            assert total_ms < LIMIT_MS * 3, "10 并发总耗时 %.0fms 过长" % total_ms
            return "%d 并发读列表：成功 %d/%d（%.0f%%）、总耗时 %.0fms、平均 %.0fms/请求" % (
                N, ok, N, rate, total_ms, total_ms / N)

        r.case("TC-M15-09", "10 并发读", "成功率 > 99%、无 500", t09, priority="P1")

        srv.sync_db()
    finally:
        srv.stop()
    prod_after = C.prod_fingerprint()
    meta = {
        "服务地址": B,
        "工作库": "test_data/m15_perf.db",
        "数据规模": "项目合同 %d / 资产合同 %d / 租金明细 %d / 用户 %d" % (
            C.query(C.BASE_DB, "SELECT COUNT(*) c FROM project_contracts")[0]["c"],
            C.query(C.BASE_DB, "SELECT COUNT(*) c FROM asset_contracts")[0]["c"],
            C.query(C.BASE_DB, "SELECT COUNT(*) c FROM asset_rent_collections")[0]["c"],
            C.query(C.BASE_DB, "SELECT COUNT(*) c FROM users")[0]["c"]),
        "判定标准": "单次 < 2s；并发成功率 > 99%",
        "生产库未改动": "是" if prod_before == prod_after else "否（指纹变化！）",
    }
    r.report("M15_性能与缓存测试.md", meta=meta)
    n, p, f, e = r.stat()
    print("\n[SUMMARY] M15 共 %d 例：通过 %d，失败 %d，错误 %d" % (n, p, f, e))
    sys.stdout.flush()
    os._exit(0 if (f == 0 and e == 0) else 1)


if __name__ == "__main__":
    main()
