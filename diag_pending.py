# -*- coding: utf-8 -*-
import os, sys, io, datetime
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
print("CWD:", os.getcwd())
print("sys.executable:", sys.executable)
try:
    print("app.py mtime:", datetime.datetime.fromtimestamp(os.path.getmtime(os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.py"))))
except Exception as e:
    print("app.py mtime error:", e)
try:
    import app
    print("DB_PATH:", app.DB_PATH)
    print("DB exists:", os.path.exists(app.DB_PATH))
    c = app.app.test_client()
    r = c.post("/api/login", json={"username": "\u5218\u52c7", "password": "123456"})
    print("login status:", r.status_code)
    if r.status_code == 200:
        d = c.get("/api/dashboard").get_json()
        p = c.get("/api/pending_payments").get_json()
        print("CARD pending_project_count =", d["pending_project_count"])
        print("REMINDS count =", p["count"], " items =", len(p["items"]))
        print("MATCH =", d["pending_project_count"] == p["count"] == len(p["items"]))
    else:
        print("login failed:", r.get_data(as_text=True)[:200])
except Exception as e:
    import traceback
    traceback.print_exc()