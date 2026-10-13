# -*- coding: utf-8 -*-
"""测试用服务启动器：以「运行时副本」的 app.py 启动 Flask，绝不加载生产目录的 app.py。
业务代码零改动；副本的 data/合同台账.db 即测试库。
"""
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RUNTIME = os.path.join(ROOT, "test_data", "runtime")
sys.path.insert(0, RUNTIME)

import app as m  # noqa: E402

real = os.path.abspath(m.__file__)
if not real.startswith(os.path.abspath(RUNTIME)):
    sys.stderr.write("拒绝启动：import app 解析到 %s（应为运行时副本）\n" % real)
    sys.exit(2)

port = int(os.environ.get("TM_TEST_PORT", "5099"))
sys.stderr.write("[SERVE] runtime=%s db=%s port=%d\n" % (RUNTIME, m.DB_PATH, port))
sys.stderr.flush()
m.app.run(host="127.0.0.1", port=port, debug=False, threaded=True, use_reloader=False)
