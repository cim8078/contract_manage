# -*- coding: utf-8 -*-
"""将 第三方许可清单.md 转换为 第三方许可清单.docx（python-docx）。"""
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _docx_common import render

BASE = Path(__file__).resolve().parent  # docs 目录
MD_PATH = BASE / "第三方许可清单.md"
OUT_PATH = BASE / "第三方许可清单.docx"

if __name__ == "__main__":
    render(MD_PATH, OUT_PATH,
           header_text="“潭合同”台账管理系统 · 第三方开源组件许可清单",
           core_title="“潭合同”台账管理系统 第三方开源组件许可清单",
           core_author="湘潭市机关事务管理局")
