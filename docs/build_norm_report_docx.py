# -*- coding: utf-8 -*-
"""将 数据库规范化评估报告.md 转换为 数据库规范化评估报告.docx（python-docx）。"""
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _docx_common import render

BASE = Path(__file__).resolve().parent  # docs 目录
MD_PATH = BASE / "数据库规范化评估报告.md"
OUT_PATH = BASE / "数据库规范化评估报告.docx"

if __name__ == "__main__":
    render(MD_PATH, OUT_PATH,
           header_text="“潭合同”台账管理系统 · 数据库规范化评估报告",
           core_title="《数据库规范化评估报告》",
           core_author="湘潭市机关事务管理局")