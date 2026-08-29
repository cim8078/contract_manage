# -*- coding: utf-8 -*-
"""将 数据字典表.md 转换为 数据字典表.docx（python-docx）。"""
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _docx_common import render

BASE = Path(__file__).resolve().parent  # docs 目录
MD_PATH = BASE / "数据字典表.md"  # 源 md 位于 docs/ 目录
OUT_PATH = BASE / "数据字典表.docx"

if __name__ == "__main__":
    render(MD_PATH, OUT_PATH,
           header_text="“潭合同”台账管理系统 · 数据字典表",
           core_title="《数据字典表》",
           core_author="湘潭市机关事务管理局")
