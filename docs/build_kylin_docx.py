# -*- coding: utf-8 -*-
"""将 系统部署手册_麒麟V10.md 转换为 系统部署手册_麒麟V10.docx（python-docx）。"""

import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE))

import _docx_common


if __name__ == "__main__":
    md_path = BASE / "系统部署手册_麒麟V10.md"
    out_path = BASE / "系统部署手册_麒麟V10.docx"
    _docx_common.render(
        md_path=md_path,
        out_path=out_path,
        header_text="潭合同 台账管理系统 · 麒麟 V10 部署操作指引",
        core_title="潭合同 台账管理系统 · 麒麟 V10 部署操作指引",
        core_author="系统管理员",
    )
    print("已生成:", out_path)
