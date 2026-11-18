# -*- coding: utf-8 -*-
"""组装临时 DOCX，并按需从同一份 DOCX 导出 PDF。"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]

FINALIZE_DOCX = ROOT / "src" / "finalize_docx.py"


def assemble_docx(out_path: Path, variant: str, pdf_path: Path | None = None) -> None:
    command = [sys.executable, str(FINALIZE_DOCX), "--variant", variant,
               "--docx", str(out_path)]
    if pdf_path is None:
        command.append("--for-dotm")
    else:
        command.extend(["--pdf", str(pdf_path)])
    subprocess.run(command, check=True)
    if not out_path.exists():
        raise RuntimeError(f"未能生成临时 docx：{out_path}")
