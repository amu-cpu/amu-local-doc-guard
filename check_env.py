from __future__ import annotations

import importlib.util
import os
import shutil
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent


def module_status(name: str) -> str:
    return "OK" if importlib.util.find_spec(name) else "MISSING"


def first_existing(paths: list[str]) -> str:
    for value in paths:
        if not value:
            continue
        path = Path(value)
        if path.exists():
            return str(path)
    return "MISSING"


def main() -> int:
    print("Project:", PROJECT_ROOT)
    print("Python:", sys.executable)
    print("Python version:", sys.version.split()[0])
    print("Virtual env:", os.environ.get("VIRTUAL_ENV", "not active"))
    print("Flask:", module_status("flask"))
    print("PyMuPDF:", module_status("fitz"))
    print("Pillow:", module_status("PIL"))
    print(
        "LibreOffice:",
        first_existing(
            [
                os.environ.get("LIBREOFFICE_PATH", ""),
                r"F:\Tools\LibreOffice\program\soffice.com",
                r"F:\Tools\LibreOffice\program\soffice.exe",
                shutil.which("soffice.com") or "",
                shutil.which("soffice") or "",
                r"C:\Program Files\LibreOffice\program\soffice.com",
                r"C:\Program Files\LibreOffice\program\soffice.exe",
            ]
        ),
    )
    print("Output dir:", PROJECT_ROOT / "output")
    print("Private samples dir:", PROJECT_ROOT / "private_samples")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
