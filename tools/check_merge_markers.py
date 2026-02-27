#!/usr/bin/env python3
from __future__ import annotations

from pathlib import Path
import re

ROOT = Path(__file__).resolve().parents[1]
MARKER = re.compile(r"^(<<<<<<<|=======|>>>>>>>)", re.MULTILINE)
IGNORE_DIRS = {".git", ".venv", "__pycache__"}
IGNORE_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".pdf", ".zip", ".xlsm", ".xlam", ".xlsx", ".sqlite", ".pyc"}


def iter_files(root: Path):
    for p in root.rglob("*"):
        if p.is_dir():
            continue
        if any(part in IGNORE_DIRS for part in p.parts):
            continue
        if p.suffix.lower() in IGNORE_EXTS:
            continue
        yield p


def main() -> int:
    failures: list[str] = []
    for path in iter_files(ROOT):
        try:
            text = path.read_text(encoding="utf-8")
        except Exception:
            continue
        if MARKER.search(text):
            failures.append(str(path.relative_to(ROOT)))

    if failures:
        print("Merge conflict markers detected in:")
        for f in failures:
            print(f" - {f}")
        return 1

    print("No merge conflict markers found.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
