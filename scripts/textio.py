"""Reading text files written by any common editor (UTF-8 with or without BOM, UTF-16, GBK)."""

from __future__ import annotations
import json
from pathlib import Path


def read_text_any(path: str | Path) -> str:
    """UTF-8 (with or without BOM); also UTF-16 and GBK files written by Windows tools."""
    raw = Path(path).read_bytes()
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return raw.decode("utf-16")
    for encoding in ("utf-8-sig", "gbk"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise ValueError(f"{Path(path).name} 的编码无法识别，请用 UTF-8 重新保存。")


def read_json(path: str | Path):
    return json.loads(read_text_any(path))
