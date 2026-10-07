#!/usr/bin/env python3
"""Cross-platform bundle verification, including manifest coverage and safe paths.

Only Python/cache/OS bookkeeping is ignored. Store task outputs outside the skill.
This detects accidental damage; it is not an authenticity signature.
"""

from __future__ import annotations
import hashlib
import re
import sys
from pathlib import Path, PurePosixPath

ROOT = Path(__file__).resolve().parents[1]
IGNORED_DIRS = {"__pycache__", ".pytest_cache", ".git"}
IGNORED_FILES = {".DS_Store", "Thumbs.db", "desktop.ini"}


def ignored(path: Path) -> bool:
    return any(part in IGNORED_DIRS for part in path.parts) or path.name in IGNORED_FILES


def verify(root: Path) -> dict:
    root = Path(root).resolve()
    manifest = root / "SHA256SUMS"
    bad: list[str] = []
    listed: set[str] = set()
    if not manifest.is_file():
        return {"passed": False, "files": 0, "errors": ["缺少SHA256SUMS，请恢复完整技能包。"]}
    try:
        lines = manifest.read_text(encoding="utf-8").splitlines()
    except (OSError, UnicodeError) as exc:
        return {"passed": False, "files": 0, "errors": ["无法读取校验清单：" + str(exc)]}
    for line in lines:
        if not line.strip():
            continue
        fields = line.split("  ", 1)
        if len(fields) != 2 or not re.fullmatch(r"[0-9a-fA-F]{64}", fields[0]):
            bad.append("校验清单语法或摘要格式错误。")
            continue
        digest, name = fields
        rel = PurePosixPath(name)
        if (
            not name
            or "\\" in name
            or rel.is_absolute()
            or ".." in rel.parts
            or ":" in name
            or rel.as_posix() != name
            or name == "SHA256SUMS"
        ):
            bad.append("非法校验路径：" + name)
            continue
        if name in listed:
            bad.append("重复校验路径：" + name)
            continue
        listed.add(name)
        target = root / name
        if not target.resolve().is_relative_to(root) or any(
            p.is_symlink()
            for p in [target, *target.parents]
            if p != root and p.is_relative_to(root)
        ):
            bad.append("校验路径越界或包含符号链接：" + name)
        elif not target.is_file():
            bad.append("缺少：" + name)
        else:
            try:
                if hashlib.sha256(target.read_bytes()).hexdigest() != digest.lower():
                    bad.append("内容变化：" + name)
            except OSError as exc:
                bad.append("无法读取：" + name + "：" + str(exc))
    if not listed:
        bad.append("校验清单为空或没有有效条目。")
    actual = {
        p.relative_to(root).as_posix()
        for p in root.rglob("*")
        if (p.is_file() or p.is_symlink()) and p != manifest and not ignored(p.relative_to(root))
    }
    for extra in sorted(actual - listed):
        bad.append("清单未覆盖的额外文件或工作产物：" + extra)
    return {"passed": not bad, "files": len(listed), "errors": bad}


def main() -> int:
    result = verify(ROOT)
    if not result["passed"]:
        print("\n".join(result["errors"]), file=sys.stderr)
        return 2
    print(f"完整性检查通过：{result['files']}个文件；清单完整覆盖。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
