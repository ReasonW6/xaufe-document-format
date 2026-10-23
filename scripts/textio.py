"""Reading text files written by any common editor (UTF-8 with or without BOM, UTF-16, GBK)."""

from __future__ import annotations
import json, re
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


# How a .txt/.md original reads as text: shared by the importer (text_import) and the
# format-only comparison (source_compare), so both agree on what the user's words are.
_CODE_SPAN = re.compile(r"(`+)(.+?)\1")
_LINK = re.compile(r"\[([^\]]+)\]\([^)]*\)")
# "*" may emphasise inside a word (CJK has no spaces); "_" only at word edges, as in CommonMark.
_EMPHASIS = (re.compile(r"(?<![\\*])\*\*(?=\S)(.+?)(?<=\S)\*\*(?!\*)"),
             re.compile(r"(?<![\\\w])__(?=\S)(.+?)(?<=\S)__(?!\w)"),
             re.compile(r"(?<![\\*])\*(?=[^\s*])(.+?)(?<=[^\s*])\*(?!\*)"))
_ESCAPED = re.compile(r"\\([\\`*_{}\[\]()#+\-.!|])")


def md_visible(line, keep_links=False):
    """A Markdown line as it reads: code spans literal (only their backticks go), paired emphasis
    markers dropped, escapes resolved; a lone * or __ is the user's character and stays.
    ``keep_links`` leaves [label](url) for the importer, which turns it into a link."""
    out, at = [], 0
    for m in _CODE_SPAN.finditer(line):
        out.append(_md_plain(line[at:m.start()], keep_links))
        out.append(m.group(2).strip(" ") if m.group(2).strip(" ") else m.group(2))
        at = m.end()
    out.append(_md_plain(line[at:], keep_links))
    return "".join(out)


def _md_plain(text, keep_links):
    if not keep_links:
        text = _LINK.sub(r"\1", text)
    for pattern in _EMPHASIS:
        text = pattern.sub(r"\1", text)
    return _ESCAPED.sub(r"\1", text)


_SENTENCE_END = tuple("。！？!?.；;：:”」』）)…")


def continues(previous, line, longest):
    """A .txt line that carries on the paragraph above it (a hard-wrapped line), not a new one.

    Most Chinese .txt files put each paragraph on one line, often without blank lines between;
    hard-wrapped files break long lines near a fixed width.  So a line continues the one above
    only when that one is near the longest line's length and does not end a sentence."""
    return bool(previous) and longest >= 30 and len(previous) >= 0.7 * longest \
        and not previous.rstrip().endswith(_SENTENCE_END)
