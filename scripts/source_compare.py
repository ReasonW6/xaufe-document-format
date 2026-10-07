"""Find omitted source prose without treating generated JSON as the original.

This is a comparison aid, not a complete Markdown parser or a claim of semantic
equivalence. Unsupported source types still need the original-file review.
"""

from pathlib import Path
import re
from docx import Document


def compact(text):
    return re.sub(r"\s+", "", text)


def fragments(path):
    path = Path(path)
    if path.suffix.lower() == ".docx":
        doc = Document(path)
        return [
            ("paragraph " + str(i + 1), "".join(p.xpath(".//w:t/text()")))
            for i, p in enumerate(doc.element.body.xpath(".//w:p"))
            if "".join(p.xpath(".//w:t/text()")).strip()
        ]
    if path.suffix.lower() not in (".md", ".txt"):
        return None
    text = path.read_text(encoding="utf-8-sig")
    result = []
    lines = []
    start = 1
    fenced = False

    def flush():
        if lines:
            result.append(("line " + str(start), " ".join(lines)))
        lines.clear()

    for index, line in enumerate(text.splitlines(), 1):
        stripped = line.strip()
        if stripped.startswith(("```", "~~~")):
            flush()
            fenced = not fenced
            continue
        if fenced:
            continue
        if not stripped or stripped.startswith(("#", "|", "![")):
            flush()
            continue
        title = re.match(r"^(?:English title|英文题目)\s*[:：]\s*(.+)$", stripped, re.I)
        words = re.match(r"^(?:Key words|Keywords|关键词)\s*[:：]\s*(.+)$", stripped, re.I)
        if title or words:
            flush()
            values = re.split(r"[;；]", words.group(1)) if words else [title.group(1)]
            result.extend(
                ("line " + str(index), value.strip()) for value in values if value.strip()
            )
            continue
        # Preserve visible link labels and inline emphasis, not Markdown syntax.
        stripped = re.sub(r"\[([^\]]+)\]\([^)]*\)", r"\1", stripped)
        stripped = re.sub(r"(?<!\\)(\*\*|__|`)", "", stripped)
        stripped = re.sub(r"^[-*+]\s+", "", stripped)
        if not lines:
            start = index
        lines.append(stripped)
    flush()
    return result


def compare(sources, docx):
    doc = Document(docx)
    target = compact("".join(doc.element.body.xpath(".//w:t/text()")))
    reports = []
    for source in sources:
        path = Path(source["path"])
        items = fragments(path)
        if items is None:
            reports.append(
                {
                    "source": str(path),
                    "status": "manual_review",
                    "reason": "此文件类型需按原件核对文字或对象。",
                }
            )
            continue
        missing = [
            {"location": loc, "text": text} for loc, text in items if compact(text) not in target
        ]
        reports.append(
            {
                "source": str(path),
                "status": "differences" if missing else "matched_prose",
                "checked_fragments": len(items),
                "missing": missing,
            }
        )
    return {
        "sources": reports,
        "notice": "只检查可提取的原始文字片段；标题映射、表格、公式、图片、链接和用户明确删改范围仍须核对。差异不是自动认定丢失，应对照用户要求逐项处理。",
    }
