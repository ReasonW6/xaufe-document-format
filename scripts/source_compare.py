"""Compare the delivered Word with the user's original, in both directions.

Missing: a paragraph of the original that does not occur in the output.
Added: an output paragraph whose text occurs nowhere in the original.

Text is compared without whitespace and punctuation, so re-spacing keywords or
full-width/half-width punctuation is not a difference; any changed, dropped or
invented wording is.  Parts the template itself generates (cover, TOC, part
labels such as 摘要/参考文献, the header) are not counted as added, and the old
cover and old TOC that the template replaces are not counted as missing.
"""

from pathlib import Path
import re
from docx import Document
from docx.oxml.ns import qn

GENERATED_ROLES = {
    "toc_title", "toc1", "toc2", "toc3", "toc4", "abstract_label_zh", "abstract_label_en",
    "reference_title", "header", "spacer", "abstract_spacer",
}
LABEL = re.compile(r"^\s*(?:关键词|key\s*words?)\s*[:：]|^\s*[\[［]\d+[\]］]", re.I)


def key(text):
    return re.sub(r"[\W_]+", "", str(text or "")).lower()


def compact(text):
    return re.sub(r"\s+", "", text)


def _text(element):
    return "".join(element.xpath(".//w:t/text()|.//m:t/text()"))


def fragments(path):
    """[(location, text, block_index)] of the original; None for file types that cannot be read."""
    path = Path(path)
    if path.suffix.lower() == ".docx":
        doc = Document(path)
        result, pno, tno = [], 0, 0
        for index, el in enumerate(doc.element.body):
            if el.tag == qn("w:p"):
                if _text(el).strip():
                    result.append((f"p{pno}", _text(el), index))
                pno += 1
            elif el.tag == qn("w:tbl"):
                for cell in el.xpath(".//w:tc"):
                    text = "".join(_text(p) for p in cell.xpath("./w:p"))
                    if text.strip():
                        result.append((f"t{tno}", text, index))
                tno += 1
        return result
    if path.suffix.lower() not in (".md", ".txt"):
        return None
    text = path.read_text(encoding="utf-8-sig")
    result = []
    lines = []
    start = 1
    fenced = False

    def flush():
        if lines:
            result.append(("line " + str(start), " ".join(lines), None))
        lines.clear()

    for index, line in enumerate(text.splitlines(), 1):
        stripped = line.strip()
        if stripped.startswith(("```", "~~~")):
            flush()
            fenced = not fenced
            continue
        if fenced:
            if stripped:
                result.append(("line " + str(index), stripped, None))
            continue
        if not stripped or stripped.startswith(("|", "![")) or re.fullmatch(r"[-*_=\s]{3,}", stripped):
            flush()
            if stripped.startswith("|"):
                cells = [c for c in stripped.strip("|").split("|") if key(c) and not re.fullmatch(r"\s*:?-+:?\s*", c)]
                result.extend(("line " + str(index), c.strip(), None) for c in cells)
            continue
        if stripped.startswith("#"):
            flush()
            result.append(("line " + str(index), stripped.lstrip("#").strip(), None))
            continue
        title = re.match(r"^(?:English title|英文题目)\s*[:：]\s*(.+)$", stripped, re.I)
        words = re.match(r"^(?:Key words|Keywords|关键词)\s*[:：]\s*(.+)$", stripped, re.I)
        if title or words:
            flush()
            values = re.split(r"[;；]", words.group(1)) if words else [title.group(1)]
            result.extend(("line " + str(index), value.strip(), None) for value in values if value.strip())
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


def _whole(path):
    path = Path(path)
    if path.suffix.lower() == ".docx":
        return "".join(_text(p) for p in Document(path).element.body.xpath(".//w:p"))
    return path.read_text(encoding="utf-8-sig")


def body_key(docx):
    """All visible text except the TOC (its page numbers may change), for "only formatting changed" checks."""
    doc = Document(docx)
    toc = {s.style_id for s in doc.styles if s.name.lower().startswith("toc ")}
    return key("".join(
        _text(p) for p in doc.element.body.xpath(".//w:p")
        if not set(p.xpath("./w:pPr/w:pStyle/@w:val")) & toc
    ))


def output_paragraphs(docx):
    """[(label, role, text)] of the output body, skipping the template's cover section."""
    from customization import ROLE_STYLES

    doc = Document(docx)
    roles = {name: role for role, name in ROLE_STYLES.items()}
    has_cover = any(
        v.get(qn("w:name")) == "XAUFE_HAS_COVER" and v.get(qn("w:val")) == "1"
        for v in doc.settings.element.xpath("./w:docVars/w:docVar")
    )
    in_cover = has_cover
    result, number = [], 0
    for el in doc.element.body:
        if el.tag == qn("w:p"):
            style = el.xpath("./w:pPr/w:pStyle/@w:val")
            name = doc.styles.get_by_id(style[0], 1).name if style else "Normal"
            if not in_cover and _text(el).strip():
                result.append((f"成品第 {number + 1} 段", roles.get(name), _text(el)))
            number += 1
            if in_cover and el.xpath("./w:pPr/w:sectPr"):
                in_cover = False
        elif el.tag == qn("w:tbl") and not in_cover:
            for cell in el.xpath(".//w:tc"):
                text = "".join(_text(p) for p in cell.xpath("./w:p"))
                if text.strip():
                    result.append((f"成品第 {number} 段后的表格", "table_text", text))
    return result


def _approved(text, approved):
    k = key(text)
    return any(len(key(a)) >= 4 and (key(a) in k or k in key(a)) for a in approved)


def compare(sources, docx, *, approved=(), generated=(), replaced=None, old_toc=()):
    """Report per source; ``approved`` texts the user agreed to add or drop, ``generated``
    texts that come from job.json (cover values) or mapping.json "add" (agreed translations),
    ``replaced`` the old-cover block range and ``old_toc`` the old TOC paragraph ids of a .docx."""
    target_items = output_paragraphs(docx)
    # Missing is judged against everything visible, cover included (a title may live only there).
    target = key(_whole(docx))
    allowed = {key(g) for g in generated if key(g)}
    reports = []
    whole = ""
    for source in sources:
        path = Path(source["path"])
        items = fragments(path)
        if items is None:
            reports.append({"source": str(path), "status": "manual_review", "reason": "此文件类型需按原件核对文字或对象。"})
            continue
        whole += _whole(path)
        missing = []
        for loc, text, block in items:
            if path.suffix.lower() == ".docx" and (
                (replaced is not None and block is not None and block < replaced) or loc in old_toc
            ):
                continue
            if key(text) and key(text) not in target and not _approved(text, approved):
                missing.append({"location": loc, "text": text})
        reports.append(
            {
                "source": str(path),
                "status": "differences" if missing else "matched_prose",
                "checked_fragments": len(items),
                "missing": missing,
            }
        )
    readable = [r for r in reports if r["status"] != "manual_review"]
    added = []
    if readable:
        original = key(whole)
        for label, role, text in target_items:
            if role in GENERATED_ROLES:
                continue
            k = key(LABEL.sub("", text))
            if k and k not in original and k not in allowed and not _approved(text, approved):
                added.append({"location": label, "role": role, "text": text})
    return {
        "sources": reports,
        "added": added,
        "notice": "按文字比对（不计空格和标点）：missing 是原稿有、成品没有的段落；added 是成品有、原稿里找不到的文字。"
        "表格结构、图片、公式排版和链接目标仍需看页图核对。",
    }
