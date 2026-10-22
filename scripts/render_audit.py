#!/usr/bin/env python3
"""Rendered-page audit: reads only the exported PDF and the delivered DOCX.

It does not call the layout engine.  The page box and margins come from the
DOCX section properties (the static audit already ties those to the template),
and the header/footer expectations come from the recorded effective profile.
Everything else is measured on the PDF a real office suite produced, so a
wrong assumption in the generator cannot hide a visible defect here.

Errors are things a person would see as broken: content outside the page or
the text column, a page with nothing in the text area, lines printed on top of
each other, missing or unexpected header/footer text, page numbers that skip
or restart in the wrong place.  The one warning is a font the renderer had to
substitute, which changes the page images but not the Word file.
"""

from __future__ import annotations
import argparse, json, re, sys
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn

PT = 20  # twips per point
EDGE_TOLERANCE = 2.0  # pt; drawing and image edges are exact
TEXT_TOLERANCE = 6.0  # pt; about half a 五号 glyph, absorbs justification and font metrics
COVER_SAFE = 28.0  # pt (~1 cm); the school cover centres 36 pt titles on its own measure
HANGING = set("，。、；：！？）》」』”’,.;:!?)")

# PDF font names that count as the requested family.
FONT_ALIASES = {
    "宋体": ("simsun", "nsimsun", "songti", "stsong", "宋体"),
    "黑体": ("simhei", "heiti", "stheiti", "黑体"),
    "仿宋": ("fangsong", "stfangsong", "仿宋"),
    "楷体": ("kaiti", "stkaiti", "楷体"),
    "微软雅黑": ("microsoftyahei", "msyh", "微软雅黑"),
    # Liberation/Tinos/Arimo are metric-compatible: same advance widths, same line breaks.
    "Times New Roman": ("timesnewroman", "timesnewromanps", "liberationserif", "tinos"),
    "Arial": ("arial", "arialmt", "liberationsans", "arimo"),
    "Courier New": ("couriernew", "couriernewps", "liberationmono", "cousine"),
    "Calibri": ("calibri",),
    "Consolas": ("consolas",),
}


def _key(name):
    # "ABCDEF+SimSun-Bold" -> "simsunbold"
    return re.sub(r"[^0-9a-z一-鿿]", "", name.split("+", 1)[-1].lower())


def requested_fonts(doc):
    """Font families the visible text of the DOCX resolves to (styles, runs, tables, headers)."""
    from style_model import run_properties
    from table_layout import walk_paragraphs

    paragraphs = list(doc.paragraphs)
    for t in doc.tables:
        paragraphs += [p for _, p in walk_paragraphs(t, "t")]
    for section in doc.sections:
        for part in (section.header, section.first_page_header, section.footer, section.first_page_footer):
            paragraphs += list(part.paragraphs)
    names = set()
    for p in paragraphs:
        for r in p._p.xpath(".//w:r"):
            text = "".join(r.xpath("./w:t/text()"))
            if not text.strip():
                continue
            try:
                fonts = run_properties(p, r).get("fonts", {})
            except ValueError:
                continue
            cjk = bool(re.search(r"[\u3400-\u9fff\uff00-\uffef]", text))
            latin = bool(re.search(r"[A-Za-z0-9]", text))
            wanted = ([fonts.get("eastAsia")] if cjk else []) + ([fonts.get("ascii")] if latin else [])
            for value in wanted:
                if value:
                    names.add(value.split(";")[-1].strip())
    return sorted(names)


def _font_present(family, used):
    aliases = FONT_ALIASES.get(family, (_key(family),))
    return any(alias and any(used_key.startswith(alias) for used_key in used) for alias in aliases)


def _expectations(doc):
    vv = {e.get(qn("w:name")): e.get(qn("w:val")) for e in doc.settings.element.xpath("./w:docVars/w:docVar")}
    header = footer = None
    template = vv.get("XAUFE_TEMPLATE")
    if template in ("A", "B"):
        from customization import effective_profile

        spec = effective_profile(template, json.loads(vv.get("XAUFE_FORMAT_REQUEST", "{}")))
        header = bool(spec["header"]["enabled"])
        footer = bool(spec["footer"].get("page_numbers"))
    return {"has_cover": vv.get("XAUFE_HAS_COVER") == "1", "header": header, "page_numbers": footer}


def _margins(sections):
    """Page box and text column; when sections differ, the most generous edge of each."""
    get = lambda s, name: getattr(s, name).twips / PT
    return {
        "width": max(get(s, "page_width") for s in sections),
        "height": max(get(s, "page_height") for s in sections),
        **{k: min(get(s, k + "_margin") for s in sections) for k in ("left", "right", "top", "bottom")},
    }


def _issue(kind, code, page, message, *, where=None, expected=None, actual=None, bbox=None):
    item = {"code": code, "severity": kind, "page": page, "message": message}
    if where:
        item["where"] = where
    if expected is not None:
        item["expected"] = expected
    if actual is not None:
        item["actual"] = actual
    if bbox is not None:
        item["bbox_pt"] = [round(v, 1) for v in bbox]
    return item


def _lines(page):
    """Visible lines; ``ink_right`` ignores one trailing punctuation mark (行末标点悬挂)."""
    result = []
    for block in page.get_text("rawdict")["blocks"]:
        for line in block.get("lines", []):
            chars = [dict(c, size=span["size"]) for span in line["spans"] for c in span["chars"] if c["c"].strip()]
            if not chars:
                continue
            boxes = [c["bbox"] for c in chars]
            body = chars[:-1] if len(chars) > 1 and chars[-1]["c"] in HANGING else chars
            result.append(
                {
                    "spans": [{"text": "".join(c["c"] for c in chars)}],
                    "bbox": (min(b[0] for b in boxes), min(b[1] for b in boxes),
                             max(b[2] for b in boxes), max(b[3] for b in boxes)),
                    "ink_right": max(c["bbox"][2] for c in body),
                    "chars": [(c["bbox"][0], c["bbox"][2], c["origin"][1], c["size"]) for c in chars],
                }
            )
    return result


def _ink(char):
    """Where a glyph actually puts ink: about 0.8 em above the baseline to 0.1 em below it."""
    x0, x1, base, size = char
    return x0, base - 0.8 * size, x1, base + 0.1 * size


def _line_clash(a, b):
    """(baseline gap, ink overlap) of the worst pair of glyphs printed into each other, or None.

    Each glyph inks roughly 0.8 em above to 0.1 em below its baseline.  Two glyphs
    clash when they share at least 30% of the narrower one's width and those ink
    bands overlap by more than 5% of the smaller size: for one size that is a
    baseline gap under about 0.85 of the size; a large title above small text is
    judged by the room actually left between them.  Superscripts beside a word
    never share its width.
    """
    if a["bbox"][1] > b["bbox"][3] + 40 or b["bbox"][1] > a["bbox"][3] + 40:
        return None
    worst = None
    for char in a["chars"]:
        x0, x1, base, size = char
        for other in b["chars"]:
            u0, u1, ubase, usize = other
            narrow = min(x1 - x0, u1 - u0)
            if narrow <= 0 or min(x1, u1) - max(x0, u0) < 0.3 * narrow:
                continue
            _, top, _, bottom = _ink(char)
            _, utop, _, ubottom = _ink(other)
            overlap = min(bottom, ubottom) - max(top, utop)
            if overlap > 0.05 * min(size, usize) and (worst is None or overlap > worst[1]):
                worst = (abs(base - ubase), overlap)
    return worst


def _covers(box, ink, share=0.2):
    width = min(box[2], ink[2]) - max(box[0], ink[0])
    height = min(box[3], ink[3]) - max(box[1], ink[1])
    area = (ink[2] - ink[0]) * (ink[3] - ink[1])
    return width > 0 and height > 0 and area > 0 and width * height > share * area


def _header_texts(doc):
    texts = set()
    for s in doc.sections:
        for part in (s.header, s.even_page_header):
            for p in part.paragraphs:
                t = re.sub(r"\s+", "", p.text)
                if t:
                    texts.add(t)
    return texts


def _header_overflow(lines, top, texts):
    """The last header line when the header wrapped past the top margin into the text area."""
    if not texts:
        return None
    chain = [l for l in lines if l["bbox"][3] <= top + 1]
    if not chain:
        return None
    last = max(chain, key=lambda l: l["bbox"][3])
    for line in sorted(lines, key=lambda l: l["bbox"][1]):
        if line in chain or line["bbox"][1] < last["bbox"][1]:
            continue
        height = last["bbox"][3] - last["bbox"][1]
        text = re.sub(r"\s+", "", line["spans"][0]["text"])
        if line["bbox"][1] - last["bbox"][3] < 0.6 * height and text and any(text in t for t in texts):
            last = line
        else:
            break
    return last if last["bbox"][3] > top + 1 else None


ROMAN = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100}


def _page_number(text):
    """("decimal", 3) / ("roman", 3) for a footer like "3", "- 3 -", "第3页", "III"; None otherwise."""
    core = re.sub(r"[\s\-—第页]", "", text)
    if core.isdigit():
        return "decimal", int(core)
    if core and re.fullmatch(r"[IVXLC]+|[ivxlc]+", core):
        values = [ROMAN[c] for c in core.upper()]
        return "roman", sum(-v if i + 1 < len(values) and v < values[i + 1] else v for i, v in enumerate(values))
    return None


def _number_sequence(numbers):
    """Front matter I, II, ...; body 1, 2, ... to the end: no skip, repeat or stray restart."""
    errors, last = [], None
    for page, text in numbers:
        value = _page_number(text)
        if value is None:  # already reported as PAGE_NUMBER_MISSING; keep counting past it
            last = (last[0], last[1] + 1, text) if last else None
            continue
        kind, n = value
        first = "1" if kind == "decimal" else "I"
        if last is None and n != 1:
            errors.append(_issue("error", "PAGE_NUMBER_SEQUENCE", page, "第一个显示页码的页面没有从 1 开始。",
                                 expected=first, actual=text))
        elif last is not None and last[0] != kind and kind == "roman":
            errors.append(_issue("error", "PAGE_NUMBER_SEQUENCE", page, "正文之后又出现了罗马数字页码。",
                                 expected="阿拉伯数字，接着上一页", actual=text))
        elif last is not None and last[0] != kind and n != 1:
            errors.append(_issue("error", "PAGE_NUMBER_SEQUENCE", page, "换成阿拉伯数字页码时没有从 1 开始。",
                                 expected=first, actual=text))
        elif last is not None and last[0] == kind and n != last[1] + 1:
            errors.append(_issue("error", "PAGE_NUMBER_SEQUENCE", page, "页码不连续（跳号、重复或中途重新开始）。",
                                 expected=f"接着上一页（{last[2]}）往下数", actual=text))
        last = (kind, n, text)
    return errors


def audit(docx, pdf):
    """Return {"errors": [...], "warnings": [...], "pages": n, "fonts": {...}} for one rendered PDF."""
    import pymupdf as fitz

    doc = Document(docx)
    want = _expectations(doc)
    sections = list(doc.sections)
    body_margins = _margins(sections[1:] if want["has_cover"] and len(sections) > 1 else sections)
    cover_margins = _margins(sections[:1])
    header_texts = _header_texts(doc)
    errors, warnings = [], []
    used = set()
    previous = ""
    numbers = []
    with fitz.open(pdf) as rendered:
        pages = len(rendered)
        for index, page in enumerate(rendered, 1):
            for font in page.get_fonts():
                used.add(_key(font[3]))
            cover = want["has_cover"] and index == 1
            m = cover_margins if cover else body_margins
            w, h = page.rect.width, page.rect.height
            if abs(w - m["width"]) > 2 or abs(h - m["height"]) > 2:
                errors.append(
                    _issue("error", "PAGE_SIZE", index, "渲染出的纸张大小和文档设置不一致。",
                           expected=f"{m['width']:.0f}×{m['height']:.0f}pt", actual=f"{w:.0f}×{h:.0f}pt")
                )
            left, right = m["left"], w - m["right"]
            top, bottom = m["top"], h - m["bottom"]
            lines = _lines(page)
            images = [i["bbox"] for i in page.get_image_info() if i["bbox"][2] > i["bbox"][0]]
            drawings = [d["rect"] for d in page.get_drawings() if d["rect"].width > 0 or d["rect"].height > 0]

            for line in lines:
                x0, y0, x1, y1 = line["bbox"]
                text = line["spans"][0]["text"]
                ink = line.get("ink_right", x1)
                if x0 < -0.5 or y0 < -0.5 or x1 > w + 0.5 or y1 > h + 0.5:
                    errors.append(_issue("error", "OFF_PAGE", index, "文字超出纸张边缘，会被裁掉。",
                                         where=text[:30], bbox=line["bbox"]))
                elif cover and (x0 < COVER_SAFE or x1 > w - COVER_SAFE):
                    errors.append(_issue("error", "TEXT_OUTSIDE_MARGINS", index, "封面文字太靠近纸张边缘（不到 1 厘米）。",
                                         where=text[:30], expected=f"{COVER_SAFE:.0f}–{w - COVER_SAFE:.0f}pt",
                                         actual=f"{x0:.1f}–{x1:.1f}pt", bbox=line["bbox"]))
                elif not cover and (x0 < left - TEXT_TOLERANCE or ink > right + TEXT_TOLERANCE):
                    errors.append(_issue("error", "TEXT_OUTSIDE_MARGINS", index, "文字伸出了左右页边距。",
                                         where=text[:30], expected=f"{left:.1f}–{right:.1f}pt",
                                         actual=f"{x0:.1f}–{x1:.1f}pt", bbox=line["bbox"]))
            for box in images:
                x0, y0, x1, y1 = box
                if x0 < -0.5 or y0 < -0.5 or x1 > w + 0.5 or y1 > h + 0.5:
                    errors.append(_issue("error", "IMAGE_OFF_PAGE", index, "图片超出纸张边缘。", bbox=box))
                elif x0 < left - EDGE_TOLERANCE or x1 > right + EDGE_TOLERANCE:
                    errors.append(_issue("error", "IMAGE_OUTSIDE_MARGINS", index, "图片伸出了左右页边距。",
                                         expected=f"{left:.1f}–{right:.1f}pt", actual=f"{x0:.1f}–{x1:.1f}pt", bbox=box))
                elif not cover and (y0 < top - EDGE_TOLERANCE or y1 > bottom + EDGE_TOLERANCE):
                    errors.append(_issue("error", "IMAGE_OUTSIDE_MARGINS", index, "图片伸进了上下页边距（压到页眉页脚区）。",
                                         expected=f"{top:.1f}–{bottom:.1f}pt", actual=f"{y0:.1f}–{y1:.1f}pt", bbox=box))
                else:
                    text = next((l["spans"][0]["text"] for l in lines
                                 if any(_covers(box, _ink(c)) for c in l["chars"])), None)
                    if text:
                        errors.append(_issue("error", "TEXT_OVER_IMAGE", index, "文字压在图片上。",
                                             where=text[:30], bbox=box))
            for rect in drawings:
                if rect.x0 < left - EDGE_TOLERANCE or rect.x1 > right + EDGE_TOLERANCE:
                    errors.append(_issue("error", "TABLE_OUTSIDE_MARGINS", index, "表格框线或图形伸出了左右页边距。",
                                         expected=f"{left:.1f}–{right:.1f}pt", actual=f"{rect.x0:.1f}–{rect.x1:.1f}pt",
                                         bbox=tuple(rect)))
                    break

            body = [l for l in lines if l["bbox"][3] > top and l["bbox"][1] < bottom]
            # Rules, table borders and shapes count as content too; a lone header rule does not.
            body_objects = [b for b in images if b[3] > top and b[1] < bottom] + [
                r for r in drawings if r.y1 > top + 2 and r.y0 < bottom - 2 and (r.height > 2 or r.y0 > top + 30)]
            if not body and not body_objects:
                errors.append(_issue("error", "BLANK_PAGE", index, "这一页正文区什么都没有（多余的空白页）。",
                                     where=("紧跟在“" + previous + "”之后") if previous else None))
            elif body:
                previous = body[-1]["spans"][0]["text"][-20:]

            head = [l for l in lines if l["bbox"][3] <= top + 1]
            foot = [l for l in lines if l["bbox"][1] >= bottom - 1]
            if want["header"] is not None:
                if cover and head:
                    errors.append(_issue("error", "COVER_HEADER_VISIBLE", index, "封面页出现了页眉文字。",
                                         where=head[0]["spans"][0]["text"][:30]))
                elif not cover and want["header"] and not head:
                    errors.append(_issue("error", "HEADER_MISSING", index, "这一页应该有页眉，渲染结果里没有。"))
                elif not cover and want["header"]:
                    spill = _header_overflow(lines, top, header_texts)
                    if spill:
                        errors.append(_issue("error", "HEADER_OVERFLOW", index, "页眉文字折成多行，伸出上边距压进了正文区。",
                                             where=spill["spans"][0]["text"][:30], expected=f"页眉在 {top:.1f}pt 以上",
                                             actual=f"页眉底边 {spill['bbox'][3]:.1f}pt", bbox=spill["bbox"]))
                elif not cover and not want["header"] and head:
                    errors.append(_issue("error", "HEADER_UNEXPECTED", index, "这一页不应该有页眉，渲染结果里有。",
                                         where=head[0]["spans"][0]["text"][:30]))
            if want["page_numbers"] is not None:
                texts = "".join(l["spans"][0]["text"] for l in foot).strip()
                if cover and texts:
                    errors.append(_issue("error", "COVER_FOOTER_VISIBLE", index, "封面页出现了页脚文字或页码。", where=texts[:30]))
                elif not cover and want["page_numbers"] and not re.fullmatch(r"[0-9IVXLCivxlc\-—\s第页]+", texts or "?"):
                    errors.append(_issue("error", "PAGE_NUMBER_MISSING", index, "这一页页脚应该显示页码，渲染结果里没有。",
                                         actual=texts[:30] or "（空）"))
                elif not cover and not want["page_numbers"] and texts:
                    errors.append(_issue("error", "FOOTER_UNEXPECTED", index, "页脚应该是空白，渲染结果里有文字。", where=texts[:30]))

            clash = next(((a, b, gap) for i, a in enumerate(lines) for b in lines[i + 1 :]
                          for gap in [_line_clash(a, b)] if gap), None)
            if clash:
                a, b, (gap, overlap) = clash
                errors.append(_issue("error", "TEXT_OVERLAP", index, "两行文字在页面上叠在一起。",
                                     where=a["spans"][0]["text"][:20] + " / " + b["spans"][0]["text"][:20],
                                     expected="上下两行的字形互不重叠（同一字号时基线至少相距约 0.9 倍字号）",
                                     actual=f"基线相距 {gap:.1f}pt，字形重叠 {overlap:.1f}pt", bbox=a["bbox"]))
            if want["page_numbers"] and not cover:
                numbers.append((index, "".join(l["spans"][0]["text"] for l in foot).strip()))
    errors += _number_sequence(numbers)
    wanted = requested_fonts(doc)
    missing = [f for f in wanted if not _font_present(f, used)]
    if missing:
        warnings.append(
            _issue(
                "warning",
                "FONT_SUBSTITUTED",
                None,
                "渲染这台机器上没有文档要求的字体（"
                + "、".join(missing)
                + "），页图用了替代字体。字宽和分页可能和装有这些字体的 Word 不同，页图只能当近似参考；"
                "Word 文件本身的字体设置没有变。",
                expected="、".join(missing),
                actual="、".join(sorted(used))[:200],
            )
        )
    return {
        "passed": not errors,
        "pages": pages,
        "errors": errors,
        "warnings": warnings,
        "fonts": {"requested": wanted, "missing": missing, "rendered": sorted(used)},
    }


def main():
    p = argparse.ArgumentParser(description="检查实际渲染出的 PDF 版面（越界、空白页、页眉页脚、重叠、字体替代）")
    p.add_argument("docx")
    p.add_argument("pdf")
    p.add_argument("--report")
    a = p.parse_args()
    result = audit(a.docx, a.pdf)
    text = json.dumps(result, ensure_ascii=False, indent=2)
    if a.report:
        Path(a.report).write_text(text, encoding="utf-8")
    print(text)
    return 0 if result["passed"] else 2


if __name__ == "__main__":
    sys.exit(main())
