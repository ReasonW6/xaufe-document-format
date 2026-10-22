#!/usr/bin/env python3
"""Rendered-page audit: reads only the exported PDF and the delivered DOCX.

It does not call the layout engine.  The page box and margins come from the
DOCX section properties (the static audit already ties those to the template),
and the header/footer expectations come from the recorded effective profile.
Everything else is measured on the PDF a real office suite produced, so a
wrong assumption in the generator cannot hide a visible defect here.

Errors are things a person would see as broken: content outside the page or
the text column, a page with nothing in the text area, missing or unexpected
header/footer text.  Warnings need a human or model look: overlapping lines
and fonts the renderer substituted.
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


def _margins(doc):
    s = doc.sections[-1]
    return {
        "width": s.page_width.twips / PT,
        "height": s.page_height.twips / PT,
        "left": s.left_margin.twips / PT,
        "right": s.right_margin.twips / PT,
        "top": s.top_margin.twips / PT,
        "bottom": s.bottom_margin.twips / PT,
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
            chars = [c for span in line["spans"] for c in span["chars"] if c["c"].strip()]
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
                }
            )
    return result


def _overlaps(a, b):
    width = min(a[2], b[2]) - max(a[0], b[0])
    height = min(a[3], b[3]) - max(a[1], b[1])
    if width <= 0 or height <= 0:
        return False
    smaller_h = min(a[3] - a[1], b[3] - b[1])
    smaller_w = min(a[2] - a[0], b[2] - b[0])
    return height > 0.5 * smaller_h and width > 0.3 * smaller_w


def audit(docx, pdf):
    """Return {"errors": [...], "warnings": [...], "pages": n, "fonts": {...}} for one rendered PDF."""
    import pymupdf as fitz

    doc = Document(docx)
    want = _expectations(doc)
    m = _margins(doc)
    errors, warnings = [], []
    used = set()
    previous = ""
    with fitz.open(pdf) as rendered:
        pages = len(rendered)
        for index, page in enumerate(rendered, 1):
            for font in page.get_fonts():
                used.add(_key(font[3]))
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
            cover = want["has_cover"] and index == 1

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
            for rect in drawings:
                if rect.x0 < left - EDGE_TOLERANCE or rect.x1 > right + EDGE_TOLERANCE:
                    errors.append(_issue("error", "TABLE_OUTSIDE_MARGINS", index, "表格框线或图形伸出了左右页边距。",
                                         expected=f"{left:.1f}–{right:.1f}pt", actual=f"{rect.x0:.1f}–{rect.x1:.1f}pt",
                                         bbox=tuple(rect)))
                    break

            body = [l for l in lines if l["bbox"][3] > top and l["bbox"][1] < bottom]
            body_objects = [b for b in images if b[3] > top and b[1] < bottom]
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

            for i, a in enumerate(lines):
                for b in lines[i + 1 :]:
                    if _overlaps(a["bbox"], b["bbox"]):
                        warnings.append(_issue("warning", "TEXT_OVERLAP", index, "两行文字在页面上叠在一起，打开这一页的页图确认。",
                                               where=a["spans"][0]["text"][:20] + " / " + b["spans"][0]["text"][:20],
                                               bbox=a["bbox"]))
                        break
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
    # Collapse repeated overlap reports to one per page.
    seen, unique = set(), []
    for w in warnings:
        key = (w["code"], w["page"])
        if key not in seen:
            seen.add(key)
            unique.append(w)
    return {
        "passed": not errors,
        "pages": pages,
        "errors": errors,
        "warnings": unique,
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
