"""Place the existing date paragraph at the cover's printable bottom, not in a footer."""

from __future__ import annotations
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# CT_PPr child order. Only the date paragraph is changed; its text/runs stay intact.
_ORDER = [
    "pStyle",
    "keepNext",
    "keepLines",
    "pageBreakBefore",
    "framePr",
    "widowControl",
    "numPr",
    "suppressLineNumbers",
    "pBdr",
    "shd",
    "tabs",
    "suppressAutoHyphens",
    "kinsoku",
    "wordWrap",
    "overflowPunct",
    "topLinePunct",
    "autoSpaceDE",
    "autoSpaceDN",
    "bidi",
    "adjustRightInd",
    "snapToGrid",
    "spacing",
    "ind",
    "contextualSpacing",
    "mirrorIndents",
    "suppressOverlap",
    "jc",
    "textDirection",
    "textAlignment",
    "textboxTightWrap",
    "outlineLvl",
    "divId",
    "cnfStyle",
    "rPr",
    "sectPr",
    "pPrChange",
]


def apply(doc, spec):
    """Idempotent positional repair. Keep the date text, font, and all non-date blocks."""
    index = spec["cover"]["date_paragraph"]
    if len(doc.paragraphs) < spec["cover"]["paragraph_count"]:
        raise ValueError("封面结构不完整，不能定位日期。")
    paragraph = doc.paragraphs[index]
    if not paragraph.text.startswith("完成日期："):
        raise ValueError("封面日期段落不符合本模板；不能移动其他内容冒充日期。")
    ppr = paragraph._p.get_or_add_pPr()
    layout = spec["cover"]["date_layout"]
    properties = layout["paragraph"]
    # A direct pageBreakBefore=0 preceding framePr triggers a spurious page
    # in the tested LibreOffice renderer. Absence means no break; audit inheritance.
    for old in list(ppr.findall(qn("w:pageBreakBefore"))):
        ppr.remove(old)
    frame = dict(layout["frame"])
    page = spec["page_twips"]
    frame["w"] = str(page["width"] - page["left"] - page["right"])
    for tag, attributes in dict(properties, framePr=frame).items():
        for old in list(ppr.findall(qn("w:" + tag))):
            ppr.remove(old)
        child = OxmlElement("w:" + tag)
        for key, value in attributes.items():
            child.set(qn("w:" + key), str(value))
        ppr.append(child)
    ordered = {qn("w:" + name): i for i, name in enumerate(_ORDER)}
    children = sorted(list(ppr), key=lambda child: ordered.get(child.tag, len(_ORDER)))
    for child in list(ppr):
        ppr.remove(child)
    ppr.extend(children)
    return paragraph
