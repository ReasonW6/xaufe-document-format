"""Read effective Word properties without calling the formatting writer.

The checker resolves document defaults, paragraph/character basedOn chains and
run overrides. Style boolean toggles and direct absolute settings differ.
Conditional table/numbering overrides that are not normalized are reported by
additional structural checks rather than declared universally supported.
"""

from __future__ import annotations
from docx.oxml.ns import qn

TOGGLES = {
    "b",
    "bCs",
    "i",
    "iCs",
    "caps",
    "smallCaps",
    "strike",
    "dstrike",
    "outline",
    "shadow",
    "emboss",
    "imprint",
    "vanish",
}


def on(value):
    return value is None or str(value).lower() in ("1", "true", "on")


def style_chain(style):
    chain = []
    seen = set()
    while style is not None:
        if style.style_id in seen:
            raise ValueError("样式继承存在循环：" + style.style_id)
        seen.add(style.style_id)
        chain.append(style)
        style = style.base_style
    return list(reversed(chain))


def style_by_id(doc, style_id):
    matches = [s for s in doc.styles if s.style_id == style_id]
    if not matches:
        raise ValueError("引用了不存在的样式：" + style_id)
    return matches[0]


def defaults(doc, kind):
    paths = doc.styles.element.xpath(f"./w:docDefaults/w:{kind}Default/w:{kind}")
    return paths[0] if paths else None


def paragraph_properties(p, tag):
    doc = p.part.package.main_document_part.document
    result = {}
    layers = (
        [defaults(doc, "pPr")]
        + [s.element.find(qn("w:pPr")) for s in style_chain(p.style)]
        + [p._p.pPr]
    )
    for pp in layers:
        if pp is None:
            continue
        node = pp.find(qn("w:" + tag))
        if node is not None:
            if tag not in ("ind", "spacing"):
                result = {}
            result.update({k.split("}")[-1]: v for k, v in node.attrib.items()})
            if not node.attrib and tag not in ("ind", "spacing", "tabs"):
                result["val"] = "1"
    if tag == "ind":
        # Modern OOXML uses start/end; resolve to physical sides for this paragraph.
        bidi = paragraph_properties(p, "bidi")
        rtl = bool(bidi) and on(bidi.get("val"))
        aliases = {
            "start": "right" if rtl else "left",
            "end": "left" if rtl else "right",
            "startChars": "rightChars" if rtl else "leftChars",
            "endChars": "leftChars" if rtl else "rightChars",
        }
        for source, target in aliases.items():
            if source in result:
                result[target] = result.pop(source)
    return result


def run_properties(p, r):
    doc = p.part.package.main_document_part.document
    result = {}
    fonts = {}

    def apply(rp, is_style=False):
        if rp is None:
            return
        for el in rp:
            tag = el.tag.split("}")[-1]
            if tag == "rFonts":
                values = {k.split("}")[-1]: v for k, v in el.attrib.items()}
                themes = {
                    "ascii": "asciiTheme",
                    "hAnsi": "hAnsiTheme",
                    "eastAsia": "eastAsiaTheme",
                    "cs": "cstheme",
                }
                for key, value in values.items():
                    if key in themes and themes[key] not in values:
                        fonts.pop(themes[key], None)
                    for direct, theme in themes.items():
                        if key == theme:
                            fonts.pop(direct, None)
                    fonts[key] = value
            elif tag != "rStyle":
                value = el.get(qn("w:val"))
                if tag in TOGGLES and is_style:
                    if on(value):
                        result[tag] = not result.get(tag, False)
                elif tag in TOGGLES or tag in ("webHidden", "rtl", "cs", "specVanish"):
                    result[tag] = on(value)
                else:
                    result[tag] = value if value is not None else "1"

    apply(defaults(doc, "rPr"))
    # Table style is lower priority than explicit paragraph/character/run settings.
    ancestor = r.getparent()
    while ancestor is not None and ancestor.tag != qn("w:tbl"):
        ancestor = ancestor.getparent()
    if ancestor is not None:
        ids = ancestor.xpath("./w:tblPr/w:tblStyle/@w:val")
        if ids:
            for s in style_chain(style_by_id(doc, ids[0])):
                apply(s.element.find(qn("w:rPr")), True)
    for s in style_chain(p.style):
        apply(s.element.find(qn("w:rPr")), True)
    rp = r.find(qn("w:rPr"))
    if rp is not None:
        rs = rp.find(qn("w:rStyle"))
        if rs is not None:
            for s in style_chain(style_by_id(doc, rs.get(qn("w:val")))):
                apply(s.element.find(qn("w:rPr")), True)
    apply(rp)
    result["fonts"] = fonts
    return result
