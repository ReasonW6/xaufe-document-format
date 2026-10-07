"""Pin TOC title/text/link/page/leader formatting; audit effective values independently."""

from __future__ import annotations
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from style_model import run_properties, style_chain
from customization import ROLE_STYLES, normalize_style_options

TOC_ROLES = ("toc_title", "toc1", "toc2", "toc3", "toc4")


def enforce(doc, spec):
    # Writer configuration is used only for writing. audit() uses the separate contract.
    import engine as E

    rules = E.load_json(E.ROOT / spec.get("paragraph_rules", "assets/paragraph-rules.json"))[
        "styles"
    ]
    width = spec["page_twips"]["width"] - spec["page_twips"]["left"] - spec["page_twips"]["right"]
    for role in TOC_ROLES:
        name = ROLE_STYLES[role]
        options = normalize_style_options(
            rules[name], spec.get("style_overrides", {}).get(name, {})
        )
        if options.get("left") == "$toc_indent":
            options["left"] = spec["toc"]["indents_twips"][int(role[-1]) - 1]
        E.configure_style(doc, name, **options)
        rpr = doc.styles[name].element.get_or_add_rPr()
        for tag in (
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
        ):
            E.element(rpr, "w:" + tag, val=0)
        E.element(rpr, "w:u", val="none")
        E.apply_text_options(rpr, spec.get("style_overrides", {}).get(name, {}))
    count = 0
    for p in doc.paragraphs:
        if p.style.name not in [ROLE_STYLES[r] for r in TOC_ROLES]:
            continue
        role = next(r for r in TOC_ROLES if ROLE_STYLES[r] == p.style.name)
        E.style_paragraph(p, role)
        pp = p._p.get_or_add_pPr()
        for tab in list(pp.findall(qn("w:tabs"))):
            pp.remove(tab)
        if role != "toc_title":
            tab = OxmlElement("w:tab")
            tab.set(qn("w:val"), "right")
            tab.set(qn("w:leader"), "dot")
            tab.set(qn("w:pos"), str(width))
            tabs = OxmlElement("w:tabs")
            tabs.append(tab)
            pp.append(tabs)
        cn, size, bold = E.role_format(role)
        # Paragraph mark controls the leader between label and page reference.
        mark = pp.find(qn("w:rPr"))
        if mark is not None:
            pp.remove(mark)
        mark = OxmlElement("w:rPr")
        pp.append(mark)
        props = [mark] + [r.get_or_add_rPr() for r in p._p.xpath(".//w:r")]
        for rp in props:
            for child in list(rp):
                rp.remove(child)
            E.font(rp, cn, size, bold, latin=E.role_latin(role))
            for tag in (
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
            ):
                E.element(rp, "w:" + tag, val=0)
            E.element(rp, "w:u", val="none")
            E.apply_text_options(rp, spec.get("style_overrides", {}).get(ROLE_STYLES[role], {}))
        # Simple fields remain live and editable; opening the file does not auto-regenerate.
        for fld in p._p.xpath(".//w:fldSimple"):
            fld.attrib.pop(qn("w:dirty"), None)
        count += 1
    E.element(doc.settings.element, "w:updateFields", val="false")
    return count


def tabs_for(p):
    result = {}
    layers = [s.element.find(qn("w:pPr")) for s in style_chain(p.style)] + [p._p.pPr]
    for pp in layers:
        if pp is None:
            continue
        for tab in pp.findall("./" + qn("w:tabs") + "/" + qn("w:tab")):
            pos = tab.get(qn("w:pos"))
            value = tab.get(qn("w:val"))
            if value == "clear":
                result.pop(pos, None)
            else:
                result[pos] = (value, tab.get(qn("w:leader"), "none"))
    return result


def _run_audit(p, r, want, label, err):
    props = run_properties(p, r)
    size = str(want["font"]["size_half_points"])
    if props.get("sz") != size or props.get("szCs") != size:
        err("TOC_FONT_SIZE", label + ": 目录实际字号或复杂文字字号不符。")
    for key in ("eastAsia", "ascii", "hAnsi"):
        if props["fonts"].get(key) != want["font"][key] or props["fonts"].get(key + "Theme"):
            err("TOC_FONT", label + ": 目录实际字体或主题字体不符。")
    if (
        props.get("b", False) != want["font"]["bold_required"]
        or props.get("bCs", False) != want["font"]["bold_required"]
    ):
        err("TOC_EMPHASIS", label + ": 目录加粗状态不符合本次要求。")
    if (
        any(
            props.get(k, False)
            for k in ("caps", "smallCaps", "strike", "dstrike", "vanish", "webHidden", "specVanish")
        )
        or props.get("i", False) != want["font"].get("italic", False)
        or props.get("iCs", False) != want["font"].get("italic", False)
        or props.get("u", "none") != want["font"].get("underline", "none")
    ):
        err("TOC_EMPHASIS", label + ": 目录受斜体、下划线、大小写转换或隐藏格式污染。")
    if (
        props.get("vertAlign", "baseline") != "baseline"
        or props.get("position", "0") != "0"
        or props.get("w", "100") != "100"
        or props.get("spacing", "0") != "0"
    ):
        err("TOC_GEOMETRY", label + ": 目录字形缩放、基线或字间距改变。")
    if props.get("color", "000000").upper() not in (
        {"000000", "AUTO"}
        if want["font"].get("color", "000000") == "000000"
        else {want["font"]["color"]}
    ):
        err("TOC_COLOR", label + ": 目录文字不是规定的黑色。")


def audit(doc, spec, err):
    from audit_helpers import contract, paragraph_check

    expected = contract(spec["id"], spec.get("_format_request"))
    width = spec["page_twips"]["width"] - spec["page_twips"]["left"] - spec["page_twips"]["right"]
    actual = [p for p in doc.paragraphs if p.style.name in [ROLE_STYLES[r] for r in TOC_ROLES]]
    if not actual:
        return

    def ignored(*args):
        pass

    # Inspect the style independently even if run-level overrides mask its corruption.
    for role in TOC_ROLES:
        if role == "toc4" and spec.get("heading_levels", 3) < 4:
            continue
        name = ROLE_STYLES[role]
        if name not in doc.styles:
            err("TOC_STYLE_MISSING", "缺少目录样式：" + name)
            continue
        p = Paragraph(OxmlElement("w:p"), doc._body)
        p.style = name
        paragraph_check(p, role, "目录样式 " + name, expected[role], err, ignored)
        r = OxmlElement("w:r")
        p._p.append(r)
        _run_audit(p, r, expected[role], "目录样式 " + name, err)
        if role != "toc_title" and tabs_for(p) != {str(width): ("right", "dot")}:
            err("TOC_TABS", "目录样式制表位不是版心右侧点引导线。")
    for index, p in enumerate(actual):
        role = next(r for r in TOC_ROLES if ROLE_STYLES[r] == p.style.name)
        label = f"目录段{index}"
        if role != "toc_title" and tabs_for(p) != {str(width): ("right", "dot")}:
            err("TOC_TABS", label + ": 实际制表位或点引导线不符。")
        for ri, r in enumerate(p._p.xpath(".//w:r")):
            _run_audit(p, r, expected[role], f"{label}/r{ri}", err)
        mark = p._p.xpath("./w:pPr/w:rPr")
        if mark:
            from copy import deepcopy

            r = OxmlElement("w:r")
            r.append(deepcopy(mark[0]))
            _run_audit(p, r, expected[role], label + "/段落标记", err)
        # Do not require mark presence in paragraphs; actual glyphs/styles still checked.
