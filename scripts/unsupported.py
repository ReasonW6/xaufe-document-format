"""Conservative boundaries shared by writer and audit, never flatten unsupported content."""

from docx.oxml.ns import qn
from style_model import style_by_id, style_chain, on


def where(doc, nodes, limit=3):
    """"p3“开头文字”、t0（表格）" for the body blocks holding ``nodes``, as mapping.json numbers them."""
    body = doc.element.body
    names, pno, tno = {}, 0, 0
    for el in body:
        if el.tag == qn("w:p"):
            names[el] = f"p{pno}“{''.join(el.xpath('.//w:t/text()')).strip()[:12]}”"
            pno += 1
        elif el.tag == qn("w:tbl"):
            names[el] = f"t{tno}（表格）"
            tno += 1
    found = []
    for node in nodes:
        while node is not None and node.getparent() is not body:
            node = node.getparent()
        if node in names and names[node] not in found:
            found.append(names[node])
    return "、".join(found[:limit]) + ("等" if len(found) > limit else "")


def unsupported_reasons(doc):
    out = []
    body = doc.element.body
    found = body.xpath(".//w:sdt|.//w:altChunk|.//w:object")
    if found:
        out.append(f"含内容控件、嵌入文档或OLE对象（{where(doc, found)}）；通用路径不自动改写，须保全对象的专门流程。")
    found = body.xpath(".//w:ins|.//w:del|.//w:moveFrom|.//w:moveTo|.//w:pPrChange|.//w:rPrChange")
    if found:
        out.append(f"含未处理修订（{where(doc, found)}）；不得静默接受或删除修订："
                   "请用户在 Word 的“审阅”里接受或拒绝这些修订后重新给原稿。")
    found = body.xpath(".//wp:anchor|.//w:pict|.//w:txbxContent")
    if found:
        out.append(f"含浮动图片、VML或文本框（{where(doc, found)}）；仅内嵌图片进入本通用排版验收。")
    # Only the designated date paragraph of a recognized skill cover may be framed.
    variables = {
        e.get(qn("w:name")): e.get(qn("w:val"))
        for e in doc.settings.element.xpath("./w:docVars/w:docVar")
    }
    allowed_date = None
    if (
        variables.get("XAUFE_HAS_COVER") == "1"
        and variables.get("XAUFE_TEMPLATE") in ("A", "B")
        and len(doc.paragraphs) > 14
    ):
        p = doc.paragraphs[14]
        if p.text.startswith("完成日期："):
            allowed_date = p._p
    for frame in body.xpath(".//w:framePr"):
        if frame.getparent().getparent() is not allowed_date:
            out.append("存在封面日期以外的段落定位框；通用路径不静默改变其位置。")
    # Paragraph/run direct formatting wins for normal font settings, but conditional
    # hidden text cannot be safely made visible without explicit content consent.
    for sid in set(body.xpath(".//w:tblPr/w:tblStyle/@w:val")):
        try:
            for style in style_chain(style_by_id(doc, sid)):
                for el in style.element.xpath(
                    "./w:tblStylePr/w:rPr/w:vanish|./w:tblStylePr/w:rPr/w:webHidden|./w:tblStylePr/w:rPr/w:specVanish"
                ):
                    if on(el.get(qn("w:val"))):
                        out.append("表格条件样式可能隐藏文字；先确认显示意图，再专门处理。")
        except ValueError as exc:
            out.append(str(exc))
    return list(dict.fromkeys(out))
