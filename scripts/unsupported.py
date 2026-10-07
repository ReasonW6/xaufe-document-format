"""Conservative boundaries shared by writer and audit, never flatten unsupported content."""

from docx.oxml.ns import qn
from style_model import style_by_id, style_chain, on


def unsupported_reasons(doc):
    out = []
    body = doc.element.body
    if body.xpath(".//w:sdt|.//w:altChunk|.//w:object"):
        out.append("含内容控件、嵌入文档或OLE对象；通用路径不自动改写，须保全对象的专门流程。")
    if body.xpath(".//w:ins|.//w:del|.//w:moveFrom|.//w:moveTo|.//w:pPrChange|.//w:rPrChange"):
        out.append("含未处理修订；不得静默接受或删除修订。")
    if body.xpath(".//wp:anchor|.//w:pict|.//w:txbxContent"):
        out.append("含浮动图片、VML或文本框；仅内嵌图片进入本通用排版验收。")
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
