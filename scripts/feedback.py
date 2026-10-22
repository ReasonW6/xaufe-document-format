"""Turn acceptance failures into issues a model can act on.

Each issue says where (page / paragraph / text / which input to edit), what was
expected and what was found, the likely cause, and a fix that edits the inputs
(job.json, content.json, mapping.json), never the generated Word.
"""

from __future__ import annotations
import json, re
from pathlib import Path

MAX_ROUNDS = 5  # consecutive failed runs before the agent must stop and report


class AcceptanceFailed(ValueError):
    """The document did not pass an acceptance check; ``issues`` explain each failure."""

    def __init__(self, stage, issues):
        self.stage = stage
        self.issues = issues
        super().__init__(f"{stage}未通过（{len(issues)} 个问题）：\n" + render_text(issues))


_STALE = ("任务记录和成品对不上，通常是改了 job.json 后没有重新 run。", "重新运行 run；不要手改 WORK 里的文件。")
_DIRECT = (
    "这段在原稿里带着直接格式（或者角色选错了），排版后仍和模板不一致。",
    "先看 mapping.json 里这一段的 role 是否正确（标题、正文、摘要、参考文献）；角色没错仍失败，说明是脚本不支持的原稿格式，按 references/special-cases.md 的“脚本做不到的要求”处理并告诉用户。",
)
_USER_STYLE = (
    "实际格式和“模板默认值 + job.json 里的特别要求”不一致。",
    "核对 job.json 的 format.styles 写法（references/job.md）；用户没要求改的项不要写。",
)
HINTS = {
    # static audit
    "PARAGRAPH": _DIRECT,
    "PARAGRAPH_EXTRA": _DIRECT,
    "STYLE_INHERITANCE": _DIRECT,
    "FONT_SIZE": _USER_STYLE,
    "FONT_CN": _USER_STYLE,
    "FONT_EN": _USER_STYLE,
    "FONT_BOLD": _USER_STYLE,
    "FONT_COLOR": _USER_STYLE,
    "FONT_ITALIC": _USER_STYLE,
    "FONT_UNDERLINE": _USER_STYLE,
    "FONT_GEOMETRY": _DIRECT,
    "HIDDEN_TEXT": ("原稿里有隐藏文字。", "问用户这段隐藏文字要不要；要就在 Word 里取消隐藏后重新给，不要就经用户同意把那段 role 设为 delete。"),
    "UNKNOWN_STYLE": ("有段落没有分到任何角色。", "在 mapping.json 里给这一段一个 role（正文用 body）。"),
    "MISPLACED_COVER_STYLE": ("封面以外的段落用了封面样式。", "把 mapping.json 里这一段改成正确的 role。"),
    "HEADING_DEPTH": ("出现了第四级标题，但没有开启四级。", "用户要四级标题就在 job.json 写 \"heading_levels\": 4；否则把这一段改成三级以内。"),
    "KEYWORD_SPACING": ("关键词之间的分隔不对。", "content.json 用 keywords_zh 列表写关键词；已有 Word 时核对这一段的 role 是 keywords_zh。"),
    "KEYWORD_RUN": ("“关键词：”标签和词项混在一个文字段里。", "重新 run；仍失败就把原稿这一段改成“关键词：词1 词2”的形式。"),
    "PAGE": ("页面尺寸或页边距和本次有效设置不一致。", "核对 job.json 的 page_twips（1 厘米 = 567）；用户没要求就删掉这一项。"),
    "HEADER": ("页眉缺失或内容不对。", "A 版页眉用专业和题目：确认 job.json 的 info 里有 major、title_zh；用户自定义页眉就核对 format.header 写法。"),
    "HEADER_TEXT": ("页眉文字和封面信息或 header 设置不一致。", "核对 job.json 的 major、title_zh 和 format.header.left/right。"),
    "HEADER_BORDER": ("页眉横线和设置不一致。", "用户不要横线写 \"header\": {\"border\": false}；否则删掉这项。"),
    "RED_HEADER": ("不该有页眉的地方出现了页眉。", "B 版默认无页眉；用户要页眉才写 format.header.enabled=true。"),
    "COVER_HEADER": ("封面出现了页眉。", "重新 run；仍有就检查原稿是不是自带页眉的旧封面（mapping 的 replace_cover_before）。"),
    "FOOTER": ("页脚或页码和设置不一致。", "要页码写 \"footer_page_numbers\": true，不要就删掉这一项。"),
    "FOOTER_FIELD": ("页脚里残留了页码域。", "重新 run；原稿自带页脚时检查是否带进来了。"),
    "COVER": ("封面结构不完整。", "重新 run；不要手改封面。"),
    "COVER_VALUE": ("封面上的信息和 job.json 不一致。", "核对 job.json 的 info 和 blank；改完重新 run。"),
    "COVER_DATE": ("封面日期和要求不一致。", "日期写成 \"date\": \"2026-10-07\"；要留空就把 date 放进 blank。"),
    "COVER_LAYOUT": ("封面版式和母版不一致。", "重新 run；仍失败说明封面适配有问题，请把题目或课程名缩短后再试，或问用户能否缩小封面字号（format.cover_fonts）。"),
    "COVER_RUNS": ("封面出现了母版以外的文字段。", "重新 run；不要手改封面。"),
    "COVER_FONT": ("封面字体和母版不一致。", "用户要改封面字体就写 format.cover_fonts；否则删掉这一项后重新 run。"),
    "COVER_EFFECTIVE_FONT": ("封面字体被其他样式覆盖。", "同上：核对 format.cover_fonts。"),
    "COVER_TITLE": ("封面题目标签不对。", "核对 format.cover_title_label。"),
    "COVER_LABEL": ("封面文种（如“期末大作业”）不对。", "核对 format.cover_label。"),
    "COVER_BREAK": ("封面没有单独成节。", "重新 run；原稿的旧封面要在 mapping 里用 replace_cover_before 替换。"),
    "COVER_DATE_POSITION": ("封面日期段落缺失。", "重新 run。"),
    "COVER_DATE_LAYOUT": ("封面日期的位置属性不对。", "重新 run；不要手改封面。"),
    "COVER_DATE_BREAK": ("日期段落带了分页。", "重新 run。"),
    "COVER_FRAME": ("封面以外出现了定位框（浮动段落）。", "原稿有浮动段落：请用户改成普通段落，或经同意删掉。"),
    "LOGO": ("校徽图片不是选定的原图。", "job.json 的 logo 只能是 red 或 green；不要替换图片文件。"),
    "LOGO_SIZE": ("校徽尺寸不对。", "用户要改大小就写 format.logo_width_cm；否则删掉这一项。"),
    "LOGO_FLOAT": ("校徽变成了浮动图片。", "重新 run；不要手改封面。"),
    "LOGO_CROP": ("校徽被裁切。", "重新 run；不要手改封面。"),
    "LOGO_ALIGN": ("校徽对齐方式不对。", "用户要居中写 format.logo_alignment；否则删掉这一项。"),
    "TABLE_WIDTH": ("表格比版心宽。", "已有 Word 时确认 mapping 的 format_tables 是 true；仍超宽就请用户减少列数或拆表。"),
    "TABLE_GRID": ("表格列宽信息无效。", "原稿表格结构异常：请用户在 Word 里重新插入这个表格。"),
    "TABLE_FLOAT": ("原稿有浮动表格。", "请用户把表格改成“嵌入文本行中”（取消文字环绕）后再给。"),
    "TABLE_ALIGNMENT": ("表格没有居中。", "确认 mapping 的 format_tables 是 true 后重新 run。"),
    "TABLE_INDENT": ("表格有额外缩进。", "确认 format_tables 是 true 后重新 run。"),
    "TABLE_SPACING": ("表格单元格外间距会扩宽表格。", "确认 format_tables 是 true；仍失败请用户去掉单元格间距。"),
    "TABLE_FIXED_HEIGHT": ("表格行是固定行高，文字可能被裁掉。", "确认 format_tables 是 true；仍失败请用户把行高改成“最小值”。"),
    "TABLE_STRUCTURE": ("表格结构脚本读不懂。", "请用户简化这个表格（拆掉嵌套或复杂合并）。"),
    "CELL_WIDTH": ("单元格宽度和表格网格不一致或太窄。", "确认 format_tables 是 true；列太多就请用户拆表。"),
    "CELL_LAYOUT": ("单元格设置了不换行或压缩文字。", "确认 format_tables 是 true 后重新 run。"),
    "CELL_IMAGE_WIDTH": ("表格里的图片比单元格宽。", "请用户把这张图改小，或从表格里移出来。"),
    "CELL_IMAGE_SPACING": ("表格里的图片留白会撑出边界。", "确认 format_tables 是 true 后重新 run。"),
    "IMAGE_WIDTH": ("图片比版心宽。", "content.json 的图片会自动缩放；已有 Word 时确认这一段 role 是 figure。"),
    "OBJECT_HEIGHT": ("大图片放在固定行距的段落里，会被裁掉。", "把这一段的 role 改成 figure。"),
    "PART_MISSING": ("完整论文结构缺了一部分（摘要、目录、正文或参考文献）。", "写作时补上缺的部分；只排版就把 mode 改成 \"document\"，不要替用户补写。"),
    "PART_ORDER": ("各部分顺序和版式不一致。", "删掉 format.parts（用默认顺序），或按用户要求核对它的写法。"),
    "PART_BREAKS": ("封面、摘要、目录、正文没有各自起页。", "重新 run；mapping 的 mode 应为 paper。"),
    "PLACEHOLDER": ("模板里的说明或示例文字留在了成品里。", "从 content.json / 原稿里删掉这段占位文字（只排版时先问用户）。"),
    "TOC_SCOPE": ("目录条目和标题对不上。", "检查标题的 role/level 是否正确后重新 run。"),
    "TOC_UNRESOLVED": ("目录页码没算出来（没有实际分页）。", "有办公软件就重新 run；没有就是静态交付，如实告诉用户。"),
    "TOC_PROOF_STALE": ("目录页码和实际分页不一致。", "重新 run。"),
    "TOC_FONT": _USER_STYLE,
    "TOC_FONT_SIZE": _USER_STYLE,
    "TOC_EMPHASIS": _USER_STYLE,
    "TOC_COLOR": _USER_STYLE,
    "TOC_GEOMETRY": _DIRECT,
    "TOC_TABS": ("目录点引导线或右侧页码位置不对。", "重新 run；改过 page_twips 时核对写法。"),
    "TOC_STYLE_MISSING": ("目录样式缺失。", "重新 run。"),
    "UNSUPPORTED_STRUCTURE": ("原稿里有脚本不能安全排版的对象（文本框、修订、批注、内容控件等）。", "按 references/special-cases.md 的“输入文件”一节处理：请用户先接受修订、把文本框改成普通文字，再重新给。"),
    "EMBEDDED_FONT": ("原稿嵌入了字体文件。", "请用户在 Word“选项 → 保存”里取消“嵌入字体”后重新给。"),
    "AUTO_FIELD_UPDATE": _STALE,
    "CUSTOMIZATION_HASH": _STALE,
    "CUSTOMIZATION_STALE": _STALE,
    "METADATA_STALE": _STALE,
    "TEMPLATE_MISMATCH": _STALE,
    "JOB_MISMATCH": _STALE,
    "WORKFLOW": _STALE,
    "INTAKE": _STALE,
    "INTAKE_REQUIRED": _STALE,
    "DOCUMENT_SCOPE": ("有没有封面和用户要求的不一致。", "核对 job.json 的 cover。"),
    "RED_ENGLISH_ABSTRACT": ("出现了英文摘要，但本次设置不要英文摘要。", "用户要英文摘要就写 \"english_abstract\": true；不要就去掉英文摘要（已有 Word 时问用户后把这些段落设为 delete）。"),
    "OPEN": ("成品打不开或不是本工具生成的。", "重新 run。"),
    "AUDIT_ABORTED": ("检查器读不懂文档结构。", "重新 run；仍失败就把错误原文告诉用户，不要绕过检查交付。"),
    # rendered-page audit
    "PAGE_SIZE": ("办公软件按别的纸张大小导出了。", "核对 job.json 的 page_twips；仍不对就换一个办公软件（--office word/wps/libreoffice）。"),
    "OFF_PAGE": ("内容超出纸张，打印会被裁掉。", "通常是超宽表格、长网址或图片：拆表、给网址加空格断行，或缩小图片；改 content.json / 原稿后重新 run。"),
    "TEXT_OUTSIDE_MARGINS": ("文字伸进了页边距：常见于不能断行的长英文、网址、公式或代码。", "在 content.json 里给长串加断行（代码块用 code 类型），或请用户同意缩短；封面就缩短这一项或问用户能否缩小封面字号。"),
    "IMAGE_OFF_PAGE": ("图片超出纸张。", "把图片改小，或确认这一段 role 是 figure。"),
    "IMAGE_OUTSIDE_MARGINS": ("图片比版心宽。", "把图片改小，或确认这一段 role 是 figure（脚本会按版心缩放）。"),
    "TABLE_OUTSIDE_MARGINS": ("表格或图形伸出了版心。", "已有 Word 时确认 format_tables 是 true；列太多就请用户拆表；浮动图形请用户改成嵌入式。"),
    "BLANK_PAGE": ("出现了多余的空白页：常见于连续分页符、段落自带的“段前分页”，或分节后又手动分页。", "删掉 content.json 里多余的 page_break；已有 Word 时把那段空段落的 role 设为 spacer 或经用户同意设为 delete。"),
    "HEADER_MISSING": ("应有页眉的页面没有页眉。", "核对 job.json 的 format.header 和封面信息（A 版页眉要 major、title_zh）。"),
    "HEADER_UNEXPECTED": ("不该有页眉的页面出现了页眉文字。", "B 版默认无页眉；检查原稿是否带页眉，以及 format.header 写法。"),
    "COVER_HEADER_VISIBLE": ("封面上出现了页眉。", "重新 run；原稿有旧封面时用 mapping 的 replace_cover_before 替换。"),
    "COVER_FOOTER_VISIBLE": ("封面上出现了页脚或页码。", "重新 run；仍出现就是原稿封面自带页脚，请用户删掉后再给。"),
    "PAGE_NUMBER_MISSING": ("用户要页码，但渲染结果里页脚没有页码。", "确认 job.json 写了 \"footer_page_numbers\": true；换办公软件再 run 一次排除渲染差异。"),
    "FOOTER_UNEXPECTED": ("页脚应该空白，渲染结果里有文字。", "用户没要页码时删掉 footer_page_numbers；原稿自带页脚文字时请用户删掉。"),
    "TEXT_OVERLAP": ("两行文字在页面上叠在一起：常见于过小的固定行距、负的段前段后，或浮动对象。", "打开这一页的页图确认；是真的重叠就检查 job.json 里 line/before/after 的写法（固定行距不能小于字号）。"),
    "FONT_SUBSTITUTED": ("渲染机器上没有文档要求的字体。", "不用改文档。在 review.json 的 warnings 里写明页图是用替代字体看的，交付时告诉用户分页以装有这些字体的 Word 为准。"),
}

_LABEL = re.compile(r"^(?P<label>(?:p\d+|t\d+(?:/[rcpt]\d+)*|header\d+/\w+/p\d+)(?:/r\d+)?)\s*[:：]\s*(?P<rest>.*)$")
_EXPECT = re.compile(r"(?P<what>[\w/]+)=(?P<actual>[^，,]+?)(?:半磅)?，应为(?P<expected>[^。]+)。?")


def _paragraph(doc, label):
    from table_layout import walk_paragraphs

    base = re.sub(r"/r\d+$", "", label)
    if re.fullmatch(r"p\d+", base):
        index = int(base[1:])
        return doc.paragraphs[index] if index < len(doc.paragraphs) else None
    if base.startswith("t"):
        top = int(re.match(r"t(\d+)", base).group(1))
        if top < len(doc.tables):
            for found, p in walk_paragraphs(doc.tables[top], f"t{top}"):
                if found == base:
                    return p
    m = re.fullmatch(r"header(\d+)/(\w+)/p(\d+)", base)
    if m and int(m.group(1)) < len(doc.sections):
        section = doc.sections[int(m.group(1))]
        part = {"default": section.header, "first": section.first_page_header, "even": section.even_page_header}.get(m.group(2))
        if part is not None and int(m.group(3)) < len(part.paragraphs):
            return part.paragraphs[int(m.group(3))]
    return None


def _input_hint(root, text):
    """Which input holds this text: content.json block or mapping.json paragraph."""
    key = re.sub(r"\s+", "", text)[:12]
    if not key or root is None:
        return None
    content, mapping = Path(root) / "content.json", Path(root) / "mapping.json"
    try:
        if content.is_file():
            data = json.loads(content.read_text(encoding="utf-8-sig"))
            for i, block in enumerate(data.get("blocks", [])):
                body = block.get("text") or "".join(r.get("text", "") for r in block.get("runs", []))
                body += " ".join(str(c) for row in [block.get("header") or []] + (block.get("rows") or []) for c in row)
                if key in re.sub(r"\s+", "", body):
                    return f"content.json 的 blocks[{i}]（{block.get('type')}）"
            for field in ("abstract_zh", "abstract_en", "keywords_zh", "keywords_en", "references"):
                if key in re.sub(r"\s+", "", json.dumps(data.get(field, ""), ensure_ascii=False)):
                    return f"content.json 的 {field}"
        elif mapping.is_file():
            data = json.loads(mapping.read_text(encoding="utf-8-sig"))
            for row in data.get("paragraphs", []):
                if key[:8] and re.sub(r"\s+", "", row.get("text", "")).startswith(key[:8]):
                    return f"mapping.json 里 id 为 {row.get('id')} 的段落（当前 role：{row.get('role')}）"
    except (OSError, ValueError):
        return None
    return None


def explain(errors, docx=None, root=None, stage="格式检查"):
    """Static-audit errors -> actionable issues."""
    doc = None
    if docx:
        try:
            from docx import Document

            doc = Document(docx)
        except Exception:
            doc = None
    issues = []
    for e in errors:
        code, message = e.get("code", "?"), e.get("message", "")
        cause, fix = HINTS.get(code, ("", "看 message 的说明修改 job.json / content.json / mapping.json 后重新 run。"))
        item = {"code": code, "stage": stage, "message": message, "cause": cause, "fix": fix}
        m = _LABEL.match(message)
        if m:
            item["where"] = {"location": m.group("label")}
            p = _paragraph(doc, m.group("label")) if doc is not None else None
            if p is not None:
                text = "".join(p._p.xpath(".//w:t/text()")).strip()
                item["where"]["text"] = text[:40] or "（空段落）"
                item["where"]["style"] = p.style.name if p.style is not None else ""
                source = _input_hint(root, text)
                if source:
                    item["where"]["edit"] = source
        x = _EXPECT.search(message)
        if x:
            item["expected"] = x.group("expected").strip()
            item["actual"] = x.group("actual").strip()
        issues.append(item)
    return issues


def from_render(audit, root=None):
    """Rendered-page audit errors -> actionable issues (they already carry page/expected/actual)."""
    issues = []
    for e in audit.get("errors", []):
        cause, fix = HINTS.get(e["code"], ("", ""))
        item = {"code": e["code"], "stage": "渲染后版面检查", "message": e["message"], "cause": cause, "fix": fix,
                "where": {"page": e.get("page")}}
        if e.get("where"):
            item["where"]["text"] = e["where"]
            quoted = re.search(r"“(.+)”", e["where"])
            source = _input_hint(root, quoted.group(1)[-12:] if quoted else e["where"])
            if source:
                item["where"]["edit"] = source
        if e.get("bbox_pt"):
            item["where"]["bbox_pt"] = e["bbox_pt"]
        for key in ("expected", "actual"):
            if key in e:
                item[key] = e[key]
        issues.append(item)
    return issues


def render_text(issues, limit=12):
    """Readable summary for the agent; the full list is in WORK/issues.json."""
    lines = []
    for n, item in enumerate(issues[:limit], 1):
        where = item.get("where") or {}
        place = "，".join(
            x for x in (
                f"第{where['page']}页" if where.get("page") else "",
                where.get("location", ""),
                (where["text"] if "“" in where["text"] else f"“{where['text']}”") if where.get("text") else "",
            ) if x
        )
        lines.append(f"{n}. [{item['code']}] {item['message']}")
        if place:
            lines.append(f"   位置：{place}")
        if where.get("edit"):
            lines.append(f"   要改的输入：{where['edit']}")
        if "expected" in item or "actual" in item:
            lines.append(f"   实际：{item.get('actual', '?')}；应为：{item.get('expected', '?')}")
        if item.get("cause"):
            lines.append(f"   可能原因：{item['cause']}")
        if item.get("fix"):
            lines.append(f"   怎么改：{item['fix']}")
    if len(issues) > limit:
        lines.append(f"……另有 {len(issues) - limit} 个问题，见 issues.json。")
    return "\n".join(lines)


def record_round(root, issues):
    """Count consecutive failed runs; returns (rounds, repeated) where repeated = same codes as last time."""
    path = Path(root) / "attempts.json"
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        data = {}
    codes = sorted({i["code"] for i in issues})
    rounds = int(data.get("failed_rounds", 0)) + 1 if issues else 0
    repeated = bool(issues) and codes == data.get("last_codes")
    path.write_text(json.dumps({"failed_rounds": rounds, "last_codes": codes}, ensure_ascii=False), encoding="utf-8")
    return rounds, repeated
