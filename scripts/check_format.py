#!/usr/bin/env python3
"""Profile-aware static audit. No implicit template; no visual-pass claims."""

from __future__ import annotations
import argparse, hashlib, json, re, sys
from pathlib import Path
from zipfile import ZipFile
from docx import Document
from docx.oxml.ns import qn
import engine as E
from intake import require as require_intake, digest
from audit_helpers import (
    contract,
    paragraph_check,
    run_check,
    tables_check,
    toc_entries,
    toc_proof_check,
)
from table_layout import walk_paragraphs
from style_model import paragraph_properties, run_properties
from cover_date_audit import static as check_cover_date
from customization import apply_cover, fingerprint


def attrs(e):
    return {} if e is None else dict(e.attrib)


def shape(e):
    if e is None:
        return None
    return (e.tag, tuple(sorted(e.attrib.items())), tuple(shape(x) for x in e))


def inherited(element, style, tag):
    pp = style.element.find(qn("w:pPr"))
    result = attrs(pp.find(qn("w:" + tag))) if pp is not None else {}
    pp = element.find(qn("w:pPr"))
    if pp is not None:
        result.update(attrs(pp.find(qn("w:" + tag))))
    return result


def _cover_values(doc, spec):
    """Read back the cover strings, ignoring wrapping breaks, tabs and placeholders."""

    def clean(text):
        text = text.replace("\n", "").replace("\t", "").strip()
        return "" if not text.strip("　") else text

    cover = spec["cover"]
    ps = doc.paragraphs
    values = {key: clean(ps[i].runs[1].text) for key, i in cover["fields"].items()}
    values["title_zh"] = clean(ps[cover["title_paragraph"]].runs[1].text)
    if cover.get("course_paragraph") is not None:
        values["course_name"] = clean(ps[cover["course_paragraph"]].text).strip("《》").strip()
        values["course_name"] = clean(values["course_name"])
    return values


def text_of(p):
    return "".join(p._p.xpath(".//w:t/text()"))


def visible_part(part):
    return bool(
        part._element.xpath(".//w:t|.//w:instrText|.//w:fldSimple|.//w:drawing|.//w:pict|.//w:pBdr")
    )


def _check(path, final=False, intake_path=None, toc_report=None, validation_scope="full"):
    if validation_scope not in ("full", "static-only"):
        raise ValueError("未知验收范围。")
    errors = []
    warnings = []

    def err(code, msg):
        errors.append({"code": code, "message": msg})

    def warn(code, msg):
        warnings.append({"code": code, "message": msg})

    try:
        doc = Document(path)
        vv = E.get_vars(doc)
        request = json.loads(vv.get("XAUFE_FORMAT_REQUEST", "{}"))
        E.select_profile(vv.get("XAUFE_TEMPLATE"), request)
        E.asset_check()
    except Exception as exc:
        return {
            "passed": False,
            "errors": [{"code": "OPEN", "message": str(exc)}],
            "warnings": [],
            "stats": {},
        }
    spec = E.SPEC
    expected = contract(spec["id"], request)
    if doc.settings.element.xpath('./w:updateFields[@w:val="true" or @w:val="1"]'):
        err(
            "AUTO_FIELD_UPDATE",
            "交付稿不得请求打开时自动更新域；关闭文档级更新请求，保留手动更新能力，不通过修改 Word 全局提示设置规避。",
        )
    if "XAUFE_FORMAT_REQUEST" in vv and vv.get("XAUFE_FORMAT_REQUEST_HASH") != fingerprint(request):
        err("CUSTOMIZATION_HASH", "任务格式差异摘要无效；请重新排版，不手改成品声明。")
    from unsupported import unsupported_reasons

    for reason in unsupported_reasons(doc):
        err("UNSUPPORTED_STRUCTURE", reason)
    has_cover = vv.get("XAUFE_HAS_COVER") == "1"
    mode = vv.get("XAUFE_MODE", "unknown")
    count = spec["cover"]["paragraph_count"]
    state = None
    if mode != "template":
        if vv.get("XAUFE_WORKFLOW") != "ready" or not vv.get("XAUFE_INTAKE_ID"):
            err("WORKFLOW", "缺少本次模板选择及封面信息已解决的执行记录。")
        if intake_path:
            try:
                state = require_intake(intake_path)
                if state.get("format_request", {}) != request:
                    err("CUSTOMIZATION_STALE", "成品没有采用本次用户的特别要求，或任务差异已改变。")
                if state["template"] != spec["id"]:
                    err("TEMPLATE_MISMATCH", "输出模板与用户选择不同。")
                if "include_cover" in state and state["include_cover"] != has_cover:
                    err("DOCUMENT_SCOPE", "有无封面与用户明确的文档范围不一致。")
                if state["job_id"] != vv.get("XAUFE_INTAKE_ID"):
                    err("JOB_MISMATCH", "输出关联了其他任务的交互记录。")
                if digest(state["metadata"]) != vv.get("XAUFE_COVER_METADATA_HASH") or sorted(
                    json.loads(vv.get("XAUFE_COVER_BLANK_FIELDS", "[]"))
                ) != sorted(state.get("blank_fields", [])):
                    err("METADATA_STALE", "封面信息、留空设置或默认日期已变化，请重新生成和检查。")
            except Exception as exc:
                err("INTAKE", str(exc))
        elif final:
            err("INTAKE_REQUIRED", "最终检查需 --intake 以核对真实选择、封面信息和留空设置。")
        else:
            warn("INTAKE_NOT_RECHECKED", "本次静态检查未提供交互记录；最终检查必须提供。")
    if has_cover and mode != "template" and len(doc.sections) < 2:
        err("COVER_BREAK", "封面必须独立分节，与后续部分分开。")
    for i, s in enumerate(doc.sections):
        for attr, key in [
            ("page_width", "width"),
            ("page_height", "height"),
            ("top_margin", "top"),
            ("bottom_margin", "bottom"),
            ("left_margin", "left"),
            ("right_margin", "right"),
            ("header_distance", "header"),
            ("footer_distance", "footer"),
            ("gutter", "gutter"),
        ]:
            actual = getattr(s, attr)
            if actual is None or actual.twips != spec["page_twips"][key]:
                err("PAGE", f"第{i+1}节 {attr} 与本次有效格式（默认值加用户覆盖）不一致。")
        for variant, f in [
            ("default", s.footer),
            ("first", s.first_page_footer),
            ("even", s.even_page_footer),
        ]:
            if spec["footer"].get("page_numbers") and not (i == 0 and has_cover):
                fields = f._element.xpath(".//w:fldSimple/@w:instr")
                if len(fields) != 1 or fields[0].strip() != "PAGE":
                    err("FOOTER", "用户要求的页码域缺失或混入其他字段。")
            elif visible_part(f):
                err("FOOTER", f"第{i+1}节 {variant} 页脚不是空白。")
        for variant, h in [
            ("default", s.header),
            ("first", s.first_page_header),
            ("even", s.even_page_header),
        ]:
            if not spec["header"]["enabled"] or (i == 0 and has_cover):
                if visible_part(h):
                    err(
                        "RED_HEADER" if not spec["header"]["enabled"] else "COVER_HEADER",
                        f"第{i+1}节 {variant} 不应有页眉、横线或字段。",
                    )
            elif not "".join(h._element.xpath(".//w:t/text()")).strip():
                err("HEADER", f"第{i+1}节 {variant} 页眉缺失。")
            elif state:
                want = (
                    spec["header"].get(
                        "left", (state["metadata"].get("major") or "＿＿") + "专业本科学年论文"
                    )
                    + "\t"
                    + spec["header"].get(
                        "right",
                        state["metadata"].get("title_zh") or doc.core_properties.title or "",
                    )
                )
                actual = [p.text.replace("\n", "") for p in h.paragraphs]
                if actual != [want] and actual != want.split("\t", 1):
                    err(
                        "HEADER_TEXT",
                        "页眉文字与本次有效设置或已核实的专业／题目不一致；以header覆盖后的左右文字为准，不按校徽颜色判断。",
                    )
    if spec["header"]["enabled"]:
        for si, section in enumerate(doc.sections):
            if si == 0 and has_cover:
                continue
            for variant, h in [
                ("default", section.header),
                ("first", section.first_page_header),
                ("even", section.even_page_header),
            ]:
                bottoms = h._element.xpath(".//w:pBdr/w:bottom")
                if not spec["header"].get("border", True):
                    if bottoms:
                        err("HEADER_BORDER", "用户要求无页眉横线。")
                    continue
                if len(bottoms) != 1 or any(
                    bottoms[0].get(qn("w:" + key)) != value
                    for key, value in {
                        "val": "single",
                        "sz": "6",
                        "space": "1",
                        "color": "000000",
                    }.items()
                ):
                    err("HEADER_BORDER", f"第{si+1}节 {variant} 页眉横线缺失或改变。")
    if has_cover:
        check_cover_date(doc, spec, err)
        if len(doc.paragraphs) < count:
            err("COVER", "封面字段段落不完整。")
        else:
            p = doc.paragraphs[0]
            blips = p._p.xpath(".//a:blip")
            if len(blips) != 1:
                err("LOGO", "封面首段须包含本次选定的原校徽；校徽来源与基础版式分别核对。")
            else:
                rid = blips[0].get(qn("r:embed"))
                if not rid or rid not in doc.part.related_parts:
                    err("LOGO", "校徽不是有效嵌入式图片。")
                elif (
                    hashlib.sha256(doc.part.related_parts[rid].blob).hexdigest()
                    != spec["logo"]["sha256"]
                ):
                    err(
                        "LOGO",
                        "校徽图像与本次选定来源不一致，或原图字节已改变。A/B均可配红／绿校徽；按logo核对，不按基础版式推断颜色。",
                    )
            ext = p._p.xpath(".//wp:inline/wp:extent")
            if (
                len(ext) != 1
                or ext[0].get("cx") != str(spec["logo"]["width_emu"])
                or ext[0].get("cy") != str(spec["logo"]["height_emu"])
            ):
                err("LOGO_SIZE", "校徽显示尺寸或嵌入方式改变。")
            if p._p.xpath(".//wp:anchor"):
                err("LOGO_FLOAT", "校徽不得使用浮动锚点。")
            if any(
                any(v not in ("0", "") for v in x.attrib.values())
                for x in p._p.xpath(".//a:srcRect")
            ):
                err("LOGO_CROP", "校徽不得裁切。")
            if p._p.xpath("./w:pPr/w:jc/@w:val") != [spec["logo"]["alignment"]]:
                err("LOGO_ALIGN", "校徽必须左对齐。")
            if not doc.paragraphs[spec["cover"]["title_paragraph"]].text.startswith(
                spec["cover"]["title_label"]
            ):
                err("COVER_TITLE", "封面题目标签与所选模板不符。")
            if (
                doc.paragraphs[spec["cover"]["label_paragraph"]].text
                != spec["cover"]["document_label"]
            ):
                err("COVER_LABEL", "封面文种与所选模板不符。")
            master = Document(E.ROOT / spec["master_file"])
            apply_cover(master, spec)
            import cover_fit

            cover_values = _cover_values(doc, spec)
            try:
                fit_level = int(vv.get("XAUFE_COVER_FIT", "0"))
            except ValueError:
                fit_level = -1
            if not 0 <= fit_level <= cover_fit.MAX_LEVEL:
                err("COVER_LAYOUT", "封面适配级别记录无效；请重新生成文档。")
                fit_level = 0
            # Expected cover = master + user overrides + the same long-text fitting.
            cover_fit.apply(master, spec, cover_fit.plan(master, spec, cover_values, level=fit_level))
            for pi in range(count):
                actual = doc.paragraphs[pi]
                baseline = master.paragraphs[pi]
                for tag in ["jc", "ind", "tabs", "spacing"]:
                    a = actual._p.pPr.find(qn("w:" + tag)) if actual._p.pPr is not None else None
                    b = (
                        baseline._p.pPr.find(qn("w:" + tag))
                        if baseline._p.pPr is not None
                        else None
                    )
                    if shape(a) != shape(b):
                        err("COVER_LAYOUT", f"封面p{pi}的{tag}与母版不符。")
                for ri, r in enumerate(actual.runs):
                    if not r.text.strip():
                        continue
                    if ri >= len(baseline.runs):
                        err("COVER_RUNS", f"封面p{pi}出现母版外运行；须确认布局。")
                        continue
                    rr = baseline.runs[ri]
                    effective = run_properties(actual, r._r)
                    original = run_properties(baseline, rr._r)
                    for key in (
                        "fonts",
                        "sz",
                        "b",
                        "i",
                        "iCs",
                        "u",
                        "color",
                        "vanish",
                        "webHidden",
                        "specVanish",
                        "spacing",
                        "w",
                        "position",
                    ):
                        default = (
                            False
                            if key in ("b", "i", "iCs", "vanish", "webHidden", "specVanish")
                            else None
                        )
                        if effective.get(key, default) != original.get(key, default):
                            err("COVER_EFFECTIVE_FONT", f"封面p{pi}/r{ri}的有效{key}被样式覆盖。")
                    for tag in ["rFonts", "sz", "b"]:
                        want = (
                            attrs(rr._r.rPr.find(qn("w:" + tag))) if rr._r.rPr is not None else {}
                        )
                        got = attrs(r._r.rPr.find(qn("w:" + tag))) if r._r.rPr is not None else {}
                        if got != want:
                            err("COVER_FONT", f"封面p{pi}/r{ri}的{tag}与母版不符。")
            if state:
                cm = E.cover_metadata(state["metadata"], state)
                for key, pi in spec["cover"]["fields"].items():
                    got = cover_values.get(key, "")
                    if got != cm.get(key, ""):
                        err("COVER_VALUE", f"{key} 未按已核实信息或明确留空设置填写。")
                title = (
                    doc.paragraphs[spec["cover"]["title_paragraph"]]
                    .runs[1]
                    .text.strip()
                    .replace("\n", "")
                )
                if title != cm.get("title_zh", ""):
                    err("COVER_VALUE", "封面题目未按已核实信息／留空设置处理。")
                cp = spec["cover"].get("course_paragraph")
                if cp is not None:
                    val = cover_values.get("course_name", "")
                    if val != cm.get("course_name", "").strip().strip("《》"):
                        err("COVER_VALUE", "课程名称与本次信息不符。")
                date = cm.get("date", "")
                dp = doc.paragraphs[spec["cover"]["date_paragraph"]].text.replace("完成日期：", "")
                if date:
                    y, m, d = map(int, date.split("-"))
                    want = f"{y}年{m}月{d}日"
                    if re.sub(r"\s+", "", dp) != want:
                        err("COVER_DATE", "日期不是指定日期或本次运行日。")
                elif re.search(r"\d", dp):
                    err("COVER_DATE", "用户要求留空，却保留了数字日期。")
    reverse = {v: k for k, v in E.ROLES.items()}
    allp = [(f"p{i}", p) for i, p in enumerate(doc.paragraphs)]
    for ti, t in enumerate(doc.tables):
        allp.extend(walk_paragraphs(t, f"t{ti}"))
    if spec["header"]["enabled"]:
        for si, sec in enumerate(doc.sections):
            if si == 0 and has_cover:
                continue
            for variant, h in [
                ("default", sec.header),
                ("first", sec.first_page_header),
                ("even", sec.even_page_header),
            ]:
                allp.extend((f"header{si}/{variant}/p{pi}", p) for pi, p in enumerate(h.paragraphs))
    checked = 0
    for label, p in allp:
        if has_cover and label.startswith("p") and int(label[1:]) < count:
            continue
        role = reverse.get(p.style.name)
        visible = text_of(p)
        if not role:
            if visible.strip():
                err("UNKNOWN_STYLE", f"{label}: 未映射的样式 {p.style.name}。")
            continue
        if role == "cover":
            err("MISPLACED_COVER_STYLE", f"{label} 在封面范围外使用封面样式。")
            continue
        if not spec["english_abstract"] and role in (
            "title_en",
            "abstract_label_en",
            "abstract_en",
            "keywords_en",
        ):
            err(
                "RED_ENGLISH_ABSTRACT",
                "本次有效设置english_abstract=false，但文档含英文摘要角色；检查用户要求是否已记录，不按A/B或校徽颜色禁用。",
            )
        if role in ("h4", "toc4") and spec.get("heading_levels", 3) < 4:
            err("HEADING_DEPTH", "第四级结构未被本次用户要求启用。")
        checked += 1
        rule = expected[role]
        if role == "header" and label.startswith("header") and len(p._parent.paragraphs) == 2:
            from copy import deepcopy

            rule = deepcopy(rule)
            rule["pPr"]["jc"]["val"] = "right" if label.endswith("/p1") else "left"
        paragraph_check(p, role, label, rule, err, warn)
        run_check(p, role, label, expected[role], err)
        if p._p.xpath(".//wp:inline") and role != "figure":
            maxh = max((int(x.get("cy", "0")) for x in p._p.xpath(".//wp:extent")), default=0)
            if maxh > 254000 and paragraph_properties(p, "spacing").get("lineRule") == "exact":
                err("OBJECT_HEIGHT", "大图片不能放在固定20磅的普通段落里。")
    with ZipFile(path) as z:
        for name in z.namelist():
            if name.startswith("word/fonts/"):
                err("EMBEDDED_FONT", "不得嵌入并分发字体文件。")
            if (
                not spec["footer"].get("page_numbers")
                and name.startswith("word/footer")
                and name.endswith(".xml")
                and re.search(r"\b(?:PAGE|NUMPAGES|SECTIONPAGES)\b", z.read(name).decode("utf-8"))
            ):
                err("FOOTER_FIELD", f"{name} 残留页码域。")
    for p in doc.paragraphs:
        if re.fullmatch(r"\s*(?:空一行|本科毕业论文（设计）模板)\s*", p.text) or "※※※" in p.text:
            err("PLACEHOLDER", "模板说明或示例占位泄漏到成品。")
    unresolved = []
    for fld in doc.element.body.xpath(".//w:fldSimple"):
        if "PAGEREF" in fld.get(qn("w:instr"), ""):
            text = "".join(x.text or "" for x in fld.findall(".//" + qn("w:t")))
            if not text.isdigit():
                unresolved.append(text)
    if unresolved:
        (err if final and validation_scope == "full" else warn)(
            "TOC_UNRESOLVED", "目录页码还没有计算（需要实际分页）。"
        )
    entries = []
    try:
        entries = toc_entries(doc)
    except ValueError as exc:
        err("TOC_SCOPE", str(exc))
    if mode == "paper" and doc.core_properties.title:
        title_key = re.sub(r"\s+", "", doc.core_properties.title)
        if any(re.sub(r"\s+", "", entry["text"]) == title_key for entry in entries):
            warn(
                "DOCUMENT_TITLE_IN_TOC",
                "全篇题目“"
                + doc.core_properties.title
                + "”同时被列为正文章节。默认全篇题目放在元数据供封面/摘要使用，不重复进入正文和目录；请检查原始材料的文档标题是否误映射进blocks。仅在用户明确要求时保留该特殊范围。",
            )
    if validation_scope == "static-only":
        from render_availability import pending_check

        pending_check(doc, err)
        warn(
            "RENDER_NOT_VERIFIED",
            "仅做静态格式检查；未完成实际分页、目录页位与逐页视觉检查，必须向用户告知。",
        )
    else:
        if final and vv.get("XAUFE_VALIDATION_SCOPE") == "static-only":
            err("RENDER_REQUIRED", "此稿尚未实际分页，不能当作完整检查；请重新运行 run。")
        if entries:
            toc_proof_check(path, entries, vv, final, err, warn, toc_report)
    from toc_format import audit as audit_toc

    audit_toc(doc, spec, err)
    table_count = tables_check(doc, E.WIDTH, err)
    if mode == "paper":
        front = []
        for part in spec["parts"]:
            front += {
                "toc": ["toc_title"],
                "abstract_zh": ["title_zh", "abstract_label_zh"],
                "abstract_en": ["title_en", "abstract_label_en"],
            }.get(part, [])
        roles = [reverse.get(p.style.name) for p in doc.paragraphs]
        pos = []
        for role in front:
            if role not in roles:
                err("PART_MISSING", f"缺少{role}部分。")
            else:
                pos.append(roles.index(role))
        if pos != sorted(pos):
            err("PART_ORDER", "前置部分顺序与用户所选模板不符。")
        body = next(
            (i for i, x in enumerate(roles) if x in ("preface", "h1", "h2", "h3", "h4", "body")),
            None,
        )
        if body is None:
            err("PART_MISSING", "缺少正文。")
        elif pos and body <= max(pos):
            err("PART_ORDER", "正文排在前置部分之前。")
        if "reference_title" not in roles:
            err("PART_MISSING", "缺少参考文献标题。")
        elif body is not None and roles.index("reference_title") <= body:
            err("PART_ORDER", "参考文献未排在正文之后。")
        minimum = (5 if spec["english_abstract"] else 4) - (0 if has_cover else 1)
        if len(doc.sections) < minimum:
            err("PART_BREAKS", "封面／摘要／目录／正文没有独立分节起页。")
    if mode == "paper":
        missing_content = []
        for role in ("abstract_zh", "keywords_zh", "reference") + (
            ("abstract_en", "keywords_en") if spec["english_abstract"] else ()
        ):
            paragraphs = [p for p in doc.paragraphs if reverse.get(p.style.name) == role]
            nonempty = any(
                re.sub(r"^(关键词\s*[:：]|Key\s*words\s*[:：])", "", text_of(p), flags=re.I).strip()
                for p in paragraphs
            )
            if not nonempty:
                missing_content.append(role)
        if missing_content:
            warn(
                "EMPTY_CONTENT",
                "以下部分是空的：用户要求写作就补写；只排版就在 review.json 说明这是原稿本来就空的："
                + ",".join(missing_content),
            )
    for i, img in enumerate(doc.inline_shapes):
        if i == 0 and has_cover:
            continue
        if img.width > E.Twips(E.WIDTH):
            err("IMAGE_WIDTH", "图片宽度超过所选模板版心。")
    warn(
        "VISUAL_REQUIRED",
        "静态检查不等于看图检查；有看图能力就逐页打开页图。",
    )
    stats = {
        "template": spec["id"],
        "mode": mode,
        "cover": has_cover,
        "sections": len(doc.sections),
        "checked_paragraphs": checked,
        "tables": len(doc.tables),
        "tables_including_nested": table_count,
        "inline_images": len(doc.inline_shapes),
        "toc_unresolved": len(unresolved),
    }
    return {
        "passed": not errors,
        "document_sha256": hashlib.sha256(Path(path).read_bytes()).hexdigest(),
        "errors": errors,
        "warnings": warnings,
        "stats": stats,
        "validation_scope": validation_scope,
        "fully_validated": False,
    }


def check(path, final=False, intake_path=None, toc_report=None, validation_scope="full"):
    try:
        return _check(path, final, intake_path, toc_report, validation_scope)
    except Exception as exc:
        return {
            "passed": False,
            "errors": [{"code": "AUDIT_ABORTED", "message": "文档结构无法安全检查：" + str(exc)}],
            "warnings": [],
            "stats": {},
        }


def main():
    p = argparse.ArgumentParser()
    p.add_argument("docx")
    p.add_argument("--report", required=True)
    p.add_argument("--intake")
    p.add_argument("--final", action="store_true")
    p.add_argument("--toc-report")
    p.add_argument("--validation-scope", choices=["full", "static-only"], default="full")
    a = p.parse_args()
    result = check(a.docx, a.final, a.intake, a.toc_report, a.validation_scope)
    Path(a.report).parent.mkdir(parents=True, exist_ok=True)
    Path(a.report).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(
        json.dumps(
            {
                "passed": result["passed"],
                "errors": len(result["errors"]),
                "warnings": len(result["warnings"]),
                "report": a.report,
            },
            ensure_ascii=False,
        )
    )
    return 0 if result["passed"] else 2


if __name__ == "__main__":
    sys.exit(main())
