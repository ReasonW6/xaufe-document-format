"""Explicit, auditable static-only delivery when no renderer can be used.

This is NOT a successful render. Pending page references are intentionally visible.
The final chat notice must disclose every skipped check. Ordinary layout defects,
invalid inputs, and failed static checks are never downgraded by this module.
"""

from __future__ import annotations
import hashlib
import json
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
import engine as E
from audit_helpers import toc_entries

MODE = "static-only"
PENDING = "待更新"
MARKER = "XAUFE_VALIDATION_SCOPE"  # Private document validation marker.
PAGE_FIELDS = {"PAGEREF", "PAGE", "NUMPAGES", "SECTIONPAGES"}


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def page_fields(doc):
    """All supported simple page-dependent fields, including header/footer parts."""
    seen = set()
    for part in doc.part.package.parts:
        root = getattr(part, "element", None)
        if root is None:
            root = getattr(part, "_element", None)
        if root is None or id(root) in seen:
            continue
        seen.add(id(root))
        for field in root.findall(".//" + qn("w:fldSimple")):
            instruction = field.get(qn("w:instr"), "").strip().split()
            if instruction and instruction[0].upper() in PAGE_FIELDS:
                yield field


def mark_pending(doc):
    """No fabricated old page numbers; retain live fields and content structure."""
    count = 0
    for field in page_fields(doc):
        texts = field.findall(".//" + qn("w:t"))
        if len(texts) != 1:
            raise ValueError("页码域缓存结构不受支持；不能在未保全结构时交付。")
        texts[0].text = PENDING
        field.attrib.pop(qn("w:dirty"), None)
        count += 1
    E.set_vars(doc, {MARKER: MODE})
    # Keep live, unlocked fields for manual updating. Do not request automatic
    # update on open: Word may present a warning even for internal PAGEREF fields.
    E.element(doc.settings.element, "w:updateFields", val="false")
    return count


def clear_scope(doc):
    """A new full preparation must not inherit an earlier unverified status."""
    for node in list(doc.settings.element.findall(".//" + qn("w:docVar"))):
        if node.get(qn("w:name")) == MARKER:
            node.getparent().remove(node)


def pending_check(doc, err):
    if E.get_vars(doc).get(MARKER) != MODE:
        err("STATIC_SCOPE_REQUIRED", "缺少明确的未渲染验收范围标记。")
    for field in page_fields(doc):
        texts = field.findall(".//" + qn("w:t"))
        if len(texts) != 1 or texts[0].text != PENDING:
            err(
                "UNVERIFIED_PAGE_CACHE",
                "没有实际分页时，页码缓存必须统一为“待更新”，不能交付旧数字。",
            )
        if field.get(qn("w:dirty")) in ("true", "1"):
            err("PENDING_FIELD_STATE", "静态稿保留手动更新的域，不请求打开时自动更新。")
    update = doc.settings.element.find(qn("w:updateFields"))
    if update is None or update.get(qn("w:val")) not in ("false", "0"):
        err("PENDING_FIELD_STATE", "交付稿须关闭打开时自动更新域，未核实页码留待手动更新。")


def skipped_steps(has_toc, page_field_count):
    rows = [
        {"code": "PAGINATION_NOT_VERIFIED", "step": "实际分页和总页数核对"},
        {"code": "RENDER_NOT_PERFORMED", "step": "最终页面预览生成"},
        {
            "code": "VISUAL_SKIPPED_NO_RENDERER",
            "step": "逐页视觉检查（裁切、重叠、分页、封面日期实际位置）",
        },
    ]
    if has_toc:
        rows.insert(1, {"code": "TOC_PAGES_NOT_VERIFIED", "step": "目录页码刷新及真实页位置核对"})
    elif page_field_count:
        rows.insert(1, {"code": "PAGE_FIELDS_NOT_VERIFIED", "step": "页码域的实际数字核对"})
    return rows


def disclosure(reason, steps, has_page_fields):
    note = (
        "已完成文档排版及静态格式检查。"
        + reason
        + "本次未完成："
        + "；".join(row["step"] for row in steps)
        + "。"
    )
    if has_page_fields:
        note += "未核实的目录/页码已标为“待更新”，请在可用办公软件中更新并核对。"
    note += "这份文件未经过完整页面验收，不能视为已确认无裁切、重叠或分页问题的成品。"
    return note


def create_receipt(root, docx, evidence, page_count):
    root = Path(root)
    has_toc = bool(toc_entries(Document(docx)))
    steps = skipped_steps(has_toc, page_count)
    reason = evidence["reason"]
    notice = disclosure(reason, steps, page_count > 0)
    receipt = {
        "version": "1.0.0",
        "validation_scope": MODE,
        "fully_validated": False,
        "document_sha256": sha(docx),
        "intake_sha256": sha(root / "active-intake.json"),
        "office_evidence_sha256": sha(root / "office-unavailable.json"),
        "visual_review": "skipped_no_renderer",
        "reason": reason,
        "skipped_steps": steps,
        "user_notice": notice,
        "completed_steps": [
            "文档生成/排版",
            "当前模板及用户覆盖参数检查",
            "静态字体字号、目录结构及对象检查",
        ],
        "pending_page_fields": page_count,
    }
    (root / "limited-validation.json").write_text(
        json.dumps(receipt, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return receipt


def verify_receipt(root, state):
    root = Path(root)
    receipt = json.loads((root / "limited-validation.json").read_text(encoding="utf-8"))
    if receipt.get("validation_scope") != MODE or receipt.get("fully_validated") is not False:
        raise ValueError("未渲染稿不能声明完整验收。")
    if sha(root / "limited-validation.json") != state.get("limited_validation_sha256"):
        raise ValueError("未完成步骤或交付说明已变化，请重新运行 run。")
    if receipt.get("document_sha256") != sha(root / "final.docx") or receipt.get(
        "intake_sha256"
    ) != sha(root / "active-intake.json"):
        raise ValueError("未渲染验收记录不属于当前文件或要求。")
    if receipt.get("office_evidence_sha256") != sha(root / "office-unavailable.json"):
        raise ValueError("软件探测证据变化，请重新运行 run。")
    evidence = json.loads((root / "office-unavailable.json").read_text(encoding="utf-8"))
    if evidence.get("probe_sha256") != sha(root / "office-probe.json"):
        raise ValueError("软件探测记录已变化，请重新运行 run。")
    if state.get("pages") != []:
        raise ValueError("未渲染稿不能持有页图。")
    return receipt


def page_neutral_signature(doc):
    """Content-preservation comparison allowing only live page-cache changes."""
    saved = []
    try:
        for field in page_fields(doc):
            texts = field.findall(".//" + qn("w:t"))
            for text in texts:
                saved.append((text, text.text))
                text.text = "<PAGE-CACHE>"
        return E.content_signature(doc)
    finally:
        for node, text in saved:
            node.text = text
