#!/usr/bin/env python3
"""Render supported TOCs, update only caches, and bind evidence to the exact DOCX."""

from __future__ import annotations
import argparse, hashlib, json, re, sys, tempfile
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from render_preview import pdf_render, font_report
from engine import safe_out, get_vars
from intake import require as require_intake, digest
from audit_helpers import toc_entries


def norm(s):
    return re.sub(r"\s+", "", s)


def logical_policy(doc, entries):
    """Do not extrapolate from the first listed heading across arbitrary page restarts.

    Supported: first substantive body paragraph is the first listed heading,
    the body's section explicitly restarts at 1, and later sections do not restart.
    This is the exact layout produced by the bundled paper builder.
    """
    sections = []
    current = []
    for el in doc.element.body:
        if el.tag == qn("w:sectPr"):
            sections.append((current, el))
            break
        current.append(el)
        if el.tag == qn("w:p"):
            ends = el.xpath("./w:pPr/w:sectPr")
            if ends:
                sections.append((current, ends[0]))
                current = []
    first = entries[0]["target"]._p
    index = next(
        (i for i, (nodes, _) in enumerate(sections) if any(el is first for el in nodes)), None
    )
    if index is None:
        raise ValueError("无法定位正文所属节。")
    nodes, sec = sections[index]
    start = sec.find(qn("w:pgNumType"))
    if start is None or start.get(qn("w:start")) != "1":
        raise ValueError("正文所在节未明确从1开始；不能从首个目录项猜测逻辑页码。")
    for el in nodes:
        if el is first:
            break
        if el.tag == qn("w:tbl") or el.xpath('.//w:t|.//w:drawing|.//w:br[@w:type="page"]'):
            raise ValueError("首目录标题前还有正文或强制分页；本自动目录路径无法可靠确定正文起页。")
    for _, sec in sections[index + 1 :]:
        page = sec.find(qn("w:pgNumType"))
        if page is not None and page.get(qn("w:start")) is not None:
            raise ValueError("正文后还有页码重启；需要专门目录验证，不自动外推。")


def rendered_positions(docx, outdir, entries):
    return pdf_positions(pdf_render(docx, outdir), entries)


def pdf_positions(pdf, entries):
    import pymupdf as fitz

    pdf = Path(pdf)
    with fitz.open(pdf) as rendered:
        clean = [(norm(row[1]), row[2]) for row in rendered.get_toc()]
        positions = []
        cursor = 0
        for entry in entries:
            text = norm(entry["text"])
            found = None
            for j in range(cursor, len(clean)):
                if clean[j][0] == text:
                    found = clean[j][1]
                    cursor = j + 1
                    break
            if found is None:
                raise ValueError(f"无法从实际渲染定位标题{text!r}；不得填入猜测页码。")
            positions.append(found)
        count = len(rendered)
    values = [x - positions[0] + 1 for x in positions]
    return (
        values,
        {"physical_pages": positions, "logical_pages": values, "total_pages": count},
        hashlib.sha256(pdf.read_bytes()).hexdigest(),
    )


def report_for(out, state, entries, passes, pdf_hash):
    return {
        "version": "1.0.0",
        "output": str(out),
        "converged": True,
        "document_sha256": hashlib.sha256(Path(out).read_bytes()).hexdigest(),
        "template": state["template"],
        "job_id": state["job_id"],
        "passes": passes,
        "format_request": state.get("format_request", {}),
        "entries": [
            {
                "anchor": e["anchor"],
                "text": norm(e["text"]),
                "logical_page": int(
                    "".join(t.text or "" for t in e["field"].findall(".//" + qn("w:t")))
                ),
            }
            for e in entries
        ],
        "rendered_pdf_sha256": pdf_hash,
        "font_resolution": font_report(),
        "office_backend": dict(__import__("office_backend").LAST_BACKEND),
        "note": "仅渲染并核对页引用，不把LibreOffice重新保存的DOCX作为成品。报告绑定本文件哈希，不是视觉验收证书。",
    }


def validate_job(input_path, intake_path):
    state = require_intake(intake_path)
    doc = Document(input_path)
    vv = get_vars(doc)
    if (
        vv.get("XAUFE_TEMPLATE") != state["template"]
        or vv.get("XAUFE_INTAKE_ID") != state["job_id"]
    ):
        raise ValueError("目录文件不属于本次模板选择／任务。")
    if vv.get("XAUFE_COVER_METADATA_HASH") != digest(state["metadata"]) or sorted(
        json.loads(vv.get("XAUFE_COVER_BLANK_FIELDS", "[]"))
    ) != sorted(state.get("blank_fields", [])):
        raise ValueError("封面信息、留空或自动日期已变化；先重新排版。")
    from customization import fingerprint

    if json.loads(vv.get("XAUFE_FORMAT_REQUEST", "{}")) != state.get("format_request", {}):
        raise ValueError("本次格式特别要求已变化；先重新排版。")
    entries = toc_entries(doc)
    if not entries:
        raise ValueError("未找到受支持的目录页引用；不对第三方目录结构伪报通过。")
    logical_policy(doc, entries)
    return state, doc, entries


def update(input_path, output, force=False, intake_path=None):
    # Validate the workflow before touching any input or creating output paths.
    state = require_intake(intake_path)
    out = safe_out(input_path, output, force)
    state, doc, entries = validate_job(input_path, intake_path)
    import engine as E
    from toc_format import enforce

    E.select_profile(state["template"], state.get("format_request"))
    enforce(doc, E.SPEC)
    last = None
    passes = []
    with tempfile.TemporaryDirectory(prefix="xaufe-toc-") as td:
        td = Path(td)
        for iteration in range(1, 5):
            candidate = td / "candidate.docx"
            doc.save(candidate)
            values, measurement, pdf_hash = rendered_positions(candidate, td / "preview", entries)
            passes.append(dict(measurement, iteration=iteration))
            caches = [
                "".join(t.text or "" for t in e["field"].findall(".//" + qn("w:t")))
                for e in entries
            ]
            if values == last and caches == list(map(str, values)):
                # Copy the exact rendered bytes; another save could produce a different ZIP/hash.
                import shutil

                shutil.copy2(candidate, out)
                report = report_for(out, state, entries, passes, pdf_hash)
                out.with_suffix(".toc-report.json").write_text(
                    json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
                )
                return report
            for entry, num in zip(entries, values):
                ts = entry["field"].findall(".//" + qn("w:t"))
                if len(ts) != 1:
                    raise ValueError("目录页引用缓存结构异常。")
                ts[0].text = str(num)
            last = values
    raise ValueError("四轮内目录页引用未稳定；未生成定稿。")


def verify_existing(input_path, intake_path, report_path=None):
    """For unchanged/reopened supported DOCX: verify, never repair wrong caches."""
    state, doc, entries = validate_job(input_path, intake_path)
    from customization import profile_for
    from toc_format import audit

    problems = []
    audit(doc, profile_for(state), lambda code, message: problems.append((code, message)))
    if problems:
        raise ValueError("目录实际格式不符，先规范再刷新；不能为错误字号签发报告：" + str(problems))
    original_hash = hashlib.sha256(Path(input_path).read_bytes()).hexdigest()
    with tempfile.TemporaryDirectory(prefix="xaufe-toc-verify-") as td:
        values, measurement, pdf_hash = rendered_positions(input_path, Path(td), entries)
    if [e["cache"] for e in entries] != list(map(str, values)):
        raise ValueError("现有目录数字与实际渲染不一致；请刷新目录，不能签发验收报告。")
    if hashlib.sha256(Path(input_path).read_bytes()).hexdigest() != original_hash:
        raise ValueError("检查期间文档发生变化，请重新验证。")
    report = report_for(input_path, state, entries, [dict(measurement, iteration=1)], pdf_hash)
    dest = Path(report_path) if report_path else Path(input_path).with_suffix(".toc-report.json")
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return report


def main():
    p = argparse.ArgumentParser()
    p.add_argument("docx")
    p.add_argument("--output")
    p.add_argument("--force", action="store_true")
    p.add_argument("--intake", required=True)
    p.add_argument("--verify-only", action="store_true")
    p.add_argument("--report")
    a = p.parse_args()
    try:
        if a.verify_only:
            if a.output:
                raise ValueError("--verify-only不改写文档，不能同时指定--output。")
            result = verify_existing(a.docx, a.intake, a.report)
        else:
            if not a.output:
                raise ValueError("更新目录必须指定--output，另存而不覆盖原件。")
            result = update(a.docx, a.output, a.force, a.intake)
        print(
            json.dumps(
                {
                    "converged": result["converged"],
                    "output": result["output"],
                    "passes": len(result["passes"]),
                },
                ensure_ascii=False,
            )
        )
        return 0
    except Exception as exc:
        print("目录处理停止：" + str(exc), file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
