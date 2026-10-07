"""Task workspace, real pagination/preview, review gate and single-DOCX delivery.

The review is the agent's honest written record of what it checked; the script
binds it to the exact document and page images but cannot prove anyone looked.
Only the owned workspace is deleted; sources and unrelated files stay untouched.
"""

from __future__ import annotations
import hashlib, json, os, re, shutil, tempfile, uuid
from pathlib import Path
from docx import Document
import engine as E
from intake import require
from check_format import check
from render_preview import render
from refresh_toc import update
from audit_helpers import toc_entries
from toc_format import enforce
from office_backend import OfficeUnavailable, OfficeExportFailed
import render_availability as RF

MAGIC = "xaufe-owned-document-job-v2"
MARKER = ".xaufe-session.json"
REVIEW_TEXTS = {
    "content_check": "需求与内容保全",
    "layout_check": "格式与版面",
    "delivery_check": "最终交付",
}
PLACEHOLDER = "请填写"


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def read_json(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def write_json(path, data):
    Path(path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def create(destination, *, parent=None, sources=()):
    dest = Path(destination).expanduser().resolve()
    if dest.suffix.lower() != ".docx":
        raise ValueError("最终交付文件必须是 .docx。")
    if dest.exists():
        raise ValueError("目标文件已存在；换一个文件名，不要覆盖用户已有文件。")
    if dest.is_relative_to(E.ROOT):
        raise ValueError("交付路径不能放在技能目录里。")
    originals = []
    for value in sources:
        path = Path(value).expanduser().resolve()
        if not path.is_file():
            raise ValueError("原始材料不是可读取的文件：" + str(path))
        if str(path) not in [x["path"] for x in originals]:
            originals.append({"path": str(path), "sha256": sha(path), "role": "original"})
    if os.name == "nt":
        # A plain directory inherits the parent's ACL, so an approved desktop-user
        # export can read it (mkdtemp would create a sandbox-private folder).
        folder = Path(parent) if parent is not None else Path(tempfile.gettempdir())
        root = (folder / ("xaufe-job-" + uuid.uuid4().hex[:16])).resolve()
        root.mkdir()
    else:
        root = Path(tempfile.mkdtemp(prefix="xaufe-job-", dir=parent)).resolve()
    state = {
        "magic": MAGIC,
        "workdir": str(root),
        "job_token": str(uuid.uuid4()),
        "destination": str(dest),
        "stage": "working",
        "source_documents": originals,
    }
    write_json(root / MARKER, state)
    return {"workspace": str(root), "destination": str(dest)}


def load(root):
    raw = Path(root).expanduser()
    if raw.is_symlink():
        raise ValueError("工作区不能是符号链接。")
    root = raw.resolve()
    marker = root / MARKER
    if (
        not root.name.startswith("xaufe-job-")
        or not root.is_dir()
        or not marker.is_file()
        or marker.is_symlink()
    ):
        raise ValueError("不是本工具创建的工作区；只能使用 xaufe.py new 返回的 WORK 路径。")
    state = read_json(marker)
    if (
        state.get("magic") != MAGIC
        or state.get("workdir") != str(root)
        or not state.get("job_token")
    ):
        raise ValueError("工作区所有权标记不匹配。")
    if root == E.ROOT or E.ROOT.is_relative_to(root) or root.is_relative_to(E.ROOT):
        raise ValueError("禁止把技能目录或其父目录当作工作区。")
    return root, state


def compare_originals(root, state, result):
    originals = [x for x in state["source_documents"] if x.get("role") == "original"]
    if not originals:
        return
    from source_compare import compare

    report = compare(originals, root / "final.docx")
    write_json(root / "source-comparison.json", report)
    gaps = [item for item in report["sources"] if item["status"] == "differences"]
    if gaps:
        result["warnings"].append(
            {
                "code": "SOURCE_DIFFERENCES",
                "message": "原稿里有些文字在成品中没找到，打开 WORK/source-comparison.json 逐条核对：没经用户同意的遗漏要先补回；用户要求删改的，在 review.json 的 warnings 里写明。",
                "sources": gaps,
            }
        )


def check_preview_toc(root):
    """Bind page caches to the PDF the model actually sees, including backend fallback."""
    from refresh_toc import pdf_positions

    doc = Document(root / "final.docx")
    entries = toc_entries(doc)
    if not entries:
        return {"required": False}
    values, measurement, pdf_hash = pdf_positions(root / "preview/final.pdf", entries)
    if [entry["cache"] for entry in entries] != list(map(str, values)):
        raise ValueError("最终预览与目录页码不一致；可能切换了渲染软件。请重新运行 run。")
    proof = read_json(root / "final.toc-report.json")
    previous = proof["passes"][-1]
    if (
        measurement["physical_pages"] != previous["physical_pages"]
        or measurement["total_pages"] != previous["total_pages"]
    ):
        raise ValueError("预览版面与目录定位结果不一致，请重新运行 run。")
    return {
        "required": True,
        "entries": len(entries),
        "preview_pdf_sha256": pdf_hash,
        **measurement,
    }


def verify_originals(state):
    for source in state["source_documents"]:
        if source.get("role") != "original":
            continue
        path = Path(source["path"])
        if not path.is_file() or sha(path) != source["sha256"]:
            raise ValueError(
                "登记的原稿在处理期间被改动：" + str(path) + "。请确认原稿后用 new 重新开始。"
            )


def _reset_for_prepare(root, s, docx, intake_path):
    verify_originals(s)
    if docx == Path(s["destination"]):
        raise ValueError("原稿不能同时作为最终交付目标。")
    state = require(intake_path)
    doc = Document(docx)
    vv = E.get_vars(doc)
    if (
        vv.get("XAUFE_TEMPLATE") != state["template"]
        or vv.get("XAUFE_INTAKE_ID") != state["job_id"]
    ):
        raise ValueError("文档与本次任务记录不匹配；请重新运行 run。")
    if json.loads(vv.get("XAUFE_FORMAT_REQUEST", "{}")) != state.get("format_request", {}):
        raise ValueError("用户特别要求已变化，请重新运行 run。")
    local_intake = root / "active-intake.json"
    if intake_path != local_intake:
        shutil.copy2(intake_path, local_intake)
    s["active_request_source"] = str(intake_path)
    s["stage"] = "preparing"
    for key in (
        "review",
        "validation_scope",
        "limited_validation_sha256",
        "skipped_steps",
        "user_notice",
        "pages",
    ):
        s.pop(key, None)
    write_json(root / MARKER, s)
    E.select_profile(state["template"], state.get("format_request"))
    RF.clear_scope(doc)
    enforce(doc, E.SPEC)
    normalized = root / "toc-normalized.docx"
    doc.save(normalized)
    before = check(normalized, intake_path=local_intake)
    write_json(root / "initial-check.json", before)
    if not before["passed"]:
        paragraphs = doc.paragraphs

        def describe(error):
            index = re.match(r"p(\d+)\b", error["message"])
            if index and int(index.group(1)) < len(paragraphs):
                text = paragraphs[int(index.group(1))].text.strip()[:30]
                return f"{error['message']}（这一段是：“{text}”）"
            return error["message"]

        raise ValueError(
            "格式初查未通过：\n  - "
            + "\n  - ".join(describe(e) for e in before["errors"])
            + "\n多半是 mapping.json 里某段的 role 标错了（例如英文题目标成了 body），或 content.json 缺了某一部分。"
            "改好后重新 run；不要换成另一种做法，也不要加 --office none 绕过。"
        )
    return doc, normalized, local_intake


def _prepare(root, s, docx, intake_path, unavailable_note=None, office="auto"):
    doc, normalized, local_intake = _reset_for_prepare(root, s, docx, intake_path)
    if office == "none":
        write_json(root / "office-probe.json", {"skipped": "user_request"})
        evidence = {
            "version": "1.0.0",
            "probe_sha256": sha(root / "office-probe.json"),
            "reason": "按用户要求，本次没有启动办公软件。",
            "status": "skipped_by_user_request",
        }
        write_json(root / "office-unavailable.json", evidence)
        return _prepare_static_only(root, s, normalized, local_intake, evidence)
    from office_discovery import discover

    # The real export below is the capability probe; no separate launch first.
    probe = discover()
    write_json(root / "office-probe.json", probe)
    final = root / "final.docx"
    try:
        if toc_entries(doc):
            update(normalized, final, force=True, intake_path=local_intake)
        else:
            shutil.copy2(normalized, final)
        result = check(final, final=True, intake_path=local_intake)
        write_json(root / "final-check.json", result)
        if not result["passed"]:
            raise ValueError("格式终查未通过：" + json.dumps(result["errors"], ensure_ascii=False))
        preview = render(final, root / "preview")
        write_json(root / "preview-toc-check.json", check_preview_toc(root))
    except OfficeUnavailable as exc:
        evidence = _renderer_unavailable(root, exc, probe, unavailable_note)
        return _prepare_static_only(root, s, normalized, local_intake, evidence)
    if sha(final) != result["document_sha256"]:
        raise ValueError("预览时成品发生变化，请重新运行 run。")
    compare_originals(root, s, result)
    write_json(root / "final-check.json", result)
    warnings = [w for w in result["warnings"] if w["code"] != "VISUAL_REQUIRED"]
    s.update(
        validation_scope="full",
        stage="awaiting_review",
        document_sha256=sha(final),
        render_report_sha256=sha(root / "preview/render-report.json"),
        pages=preview["pages"],
        warning_codes=sorted({w["code"] for w in warnings}),
    )
    write_json(root / MARKER, s)
    return {
        "workspace": str(root),
        "validation_scope": "full",
        "pages": [dict(p, path=str(root / "preview" / p["png"])) for p in preview["pages"]],
        "warnings": warnings,
    }


def _renderer_unavailable(root, exc, probe, note):
    """Only real unavailability of the environment can skip rendering, never bad layout."""
    from office_backend import export_pdf

    installed = [c for c in probe.get("candidates", []) if c.get("found")]
    smoke_result = None
    if isinstance(exc, OfficeExportFailed):
        smoke = root / "renderer-smoke.docx"
        test = Document()
        test.add_paragraph("Renderer availability check.")
        test.save(smoke)
        try:
            export_pdf(smoke, root / "renderer-smoke.pdf")
        except OfficeUnavailable as smoke_exc:
            smoke_result = {
                "status": "failed",
                "error": str(smoke_exc),
                "attempts": smoke_exc.failures,
            }
        else:
            raise ValueError(
                "办公软件能导出一个测试文档，说明是当前稿件导出失败，不是缺软件。请根据下面的错误修稿后重试："
                + str(exc)
            )
    elif any(c.get("automation") for c in installed):
        raise ValueError("还有可自动调用的办公软件没试过；请直接重新运行 run。")
    if installed and (not isinstance(note, str) or not note.strip()):
        raise ValueError(
            "RENDER_ALTERNATIVE_REQUIRED：电脑上找到了办公软件，但自动调用失败。"
            "如果你的运行环境支持经用户批准后在本机会话里执行命令，请用那种方式重跑同一条 run；"
            '否则用 run --office-note "失败的实际原因" 继续，进入静态交付（会告诉用户哪些检查没做）。'
            "不要绕过权限，也不要让用户另装软件。"
        )
    if installed:
        reason = (
            "已找到办公软件，但当前环境无法完成实际导出：" + note.strip().rstrip("。.;；") + "。"
        )
    else:
        reason = "当前运行环境里没有找到可用的 Word、WPS 或 LibreOffice（不代表用户电脑上没有）。"
    evidence = {
        "version": "1.0.0",
        "probe_sha256": sha(root / "office-probe.json"),
        "probe": probe,
        "export_discovery": exc.report,
        "export_attempts": exc.failures,
        "smoke_test": smoke_result,
        "alternative_note": note or "",
        "reason": reason,
        "status": "unavailable_in_current_runtime",
    }
    write_json(root / "office-unavailable.json", evidence)
    return evidence


def _prepare_static_only(root, s, normalized, intake, evidence):
    # Always start from the statically checked source, not a partly rendered output.
    doc = Document(normalized)
    preserved = RF.page_neutral_signature(doc)
    count = RF.mark_pending(doc)
    if RF.page_neutral_signature(doc) != preserved:
        raise ValueError("标记待更新页码时内容保全检查失败。")
    final = root / "final.docx"
    doc.save(final)
    for name in ("final.toc-report.json", "preview-toc-check.json", "limited-validation.json"):
        path = root / name
        if path.is_file():
            path.unlink()
    preview = root / "preview"
    if preview.exists():
        if preview.is_symlink():
            raise ValueError("预览目录不能是符号链接。")
        shutil.rmtree(preview)
    result = check(final, final=True, intake_path=intake, validation_scope="static-only")
    compare_originals(root, s, result)
    write_json(root / "final-check.json", result)
    if not result["passed"]:
        raise ValueError(
            "没有排版软件也必须先通过静态检查：" + json.dumps(result["errors"], ensure_ascii=False)
        )
    receipt = RF.create_receipt(root, final, evidence, count)
    warnings = [
        w
        for w in result["warnings"]
        # These are already part of the user notice for a static delivery.
        if w["code"] not in ("VISUAL_REQUIRED", "RENDER_NOT_VERIFIED", "TOC_UNRESOLVED")
    ]
    s.update(
        stage="awaiting_review",
        validation_scope="static-only",
        document_sha256=sha(final),
        render_report_sha256=None,
        pages=[],
        limited_validation_sha256=sha(root / "limited-validation.json"),
        warning_codes=sorted({w["code"] for w in warnings}),
        skipped_steps=receipt["skipped_steps"],
        user_notice=receipt["user_notice"],
    )
    write_json(root / MARKER, s)
    return {
        "workspace": str(root),
        "validation_scope": "static-only",
        "pages": [],
        "warnings": warnings,
        "skipped_steps": receipt["skipped_steps"],
        "user_notice": receipt["user_notice"],
    }


def configure_office(workspace, engine="auto", wps_mode="auto"):
    root, s = load(workspace)
    if engine not in ("auto", "word", "libreoffice", "wps"):
        raise ValueError("未知办公软件后端。")
    if wps_mode not in ("auto", "manual") or wps_mode == "manual" and engine != "wps":
        raise ValueError("手动界面导出只用于 WPS。")
    s.update(office_engine=engine, wps_mode=wps_mode, stage="working")
    s.pop("review", None)
    write_json(root / MARKER, s)
    return {"workspace": str(root), "engine": engine, "wps_mode": wps_mode}


def prepare(workspace, docx, intake_path, unavailable_note=None, office="auto"):
    root, s = load(workspace)
    docx = Path(docx).resolve()
    intake_path = Path(intake_path).resolve()
    keys = ("XAUFE_RENDER_ENGINE", "XAUFE_WPS_MODE", "XAUFE_WPS_EXCHANGE")
    saved = {key: os.environ.get(key) for key in keys}
    try:
        if "office_engine" in s:
            os.environ["XAUFE_RENDER_ENGINE"] = s["office_engine"]
        if "wps_mode" in s:
            os.environ["XAUFE_WPS_MODE"] = s["wps_mode"]
        if s.get("wps_mode") == "manual":
            os.environ["XAUFE_WPS_EXCHANGE"] = str(root / "wps-exchange")
        try:
            return _prepare(root, s, docx, intake_path, unavailable_note, office)
        except Exception as exc:
            from wps_exchange import ManualExportRequired

            if not isinstance(exc, ManualExportRequired) or not unavailable_note:
                raise
            # A note about an unusable GUI is not a skip switch: try every
            # callable automatic engine before entering the static-only branch.
            os.environ["XAUFE_WPS_MODE"] = "auto"
            os.environ["XAUFE_RENDER_ENGINE"] = "auto"
            root, s = load(workspace)
            return _prepare(root, s, docx, intake_path, unavailable_note, office)
    finally:
        for key, val in saved.items():
            if val is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = val


def _real_text(value, minimum=4):
    return isinstance(value, str) and len(value.strip()) >= minimum and PLACEHOLDER not in value


def attest(workspace, review):
    """Accept the agent's written review of the current final document."""
    root, s = load(workspace)
    if s.get("stage") not in ("awaiting_review", "ready"):
        raise ValueError("还没有可检查的成品；先运行 run。")
    if sha(root / "final.docx") != s.get("document_sha256"):
        raise ValueError("成品在 run 之后被改动；请重新运行 run 再检查。")
    if not isinstance(review, dict):
        raise ValueError("review.json 必须是 JSON 对象。")
    unknown = set(review) - {"vision", "no_vision_reason", "pages", *REVIEW_TEXTS, "warnings"}
    if unknown:
        raise ValueError("review.json 有不认识的项：" + "、".join(sorted(unknown)))
    for key, label in REVIEW_TEXTS.items():
        if not _real_text(review.get(key), 8):
            raise ValueError(f"review.json 的 {key}（{label}）没有写实际检查结果。")
    pages = review.get("pages", {})
    if not isinstance(pages, dict):
        raise ValueError('review.json 的 pages 必须写成 {"1": "第1页看到的情况", ...}。')
    vision = review.get("vision")
    if not isinstance(vision, bool):
        raise ValueError("review.json 的 vision 只能是 true（我能看图并看过了）或 false（我不能看图）。")
    if s.get("validation_scope") == "static-only":
        if any(_real_text(v) for v in pages.values()):
            raise ValueError("这次没有生成页图，pages 必须为空，不能写看图结果。")
        visual = "skipped_no_renderer"
    elif vision:
        expected = {str(p["page"]) for p in s["pages"]}
        missing = sorted(expected - set(pages), key=int)
        if missing:
            raise ValueError("每一页都要打开看并写一句观察，还缺第 " + "、".join(missing) + " 页。")
        extra = set(pages) - expected
        if extra:
            raise ValueError("pages 里有不存在的页码：" + "、".join(sorted(extra)))
        weak = sorted((k for k, v in pages.items() if not _real_text(v)), key=int)
        if weak:
            raise ValueError("这些页没有写具体看到的情况：第 " + "、".join(weak) + " 页。")
        for page in s["pages"]:
            if sha(root / "preview" / page["png"]) != page["png_sha256"]:
                raise ValueError("页图已变化；请重新运行 run 后再看图。")
        visual = "passed_agent_review"
    else:
        if any(_real_text(v) for v in pages.values()):
            raise ValueError("vision=false 时 pages 必须为空，不能写看图结果。")
        if not _real_text(review.get("no_vision_reason")):
            raise ValueError("vision=false 时要在 no_vision_reason 写明为什么不能看图。")
        visual = "skipped_no_vision"
    notes = review.get("warnings", {})
    if not isinstance(notes, dict):
        raise ValueError("review.json 的 warnings 必须是 {\"警告代码\": \"怎么处理的\"}。")
    missing = [code for code in s.get("warning_codes", []) if not _real_text(notes.get(code))]
    if missing:
        raise ValueError("这些警告还没写处理说明：" + "、".join(missing))
    s["review"] = dict(review, document_sha256=s["document_sha256"], visual_review=visual)
    s["stage"] = "ready"
    write_json(root / MARKER, s)
    return {"visual_review": visual}


def clean(workspace):
    root, s = load(workspace)
    dest = Path(s["destination"]).resolve()
    if dest.is_relative_to(root):
        raise ValueError("最终 Word 不能位于将要清理的工作区内。")
    # rmtree does not follow interior symlinks; root ownership was checked above.
    shutil.rmtree(root)
    if root.exists():
        raise OSError("工作区没有清理干净，不能说只剩最终 Word。")


def user_notice(s):
    review = s["review"]
    if s.get("validation_scope") == "static-only":
        return s["user_notice"]
    if review["visual_review"] == "skipped_no_vision":
        return (
            "已完成排版、格式检查和实际分页；没有逐页看图检查，原因："
            + review["no_vision_reason"].strip().rstrip("。")
            + "。建议打开文件快速翻看一遍。"
        )
    return ""


def publish(workspace):
    root, s = load(workspace)
    if s.get("stage") != "ready" or not s.get("review"):
        raise ValueError("还没有通过 review.json 检查，不能交付。")
    final = root / "final.docx"
    if sha(final) != s.get("document_sha256") or s["review"]["document_sha256"] != sha(final):
        raise ValueError("检查之后 Word 又被改动；请重新运行 run 并重新检查。")
    limited = s.get("validation_scope") == "static-only"
    receipt = RF.verify_receipt(root, s) if limited else None
    if not limited:
        if sha(root / "preview/render-report.json") != s.get("render_report_sha256"):
            raise ValueError("预览已变化，请重新运行 run。")
        for page in s["pages"]:
            if sha(root / "preview" / page["png"]) != page["png_sha256"]:
                raise ValueError("页图已变化，请重新运行 run。")
        check_preview_toc(root)
    verify_originals(s)
    source = s.get("active_request_source")
    if source and require(source)["integrity"] != require(root / "active-intake.json")["integrity"]:
        raise ValueError("用户要求或日期已变化，请重新运行 run。")
    result = check(
        final,
        final=True,
        intake_path=root / "active-intake.json",
        validation_scope="static-only" if limited else "full",
    )
    if not result["passed"]:
        raise ValueError("交付前复查未通过：" + json.dumps(result["errors"], ensure_ascii=False))
    dest = Path(s["destination"])
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.is_relative_to(root) or dest.is_relative_to(E.ROOT):
        raise ValueError("交付位置不安全。")
    created = False
    try:
        # Exclusive creation: never overwrite an existing output or source file.
        with dest.open("xb") as out, final.open("rb") as src:
            created = True
            shutil.copyfileobj(src, out)
            out.flush()
            os.fsync(out.fileno())
        if sha(dest) != s["document_sha256"]:
            raise OSError("复制后的 Word 校验不一致。")
    except Exception:
        if created and dest.exists():
            dest.unlink()
        raise
    notice = user_notice(s)
    clean(root)
    return {
        "output": str(dest),
        "workspace_removed": not root.exists(),
        "validation_scope": "static-only" if limited else "full",
        "visual_review": s["review"]["visual_review"],
        "skipped_steps": receipt["skipped_steps"] if limited else [],
        "user_notice": notice,
    }
