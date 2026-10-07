"""Word/LibreOffice/WPS PDF exporters; source DOCX is never resaved or overwritten."""

from __future__ import annotations
import json, os, re, shutil, subprocess, tempfile
from pathlib import Path
from office_discovery import discover, usable_candidates, ps_run

LAST_BACKEND = {}


class OfficeUnavailable(RuntimeError):
    """No callable automatic backend; carries actual discovery, never host claims."""

    def __init__(self, message, report, failures=None):
        super().__init__(message)
        self.report = report
        self.failures = failures or []


class OfficeExportFailed(OfficeUnavailable):
    """All exporters failed; delivery must test an independent smoke document."""

    pass


def _lo(docx, pdf, candidate, tmp):
    profile = tmp / "profile"
    profile.mkdir()
    od = tmp / "output"
    od.mkdir()
    cmd = [
        candidate["path"],
        "-env:UserInstallation=" + profile.as_uri(),
        "--headless",
        "--norestore",
        "--convert-to",
        "pdf:writer_pdf_Export",
        "--outdir",
        str(od),
        str(docx),
    ]
    env = os.environ.copy()
    # Do not change Windows user identity/profile; LO gets its own explicit profile.
    if os.name != "nt":
        env.update(
            HOME=str(tmp), XDG_CONFIG_HOME=str(tmp / "config"), XDG_CACHE_HOME=str(tmp / "cache")
        )
    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=180, env=env)
    produced = od / (docx.stem + ".pdf")
    if not produced.is_file() or not produced.stat().st_size:
        raise RuntimeError("LibreOffice已找到但导出失败：" + proc.stdout + "\n" + proc.stderr)
    shutil.copy2(produced, pdf)


def _word(docx, pdf, candidate, tmp):
    from word_com import run as run_word

    proc = run_word(ps_run, docx, pdf)
    if proc.returncode != 0 or not pdf.is_file() or not pdf.stat().st_size:
        context_hint = ""
        if "80070520" in (proc.stdout + proc.stderr).lower():
            context_hint = (
                "当前登录会话（常见于沙盒）无法创建 Word 的 COM 对象。若运行环境支持经用户批准后在本机用户会话里执行命令，"
                "就用那种方式重试同一条命令；这是调用文档接口，不需要屏幕操作权限。"
                "不要绕过权限；做不到时按“当前环境不能调用”处理。\n"
            )
        raise RuntimeError(
            "Word安装已找到，但自动化导出失败。检查激活/首次启动/权限/受保护视图；不要断言未安装。\n"
            + context_hint
            + proc.stdout
            + "\n"
            + proc.stderr
        )
    return proc.stdout.strip()


def export_pdf(docx, pdf):
    docx = Path(docx).resolve()
    pdf = Path(pdf).resolve()
    if docx == pdf:
        raise ValueError("预览输出不能覆盖原稿。")
    if os.environ.get("XAUFE_WPS_MODE") == "manual":
        if os.environ.get("XAUFE_RENDER_ENGINE", "auto") not in ("wps", "auto"):
            raise ValueError("WPS界面导出与显式选择的其他引擎冲突。")
        from wps_exchange import export_cached

        result = export_cached(docx, pdf, os.environ.get("XAUFE_WPS_EXCHANGE"))
        LAST_BACKEND.clear()
        LAST_BACKEND.update(result)
        return result
    report = discover()
    candidates = usable_candidates(report)
    failures = []
    # A missing selected application must not hide another installed renderer.
    # Keep the selected engine first, then try remaining real candidates.
    if report.get("candidates"):
        preference = os.environ.get("XAUFE_RENDER_ENGINE")
        try:
            os.environ["XAUFE_RENDER_ENGINE"] = "auto"
            for candidate in usable_candidates(report):
                if candidate not in candidates:
                    candidates.append(candidate)
        finally:
            if preference is None:
                os.environ.pop("XAUFE_RENDER_ENGINE", None)
            else:
                os.environ["XAUFE_RENDER_ENGINE"] = preference
    if not candidates:
        installed = [x for x in report["candidates"] if x["found"]]
        raise OfficeUnavailable(
            (
                "已发现Office安装，但当前运行环境不能自动调用。"
                if installed
                else "当前可访问环境未定位到可调用的Office。"
            )
            + "先运行office_discovery.py --probe并查看 references/special-cases.md；不能仅凭PATH失败认定用户电脑未安装。",
            report,
        )
    for candidate in candidates:
        try:
            with tempfile.TemporaryDirectory(prefix="xaufe-office-") as td:
                tmp = Path(td)
                produced = tmp / "render.pdf"
                if candidate["engine"] == "wps":
                    from wps_support import export as export_wps

                    version = export_wps(docx, produced, candidate, tmp, ps_run)
                else:
                    version = (
                        _word(docx, produced, candidate, tmp)
                        if candidate["engine"] == "word"
                        else _lo(docx, produced, candidate, tmp)
                    )
                import pymupdf as fitz

                # Failed path-based opens can retain a Windows file handle in
                # the parser traceback, preventing this temporary dir cleanup.
                with fitz.open(stream=produced.read_bytes(), filetype="pdf") as verified:
                    if not verified.is_pdf or verified.page_count < 1 or verified.needs_pass:
                        raise RuntimeError("Office导出物不是有效的未加密PDF。")
                pdf.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(produced, pdf)
            LAST_BACKEND.clear()
            LAST_BACKEND.update(
                engine=candidate["engine"],
                path=candidate["path"],
                version=version or "PDF export succeeded",
            )
            return dict(LAST_BACKEND)
        except (OSError, RuntimeError, subprocess.SubprocessError) as exc:
            failures.append(
                {"engine": candidate["engine"], "path": candidate["path"], "error": str(exc)}
            )
    raise OfficeExportFailed(
        "已找到Office但所有可用调用均失败；这是启动/权限/导出问题，不是已证实未安装。\n"
        + json.dumps(failures, ensure_ascii=False),
        report,
        failures,
    )
