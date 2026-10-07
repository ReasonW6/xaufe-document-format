#!/usr/bin/env python3
"""Portable DOCX -> temporary PDF -> per-page PNG. Never resaves the source DOCX."""

from __future__ import annotations
import argparse, json, os, shutil, subprocess, sys, tempfile, hashlib
from pathlib import Path


def office_binary():
    """Locate LibreOffice for development checks that require its CLI."""
    from office_discovery import discover, usable_candidates

    candidates = [c for c in usable_candidates(discover()) if c["engine"] == "libreoffice"]
    if candidates:
        return candidates[0]["path"]
    raise RuntimeError(
        "当前没有可调用的LibreOffice；请先运行office_discovery.py。pdf_render也支持Windows Word/WPS及WPS界面导出交换路径。"
    )


def pdf_render(docx, outdir):
    from office_backend import export_pdf

    docx = Path(docx).resolve()
    outdir = Path(outdir).resolve()
    outdir.mkdir(parents=True, exist_ok=True)
    target = outdir / (docx.stem + ".pdf")
    source_hash = hashlib.sha256(docx.read_bytes()).hexdigest()
    backend = export_pdf(docx, target)
    if hashlib.sha256(docx.read_bytes()).hexdigest() != source_hash:
        raise RuntimeError("渲染过程改变原Word，必须停止；不得以重新保存的文件冒充仅预览。")
    from cover_date_audit import verify_pdf

    cover_result = verify_pdf(docx, target)
    if cover_result.get("required"):
        (outdir / "cover-date-layout.json").write_text(
            json.dumps(cover_result, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    (outdir / "office-backend.json").write_text(
        json.dumps(backend, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return target


def font_report():
    result = {}
    if shutil.which("fc-match"):
        for fam in ("宋体", "黑体", "Times New Roman", "Arial"):
            r = subprocess.run(
                ["fc-match", "-f", "%{family}", fam], text=True, capture_output=True, timeout=10
            )
            result[fam] = r.stdout.strip()
    else:
        result = {
            "notice": "请在 Word/WPS 字体列表中核对宋体、黑体、Times New Roman。未自动验证本机字体。"
        }
    return result


def render(docx, outdir, dpi=120):
    import pymupdf as fitz

    docx = Path(docx)
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    pdf = pdf_render(docx, outdir)
    d = fitz.open(pdf)
    for f in outdir.glob("page-*.png"):
        f.unlink()
    pages = []
    for i, page in enumerate(d):
        fn = outdir / f"page-{i+1:03d}.png"
        page.get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72), alpha=False).save(fn)
        pages.append(
            {
                "page": i + 1,
                "png": fn.name,
                "png_sha256": hashlib.sha256(fn.read_bytes()).hexdigest(),
                "text_chars": len(page.get_text()),
                "width_pt": page.rect.width,
                "height_pt": page.rect.height,
            }
        )
    d.close()
    result = {
        "source_sha256": hashlib.sha256(docx.read_bytes()).hexdigest(),
        "pages": pages,
        "font_resolution": font_report(),
        "visual_review": "pending",
        "notice": "这些 PDF/PNG 仅供检查。字体替代时不得称像素级一致；不得把字体文件装进技能包。",
    }
    (outdir / "render-report.json").write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return result


def main():
    p = argparse.ArgumentParser()
    p.add_argument("docx")
    p.add_argument("--output-dir", required=True)
    p.add_argument("--dpi", type=int, default=120)
    a = p.parse_args()
    try:
        r = render(a.docx, a.output_dir, a.dpi)
        print(
            json.dumps(
                {
                    "pages": len(r["pages"]),
                    "output_dir": a.output_dir,
                    "fonts": r["font_resolution"],
                },
                ensure_ascii=False,
            )
        )
        return 0
    except Exception as exc:
        print(str(exc), file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
