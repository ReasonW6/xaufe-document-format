#!/usr/bin/env python3
"""Offline WPS GUI export bridge with resumable, content-addressed requests.

The agent (with actual desktop tools) or the user exports the requested DOCX in
WPS and accepts the PDF. This is not an automatic GUI controller. An acceptance
is an exporter attestation plus byte checks, not cryptographic proof of WPS use.
"""

from __future__ import annotations
import argparse
import hashlib
import json
from pathlib import Path
import shutil
import sys
from zipfile import ZipFile


class ManualExportRequired(RuntimeError):
    def __init__(self, request):
        self.request = Path(request)
        super().__init__(
            "WPS_EXPORT_REQUIRED: "
            + str(self.request)
            + "。使用已授权的桌面工具在WPS打开该请求的source.docx，导出含标题书签的PDF；运行 python -X utf8 scripts/wps_exchange.py accept，然后重新运行同一条 xaufe.py run。不猜页码，不改源DOCX。"
        )


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def read(path):
    return json.loads(Path(path).read_text(encoding="utf-8"))


def write(path, value):
    Path(path).write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def package_digest(path):
    """Ignore only ZIP container timestamps; compare all uncompressed member bytes."""
    h = hashlib.sha256()
    with ZipFile(path) as z:
        names = z.namelist()
        if len(names) != len(set(names)):
            raise ValueError("DOCX含重复ZIP成员。")
        if sum(info.file_size for info in z.infolist()) > 512 * 1024 * 1024:
            raise ValueError("DOCX超出安全解包检查上限。")
        for name in sorted(names):
            data = z.read(name)
            encoded = name.encode("utf-8")
            h.update(len(encoded).to_bytes(8, "big"))
            h.update(encoded)
            h.update(len(data).to_bytes(8, "big"))
            h.update(data)
    return h.hexdigest()


def exchange_root(value):
    if not value:
        raise ValueError(
            "WPS界面导出需要任务工作区；请用 xaufe.py run --office wps-manual。"
        )
    raw = Path(value).expanduser()
    if raw.is_symlink():
        raise ValueError("交换目录不能是符号链接。")
    root = raw.resolve()
    from delivery import load

    ancestor = next((p for p in root.parents if (p / ".xaufe-session.json").is_file()), None)
    if ancestor is None:
        raise ValueError("交换目录必须在本次delivery独占工作区内。")
    load(ancestor)
    root.mkdir(parents=True, exist_ok=True)
    return root


def validate_request(path):
    path = Path(path).resolve()
    root = exchange_root(path.parent.parent)
    if path.name != "request.json" or path.parent.parent != root:
        raise ValueError("请求路径无效。")
    request = read(path)
    source = path.parent / "source.docx"
    if request.get("version") != 1 or request.get("package_sha256") != path.parent.name:
        raise ValueError("请求身份不符。")
    if (
        source.is_symlink()
        or not source.is_file()
        or sha(source) != request.get("source_sha256")
        or package_digest(source) != request["package_sha256"]
    ):
        raise ValueError("请求源文件被修改，不能使用旧导出证明。")
    return request, source


def validate_pdf(source, pdf):
    import pymupdf as fitz
    from docx import Document
    from audit_helpers import toc_entries
    from refresh_toc import pdf_positions, norm
    from cover_date_audit import verify_pdf

    with fitz.open(pdf) as rendered:
        if not rendered.is_pdf or rendered.needs_pass or rendered.page_count < 1:
            raise ValueError("需要完整、未加密的PDF。")
        pages = len(rendered)
        entries = toc_entries(Document(source))
        if entries:
            _, measurement, _ = pdf_positions(pdf, entries)
            for entry, page in zip(entries, measurement["physical_pages"]):
                if (
                    page < 1
                    or page > pages
                    or norm(entry["text"]) not in norm(rendered[page - 1].get_text())
                ):
                    raise ValueError(
                        "PDF书签指向的页面没有对应标题；请检查WPS导出设置，不补造书签或页码。"
                    )
    cover = verify_pdf(source, pdf)
    return {"pages": pages, "toc_entries": len(entries), "cover": cover}


def export_cached(docx, pdf, exchange):
    root = exchange_root(exchange)
    key = package_digest(docx)
    folder = root / key
    if folder.is_symlink():
        raise ValueError("交换记录目录不能是符号链接。")
    folder.mkdir(exist_ok=True)
    request_path = folder / "request.json"
    source = folder / "source.docx"
    if not request_path.exists():
        shutil.copy2(docx, source)
        write(
            request_path,
            {
                "version": 1,
                "package_sha256": key,
                "source_sha256": sha(source),
                "source": str(source),
                "pdf_destination": str(folder / "export.pdf"),
                "instruction": "在WPS只打开本source.docx，不重新保存Word；全篇导出PDF，启用标题书签。随后运行 accept，再重新运行同一条 xaufe.py run。",
            },
        )
    request, source = validate_request(request_path)
    proof_path = folder / "accepted.json"
    cached = folder / "accepted.pdf"
    if not proof_path.is_file():
        raise ManualExportRequired(request_path)
    proof = read(proof_path)
    if (
        proof.get("source_sha256") != request["source_sha256"]
        or proof.get("package_sha256") != key
        or not cached.is_file()
        or sha(cached) != proof.get("pdf_sha256")
    ):
        raise ValueError("WPS导出记录/文件已变化，不能复用。")
    validate_pdf(source, cached)
    pdf = Path(pdf)
    pdf.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(cached, pdf)
    return {
        "engine": "wps-manual",
        "exporter": proof["exporter"],
        "evidence": proof["evidence"],
        "source_package_sha256": key,
        "source_export_sha256": request["source_sha256"],
        "pdf_sha256": proof["pdf_sha256"],
        "native_automation": False,
        "note": "实际导出来源由执行者如实声明；字节与完整布局结构已核对，不声称程序自动操控了WPS。",
    }


def accept(request_path, pdf, exporter, evidence):
    request, source = validate_request(request_path)
    if (
        not isinstance(exporter, str)
        or not exporter.strip()
        or not isinstance(evidence, str)
        or not evidence.strip()
    ):
        raise ValueError("必须如实记录导出软件/版本和实际操作；不能用示例记录伪造WPS实测。")
    pdf = Path(pdf)
    if pdf.is_symlink():
        raise ValueError("导出PDF不能是符号链接。")
    pdf = pdf.resolve()
    folder = Path(request_path).resolve().parent
    if not pdf.is_file():
        raise ValueError("导出PDF不存在或是符号链接。")
    stats = validate_pdf(source, pdf)
    source_hash = sha(source)
    pdf_hash = sha(pdf)
    if source_hash != request["source_sha256"]:
        raise ValueError("验收期间请求源文档改变。")
    target = folder / "accepted.pdf"
    if target != pdf:
        shutil.copy2(pdf, target)
    proof = {
        "version": 1,
        "package_sha256": request["package_sha256"],
        "source_sha256": source_hash,
        "pdf_sha256": pdf_hash,
        "exporter": exporter,
        "evidence": evidence,
        "measurements": stats,
    }
    write(folder / "accepted.json", proof)
    return {
        "accepted": True,
        **proof,
        "next": "重新运行同一条 xaufe.py run；如目录页码变化，会产生下一份导出请求，照做直到不再提示。",
    }


def main():
    p = argparse.ArgumentParser(
        description="接收WPS真实导出的PDF，继续相同目录、视觉和单Word交付验收。"
    )
    sub = p.add_subparsers(dest="command", required=True)
    a = sub.add_parser("accept")
    a.add_argument("--request", required=True)
    a.add_argument("--pdf", required=True)
    a.add_argument("--exporter", required=True)
    a.add_argument("--evidence", required=True)
    a = p.parse_args()
    try:
        print(
            json.dumps(
                accept(a.request, a.pdf, a.exporter, a.evidence), ensure_ascii=False, indent=2
            )
        )
        return 0
    except Exception as exc:
        print("WPS导出验收未通过：" + str(exc), file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
