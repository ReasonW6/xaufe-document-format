#!/usr/bin/env python3
"""Real-render acceptance through the agent-facing commands (synthetic data only).

Each case runs `xaufe.py new / run` with the Office software available on this
machine, copies the final Word and page images to --output-dir for a person (or
model) to inspect page by page, then cancels its own workspace. Nothing here
signs a visual review.

    python -X utf8 tests/run_acceptance.py --output-dir 技能目录外的空文件夹
"""

from __future__ import annotations
import argparse, contextlib, io, json, shutil, sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import xaufe as X
import delivery as D

LONG = {
    "title_zh": "基于多源异构数据融合与深度强化学习的城市交通信号自适应协同控制方法及其在大规模路网中的仿真验证研究",
    "student_name": "欧阳合成",
    "student_id": "2023123456789",
    "major": "计算机科学与技术（人工智能与大数据方向）",
    "class_name": "计算机科学与技术2023级卓越工程师实验班",
    "supervisor": "合成导师",
    "course_name": "面向对象程序设计与软件工程综合实践（双语）",
}


def example(template):
    job = json.loads((ROOT / f"examples/{template}/job.json").read_text(encoding="utf-8"))
    content = json.loads((ROOT / f"examples/{template}/content.json").read_text(encoding="utf-8"))
    for block in content["blocks"]:
        if block["type"] == "image":
            block["path"] = str(ROOT / "examples/layout-test.png")
    return job, content


def long_job(template, logo=None):
    job, content = example(template)
    fields = {"A": ("title_zh", "student_name", "student_id", "major", "class_name", "supervisor"),
              "B": ("course_name", "title_zh", "student_name", "student_id", "major", "class_name")}[template]
    job = {"task": "format", "template": template, "info": {k: LONG[k] for k in fields}}
    if logo:
        job["logo"] = logo
    content["metadata"]["title_zh"] = LONG["title_zh"]
    return job, content


def existing_paper(path):
    """A plain student paper as it might arrive: Word headings, abstract, table, references."""
    from docx import Document

    doc = Document()
    doc.add_paragraph("高校图书馆座位预约系统的设计与实现")
    doc.add_paragraph("摘要")
    doc.add_paragraph("本文设计并实现了一个基于Web的座位预约系统，用于合成排版验收。")
    doc.add_paragraph("关键词：座位预约 Web系统 图书馆管理")
    for heading, text in (("一、引言", "合成正文第一段。"), ("二、系统设计", "合成正文第二段。"), ("三、结论", "合成结论。")):
        doc.add_heading(heading, level=1)
        doc.add_paragraph(text)
    table = doc.add_table(rows=2, cols=2)
    table.style = "Table Grid"
    for i, row in enumerate((("表名", "说明"), ("seat", "座位信息"))):
        for j, value in enumerate(row):
            table.cell(i, j).text = value
    doc.add_paragraph("参考文献")
    doc.add_paragraph("[1] 合成作者. 合成文献[M]. 北京: 合成出版社, 2020.")
    doc.save(path)
    return path


def call(*args):
    out, err = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        code = X.main([str(a) for a in args])
    return code, out.getvalue(), err.getvalue()


def run_case(out_dir, name, job, content=None, docx=None, mapping_edit=None, office=None):
    target = out_dir / name
    target.mkdir()
    code, text, err = call("new", "--out", target / "成品.docx", *(["--source", docx] if docx else []))
    if code:
        raise RuntimeError(err)
    work = Path(text.split("WORK = ")[1].splitlines()[0].strip())
    try:
        (work / "job.json").write_text(json.dumps(job, ensure_ascii=False, indent=2), encoding="utf-8")
        if content is not None:
            (work / "content.json").write_text(json.dumps(content, ensure_ascii=False, indent=2), encoding="utf-8")
        else:
            code, text, err = call("inspect", "--work", work, "--docx", docx)
            if code:
                raise RuntimeError(err)
            if mapping_edit:
                mapping = json.loads((work / "mapping.json").read_text(encoding="utf-8"))
                mapping_edit(mapping)
                (work / "mapping.json").write_text(json.dumps(mapping, ensure_ascii=False), encoding="utf-8")
        code, text, err = call("run", "--work", work, *(["--office", office] if office else []))
        (target / "run-output.txt").write_text(text + err, encoding="utf-8")
        if code:
            raise RuntimeError(err or text)
        shutil.copy2(work / "final.docx", target / "final.docx")
        for page in sorted((work / "preview").glob("page-*.png")) if (work / "preview").exists() else []:
            shutil.copy2(page, target / page.name)
        state = D.load(work)[1]
        row = {"case": name, "pages": len(state.get("pages", [])), "scope": state["validation_scope"],
               "cover_fit": __import__("engine").get_vars(__import__("docx").Document(work / "final.docx")).get("XAUFE_COVER_FIT")}
        print(json.dumps(row, ensure_ascii=False), flush=True)
        return row
    finally:
        call("cancel", "--work", work)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", required=True)
    a = parser.parse_args()
    out = Path(a.output_dir).resolve()
    if out.is_relative_to(ROOT) or (out.exists() and any(out.iterdir())):
        raise ValueError("输出须在技能目录之外且为空。")
    out.mkdir(parents=True, exist_ok=True)
    rows = []
    job, content = example("A")
    rows.append(run_case(out, "01-A-default", job, content))
    job, content = example("B")
    rows.append(run_case(out, "02-B-green-logo-pagenum-blank-id", job, content))
    job, content = long_job("A", "red")
    rows.append(run_case(out, "03-A-red-logo-long-cover", job, content))
    job, content = long_job("B")
    rows.append(run_case(out, "04-B-long-cover", job, content))
    job, content = example("B")
    job["format"] = {
        "footer_page_numbers": True,
        "styles": {"body": {"cn": "仿宋", "size": 12}, "h1": {"color": "0070C0"}, "toc1": {"size": 12}, "toc2": {"size": 12}},
    }
    rows.append(run_case(out, "05-B-custom-formats", job, content))
    previous = out / "02-B-green-logo-pagenum-blank-id" / "final.docx"
    job, _ = example("B")
    job["logo"] = "red"
    rows.append(run_case(out, "06-revise-previous-logo-red", job, docx=previous))
    job, content = example("A")
    rows.append(run_case(out, "07-A-static-no-office", job, content, office="none"))
    existing = existing_paper(out / "existing-paper.docx")
    job, _ = example("A")
    job["info"]["title_zh"] = "高校图书馆座位预约系统的设计与实现"

    def translated(mapping):
        mapping["add"] = {
            "title_en": "Design and Implementation of a Library Seat Reservation System",
            "abstract_en": "Synthetic English abstract used only for layout acceptance.",
            "keywords_en": ["seat reservation", "web system"],
        }

    rows.append(run_case(out, "08-A-existing-paper-structured", job, docx=existing, mapping_edit=translated))
    (out / "summary.json").write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    print("页图和成品已复制到", out, "；请逐页查看。工作区已全部清理。")


if __name__ == "__main__":
    main()
