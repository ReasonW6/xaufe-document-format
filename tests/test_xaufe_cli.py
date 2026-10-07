"""The single agent-facing entry point. All names and data below are synthetic.

Rendering is replaced by a fake page image here; real Office rendering is
exercised by run_acceptance.py. Review texts are unit-test records, not real checks.
"""

from __future__ import annotations
import contextlib, hashlib, io, json, sys, tempfile, unittest
from pathlib import Path
from unittest.mock import patch
from docx import Document

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import xaufe as X
import delivery as D
import intake as I

JOB_B = {
    "task": "format",
    "template": "B",
    "info": {
        "course_name": "数据结构",
        "title_zh": "排序算法比较",
        "student_name": "合成同学",
        "student_id": "20230001",
        "major": "软件工程",
        "class_name": "软工2301",
    },
}
CONTENT = {
    "mode": "paper",
    "metadata": {"title_zh": "排序算法比较"},
    "abstract_zh": "这是合成摘要。",
    "keywords_zh": ["排序", "算法"],
    "references": ["合成作者. 合成文献[M]. 北京: 合成出版社, 2020."],
    "blocks": [
        {"type": "heading", "level": 1, "text": "一、引言"},
        {"type": "paragraph", "text": "这是合成正文。"},
        {"type": "heading", "level": 2, "text": "（一）背景"},
        {"type": "paragraph", "text": "这是合成正文第二段。"},
    ],
}


def fake_render(docx, outdir):
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    png = outdir / "page-001.png"
    png.write_bytes(b"unit-test-only-not-a-real-page")
    result = {
        "source_sha256": hashlib.sha256(Path(docx).read_bytes()).hexdigest(),
        "pages": [{"page": 1, "png": png.name, "png_sha256": hashlib.sha256(png.read_bytes()).hexdigest()}],
    }
    D.write_json(outdir / "render-report.json", result)
    return result


class CLI(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.d = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def call(self, *args):
        out, err = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            code = X.main([str(a) for a in args])
        return code, out.getvalue(), err.getvalue()

    def new(self, job=None, content=None, sources=()):
        extra = []
        for s in sources:
            extra += ["--source", s]
        with patch("tempfile.gettempdir", return_value=str(self.d)):
            code, out, err = self.call("new", "--out", self.d / "out" / "成品.docx", *extra)
        self.assertEqual(code, 0, err)
        work = Path(out.split("WORK = ")[1].splitlines()[0].strip())
        if job is not None:
            job = {"task": "format", **job}
            (work / "job.json").write_text(json.dumps(job, ensure_ascii=False), encoding="utf-8")
        if content is not None:
            (work / "content.json").write_text(json.dumps(content, ensure_ascii=False), encoding="utf-8")
        return work

    def fill_review(self, work, vision=True):
        review = json.loads((work / "review.json").read_text(encoding="utf-8"))
        review["vision"] = vision
        review["pages"] = {k: "单元测试合成记录：这一页正常" for k in review["pages"]} if vision else {}
        if not vision:
            review["no_vision_reason"] = "单元测试：模拟不能看图"
        for key in ("content_check", "layout_check", "delivery_check"):
            review[key] = "单元测试合成记录：已核对"
        review["warnings"] = {k: "单元测试合成记录：已处理" for k in review["warnings"]}
        (work / "review.json").write_text(json.dumps(review, ensure_ascii=False), encoding="utf-8")

    # ------------------------------------------------------------ choosing
    def test_compare_prints_user_table_without_internal_ids(self):
        code, out, _ = self.call("compare")
        self.assertEqual(code, 0)
        self.assertTrue(out.startswith("| 对比项目 |"))
        self.assertIn("请选择 **A（学年论文）** 或 **B（课程作业）**", out)
        for internal in ("green-paper", "red-course", "模板标识", "template"):
            self.assertNotIn(internal, out)

    def test_check_asks_only_missing_fields(self):
        job = {"template": "A", "logo": "red", "info": {"title_zh": "合成题目", "student_name": "合成同学"}}
        work = self.new(job)
        code, out, _ = self.call("check", "--work", work)
        self.assertEqual(code, X.EXIT_ASK_USER)
        self.assertIn("采用A（本科学年论文），配红色校徽。", out)
        self.assertIn("还缺：学号、专业、班级、指导教师。这些项也可以留空", out)
        self.assertNotIn("题目", out.split("还缺：")[1].split("。")[0])

    def test_blank_and_default_date(self):
        job = dict(JOB_B, blank=["student_id"])
        job["info"] = {k: v for k, v in JOB_B["info"].items() if k != "student_id"}
        work = self.new(job)
        code, out, _ = self.call("check", "--work", work)
        self.assertEqual(code, 0)
        self.assertIn("按要求留空：学号", out)
        self.assertIn("日期采用执行当日", out)

    def test_job_errors_are_plain_and_actionable(self):
        cases = [
            ({"template": "C"}, '"A" 或 "B"'),
            ({"template": "B", "info": {"supervisor": "某老师"}}, "cover_supervisor"),
            ({"template": "A", "format": {"logo": "red"}}, "顶层的 logo"),
            ({"template": "A", "colour": "red"}, "不认识的项"),
            ({"template": "A", "info": {"student_name": "甲"}, "blank": ["student_name"]}, "既填写又留空"),
        ]
        for job, message in cases:
            with self.subTest(message=message):
                work = self.new(job)
                code, _, err = self.call("check", "--work", work)
                self.assertEqual(code, 2)
                self.assertIn(message, err)

    def test_job_must_say_write_or_format(self):
        work = self.new()
        job = {k: v for k, v in JOB_B.items() if k != "task"}
        (work / "job.json").write_text(json.dumps(job, ensure_ascii=False), encoding="utf-8")
        code, _, err = self.call("check", "--work", work)
        self.assertEqual(code, 2)
        self.assertIn('"task": "write"', err)

    def test_job_written_as_utf16_or_gbk_is_read(self):
        for encoding in ("utf-16", "gbk"):
            with self.subTest(encoding=encoding):
                work = self.new()
                (work / "job.json").write_bytes(json.dumps(JOB_B, ensure_ascii=False).encode(encoding))
                code, out, err = self.call("check", "--work", work)
                self.assertEqual(code, 0, err)

    def test_null_in_format_means_default(self):
        state = I.from_job(
            dict(JOB_B, format={"styles": {"body": {"cn": "仿宋", "size": None}}}), self.d / "s.json"
        )
        self.assertEqual(state["format_request"], {"styles": {"body": {"cn": "仿宋"}}})

    def test_same_workspace_gives_same_task_id(self):
        a = I.from_job(JOB_B, self.d / "x.json")["job_id"]
        b = I.from_job(JOB_B, self.d / "x.json")["job_id"]
        self.assertEqual(a, b)

    # ------------------------------------------------------------ run / finish
    def test_run_refuses_when_info_missing(self):
        work = self.new({"template": "B"}, CONTENT)
        code, out, _ = self.call("run", "--work", work)
        self.assertEqual(code, X.EXIT_ASK_USER)
        self.assertIn("还缺", out)
        self.assertFalse((work / "draft.docx").exists())

    def test_static_run_and_finish_deliver_one_word(self):
        work = self.new(JOB_B, CONTENT)
        with patch("office_discovery.discover") as discover:
            code, out, err = self.call("run", "--work", work, "--office", "none")
        discover.assert_not_called()
        self.assertEqual(code, 0, err)
        self.assertIn("需要在最终回复里告诉用户", out)
        self.assertIn("正文约 15 字", out)
        self.assertIn("参考文献 1 条", out)
        review = json.loads((work / "review.json").read_text(encoding="utf-8"))
        self.assertEqual(review["pages"], {})
        code, out, err = self.call("finish", "--work", work)
        self.assertEqual(code, 2)
        self.assertIn("content_check", err)
        self.fill_review(work, vision=False)
        code, out, err = self.call("finish", "--work", work)
        self.assertEqual(code, 0, err)
        self.assertIn("必须在给用户的回复里说明", out)
        self.assertFalse(work.exists())
        self.assertEqual([p.name for p in (self.d / "out").iterdir()], ["成品.docx"])

    def test_rendered_run_lists_pages_and_needs_every_page(self):
        # No table of contents here: real TOC paging is covered by run_acceptance.py.
        document = {"mode": "document", "metadata": {"title_zh": "排序算法比较"}, "blocks": CONTENT["blocks"]}
        work = self.new(JOB_B, document)
        with patch("delivery.render", side_effect=fake_render), patch("office_discovery.discover", return_value={}):
            code, out, err = self.call("run", "--work", work)
        self.assertEqual(code, 0, err)
        self.assertIn("第1页：", out)
        review = json.loads((work / "review.json").read_text(encoding="utf-8"))
        self.assertIn("请填写", review["pages"]["1"])
        self.fill_review(work)
        with patch("delivery.check_preview_toc", return_value={"required": False}):
            code, out, err = self.call("finish", "--work", work)
        self.assertEqual(code, 0, err)
        self.assertIn("已完成排版、分页和逐页检查", out)

    def test_cover_overflow_retries_with_next_fit_level(self):
        from cover_date_audit import CoverOverflow

        work = self.new(JOB_B, CONTENT)
        levels = []
        real_build = X._build

        def build(root, level):
            levels.append(level)
            return real_build(root, level)

        calls = {"n": 0}

        def prepare(*args, **kwargs):
            calls["n"] += 1
            if calls["n"] == 1:
                raise CoverOverflow("合成：日期被挤到第二页")
            return {"validation_scope": "static-only", "pages": [], "warnings": [], "user_notice": "合成说明"}

        with patch("xaufe._build", side_effect=build), patch("delivery.prepare", side_effect=prepare):
            code, out, err = self.call("run", "--work", work)
        self.assertEqual(code, 0, err)
        self.assertEqual(levels, [0, 1])
        self.assertIn("自动收紧封面空白", out)

    def test_both_content_and_mapping_is_refused(self):
        work = self.new(JOB_B, CONTENT)
        (work / "mapping.json").write_text("{}", encoding="utf-8")
        code, _, err = self.call("run", "--work", work, "--office", "none")
        self.assertEqual(code, 2)
        self.assertIn("只能保留一个", err)

    # ------------------------------------------------------------ existing Word
    def sample_docx(self):
        doc = Document()
        for text in [
            "排序算法比较",
            "摘要",
            "本文比较了几种排序算法。",
            "关键词：排序   算法",
            "一、引言",
            "这是正文。",
            "（一）背景",
            "1.问题",
            "正文内容很长，这一段不是标题。",
            "英文摘要待删除的合成段落",
            "参考文献",
            "[1] 合成作者. 合成文献[M]. 2020.",
        ]:
            doc.add_paragraph(text)
        path = self.d / "原稿.docx"
        doc.save(path)
        return path

    def test_read_prints_document_text_once(self):
        source = self.sample_docx()
        code, out, err = self.call("read", "--docx", source)
        self.assertEqual(code, 0, err)
        self.assertEqual(out.count("排序算法比较"), 1)
        self.assertIn("一、引言", out)
        self.assertIn("[1] 合成作者", out)

    def test_inspect_guesses_roles(self):
        work = self.new(JOB_B)
        source = self.sample_docx()
        code, out, err = self.call("inspect", "--work", work, "--docx", source)
        self.assertEqual(code, 0, err)
        roles = [r["role"] for r in json.loads((work / "mapping.json").read_text(encoding="utf-8"))["paragraphs"]]
        self.assertEqual(
            roles,
            ["title_zh", "abstract_label_zh", "abstract_zh", "keywords_zh", "h1", "body", "h2", "h3", "body", "body", "reference_title", "reference"],
        )
        self.assertIn("加上所选模板的封面", out)
        mapping = json.loads((work / "mapping.json").read_text(encoding="utf-8"))
        self.assertEqual(mapping["replace_cover_before"], 0)
        self.assertIn(str(source), [x["path"] for x in D.load(work)[1]["source_documents"]])

    def test_mapping_delete_role_removes_only_marked_paragraph(self):
        work = self.new(JOB_B)
        source = self.sample_docx()
        self.call("inspect", "--work", work, "--docx", source)
        mapping = json.loads((work / "mapping.json").read_text(encoding="utf-8"))
        mapping["paragraphs"][9]["role"] = "delete"
        (work / "mapping.json").write_text(json.dumps(mapping, ensure_ascii=False), encoding="utf-8")
        code, out, err = self.call("run", "--work", work, "--office", "none")
        self.assertEqual(code, 0, err)
        self.assertIn("删除了 1 段", out)
        text = "\n".join(p.text for p in Document(work / "final.docx").paragraphs)
        self.assertNotIn("英文摘要待删除", text)
        self.assertIn("正文内容很长", text)
        self.assertIn("SOURCE_DIFFERENCES", out)

    def paper_run(self, template, add=None, edit=None):
        info = dict(JOB_B["info"])
        if template == "A":
            info.pop("course_name")
        job = {"template": template, "info": info}
        if template == "A":
            job["blank"] = ["supervisor"]
        work = self.new(job)
        source = self.sample_docx()
        code, out, err = self.call("inspect", "--work", work, "--docx", source)
        self.assertEqual(code, 0, err)
        mapping = json.loads((work / "mapping.json").read_text(encoding="utf-8"))
        self.assertEqual(mapping["mode"], "paper")
        mapping["paragraphs"][9]["role"] = "body"
        if add:
            mapping["add"] = add
        if edit:
            edit(mapping)
        (work / "mapping.json").write_text(json.dumps(mapping, ensure_ascii=False), encoding="utf-8")
        code, run_out, err = self.call("run", "--work", work, "--office", "none")
        return work, out, code, run_out, err

    def test_existing_paper_gets_template_structure_b(self):
        work, _, code, out, err = self.paper_run("B")
        self.assertEqual(code, 0, err)
        doc = Document(work / "final.docx")
        styles = [p.style.name for p in doc.paragraphs if p.text.strip()]
        # cover, abstract page, TOC, body: abstract before TOC in B
        self.assertLess(styles.index("XAUFE Abstract Chinese"), styles.index("XAUFE TOC Title"))
        self.assertLess(styles.index("XAUFE TOC Title"), styles.index("Heading 1"))
        self.assertEqual(len(doc.sections), 4)
        toc = [p.text for p in doc.paragraphs if p.style.name.startswith("TOC ")]
        self.assertTrue(any("一、引言" in t for t in toc))
        self.assertTrue(any("参考文献" in t for t in toc))
        text = "\n".join(p.text for p in doc.paragraphs)
        for original in ("这是正文。", "正文内容很长，这一段不是标题。", "[1] 合成作者"):
            self.assertIn(original, text)

    def test_existing_paper_a_needs_english_abstract_decision(self):
        work, inspect_out, code, out, err = self.paper_run("A")
        self.assertIn("先问用户", inspect_out)
        self.assertEqual(code, 2)
        self.assertIn("english_abstract", err)

    def test_existing_paper_a_with_translated_english_abstract(self):
        add = {"title_en": "Comparison of Sorting Algorithms", "abstract_en": "A synthetic English abstract.", "keywords_en": ["sorting", "algorithms"]}
        work, _, code, out, err = self.paper_run("A", add=add)
        self.assertEqual(code, 0, err)
        doc = Document(work / "final.docx")
        styles = [p.style.name for p in doc.paragraphs if p.text.strip()]
        order = [styles.index(s) for s in ("XAUFE TOC Title", "XAUFE Abstract Chinese", "XAUFE Abstract English", "Heading 1")]
        self.assertEqual(order, sorted(order))
        self.assertEqual(len(doc.sections), 5)

    def test_old_manual_toc_is_replaced_by_generated_one(self):
        doc = Document()
        for text in ["排序算法比较", "摘要", "本文比较了几种排序算法。", "关键词：排序   算法", "目录",
                     "一、引言……………1", "二、方法……………2", "一、引言", "这是正文。", "二、方法", "方法正文。", "参考文献",
                     "[1] 合成作者. 合成文献[M]. 2020."]:
            doc.add_paragraph(text)
        source = self.d / "带目录.docx"
        doc.save(source)
        work = self.new(JOB_B)
        code, out, err = self.call("inspect", "--work", work, "--docx", source)
        self.assertEqual(code, 0, err)
        roles = [r["role"] for r in json.loads((work / "mapping.json").read_text(encoding="utf-8"))["paragraphs"]]
        self.assertEqual(roles[4:7], ["delete", "delete", "delete"])
        self.assertIn("旧目录", out)
        code, run_out, err = self.call("run", "--work", work, "--office", "none")
        self.assertEqual(code, 0, err)
        final = Document(work / "final.docx")
        self.assertEqual(sum(1 for p in final.paragraphs if p.style.name == "XAUFE TOC Title"), 1)
        self.assertNotIn("……………", "\n".join(p.text for p in final.paragraphs))

    def test_english_title_and_spaced_keyword_labels_are_kept(self):
        doc = Document()
        for text in ["排序算法比较", "摘要", "本文比较了几种排序算法。", "关键词：排序 算法 ", "Comparison of Sorting Algorithms",
                     "Abstract", "This paper compares sorting algorithms.", "Key words: sorting; algorithms",
                     "一、引言", "这是正文。", "参考文献", "[1] 合成作者. 合成文献[M]. 2020."]:
            doc.add_paragraph(text)
        source = self.d / "英文题目.docx"
        doc.save(source)
        work = self.new({**JOB_B, "template": "A", "info": {k: v for k, v in JOB_B["info"].items() if k != "course_name"}, "blank": ["supervisor"]})
        code, out, err = self.call("inspect", "--work", work, "--docx", source)
        self.assertEqual(code, 0, err)
        roles = [r["role"] for r in json.loads((work / "mapping.json").read_text(encoding="utf-8"))["paragraphs"]]
        self.assertEqual(roles[4], "title_en")
        code, out, err = self.call("run", "--work", work, "--office", "none")
        self.assertEqual(code, 0, err)
        text = [p.text for p in Document(work / "final.docx").paragraphs]
        self.assertIn("Key words: sorting; algorithms", text)
        self.assertIn("关键词：排序   算法", text)

    # ------------------------------------------------------------ writing mode
    def written(self, content, refs=None):
        work = self.new({**JOB_B, "task": "write"}, content)
        if refs is not None:
            (work / "refs.json").write_text(json.dumps(refs, ensure_ascii=False), encoding="utf-8")
            book = {"title": "Python for Data Analysis", "publish_date": "2017"}
            with patch("refcheck.fetch", return_value=book):
                code, out, err = self.call("refs", "--work", work)
            self.assertEqual(code, 0, err)
        return work

    def test_writing_mode_blocks_unchecked_refs_unsourced_numbers_and_numbering_gaps(self):
        content = json.loads(json.dumps(CONTENT))
        content["blocks"] += [
            {"type": "paragraph", "text": "调查显示，女生月均消费为 710 元。"},
            {"type": "heading", "level": 1, "text": "三、结论"},
        ]
        work = self.written(content)
        code, out, err = self.call("run", "--work", work, "--office", "none")
        self.assertEqual(code, 2)
        self.assertIn("参考文献还没有核对", err)
        self.assertIn("710 元", err)
        self.assertIn("标题编号不连续", err)
        self.assertFalse((work / "draft.docx").exists())

    def test_writing_mode_passes_and_finish_prints_checklist(self):
        entry = {"text": "MCKINNEY W. Python for Data Analysis[M]. Sebastopol: O'Reilly Media, 2017.",
                 "isbn": "9781491957660", "title": "Python for Data Analysis", "year": 2017,
                 "supports": "pandas 适合做数据清洗", "relevance": "高"}
        content = json.loads(json.dumps(CONTENT))
        content["references"] = [entry["text"]]
        content["blocks"].append({"type": "paragraph", "runs": [
            {"text": "pandas 的使用者超过 100 万人"}, {"text": "[1]", "superscript": True}, {"text": "。"}]})
        work = self.written(content, [entry])
        code, out, err = self.call("run", "--work", work, "--office", "none")
        self.assertEqual(code, 0, err)
        self.fill_review(work, vision=False)
        code, out, err = self.call("finish", "--work", work)
        self.assertEqual(code, 0, err)
        self.assertIn("参考文献核实清单", out)
        self.assertIn("| [1] |", out)
        self.assertIn("已核对", out)
        self.assertIn("pandas 适合做数据清洗", out)
        self.assertIn("少 5 篇", out)

    def test_writing_mode_rejects_reference_not_in_refs_json(self):
        entry = {"text": "合成作者. 合成书[M]. 2017.", "isbn": "9781491957660", "title": "Python for Data Analysis", "year": 2017}
        content = json.loads(json.dumps(CONTENT))
        content["references"] = ["编造作者. 编造的书[M]. 2020."]
        work = self.written(content, [entry])
        code, out, err = self.call("run", "--work", work, "--office", "none")
        self.assertEqual(code, 2)
        self.assertIn("参考文献第 1 条不在 refs 核对可用的条目里", err)

    def test_existing_paper_keyword_spaces_normalized(self):
        work, _, code, out, err = self.paper_run("B")
        self.assertEqual(code, 0, err)
        keywords = next(p.text for p in Document(work / "final.docx").paragraphs if p.text.startswith("关键词"))
        self.assertEqual(keywords, "关键词：排序   算法")

    def test_inspect_recognises_this_skills_own_output(self):
        work = self.new(JOB_B, CONTENT)
        self.call("run", "--work", work, "--office", "none")
        previous = self.d / "上一版.docx"
        previous.write_bytes((work / "final.docx").read_bytes())
        work2 = self.new(dict(JOB_B, logo="green"))
        code, out, err = self.call("inspect", "--work", work2, "--docx", previous)
        self.assertEqual(code, 0, err)
        mapping = json.loads((work2 / "mapping.json").read_text(encoding="utf-8"))
        roles = {r["text"]: r["role"] for r in mapping["paragraphs"] if r["text"]}
        self.assertEqual(roles["一、引言"], "h1")
        self.assertEqual(roles["这是合成正文。"], "body")
        self.assertEqual(roles["（一）背景"], "h2")
        self.assertGreater(mapping["replace_cover_before"], 0)
        self.assertIn("旧封面", out)

    # ------------------------------------------------------------ instructions
    def test_skill_md_keeps_the_core_rules(self):
        text = (ROOT / "SKILL.md").read_text(encoding="utf-8")
        for phrase in (
            "版式和校徽是两回事",
            "只排版",
            "默认写作",
            "先问清主题、方向和内容",
            "留空自己填",
            "今天（北京时间）",
            "逐页打开查看",
            "只改用户这次说的那一项",
            "scripts/xaufe.py compare",
            "finish",
        ):
            self.assertIn(phrase, text)


if __name__ == "__main__":
    unittest.main()
