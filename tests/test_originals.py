"""Originals the user gives (.txt, .md, Word): read as the reader sees them, numbers and all."""

from __future__ import annotations
import json, unittest
from docx import Document
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

import test_xaufe_cli as base  # a module import: its CLI tests are not collected here a second time
from test_xaufe_cli import CONTENT, JOB_B
import xaufe as X


class Originals(base.CLI):
    def run_docx(self, doc, name, edit=None):
        source = self.d / name
        doc.save(source)
        work = self.new(JOB_B)
        code, inspected, err = self.call("inspect", "--work", work, "--docx", source)
        self.assertEqual(code, 0, inspected + err)
        if edit:
            mapping = json.loads((work / "mapping.json").read_text(encoding="utf-8"))
            edit(mapping)
            (work / "mapping.json").write_text(json.dumps(mapping, ensure_ascii=False), encoding="utf-8")
        code, out, err = self.call("run", "--work", work, "--office", "none")
        return work, code, inspected, out + err

    # ------------------------------------------------------------ .txt / .md
    def test_plain_text_keeps_every_character_and_its_line_breaks(self):
        source = self.d / "原稿.txt"
        source.write_text("\n".join([
            "排序算法比较", "", "摘要", "这是合成摘要。", "", "关键词：合成；测试", "", "一、引言",
            "第一行是一个段落", "第二行是另一个段落", "（一）背景", "需要的步骤：", "1. 需求分析", "2. 系统设计",
            "表1 合成表", "编号\t名称\t密码", "1\t甲\t***", "二、结论", "合成结论。", "", "参考文献",
            "[1] 合成作者. 合成文献[M]. 北京: 合成出版社, 2020.",
        ]), encoding="utf-8")
        work = self.new(JOB_B, sources=[source])
        code, out, err = self.call("inspect", "--work", work, "--text", source)
        self.assertEqual(code, 0, out + err)
        blocks = json.loads((work / "content.json").read_text(encoding="utf-8"))["blocks"]
        texts = [b.get("text") or "".join(r["text"] for r in b.get("runs", [])) for b in blocks]
        self.assertIn("第一行是一个段落", texts)  # a short line is its own paragraph, not glued to the next
        self.assertIn("第二行是另一个段落", texts)
        self.assertIn("1. 需求分析", texts)  # numbered lines in a row are a list in the text, not headings
        self.assertEqual([b["level"] for b in blocks if b["type"] == "heading"], [1, 2, 1])
        table = next(b for b in blocks if b["type"] == "table")
        self.assertIn(["1", "甲", "***"], [table["header"]] + table["rows"])  # *** is the user's text
        code, out, err = self.call("run", "--work", work, "--office", "none")
        self.assertEqual(code, 0, out + err)

    def test_markdown_reads_as_shown(self):
        from textio import md_visible, continues
        from text_import import _pieces

        self.assertEqual(md_visible("用 `a*b*c` 和 **加粗** 以及 snake_case_name"), "用 a*b*c 和 加粗 以及 snake_case_name")
        self.assertEqual(md_visible("单独的 * 号和 2*3"), "单独的 * 号和 2*3")
        self.assertEqual(len(_pieces("- 第一项\n- 第二项", markdown=True)), 2)
        long = "这是一行被硬折行的很长的合成正文，它在固定宽度处断开了没有标点"
        self.assertTrue(continues(long, "下一行接着写。", len(long)))
        self.assertFalse(continues("短行", "下一行。", len(long)))
        self.assertFalse(continues(long + "。", "下一行。", len(long) + 1))

    # ------------------------------------------------------------ Word numbering
    def numbered_doc(self, fmt="chineseCounting", text="%1、", headings=("引言", "方法")):
        """Headings numbered by a Word list: the label is not in the paragraph text."""
        doc = Document()
        doc.add_paragraph("排序算法比较")
        doc.add_paragraph("摘要")
        doc.add_paragraph("这是合成摘要。")
        doc.add_paragraph("关键词：编号；标题")
        numbering = doc.part.numbering_part.element
        numbering.insert(0, parse_xml(
            f'<w:abstractNum {nsdecls("w")} w:abstractNumId="90"><w:lvl w:ilvl="0"><w:start w:val="1"/>'
            f'<w:numFmt w:val="{fmt}"/><w:lvlText w:val="{text}"/><w:suff w:val="nothing"/></w:lvl></w:abstractNum>'))
        numbering.append(parse_xml(f'<w:num {nsdecls("w")} w:numId="90"><w:abstractNumId w:val="90"/></w:num>'))
        for title in headings:
            h = doc.add_heading(title, 1)
            h._p.get_or_add_pPr().append(parse_xml(
                f'<w:numPr {nsdecls("w")}><w:ilvl w:val="0"/><w:numId w:val="90"/></w:numPr>'))
            doc.add_paragraph(f"{title}的合成正文。")
        doc.add_paragraph("参考文献")
        doc.add_paragraph("[1] 合成作者. 合成文献[M]. 北京: 合成出版社, 2020.")
        return doc

    def test_list_labels_follow_word_counters(self):
        from numbering import _chinese, _roman, list_labels

        self.assertEqual([_chinese(n) for n in (1, 10, 12, 20, 105, 110)], ["一", "十", "十二", "二十", "一百零五", "一百一十"])
        self.assertEqual(_roman(14), "XIV")
        self.assertEqual(list(list_labels(self.numbered_doc()).values()), ["一、", "二、"])
        doc = self.numbered_doc(fmt="decimal", text="%1.", headings=("引言", "方法", "结论"))
        self.assertEqual(list(list_labels(doc).values()), ["1.", "2.", "3."])

    def test_auto_numbered_headings_keep_their_visible_number_as_text(self):
        work, code, _, out = self.run_docx(self.numbered_doc(), "自动编号.docx")
        self.assertEqual(code, 0, out)
        self.assertIn("自动编号转成了同样的文字", out)
        heads = [p for p in Document(work / "final.docx").paragraphs if p.style.name == "Heading 1"]
        self.assertEqual([p.text for p in heads], ["一、引言", "二、方法"])
        self.assertFalse(any(p._p.xpath("./w:pPr/w:numPr") for p in heads))

    def test_style_numbered_list_items_keep_their_numbers(self):
        doc = Document()
        doc.add_paragraph("合成正文，下面是列表：")
        doc.add_paragraph("第一项", style="List Number")
        doc.add_paragraph("第二项", style="List Number")
        work, code, _, out = self.run_docx(doc, "列表.docx")
        self.assertEqual(code, 0, out)
        items = [p for p in Document(work / "final.docx").paragraphs if p.text in ("第一项", "第二项")]
        self.assertEqual(len(items), 2)
        self.assertTrue(all(p._p.xpath("./w:pPr/w:numPr/w:numId") for p in items))

    # ------------------------------------------------------------ Word structure
    def test_inline_abstract_label_is_kept_and_the_generated_label_follows_the_title(self):
        doc = Document()
        doc.add_paragraph("排序算法比较")
        doc.add_paragraph("摘要：本文是一段合成摘要，说明了研究内容。")
        doc.add_paragraph("关键词：合成；摘要")
        doc.add_heading("一、绪论", 1)
        doc.add_paragraph("合成正文。")
        doc.add_paragraph("参考文献")
        doc.add_paragraph("[1] 合成作者. 合成文献[M]. 北京: 合成出版社, 2020.")
        work, code, inspected, out = self.run_docx(doc, "行内摘要.docx")
        self.assertIn("p1 以“摘要：”开头", inspected)
        self.assertEqual(code, 0, out)
        shown = [p.text for p in Document(work / "final.docx").paragraphs if p.text.strip()]
        self.assertLess(shown.index("排序算法比较"), shown.index("内 容 摘 要"))
        self.assertLess(shown.index("内 容 摘 要"), shown.index("摘要：本文是一段合成摘要，说明了研究内容。"))

    def test_errors_after_a_delete_name_the_ids_in_mapping_json(self):
        doc = Document()
        for text in ("要删掉的合成段落", "合成正文。", "另一段合成正文。"):
            doc.add_paragraph(text)

        def edit(mapping):
            mapping["paragraphs"][0]["role"] = "delete"
            mapping["paragraphs"][2]["role"] = "preserve"
        _, code, _, out = self.run_docx(doc, "删段.docx", edit)
        self.assertNotEqual(code, 0)
        self.assertIn("p2“另一段合成正文。”的 role 是 preserve，但它在 p1", out)
        self.assertIn("旧目录用 old_toc", out)

    def test_table_caption_below_its_table_is_told(self):
        doc = Document()
        doc.add_paragraph("合成正文，下表是合成数据。")
        table = doc.add_table(rows=2, cols=2)
        for row, values in zip(table.rows, (("甲", "乙"), ("1", "2"))):
            for cell, value in zip(row.cells, values):
                cell.text = value
        doc.add_paragraph("表1 合成表")
        doc.add_paragraph("合成正文结尾。")
        _, code, _, out = self.run_docx(doc, "表题在下.docx")
        self.assertEqual(code, 0, out)
        self.assertIn("表题“表1 合成表”在表格下面", out)

    def test_tracked_changes_are_located(self):
        from unsupported import unsupported_reasons

        doc = Document()
        p = doc.add_paragraph("合成正文，")
        ins = OxmlElement("w:ins")
        ins.set(qn("w:id"), "1")
        ins.set(qn("w:author"), "合成")
        run, t = OxmlElement("w:r"), OxmlElement("w:t")
        t.text = "修订加的字"
        run.append(t)
        ins.append(run)
        p._p.append(ins)
        self.assertIn("p0“合成正文，修订加的字”", unsupported_reasons(doc)[0])

    # ------------------------------------------------------------ job.json
    def test_json_error_shows_the_spot_and_says_escape_not_change_quotes(self):
        path = self.d / "job.json"
        path.write_text('{"info": {"title_zh": "关于"AI"的研究"}}', encoding="utf-8")
        with self.assertRaises(ValueError) as caught:
            X._read(path, "job.json")
        message = str(caught.exception)
        self.assertIn("第 1 行第", message)
        self.assertIn('\\"', message)
        self.assertIn("不要换成中文引号", message)

    def test_line_needs_line_rule_and_font_points_to_cn_latin(self):
        work = self.new({**JOB_B, "format": {"styles": {"body": {"line": 360}}}}, CONTENT)
        code, out, err = self.call("run", "--work", work, "--office", "none")
        self.assertNotEqual(code, 0)
        self.assertIn("也要写 line_rule", out + err)
        from customization import locate_format
        with self.assertRaises(ValueError) as caught:
            locate_format({"styles": {"body": {"font": "宋体"}}})
        self.assertIn("cn（中文字体）", str(caught.exception))


# The inherited CLI tests already run in test_xaufe_cli.
for _name in [n for n in vars(base.CLI) if n.startswith("test_")]:
    setattr(Originals, _name, None)

if __name__ == "__main__":
    unittest.main()
