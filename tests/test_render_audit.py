"""Rendered-page audit on synthetic PDFs (no office suite needed) and, when
LibreOffice is installed, on a really rendered defective document."""

from pathlib import Path
import json, shutil, sys, tempfile, unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import pymupdf as fitz
import engine as E
import intake as I
import feedback as F
from render_audit import audit

A4 = (595.3, 841.9)


def value(text):
    return {"value": text, "source": "user", "evidence": "SYNTHETIC test statement"}


class RenderAudit(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.folder = Path(self.temp.name)

    def tearDown(self):
        self.temp.cleanup()

    def docx(self, template="B", patch=None, cover=True):
        folder = Path(tempfile.mkdtemp(dir=self.folder))
        state = folder / "intake.json"
        I.start(state, "SYNTHETIC render audit")
        I.select(state, template, "SYNTHETIC choice")
        if patch:
            I.customize(state, patch, "SYNTHETIC customization")
        if not cover:
            I.document_scope(state, False, "SYNTHETIC no cover")
        fields = I.profile_for(I.read(state))["cover"]["required_fields"] if cover else []
        meta = {k: value("2030-01-02" if k == "date" else "测试") for k in fields}
        if not cover and template == "A":
            meta = {"title_zh": value("测试"), "major": value("测试")}
        I.resolve(state, {"metadata": meta})
        content = folder / "content.json"
        content.write_text(json.dumps({"mode": "document", "metadata": {}, "blocks": [
            {"type": "heading", "level": 1, "text": "一、合成"},
            {"type": "paragraph", "text": "合成正文。"}]}, ensure_ascii=False), encoding="utf-8")
        out = folder / "doc.docx"
        E.build(content, out, intake_path=state)
        return out

    def pdf(self, pages):
        """pages: list of lists of (kind, x, y, payload)."""
        path = self.folder / "doc.pdf"
        doc = fitz.open()
        for items in pages:
            page = doc.new_page(width=A4[0], height=A4[1])
            for kind, x, y, payload in items:
                if kind == "text":
                    page.insert_text((x, y), payload, fontname="china-s", fontsize=10.5)
                elif kind == "image":
                    pix = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 4, 4), 0)
                    page.insert_image(fitz.Rect(x, y, x + payload, y + 40), pixmap=pix)
                elif kind == "line":
                    page.draw_line((x, y), (x + payload, y))
        doc.save(path)
        return path

    def codes(self, result, key="errors"):
        return [e["code"] for e in result[key]]

    # B: left 3.00 cm (85 pt), right 2.20 cm -> text column 85..533 pt, no header, no page numbers.
    def body(self, text="合成正文。"):
        return [("text", 90, 200, text)]

    def cover(self):
        return [("text", 120, 300, "封面")]

    def test_clean_pages_pass(self):
        result = audit(self.docx(), self.pdf([self.cover(), self.body()]))
        self.assertTrue(result["passed"], result["errors"])

    def test_text_beyond_right_margin(self):
        result = audit(self.docx(), self.pdf([self.cover(), [("text", 480, 200, "伸出版心的文字文字文字")]]))
        self.assertIn("TEXT_OUTSIDE_MARGINS", self.codes(result))
        issue = result["errors"][0]
        self.assertEqual(issue["page"], 2)
        self.assertIn("actual", issue)

    def test_one_trailing_punctuation_may_hang(self):
        # The last glyph is a comma just past the column: 行末标点悬挂 is not a defect.
        x = 533 - 10.5 * 3
        result = audit(self.docx(), self.pdf([self.cover(), [("text", x, 200, "文字字，")]]))
        self.assertTrue(result["passed"], result["errors"])

    def test_blank_page_is_an_error_with_context(self):
        result = audit(self.docx(), self.pdf([self.cover(), self.body("前一页最后一句。"), []]))
        blank = [e for e in result["errors"] if e["code"] == "BLANK_PAGE"]
        self.assertEqual([e["page"] for e in blank], [3])
        self.assertIn("前一页最后一句", blank[0]["where"])

    def test_page_with_only_a_page_number_is_blank(self):
        doc = self.docx(patch={"footer_page_numbers": True})
        result = audit(doc, self.pdf([self.cover(), self.body() + [("text", 295, 830, "1")], [("text", 295, 830, "2")]]))
        self.assertIn("BLANK_PAGE", self.codes(result))

    def test_unexpected_header_on_b(self):
        result = audit(self.docx(), self.pdf([self.cover(), self.body() + [("text", 90, 50, "多余页眉")]]))
        self.assertIn("HEADER_UNEXPECTED", self.codes(result))

    def test_missing_header_on_a(self):
        result = audit(self.docx("A"), self.pdf([self.cover(), [("text", 95, 200, "正文")]]))
        self.assertIn("HEADER_MISSING", self.codes(result))

    def test_cover_must_not_show_footer(self):
        result = audit(self.docx(), self.pdf([self.cover() + [("text", 295, 830, "1")], self.body()]))
        self.assertIn("COVER_FOOTER_VISIBLE", self.codes(result))

    def test_requested_page_numbers_must_render(self):
        doc = self.docx(patch={"footer_page_numbers": True})
        result = audit(doc, self.pdf([self.cover(), self.body()]))
        self.assertIn("PAGE_NUMBER_MISSING", self.codes(result))
        result = audit(doc, self.pdf([self.cover(), self.body() + [("text", 295, 830, "1")]]))
        self.assertTrue(result["passed"], result["errors"])

    def test_footer_text_when_page_numbers_are_off(self):
        result = audit(self.docx(), self.pdf([self.cover(), self.body() + [("text", 295, 830, "1")]]))
        self.assertIn("FOOTER_UNEXPECTED", self.codes(result))

    def test_image_and_table_rules_beyond_the_column(self):
        result = audit(self.docx(), self.pdf([self.cover(), self.body() + [("image", 300, 300, 260), ("line", 85, 500, 470)]]))
        self.assertIn("IMAGE_OUTSIDE_MARGINS", self.codes(result))
        self.assertIn("TABLE_OUTSIDE_MARGINS", self.codes(result))

    def test_off_page_text(self):
        result = audit(self.docx(), self.pdf([self.cover(), self.body() + [("text", 560, 300, "出界出界出界出界")]]))
        self.assertIn("OFF_PAGE", self.codes(result))

    def test_overlapping_lines_are_a_warning(self):
        result = audit(self.docx(), self.pdf([self.cover(), [("text", 90, 200, "第一行文字"), ("text", 92, 202, "第二行文字")]]))
        self.assertTrue(result["passed"])
        self.assertIn("TEXT_OVERLAP", self.codes(result, "warnings"))

    def test_cover_may_use_its_own_measure_but_not_the_page_edge(self):
        ok = audit(self.docx(), self.pdf([[("text", 90, 300, "《很长的课程名称在封面上居中》")], self.body()]))
        self.assertTrue(ok["passed"], ok["errors"])
        bad = audit(self.docx(), self.pdf([[("text", 10, 300, "贴边")], self.body()]))
        self.assertIn("TEXT_OUTSIDE_MARGINS", self.codes(bad))

    def test_font_substitution_names_the_missing_families(self):
        result = audit(self.docx(), self.pdf([self.cover(), self.body()]))
        warning = [w for w in result["warnings"] if w["code"] == "FONT_SUBSTITUTED"]
        self.assertTrue(warning)
        self.assertIn("宋体", result["fonts"]["missing"])

    def test_issues_carry_location_cause_and_fix(self):
        result = audit(self.docx(), self.pdf([self.cover(), self.body(), []]))
        issues = F.from_render(result)
        self.assertEqual(issues[0]["where"]["page"], 3)
        self.assertTrue(issues[0]["cause"] and issues[0]["fix"])
        text = F.render_text(issues)
        for word in ("位置", "可能原因", "怎么改"):
            self.assertIn(word, text)


@unittest.skipUnless(shutil.which("soffice") or shutil.which("libreoffice"), "needs LibreOffice")
class RenderAuditRealOffice(unittest.TestCase):
    """A real export of a defective document must be caught from the PDF alone."""

    def test_blank_page_from_double_page_break(self):
        from office_backend import export_pdf
        from docx import Document
        from docx.enum.text import WD_BREAK

        with tempfile.TemporaryDirectory() as folder:
            folder = Path(folder)
            case = RenderAudit("test_clean_pages_pass")
            case.folder = folder
            out = case.docx(cover=False)
            doc = Document(out)
            for _ in range(2):
                empty = doc.add_paragraph("", style=E.ROLES["body"])
                empty.add_run().add_break(WD_BREAK.PAGE)
            doc.add_paragraph("尾页", style=E.ROLES["body"])
            doc.save(out)
            pdf = folder / "real.pdf"
            export_pdf(out, pdf)
            result = audit(out, pdf)
            self.assertIn("BLANK_PAGE", [e["code"] for e in result["errors"]])


if __name__ == "__main__":
    unittest.main()
