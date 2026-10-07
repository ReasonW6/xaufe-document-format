"""Offline regression tests. No network; no external renderer needed for these unit tests."""

from __future__ import annotations
import hashlib, json, sys, tempfile, unittest
from pathlib import Path
from copy import deepcopy
from unittest.mock import patch
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Cm

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import engine
import intake
from check_format import check


class FormatTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.d = Path(self.tmp.name)
        engine.select_profile("A")
        self.data = {
            "mode": "document",
            "cover": True,
            "metadata": {"title_zh": "格式回归测试"},
            "blocks": [
                {"type": "heading", "level": 1, "text": "一、测试标题"},
                {"type": "paragraph", "text": "保留原文和数字123，不写新内容。"},
                {"type": "heading", "level": 2, "text": "（一）二级标题"},
                {
                    "type": "paragraph",
                    "runs": [
                        {"text": "引用"},
                        {"text": "[1]", "superscript": True},
                        {"text": "及公式x²。"},
                    ],
                },
            ],
        }

    def tearDown(self):
        self.tmp.cleanup()

    def session_for(self, meta, template="A", name="intake.json"):
        path = self.d / name
        intake.start(path, "独立回归测试夹具")
        intake.select(path, template, "测试夹具：明确选择 " + template)
        spec = engine.load_profile(template)
        fields = spec["cover"]["required_fields"]
        known = {
            k: {"value": v, "source": "document", "evidence": "测试样本文档明确字段"}
            for k, v in meta.items()
            if k in fields and v
        }
        blanks = [k for k in fields if k not in known and k != "date"]
        intake.resolve(
            path,
            {
                "metadata": known,
                "blank_fields": blanks,
                "blank_instruction": "测试夹具：明确选择未提供项留空。",
            },
        )
        return path

    def make(self, data=None, name="out.docx"):
        src = self.d / (name + ".json")
        src.write_text(json.dumps(data or self.data, ensure_ascii=False), encoding="utf-8")
        out = self.d / name
        self.session = self.session_for(
            (data or self.data).get("metadata", {}), name=name + ".intake.json"
        )
        engine.build(src, out, intake_path=self.session)
        return out

    def changed(self, fn):
        p = self.make()
        d = Document(p)
        fn(d)
        d.save(p)
        return check(p)

    def assertCode(self, result, code):
        self.assertIn(code, {e["code"] for e in result["errors"]})

    def test_build_passes_static(self):
        self.assertTrue(check(self.make())["passed"])

    def test_without_cover(self):
        data = deepcopy(self.data)
        data["cover"] = False
        p = self.make(data)
        res = check(p)
        self.assertTrue(res["passed"])
        self.assertFalse(res["stats"]["cover"])
        self.assertEqual(len(Document(p).inline_shapes), 0)

    def test_original_logo_hash(self):
        engine.asset_check()
        self.assertEqual(
            hashlib.sha256((ROOT / "assets/logos/green.png").read_bytes()).hexdigest(),
            engine.SPEC["logo"]["sha256"],
        )

    def test_footer_default_rejected(self):
        self.assertCode(
            self.changed(lambda d: d.sections[-1].footer.paragraphs[0].add_run("1")), "FOOTER"
        )

    def test_footer_first_rejected(self):
        self.assertCode(
            self.changed(lambda d: d.sections[-1].first_page_footer.paragraphs[0].add_run("1")),
            "FOOTER",
        )

    def test_footer_even_rejected(self):
        self.assertCode(
            self.changed(lambda d: d.sections[-1].even_page_footer.paragraphs[0].add_run("1")),
            "FOOTER",
        )

    def test_footer_page_field_rejected(self):
        def mutate(d):
            fld = OxmlElement("w:fldSimple")
            fld.set(qn("w:instr"), " PAGE ")
            d.sections[-1].footer.paragraphs[0]._p.append(fld)

        r = self.changed(mutate)
        self.assertCode(r, "FOOTER")
        self.assertCode(r, "FOOTER_FIELD")

    def test_wrong_body_size_rejected(self):
        def mutate(d):
            next(p for p in d.paragraphs if p.style.name == "XAUFE Body").runs[0].font.size = Pt(12)

        self.assertCode(self.changed(mutate), "FONT_SIZE")

    def test_wrong_body_chinese_font_rejected(self):
        def mutate(d):
            r = next(p for p in d.paragraphs if p.style.name == "XAUFE Body").runs[0]
            engine.element(r._r.get_or_add_rPr(), "w:rFonts", eastAsia="Arial")

        self.assertCode(self.changed(mutate), "FONT_CN")

    def test_center_logo_rejected(self):
        self.assertCode(
            self.changed(lambda d: setattr(d.paragraphs[0], "alignment", 1)), "LOGO_ALIGN"
        )

    def test_enlarged_logo_rejected(self):
        self.assertCode(
            self.changed(lambda d: setattr(d.inline_shapes[0], "width", Cm(14))), "LOGO_SIZE"
        )

    def test_wrong_margin_rejected(self):
        self.assertCode(
            self.changed(lambda d: setattr(d.sections[-1], "left_margin", Cm(2.54))), "PAGE"
        )

    def test_heading_four_stops(self):
        data = deepcopy(self.data)
        data["blocks"][0]["level"] = 4
        with self.assertRaises(ValueError):
            self.make(data)
        self.assertFalse((self.d / "out.docx").exists())

    def test_missing_image_stops(self):
        data = deepcopy(self.data)
        data["blocks"].append({"type": "image", "path": "absent.png"})
        with self.assertRaises(ValueError):
            self.make(data)
        self.assertFalse((self.d / "out.docx").exists())

    def test_refuses_in_place(self):
        p = self.make()
        with self.assertRaises(ValueError):
            engine.safe_out(p, p, True)

    def test_paper_toc_requires_refresh(self):
        data = json.loads((ROOT / "examples/A/content.json").read_text())
        data["blocks"] = [b for b in data["blocks"] if b["type"] != "image"]
        p = self.make(data)
        self.assertTrue(check(p)["passed"])
        self.assertCode(check(p, final=True), "TOC_UNRESOLVED")

    def test_master_contains_no_filled_personal_fields(self):
        d = Document(ROOT / "assets/A/format-master.docx")
        for i in engine.SPEC["cover"]["fields"].values():
            self.assertEqual(d.paragraphs[i].runs[1].text.strip(), "")
        self.assertEqual(d.core_properties.author, "")

    def source_map(self, complex_math=False):
        d = Document()
        d.add_paragraph("一、已有标题")
        p = d.add_paragraph("原文保留，含链接")
        h = OxmlElement("w:hyperlink")
        h.set(qn("w:anchor"), "test")
        r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = "链接文字"
        r.append(t)
        h.append(r)
        p._p.append(h)
        d.add_paragraph("关键词：原词甲   原词乙")
        t = d.add_table(rows=2, cols=2)
        t.cell(0, 0).text = "列甲"
        t.cell(0, 1).text = "列乙"
        t.cell(1, 0).text = "值甲"
        t.cell(1, 1).text = "值乙"
        p = d.add_paragraph()
        p.add_run().add_picture(str(ROOT / "examples/layout-test.png"), width=Cm(5))
        p = d.add_paragraph()
        m = OxmlElement("m:oMath")
        r = OxmlElement("m:r")
        t = OxmlElement("m:t")
        t.text = "x+1"
        r.append(t)
        m.append(r)
        p._p.append(m)
        src = self.d / "source.docx"
        d.save(src)
        inv = engine.inventory(src)
        mapping = {
            "input_sha256": inv["input_sha256"],
            "mode": "document",
            "has_cover": False,
            "metadata": {"title_zh": "保真测试"},
            "format_tables": True,
            "roles": {
                "p0": "h1",
                "p1": "body",
                "p2": "keywords_zh",
                "p3": "figure",
                "p4": "equation",
            },
        }
        mp = self.d / "map.json"
        mp.write_text(json.dumps(mapping, ensure_ascii=False), encoding="utf-8")
        self.session = self.session_for(mapping["metadata"])
        return src, mp

    def test_restyle_preserves_text_math_tables_images_fields(self):
        src, mp = self.source_map()
        before = engine.content_signature(Document(src))
        out = self.d / "restyled.docx"
        engine.restyle(src, mp, out, intake_path=self.session)
        self.assertEqual(before, engine.content_signature(Document(out)))
        self.assertTrue(check(out)["passed"])

    def test_stale_map_stops(self):
        src, mp = self.source_map()
        d = Document(src)
        d.add_paragraph("改过")
        d.save(src)
        with self.assertRaises(ValueError):
            engine.restyle(src, mp, self.d / "result.docx", intake_path=self.session)

    def test_unmapped_paragraph_stops(self):
        src, mp = self.source_map()
        m = json.loads(mp.read_text())
        del m["roles"]["p1"]
        mp.write_text(json.dumps(m))
        with self.assertRaises(ValueError):
            engine.restyle(src, mp, self.d / "result.docx", intake_path=self.session)

    def test_tracked_change_stops(self):
        src, mp = self.source_map()
        d = Document(src)
        ins = OxmlElement("w:ins")
        d.paragraphs[1]._p.append(ins)
        d.save(src)
        m = json.loads(mp.read_text())
        m["input_sha256"] = engine.inventory(src)["input_sha256"]
        mp.write_text(json.dumps(m))
        with self.assertRaises(ValueError):
            engine.restyle(src, mp, self.d / "result.docx", intake_path=self.session)

    def test_top_level_content_control_stops(self):
        src, mp = self.source_map()
        d = Document(src)
        d.element.body.insert(0, OxmlElement("w:sdt"))
        d.save(src)
        m = json.loads(mp.read_text())
        m["input_sha256"] = engine.inventory(src)["input_sha256"]
        mp.write_text(json.dumps(m))
        with self.assertRaises(ValueError):
            engine.restyle(src, mp, self.d / "result.docx", intake_path=self.session)

    def test_figure_not_exact_line_height(self):
        d = Document()
        engine.install_styles(d)
        e = d.styles["XAUFE Figure"].element.pPr.find(qn("w:spacing"))
        self.assertEqual(e.get(qn("w:lineRule")), "auto")
        self.assertGreaterEqual(int(e.get(qn("w:before"))), 120)

    def test_wrong_header_size_rejected(self):
        def mutate(d):
            d.sections[-1].header.paragraphs[0].runs[0].font.size = Pt(14)

        self.assertCode(self.changed(mutate), "FONT_SIZE")

    def test_wrong_cover_field_size_rejected(self):
        data = deepcopy(self.data)
        data["metadata"]["student_name"] = "测试姓名"
        p = self.make(data)
        d = Document(p)
        d.paragraphs[9].runs[1].font.size = Pt(12)
        d.save(p)
        self.assertCode(check(p), "COVER_FONT")

    def test_tampered_master_stops(self):
        (self.d / "assets/A").mkdir(parents=True)
        (self.d / "assets/logos").mkdir(parents=True)
        (self.d / "assets/logos/green.png").write_bytes(
            (ROOT / "assets/logos/green.png").read_bytes()
        )
        (self.d / "assets/A/format-master.docx").write_bytes(b"tampered")
        with patch.object(engine, "ROOT", self.d):
            with self.assertRaises(ValueError):
                engine.asset_check()

    def test_missing_original_logo_stops(self):
        with patch.object(engine, "ROOT", self.d):
            with self.assertRaises(ValueError):
                engine.asset_check()

    def test_add_cover_to_existing_keeps_original_text(self):
        src, mp = self.source_map()
        m = json.loads(mp.read_text())
        m["replace_cover_before"] = 0
        mp.write_text(json.dumps(m, ensure_ascii=False))
        before = engine.content_signature(Document(src))
        out = self.d / "covered.docx"
        engine.restyle(src, mp, out, intake_path=self.session)
        d = Document(out)
        self.assertTrue(engine.content_signature(d)["text"].endswith(before["text"]))
        self.assertTrue(check(out)["passed"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
