from pathlib import Path
import json, sys, tempfile, unittest
from docx import Document

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import source_compare as S, delivery as D


class SourceProseTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def test_missing_original_paragraph_is_reported(self):
        source = self.root / "raw.md"
        source.write_text("# 标题\n\n保留的第一段。\n\n转换中不能漏掉的第二段。", encoding="utf-8")
        out = self.root / "output.docx"
        doc = Document()
        doc.add_paragraph("保留的第一段。")
        doc.save(out)
        report = S.compare([{"path": str(source)}], out)
        self.assertEqual(
            report["sources"][0]["missing"],
            [{"location": "line 5", "text": "转换中不能漏掉的第二段。"}],
        )

    def test_known_metadata_labels_do_not_create_false_omissions(self):
        source = self.root / "raw.md"
        source.write_text(
            "English title: Room Reservation\n\nKey words: room reservation; data records\n\n保留**重点**和[链接名](https://example.com)。",
            encoding="utf-8",
        )
        out = self.root / "out.docx"
        doc = Document()
        doc.add_paragraph("Room Reservation")
        doc.add_paragraph("Key words：room reservation   data records")
        doc.add_paragraph("保留重点和链接名。")
        doc.save(out)
        self.assertEqual(S.compare([{"path": str(source)}], out)["sources"][0]["missing"], [])

    def test_create_registers_actual_original_file(self):
        source = self.root / "original.txt"
        source.write_text("原始材料", encoding="utf-8")
        job = D.create(
            self.root / "out.docx",
            parent=self.root,
            sources=[source],
        )
        work, state = D.load(job["workspace"])
        self.assertEqual(state["source_documents"][0]["sha256"], D.sha(source))
        self.assertEqual(state["source_documents"][0]["role"], "original")
        D.clean(work)
        self.assertEqual(source.read_text(encoding="utf-8"), "原始材料")

    def test_missing_source_stops_before_workspace_creation(self):
        with self.assertRaises(ValueError):
            D.create(
                self.root / "out.docx",
                parent=self.root,
                sources=[self.root / "missing.md"],
            )
        self.assertEqual(list(self.root.iterdir()), [])

    def test_explicit_workspace_parent_is_respected(self):
        job = D.create(
            self.root / "out.docx",
            parent=self.root,
        )
        work, state = D.load(job["workspace"])
        self.assertEqual(work.parent, self.root.resolve())
        self.assertTrue(work.name.startswith("xaufe-job-"))
        D.clean(work)
        self.assertFalse(work.exists())

    def test_source_gap_becomes_review_warning(self):
        source = self.root / "raw.txt"
        source.write_text("需要保留的原始段落。", encoding="utf-8")
        doc = Document()
        doc.add_paragraph("只有其他文字。")
        doc.save(self.root / "final.docx")
        result = {"warnings": []}
        D.compare_originals(
            self.root, {"source_documents": [{"path": str(source), "role": "original"}]}, result
        )
        self.assertEqual(result["warnings"][0]["code"], "SOURCE_DIFFERENCES")
        self.assertTrue((self.root / "source-comparison.json").is_file())


if __name__ == "__main__":
    unittest.main()
