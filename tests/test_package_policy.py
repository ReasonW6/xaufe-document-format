"""Package structure, resource integrity and documentation limits."""

from pathlib import Path
from zipfile import ZipFile
import hashlib, json, re, sys, tempfile, unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import delivery as D


class PackagePolicy(unittest.TestCase):
    def test_skill_name(self):
        self.assertIn("name: xaufe-document-format", (ROOT / "SKILL.md").read_text(encoding="utf-8"))

    def test_current_code_has_no_obsolete_prefix(self):
        old = "x" + "ufe"
        for directory in ("scripts", "references", "assets", "tests"):
            for p in (ROOT / directory).rglob("*"):
                if p.suffix in (".py", ".md", ".json"):
                    self.assertNotIn(old, p.read_text(encoding="utf-8").lower(), str(p))

    def test_no_vendor_or_history_specific_text(self):
        banned = ("codex", "require_escalated", "exec_command", "gpt-", "luna", "green-paper", "red-course", "logo_template")
        for p in ROOT.rglob("*"):
            rel = p.relative_to(ROOT).parts
            if p.suffix in (".py", ".md", ".json", ".txt") and rel[0] not in ("tests", ".git") and "__pycache__" not in rel:
                text = p.read_text(encoding="utf-8").lower()
                for word in banned:
                    self.assertNotIn(word, text, f"{p}: {word}")

    def test_master_xml_uses_native_names(self):
        token = ("X" + "UFE").encode()
        for template in ("A", "B"):
            with ZipFile(ROOT / f"assets/{template}/format-master.docx") as z:
                for name in z.namelist():
                    if name.endswith((".xml", ".rels")):
                        self.assertNotIn(token, z.read(name), name)

    def test_original_sources_masters_and_logos_have_declared_hashes(self):
        registry = json.loads((ROOT / "assets/templates.json").read_text(encoding="utf-8"))
        items = list(registry["logos"].values())
        for template in ("A", "B"):
            spec = json.loads((ROOT / f"assets/{template}/layout.json").read_text(encoding="utf-8"))
            items += spec["source_files"]
            items.append({"file": spec["master_file"], "sha256": spec["master_sha256"]})
        for item in items:
            self.assertEqual(
                hashlib.sha256((ROOT / item["file"]).read_bytes()).hexdigest(), item["sha256"], item["file"]
            )

    def test_layout_and_logo_are_independent_entries(self):
        registry = json.loads((ROOT / "assets/templates.json").read_text(encoding="utf-8"))
        self.assertEqual(set(registry["templates"]), {"A", "B"})
        self.assertEqual(set(registry["logos"]), {"red", "green"})
        for template, color in (("A", "green"), ("B", "red")):
            spec = json.loads((ROOT / f"assets/{template}/layout.json").read_text(encoding="utf-8"))
            self.assertEqual(spec["id"], template)
            self.assertEqual(spec["logo"]["default"], color)
            self.assertNotIn("file", spec["logo"])

    def test_no_preapproved_tasks_or_output_reports(self):
        for p in ROOT.rglob("*"):
            rel = p.relative_to(ROOT)
            if "__pycache__" in rel.parts or ".git" in rel.parts:
                continue
            self.assertNotIn(p.suffix.lower(), (".log", ".pdf", ".pyc", ".zip"))
            self.assertNotIn(p.name, ("intake.json", "final.docx", "review.json", ".xaufe-session.json"))

    def test_examples_contain_only_minimal_inputs(self):
        expected = {
            "document.json",
            "layout-test.png",
            "A/job.json",
            "A/content.json",
            "B/job.json",
            "B/content.json",
        }
        actual = {
            p.relative_to(ROOT / "examples").as_posix()
            for p in (ROOT / "examples").rglob("*")
            if p.is_file()
        }
        self.assertEqual(actual, expected)

    def test_markdown_links_resolve(self):
        for p in ROOT.rglob("*.md"):
            for target in re.findall(r"\[[^\]]*\]\(([^)]+)\)", p.read_text(encoding="utf-8")):
                if "://" not in target and not target.startswith("#"):
                    self.assertTrue((p.parent / target.split("#")[0]).exists(), (p, target))

    def test_docs_stay_short_enough_for_small_models(self):
        self.assertLess(len((ROOT / "SKILL.md").read_bytes()), 12000)
        refs = list((ROOT / "references").glob("*.md"))
        self.assertLessEqual(len(refs), 6)
        self.assertLess(sum(len(p.read_bytes()) for p in refs), 40000)

    def test_examples_are_valid_inputs(self):
        import intake as I

        with tempfile.TemporaryDirectory() as folder:
            for template in ("A", "B"):
                job = json.loads((ROOT / f"examples/{template}/job.json").read_text(encoding="utf-8"))
                state = I.from_job(job, Path(folder) / f"{template}.json")
                self.assertEqual(state["stage"], "ready")
                self.assertEqual(state["template"], template)

    def test_unowned_workspace_is_rejected(self):
        with tempfile.TemporaryDirectory() as folder:
            task = D.create(Path(folder) / "out.docx", parent=folder)
            w = Path(task["workspace"])
            record = D.read_json(w / D.MARKER)
            record["magic"] = "someone-else"
            D.write_json(w / D.MARKER, record)
            with self.assertRaises(ValueError):
                D.load(w)


if __name__ == "__main__":
    unittest.main()
