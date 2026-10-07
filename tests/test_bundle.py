"""Package integrity tests in isolated directories; no real skill file is mutated."""

from __future__ import annotations
import hashlib
from pathlib import Path
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
from verify_bundle import verify


class BundleIntegrity(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.source = self.root / "SKILL.md"
        self.source.write_text("synthetic fixture", encoding="utf-8")
        self.line = hashlib.sha256(self.source.read_bytes()).hexdigest() + "  SKILL.md\n"
        self.manifest = self.root / "SHA256SUMS"
        self.manifest.write_text(self.line, encoding="utf-8")

    def tearDown(self):
        self.tmp.cleanup()

    def test_valid_manifest(self):
        self.assertTrue(verify(self.root)["passed"])

    def test_missing_file(self):
        self.source.unlink()
        self.assertFalse(verify(self.root)["passed"])

    def test_changed_file(self):
        self.source.write_text("changed")
        self.assertFalse(verify(self.root)["passed"])

    def test_unlisted_old_script(self):
        (self.root / "old_engine.py").write_text("# old")
        self.assertFalse(verify(self.root)["passed"])

    def test_empty_manifest(self):
        self.manifest.write_text("")
        self.assertFalse(verify(self.root)["passed"])

    def test_duplicate_entry(self):
        self.manifest.write_text(self.line * 2)
        self.assertFalse(verify(self.root)["passed"])

    def test_path_escape(self):
        self.manifest.write_text("a" * 64 + "  ../outside\n")
        self.assertFalse(verify(self.root)["passed"])

    def test_runtime_cache_is_ignored(self):
        d = self.root / "__pycache__"
        d.mkdir()
        (d / "cache.pyc").write_bytes(b"cache")
        self.assertTrue(verify(self.root)["passed"])


if __name__ == "__main__":
    unittest.main()
