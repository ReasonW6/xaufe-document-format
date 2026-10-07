"""Reference checking against DOI records; the network is replaced by fixed records."""

from __future__ import annotations
import contextlib, io, json, sys, tempfile, unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import refcheck as R
import xaufe as X

DOI = "10.1109/CVPR.2016.90"
CSL_RECORD = {
    "title": "Deep Residual Learning for Image Recognition",
    "author": [{"family": "He", "given": "Kaiming"}, {"family": "Zhang", "given": "Xiangyu"}],
    "issued": {"date-parts": [[2016, 6]]},
    "container-title": "2016 IEEE Conference on Computer Vision and Pattern Recognition (CVPR)",
    "publisher": "IEEE",
}
CROSSREF_RECORD = {"message": {"title": [CSL_RECORD["title"]], "type": "proceedings-article"}}
OPENALEX_RECORD = {"publication_year": 2016, "cited_by_count": 1000, "type": "conference-paper", "is_retracted": False}
GOOD = {
    "text": "HE K, ZHANG X, REN S, et al. Deep residual learning for image recognition[C]. CVPR, 2016: 770-778.",
    "doi": DOI,
    "title": "Deep residual learning for image recognition",
    "authors": ["He K", "Zhang X"],
    "year": 2016,
    "journal": "2016 IEEE Conference on Computer Vision and Pattern Recognition (CVPR)",
}


def network(csl=CSL_RECORD, crossref=CROSSREF_RECORD, openalex=OPENALEX_RECORD, ra="Crossref", search=()):
    def fake(url, accept=None):
        if url.startswith(R.CROSSREF_SEARCH):
            return {"message": {"items": list(search)}}
        if url.startswith(R.RA_URL):
            return [{"RA": ra}] if ra else [{"status": "DOI does not exist"}]
        if url.startswith(R.DOI_URL):
            return csl
        if url.startswith(R.CROSSREF_URL):
            return crossref
        return openalex

    return patch("refcheck.fetch", side_effect=fake)


class CheckEntry(unittest.TestCase):
    def check(self, entry, **records):
        with network(**records):
            return R.check_all([entry], this_year=2026)[0]

    def test_matching_record_passes_with_citations_per_year(self):
        r = self.check(GOOD)
        self.assertEqual(r["status"], R.PASS, r["notes"])
        self.assertEqual(r["per_year"], 100.0)

    def test_unknown_doi_or_prefix_is_dropped(self):
        self.assertEqual(self.check(GOOD, csl=None)["status"], R.DROP)
        self.assertEqual(self.check(GOOD, ra=None)["status"], R.DROP)

    def test_agency_without_metadata_service_needs_a_page_link(self):
        with network(ra="ISTIC") as fetch:
            r = R.check_all([GOOD])[0]
        self.assertEqual(r["status"], R.FIX)
        self.assertIn("url", r["notes"][0])
        self.assertEqual(fetch.call_count, 1)
        with network(ra="ISTIC"), patch("refcheck.fetch_page", return_value="<h1>" + GOOD["title"] + "</h1>"):
            self.assertEqual(R.check_all([{**GOOD, "url": "https://example.org/a"}])[0]["status"], R.PASS)

    def test_doi_of_another_paper_is_dropped_not_fixed(self):
        r = self.check({**GOOD, "title": "A completely different study of sorting"})
        self.assertEqual(r["status"], R.DROP)
        self.assertIn("不要改成查到的那篇", r["notes"][0])

    def test_title_matching_tolerates_shortening_but_not_other_papers(self):
        full = "An Empirical Study on the Energy Usage and Performance of Pandas and Polars Data Analysis Python Libraries"
        self.assertTrue(R._similar("An Empirical Study on Energy Usage and Performance of Pandas and Polars", full))
        self.assertTrue(R._similar("Python for Data Analysis", "Python for Data Analysis: Data Wrangling with Pandas"))
        self.assertFalse(R._similar("Deep learning", "Deep residual learning for image recognition"))
        self.assertFalse(R._similar("A survey on deep learning for image classification",
                                    "A survey on deep learning for image segmentation"))

    def test_wrong_year_author_or_journal_must_be_fixed(self):
        r = self.check({**GOOD, "year": 2015, "authors": ["Zhang X"], "journal": "Nature"})
        self.assertEqual(r["status"], R.FIX)
        joined = " ".join(r["notes"])
        self.assertIn("2016", joined)
        self.assertIn("He", joined)
        self.assertIn("刊名", joined)

    def test_chinese_entry_is_not_rewritten_to_pinyin_or_english(self):
        cn = {**GOOD, "title": "图像分类的深度卷积神经网络模型综述", "authors": ["张珂"], "journal": "中国图象图形学报"}
        csl = {**CSL_RECORD, "title": cn["title"], "author": [{"family": "Ke", "given": "Zhang"}], "container-title": "Journal of Image and Graphics"}
        r = self.check(cn, csl=csl, crossref={"message": {}})
        self.assertEqual(r["status"], R.PASS, r["notes"])
        self.assertIn("自己核对", r["notes"][0])
        english_only = {**csl, "title": "A review of deep CNN models for image classification"}
        r = self.check(cn, csl=english_only, crossref={"message": {}}, openalex={})
        self.assertEqual(r["status"], R.PENDING)

    def test_retraction_from_crossref_or_openalex_drops(self):
        retracted = {"message": {"updated-by": [{"type": "retraction"}]}}
        self.assertEqual(self.check(GOOD, crossref=retracted)["status"], R.DROP)
        self.assertEqual(self.check(GOOD, openalex={**OPENALEX_RECORD, "is_retracted": True})["status"], R.DROP)
        self.assertEqual(self.check(GOOD, csl={**CSL_RECORD, "title": "RETRACTED: " + CSL_RECORD["title"]})["status"], R.DROP)

    def test_preprint_is_flagged(self):
        r = self.check({**GOOD, "doi": "10.48550/arXiv.1706.03762"})
        self.assertIn("预印本", " ".join(r["notes"]))

    def test_entries_without_doi_are_checked_on_their_page(self):
        book = {"text": "某作者. 某书[M]. 北京: 某出版社, 2020.", "title": "数据分析方法", "url": "https://example.org/book"}
        with patch("refcheck.fetch") as fetch, patch("refcheck.fetch_page") as page:
            self.assertEqual(R.check_all([{"text": book["text"]}])[0]["status"], R.DROP)
            self.assertEqual(R.check_all([{"text": "用户给的文献", "from_user": True}])[0]["status"], R.USER)
            self.assertEqual(R.check_all([{**book, "title": ""}])[0]["status"], R.FIX)
            page.return_value = "数据 分析 方法 - 出版社"
            self.assertEqual(R.check_all([book])[0]["status"], R.PASS)
            page.return_value = "某出版社首页"
            r = R.check_all([book])[0]
            self.assertEqual(r["status"], R.DROP)
            self.assertIn("不要编链接", r["notes"][0])
            page.side_effect = R.Offline("HTTP 403")
            fetch.return_value = [{"RA": "Crossref"}]
            self.assertEqual(R.check_all([book])[0]["status"], R.DROP)
            fetch.side_effect = R.Offline("no route", down=True)
            self.assertEqual(R.check_all([book])[0]["status"], R.PENDING)

    def test_entry_without_doi_is_checked_against_the_doi_found_by_title_and_author(self):
        entry = {k: v for k, v in GOOD.items() if k != "doi"}
        entry.update(url="https://example.org/paper", journal="Nature")
        hit = {"DOI": DOI, "title": [CSL_RECORD["title"]], "author": [{"family": "He"}]}
        r = self.check(entry, search=[hit])
        self.assertEqual(r["status"], R.FIX)
        self.assertIn(DOI, r["notes"][0])
        self.assertIn("刊名", " ".join(r["notes"]))
        other_author = {**hit, "author": [{"family": "Smith"}]}
        with network(search=[other_author]), patch("refcheck.fetch_page", return_value=CSL_RECORD["title"]):
            r = R.check_all([entry])[0]
        self.assertEqual(r["status"], R.PASS)
        self.assertEqual(r["notes"], ["链接页面上有这个题目"])

    def test_books_are_checked_by_isbn(self):
        book = {"text": "McKinney W. Python for Data Analysis[M]. 2017.", "isbn": "978-1-4919-5766-0",
                "title": "Python for data analysis", "year": 2017}
        record = {"title": "Python for Data Analysis: Data Wrangling with Pandas, NumPy, and IPython", "publish_date": "Oct 20, 2017"}
        with patch("refcheck.fetch", return_value=record):
            self.assertEqual(R.check_all([book])[0]["status"], R.PASS)
            self.assertEqual(R.check_all([{**book, "year": 2012}])[0]["status"], R.FIX)
            cn = {**book, "title": "数据结构(C语言版)"}
            self.assertEqual(R.check_all([cn])[0]["status"], R.PENDING)
        with patch("refcheck.fetch", return_value=None):
            self.assertEqual(R.check_all([book])[0]["status"], R.DROP)
        with patch("refcheck.fetch") as fetch:
            r = R.check_all([{**book, "isbn": "978-1-4919-5766-1"}])[0]
        fetch.assert_not_called()
        self.assertEqual(r["status"], R.DROP)
        self.assertIn("校验位", r["notes"][0])

    def test_registry_without_metadata_needs_manual_check(self):
        def fake(url, accept=None):
            if url.startswith(R.RA_URL):
                return [{"RA": "Crossref"}]
            raise R.Unsupported(url)

        with patch("refcheck.fetch", side_effect=fake):
            self.assertEqual(R.check_all([GOOD])[0]["status"], R.FIX)

    def test_no_network_marks_pending_and_stops_retrying(self):
        other = {**GOOD, "doi": "10.1000/other", "title": "Other"}
        with patch("refcheck.fetch", side_effect=R.Offline("no route", down=True)) as fetch:
            results = R.check_all([GOOD, other])
        self.assertEqual([r["status"] for r in results], [R.PENDING, R.PENDING])
        self.assertEqual(fetch.call_count, 1)
        with patch("refcheck.fetch", side_effect=R.Offline("timed out")) as fetch:
            R.check_all([GOOD, other])
        self.assertEqual(fetch.call_count, 2)

    def test_duplicates_are_dropped(self):
        with network():
            results = R.check_all([GOOD, {**GOOD, "doi": "https://doi.org/" + DOI.lower()}], this_year=2026)
        self.assertEqual(results[1]["status"], R.DROP)
        self.assertIn("重复", results[1]["notes"][0])


class Report(unittest.TestCase):
    def test_counts_minimum_and_recent_share(self):
        entries = [
            {"text": "a", "from_user": True, "year": 2024},
            {"text": "b", "from_user": True, "year": 2010},
            {"text": "c", "from_user": True, "year": 1990, "classic": True},
        ]
        text = R.report(R.check_all(entries, this_year=2026), entries, 6, this_year=2026)
        self.assertIn("可用 3 条", text)
        self.assertIn("1/2 条", text)
        self.assertIn("还差 3 条，不要凑数", text)


class Command(unittest.TestCase):
    def test_refs_command_reads_work_refs_json(self):
        import delivery as D

        with tempfile.TemporaryDirectory() as tmp:
            with patch("tempfile.gettempdir", return_value=tmp):
                job = D.create(Path(tmp) / "out" / "成品.docx")
            work = Path(job["workspace"])
            (work / "refs.json").write_text(json.dumps([GOOD], ensure_ascii=False), encoding="utf-8")
            out = io.StringIO()
            with network(), contextlib.redirect_stdout(out):
                code = X.main(["refs", "--work", str(work), "--min", "1"])
            self.assertEqual(code, 0)
            self.assertIn("第 1 条：通过", out.getvalue())
            self.assertIn("可用 1 条", out.getvalue())


if __name__ == "__main__":
    unittest.main()
