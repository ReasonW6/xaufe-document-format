"""Check reference entries against public DOI records (doi.org, Crossref, OpenAlex).

The agent writes WORK/refs.json; ``xaufe.py refs`` prints one verdict per entry
plus the totals that references/citations.md asks for.

Two questions are answered separately, so "the paper exists" is never read as
"the citation is right":

* ``status``  - does the work exist and do author/title/year/venue/DOI/ISBN match
  the public record (已验证 / 信息不一致 / 无法验证 / 排除 / 用户提供);
* ``support`` - does the excerpt the agent quotes for its claim actually occur in
  source text the script could read (abstract or the linked page).
"""

from __future__ import annotations
import datetime as dt
import difflib
import html
import json
import re
import socket
import urllib.error
import urllib.parse
import urllib.request

UA = "xaufe-document-format/2 (reference check)"
TIMEOUT = 20
CSL = "application/vnd.citationstyles.csl+json"
DOI_URL = "https://doi.org/"
RA_URL = "https://doi.org/ra/"
CROSSREF_URL = "https://api.crossref.org/works/"
CROSSREF_SEARCH = "https://api.crossref.org/works?rows=5&query.bibliographic="
OPENALEX_URL = "https://api.openalex.org/works/doi:"
ISBN_URL = "https://openlibrary.org/isbn/"
OPENLIBRARY = "https://openlibrary.org"
RECENT_YEARS = 5
DEFAULT_MIN = 6

# Registration agencies whose records doi.org can return as CSL JSON.
NEGOTIABLE = {"crossref", "datacite", "medra", "jalc", "kisti"}

# 已验证: matches the public record.  信息不一致: same work, some field differs - fix
# before use.  无法验证: could not be checked (blocked site, no network, other-language
# record) - usable but never shown as verified.  排除: does not exist, is retracted,
# or is another work - never use.  用户提供: from the user, not checked by the script.
PASS, FIX, DROP, PENDING, USER = "已验证", "信息不一致", "排除", "无法验证", "用户提供"
SUPPORTED, NOT_FOUND, UNCHECKED, NO_QUOTE = "摘录已在原文中找到", "摘录在能读到的原文里没找到", "取不到原文，未能对照", "没有提供原文摘录"
BLOCKED_HTTP = {401, 402, 403, 405, 406, 418, 429, 451, 500, 502, 503, 504}
BLOCKED_PAGE = re.compile(r"验证码|安全验证|人机验证|请登录|登录后|captcha|access denied|are you a robot|enable javascript|cloudflare", re.I)


class Offline(Exception):
    """The record service could not be reached; ``down`` means no network at all."""

    def __init__(self, message, down=False):
        super().__init__(message)
        self.down = down


class Unsupported(Exception):
    """The DOI exists but its registry does not return machine-readable metadata."""


def fetch(url, accept=None):
    """JSON from ``url``; None when the record does not exist (HTTP 404)."""
    headers = {"User-Agent": UA}
    if accept:
        headers["Accept"] = accept
    try:
        with urllib.request.urlopen(urllib.request.Request(url, headers=headers), timeout=TIMEOUT) as response:
            body = response.read()
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            return None
        raise Offline(f"HTTP {exc.code}") from exc
    except urllib.error.URLError as exc:
        down = isinstance(exc.reason, (socket.gaierror, ConnectionError))
        raise Offline(str(exc.reason), down) from exc
    except OSError as exc:
        raise Offline(str(exc)) from exc
    try:
        return json.loads(body)
    except ValueError as exc:
        raise Unsupported(url) from exc


class Page(str):
    """Visible text of a page; ``headline`` is what the page says it is (<title>, <h1>, citation_title)."""

    headline = None


def _headline(raw_html):
    parts = re.findall(r"<title[^>]*>(.*?)</title>|<h1[^>]*>(.*?)</h1>", raw_html, re.I | re.S)
    parts = [a or b for a, b in parts]
    parts += re.findall(r"<meta[^>]+(?:name|property)=[\"'](?:citation_title|dc\.title|og:title)[\"'][^>]*content=[\"']([^\"']*)",
                        raw_html, re.I)
    return html.unescape(re.sub(r"<[^>]+>", " ", " ".join(parts)))


def fetch_page(url):
    """Visible text of a web page, tags removed. Raises Offline when it cannot be read."""
    if not re.match(r"https?://", url, re.I):
        raise Offline("不是 http(s) 链接")
    headers = {"User-Agent": "Mozilla/5.0 (compatible; " + UA + ")"}
    try:
        with urllib.request.urlopen(urllib.request.Request(url, headers=headers), timeout=TIMEOUT) as response:
            raw = response.read(4_000_000)
            charset = response.headers.get_content_charset()
    except urllib.error.HTTPError as exc:
        error = Offline(f"HTTP {exc.code}")
        error.code = exc.code
        raise error from exc
    except (urllib.error.URLError, OSError) as exc:
        raise Offline(str(getattr(exc, "reason", exc))) from exc
    meta = re.search(rb"charset=[\"']?([\w-]+)", raw[:4000])
    for encoding in (charset, meta and meta.group(1).decode("ascii", "ignore"), "utf-8", "gb18030"):
        try:
            text = raw.decode(encoding) if encoding else None
        except (LookupError, UnicodeDecodeError):
            continue
        if text is not None:
            break
    else:
        text = raw.decode("utf-8", "ignore")
    page = Page(html.unescape(re.sub(r"<[^>]+>", " ", text)))
    page.headline = _headline(text)
    return page


def _online():
    try:
        fetch(RA_URL + "10.1038")
        return True
    except (Offline, Unsupported):
        return False


SEARCH_URL = re.compile(
    r"(?:google\.[a-z.]+/(?:search|scholar)|scholar\.google|baidu\.com/s\b|xueshu\.baidu\.com/s\b|bing\.com/search"
    r"|sogou\.com/|so\.com/s\b|duckduckgo\.com|search\.cnki|kns\.cnki\.net/kns8?s?/(?:search|defaultresult)"
    r"|[?&](?:q|wd|word|query|keyword|kw|searchword)=)",
    re.I,
)


def _surname(author):
    """The part of an author string a page or reference line is sure to contain."""
    author = str(author or "").strip()
    if _cjk(author):
        return re.sub(r"\s+", "", author)
    words = sorted(re.findall(r"[A-Za-zÀ-ÿ'-]{2,}", author), key=len, reverse=True)
    return words[0] if words else ""


def _contains_word(haystack, word):
    if not word:
        return True
    if _cjk(word):
        return word in re.sub(r"\s+", "", haystack)
    return re.search(r"(?<![A-Za-z])" + re.escape(word) + r"(?![A-Za-z])", haystack, re.I) is not None


def _page_check(entry, result, why=""):
    """Entries without a usable DOI record: the linked page must show title, first author and year.

    A page that refuses automated access (403/429, captcha, login wall) cannot
    be checked either way, so the entry is 无法验证, not 排除; only a page that
    loads and lacks the title, or a dead link, is evidence against the entry.
    Search-result and aggregator links prove nothing about the work itself.
    """
    url = str(entry.get("url") or "").strip()
    title = str(entry.get("title") or "").strip()
    if not url:
        result.update(status=FIX, notes=[why + "请加上原文页面的 url（能打开、页面上有题目）"])
    elif SEARCH_URL.search(url):
        result.update(status=FIX, notes=["url 是搜索结果页或检索入口，证明不了这篇文献存在：换成文献本身的页面（期刊官网文章页、出版社书目页、学位论文库详情页）"])
    elif not title:
        result.update(status=FIX, notes=["缺 title：写上文献的题目，脚本要在链接页面上找它"])
    else:
        try:
            page = fetch_page(url)
        except Offline as exc:
            code = getattr(exc, "code", None)
            if code in BLOCKED_HTTP:
                result.update(status=PENDING, notes=[why + f"网站拒绝自动访问（{exc}），无法验证；请用户自己打开链接核对"])
                return result
            if not _online():
                raise Offline(str(exc), down=True) from exc
            if code is None and re.search(r"timed? ?out|超时", str(exc), re.I):
                result.update(status=PENDING, notes=[why + f"网站响应超时（{exc}），无法验证；请用户自己打开链接核对"])
                return result
            result.update(status=DROP, notes=[f"链接打不开（{exc}）：换一个能打开的原文页面链接，或用 DOI"])
            return result
        if _norm(title) in _norm(page):
            authors = entry.get("authors") or []
            first = _surname(authors[0] if isinstance(authors, list) and authors else authors)
            year = str(entry.get("year") or "").strip()[:4]
            missing = [x for x, ok in ((f"第一作者 {first}", _contains_word(page, first)), (f"年份 {year}", not year or year in page)) if not ok]
            if not first:
                missing.insert(0, "authors（refs.json 里没写作者）")
            headline = getattr(page, "headline", None)
            if missing:
                result.update(status=FIX, notes=["链接页面上有这个题目，但找不到" + "、".join(missing)
                                                 + "：可能是引用了它的别的页面。换成文献本身的页面，或核对作者和年份"])
            elif headline is not None and _norm(title) not in _norm(headline):
                # A page that only mentions the work (a citing paper, a reading list) proves less than its own page.
                result.update(status=PENDING, notes=[why + "链接页面提到了这篇文献，但页面标题不是它，可能是引用它的别的页面，无法验证；"
                                                     "换成文献本身的页面（页面标题就是这篇文献）"])
            else:
                result["notes"].append("链接页面上有题目、第一作者和年份")
                result["_source_text"] = page
        elif BLOCKED_PAGE.search(page[:20000]) or len(_norm(page)) < 200:
            result.update(status=PENDING, notes=[why + "链接打开的是验证码、登录页或空白页，无法验证；请用户自己打开链接核对"])
        else:
            result.update(
                status=DROP,
                notes=["链接页面上找不到这个题目：链接要指向这篇文献本身（不是网站首页），不要编链接"],
            )
    return result


def _abstract(crossref, openalex):
    """Abstract text from Crossref (JATS) or OpenAlex (inverted index); '' when neither has one."""
    index = openalex.get("abstract_inverted_index") if isinstance(openalex, dict) else None
    if isinstance(index, dict) and index:
        words = sorted((pos, word) for word, positions in index.items() for pos in positions)
        return " ".join(word for _, word in words)
    text = (crossref or {}).get("abstract") or ""
    return html.unescape(re.sub(r"<[^>]+>", " ", str(text)))


def _quote_key(text):
    return re.sub(r"[\W_]+", "", str(text or "").lower())


def support(entry, source_text):
    """Does the excerpt quoted for the claim occur in source text we could read?"""
    quote = str(entry.get("quote") or "").strip()
    if not quote:
        return {"support": NO_QUOTE}
    key = _quote_key(quote)
    if len(key) < 8:
        return {"support": NO_QUOTE, "support_note": "摘录太短，至少写一整句原文"}
    if not source_text:
        return {"support": UNCHECKED}
    haystack = _quote_key(source_text)
    if key in haystack:
        return {"support": SUPPORTED}
    match = difflib.SequenceMatcher(None, key, haystack, autojunk=False).find_longest_match(0, len(key), 0, len(haystack))
    if match.size >= 0.9 * len(key):
        return {"support": SUPPORTED}
    return {"support": NOT_FOUND, "support_note": "摘录可能来自全文；能读到的只有摘要或页面。确认摘录是原文原句，不是自己的概括"}


def clean_isbn(value):
    """ISBN digits when the check digit is right, else ''."""
    digits = re.sub(r"[^0-9Xx]", "", str(value or "")).upper()
    if len(digits) == 13 and digits.isdigit():
        total = sum(int(d) * (1 if i % 2 == 0 else 3) for i, d in enumerate(digits[:12]))
        return digits if (10 - total % 10) % 10 == int(digits[12]) else ""
    if len(digits) == 10 and digits[:9].isdigit():
        total = sum(int(d) * (10 - i) for i, d in enumerate(digits[:9]))
        check = (11 - total % 11) % 11
        return digits if digits[9] == ("X" if check == 10 else str(check)) else ""
    return ""


def _isbn_check(entry, result):
    isbn = clean_isbn(entry.get("isbn"))
    if not isbn:
        result.update(status=DROP, notes=[f"ISBN {entry.get('isbn')} 校验位不对：ISBN 写错了或是编的"])
        return result
    try:
        book = fetch(ISBN_URL + isbn + ".json")
    except Unsupported:
        book = None
    if book is None:
        if entry.get("url"):
            return _page_check(entry, result)
        # Open Library covers few Chinese books: a valid ISBN with no record is unverified, not fake.
        result.update(status=PENDING, notes=[f"ISBN {isbn} 校验位正确，但 Open Library 里没有这本书的记录，没法自动核对；"
                                             "有出版社或图书馆的书目页面就写进 url 再试，否则按无法验证处理"])
        return result
    title = str(entry.get("title") or "").strip()
    record = str(book.get("title") or "")
    if not title:
        result.update(status=FIX, notes=[f"缺 title：记录里的书名是《{record}》，确认是同一本再写上"])
    elif _cjk(title) != _cjk(record):
        result.update(status=PENDING, notes=[f"ISBN 存在，但记录里的书名是《{record}》（拼音或外文），没法自动比对"])
    elif not _similar(title, record):
        result.update(status=DROP, notes=[f"ISBN 对应的是另一本书《{record}》，和你写的《{title}》对不上"])
        return result
    year = _int((re.search(r"\d{4}", str(book.get("publish_date") or "")) or [None])[0])
    if year and result["year"] and year != result["year"]:
        result["status"] = FIX
        result["notes"].append(f"出版年份对不上：记录里是 {year}，你写的是 {result['year']}（版次不同就写你用的那一版的 ISBN）")
    _book_authors(entry, book, result)
    return result


def _book_authors(entry, book, result):
    """The first author written for a book must be on the Open Library record."""
    authors = entry.get("authors") or []
    first = str(authors[0] if isinstance(authors, list) and authors else authors or "").strip()
    names = [str(book.get("by_statement") or "")]
    for ref in (book.get("authors") or [])[:3]:
        key = str((ref or {}).get("key") or "")
        if re.fullmatch(r"/authors/OL\w+", key):
            try:
                names.append(str((fetch(OPENLIBRARY + key + ".json") or {}).get("name") or ""))
            except (Offline, Unsupported):
                pass
    names = [n for n in names if n.strip()]
    if not first:
        result["status"] = FIX if result["status"] == PASS else result["status"]
        result["notes"].append("refs.json 里没写作者" + (f"：记录里是 {names[-1]}" if names else ""))
    elif not names:
        if result["status"] == PASS:
            result["status"] = PENDING
        result["notes"].append("书目记录里没有作者，作者没法自动核对")
    elif _cjk(first) != _cjk(" ".join(names)):
        if result["status"] == PASS:
            result["status"] = PENDING
        result["notes"].append(f"记录里作者写作 {names[-1]}（拼音或外文），和你写的 {first} 没法自动比对；请核对后再告诉用户")
    elif not any(_contains_word(n, _surname(first)) or _norm(_surname(first)) in _norm(n) for n in names):
        result["status"] = FIX if result["status"] != DROP else DROP
        result["notes"].append(f"第一作者对不上：书目记录里是 {names[-1]}，你写的是 {first}")


def _norm(text):
    return re.sub(r"[\W_]+", "", str(text or "").lower())


def _similar(a, b):
    """Same title, allowing for a subtitle that only one side writes."""
    main = lambda t: re.split(r"[:：]|\s[-–—]\s", str(t or ""), maxsplit=1)[0]
    return any(_same(x, y) for x, y in ((a, b), (main(a), b), (a, main(b))))


def _same(a, b):
    # Shortened English titles ("... of Pandas and Polars" for "... of the ... Pandas
    # and Polars Data Analysis Python Libraries"): nearly all words of the shorter
    # title appear in the longer one, and they cover most of it.
    words_a, words_b = (set(re.findall(r"[a-z0-9]+", str(t).lower())) for t in (a, b))
    few, many = sorted((words_a, words_b), key=len)
    if len(few) >= 4 and len(few & many) >= 0.9 * len(few) and len(few) >= 0.6 * len(many):
        return True
    a, b = _norm(a), _norm(b)
    if not a or not b:
        return False
    short, long = sorted((a, b), key=len)
    return a == b or difflib.SequenceMatcher(None, a, b).ratio() >= 0.9 or (
        short in long and len(short) >= 0.6 * len(long)
    )


def _cjk(text):
    return bool(re.search(r"[一-鿿]", str(text or "")))


def _first(value):
    if isinstance(value, list):
        return value[0] if value else ""
    return value or ""


def _int(value):
    try:
        return int(str(value).strip()[:4])
    except (TypeError, ValueError):
        return None


def clean_doi(value):
    doi = str(value or "").strip()
    doi = re.sub(r"^(https?://(dx\.)?doi\.org/|doi:\s*)", "", doi, flags=re.I)
    return doi if doi.startswith("10.") and "/" in doi else ""


def _years(*records):
    years = set()
    for record in records:
        for key in ("issued", "published-print", "published-online", "published"):
            parts = ((record or {}).get(key) or {}).get("date-parts") or []
            if parts and parts[0] and parts[0][0]:
                years.add(int(parts[0][0]))
    return years


def _families(csl):
    return [str(a.get("family") or a.get("literal") or a.get("name") or "") for a in csl.get("author") or []]


def _record_authors(entry, csl, result):
    """Every author written in refs.json must match the record, in order; at least the first three."""
    record = [f for f in _families(csl) if f]
    authors = entry.get("authors") or []
    authors = authors if isinstance(authors, list) else [authors]
    if not record or _cjk(record[0]) != _cjk(str(authors[0] if authors else "")):
        return
    for i, (given, family) in enumerate(zip(authors, record), 1):
        if _norm(family) not in _norm(given):
            result["status"] = FIX
            result["notes"].append(f"第 {i} 位作者对不上：记录里是 {family}，你写的是 {given}")
            return
    if len(authors) < min(3, len(record)):
        result["status"] = FIX
        result["notes"].append(f"作者少写了：记录里有 {len(record)} 位作者，至少要写前 {min(3, len(record))} 位（{'、'.join(record[:3])}）")
    elif len(authors) > len(record):
        result["status"] = FIX
        result["notes"].append(f"作者多写了：记录里只有 {len(record)} 位（{'、'.join(record)}）")


def text_problems(entry):
    """The printed reference line (text) must say what was checked: title, authors, year, venue, DOI."""
    text = str(entry.get("text") or "")
    flat = _norm(text)
    problems = []
    title = str(entry.get("title") or "")
    main = re.split(r"[:：]|\s[-–—]\s", title, maxsplit=1)[0]
    if title and _norm(main) not in flat:
        problems.append(f"text 里没有题目“{title[:30]}”")
    authors = entry.get("authors") or []
    authors = authors if isinstance(authors, list) else [authors]
    absent = [a for a in authors[:3] if not _contains_word(text, _surname(a))]
    if absent:
        problems.append("text 里没有作者 " + "、".join(map(str, absent)))
    year = str(entry.get("year") or "").strip()[:4]
    if year and year not in text:
        problems.append(f"text 里没有年份 {year}")
    journal = str(entry.get("journal") or "")
    if journal and not _venue_named(text, journal):
        problems.append(f"text 里没有刊名/会议名“{journal[:30]}”")
    written = re.search(r"10\.\d{4,9}/\S+", text)
    doi = clean_doi(entry.get("doi"))
    if written and doi and written.group(0).rstrip(".,;，。").lower() != doi.lower():
        problems.append(f"text 里的 DOI {written.group(0)} 和 doi 字段 {doi} 不一样")
    return problems


def _venue_named(text, venue):
    """Does the printed line name this journal or conference (abbreviations allowed)?"""
    if _cjk(venue):
        return _norm(venue) in _norm(text)
    # Venues are often abbreviated in the printed line; most distinctive words must be there.
    words = [w for w in re.findall(r"[a-z]{4,}", venue.lower()) if w not in ("proceedings", "conference", "journal", "international")]
    short = re.findall(r"\(([A-Za-z]{2,10})\)", venue)  # "... Recognition (CVPR)" printed as "CVPR"
    return not words or sum(w in text.lower() for w in words) >= 0.6 * len(words) or any(_contains_word(text, a) for a in short)


GBT = re.compile(r"(?P<authors>[^.．]+?)[.．]\s*(?P<title>.+?)\s*\[(?P<kind>[A-Z]{1,2})(?:/OL)?\]\s*[.．]?\s*(?P<rest>.*)")


def parse_line(line):
    """refs.json fields guessed from a printed reference line (GB/T 7714 style); ``text`` is kept as is."""
    text = str(line or "").strip()
    body = re.sub(r"^\s*[\[［]\d+[\]］]\s*", "", text)
    entry = {"text": text}
    doi = re.search(r"10\.\d{4,9}/[^\s，。；;]+", body)
    if doi:
        entry["doi"] = doi.group(0).rstrip(".,")
    url = re.search(r"https?://[^\s，。；;]+", body)
    if url and not doi:
        entry["url"] = url.group(0).rstrip(".,")
    isbn = re.search(r"ISBN[:：\s]*([\dXx-]{10,17})", body)
    if isbn:
        entry["isbn"] = isbn.group(1)
    year = re.search(r"(?<!\d)(?:19|20)\d{2}(?!\d)", re.sub(r"10\.\d{4,9}/\S+|https?://\S+", "", body))
    if year:
        entry["year"] = int(year.group(0))
    m = GBT.match(body)
    if m:
        names = [a.strip() for a in re.split(r"[,，、;；]", m.group("authors")) if a.strip()]
        entry["authors"] = [a for a in names if a not in ("等", "et al", "et al.")]
        entry["title"] = m.group("title").strip()
        venue = re.split(r"[,，.．]", m.group("rest"), maxsplit=1)[0].strip()
        if m.group("kind") in ("J", "C") and venue and not re.match(r"\d", venue):
            entry["journal"] = venue
    return entry


def _first_author(csl):
    authors = csl.get("author") or []
    if not authors:
        return ""
    head = authors[0]
    return head.get("family") or head.get("literal") or head.get("name") or ""


def _find_doi(entry):
    """DOI of an entry written without one: same title and same first author on Crossref."""
    title = str(entry.get("title") or "").strip()
    authors = entry.get("authors") or []
    first = str(authors[0] if isinstance(authors, list) and authors else authors or "")
    if len(_norm(title)) < 6 or not first:
        return ""
    try:
        found = fetch(CROSSREF_SEARCH + urllib.parse.quote(f"{title} {first}"))
    except Unsupported:
        return ""
    except Offline as exc:
        if exc.down:
            raise
        return ""
    items = (found.get("message") or {}).get("items") if isinstance(found, dict) else None
    year = _int(entry.get("year"))
    for item in items if isinstance(items, list) else []:
        # A preprint, repost or later edition under the same title is not the work cited.
        issued = ((item.get("issued") or {}).get("date-parts") or [[None]])[0][0]
        if item.get("type") == "posted-content" or (year and issued and abs(int(issued) - year) > 1):
            continue
        family = str(((item.get("author") or [{}])[0]).get("family") or "")
        same_author = _cjk(family) != _cjk(first) or (family and _norm(family) in _norm(first))
        if same_author and any(_similar(title, t) for t in item.get("title") or []):
            return clean_doi(item.get("DOI"))
    return ""


def check_entry(entry, this_year):
    """Verdict for one refs.json entry. May raise Offline."""
    text = str(entry.get("text") or "").strip()
    result = {"text": text, "status": PASS, "notes": [], "year": _int(entry.get("year")), "per_year": None}
    if not text:
        result.update(status=DROP, notes=["没有 text（写进参考文献列表的那一行）"])
        return result
    doi = clean_doi(entry.get("doi"))
    if not doi and entry.get("isbn"):
        return _isbn_check(entry, result)
    found = "" if doi or entry.get("doi") else _find_doi(entry)
    if found:
        result = check_entry({**entry, "doi": found}, this_year)
        result["notes"].insert(0, f"没写 DOI，按题目和作者查到 {found}，已按它的记录核对（把它写进 doi）")
        return result
    if not doi:
        if entry.get("doi"):
            result.update(status=FIX, notes=[f"doi 写法不对：{entry.get('doi')}（应以 10. 开头，例如 10.1109/CVPR.2016.90）"])
        elif entry.get("url"):
            return _page_check(entry, result)
        else:
            # Not findable is not "does not exist": many real books and Chinese papers have no DOI.
            result.update(status=PENDING, notes=["没有 DOI、ISBN 或链接，自动核对不了（不是查到它不存在）："
                                                 "图书补上 ISBN、文章补上 DOI 或官方页面链接后可以再核对"])
        return result

    agency = (fetch(RA_URL + doi.split("/", 1)[0]) or [{}])[0]
    if agency.get("status") == "DOI does not exist":
        result.update(status=DROP, notes=[f"DOI 前缀 {doi.split('/', 1)[0]} 不存在，DOI 写错了"])
        return result
    ra = str(agency.get("RA") or "")
    if not ra:
        raise Offline("查不到 DOI 的登记机构")
    if ra.lower() not in NEGOTIABLE:
        return _page_check(entry, result, f"这个 DOI 由 {ra} 登记，脚本取不到它的记录；")
    quoted = urllib.parse.quote(doi, safe="/")
    try:
        csl = fetch(DOI_URL + quoted, CSL)
    except Unsupported:
        return _page_check(entry, result, "这个 DOI 取不到记录；")
    if csl is None:
        result.update(status=DROP, notes=[f"DOI {doi} 查不到，可能写错或不存在"])
        return result
    try:
        crossref = (fetch(CROSSREF_URL + quoted) or {}).get("message") or {}
    except Unsupported:
        crossref = {}
    try:
        openalex = fetch(OPENALEX_URL + quoted) or {}
    except Unsupported:
        openalex = {}

    record_title = str(_first(csl.get("title")) or _first(crossref.get("title")))
    updates = crossref.get("updated-by") or []
    if (
        openalex.get("is_retracted")
        or any(re.search(r"retract|withdraw|removal", str(u.get("type", "")), re.I) for u in updates)
        or re.match(r"\s*(retracted|withdrawn)\b", record_title, re.I)
    ):
        result.update(status=DROP, notes=["已撤稿"])
        return result

    # Chinese papers are often registered with English titles, pinyin authors and
    # English journal names; those cannot be compared and are left to the agent.
    title = str(entry.get("title") or "").strip()
    strip = lambda t: re.sub(r"^\s*(retracted|withdrawn)\s*:\s*", "", str(t), flags=re.I)
    bare = strip(record_title)
    titles = [strip(t) for t in [record_title, *(crossref.get("title") or []), *(crossref.get("original-title") or []), openalex.get("title")] if t]
    comparable = [t for t in titles if _cjk(t) == _cjk(title)]
    if not title:
        result["status"] = FIX
        result["notes"].append(f"缺 title：记录里的题目是《{bare}》，确认是你读过的那篇再写上")
    elif not comparable:
        if entry.get("url"):
            return _page_check(entry, result)
        result["status"] = PENDING
        result["notes"].append(f"记录里只有另一种语言的题目《{bare}》，没法自动比对；加上原文页面的 url 可以自动核对")
    elif not any(_similar(title, t) for t in comparable):
        result.update(
            status=DROP,
            notes=[f"DOI 对应的是另一篇《{bare}》，和你写的《{title}》对不上。删掉这条，不要改成查到的那篇"],
        )
        return result

    record_author = _first_author(csl)
    authors = entry.get("authors") or []
    given_first = authors[0] if isinstance(authors, list) and authors else str(authors or "")
    if record_author and _cjk(record_author) != _cjk(given_first):
        if result["status"] == PASS:
            result["status"] = PENDING
        result["notes"].append(f"记录里第一作者写作 {record_author}（拼音或外文），和你写的 {given_first} 没法自动比对，"
                               "所以标为无法验证；请核对后再告诉用户")
    elif record_author and _norm(record_author) not in _norm(given_first):
        result["status"] = FIX
        result["notes"].append(f"第一作者对不上：记录里是 {record_author}，你写的是 {given_first or '（空）'}")
    else:
        _record_authors(entry, csl, result)

    years = _years(csl, crossref)
    if openalex.get("publication_year"):
        years.add(int(openalex["publication_year"]))
    if years and result["year"] not in years:
        result["status"] = FIX
        result["notes"].append(f"年份对不上：记录里是 {min(years)}，你写的是 {entry.get('year') or '（空）'}")
    record_year = result["year"] if result["year"] in years else (min(years) if years else result["year"])

    journal = str(entry.get("journal") or "").strip()
    containers = [
        str(_first(csl.get("container-title"))),
        str(_first(crossref.get("container-title"))),
        str(((openalex.get("primary_location") or {}).get("source") or {}).get("display_name") or ""),
    ]
    if journal:
        containers = [c for c in containers if c and _cjk(c) == _cjk(journal)]
        if containers and not any(_similar(journal, c) or _venue_named(journal, c) for c in containers):
            result["status"] = FIX
            result["notes"].append(f"刊名/会议名和记录不一样：记录里是 {containers[0]}，确认后按记录写")
    else:  # no journal field: the printed line itself must name the record's venue
        containers = [c for c in containers if c and _cjk(c) == _cjk(title or text) and _norm(c) != "arxiv"]
        if containers and not any(_venue_named(text, c) for c in containers):
            result["status"] = FIX
            result["notes"].append(f"text 里的刊名/会议名和记录对不上：记录里是 {containers[0]}，按记录写进 text 和 journal")

    if (
        doi.lower().startswith("10.48550/")
        or str(csl.get("publisher", "")).lower() == "arxiv"
        or openalex.get("type") == "preprint"
        or crossref.get("type") == "posted-content"
    ):
        result["notes"].append("预印本（质量分 1）：查有没有正式发表的版本，有就换成正式版")

    result["_source_text"] = _abstract(crossref, openalex)
    if "cited_by_count" in openalex and record_year:
        result["per_year"] = round(openalex["cited_by_count"] / max(1, this_year - record_year), 1)
    result["year"] = record_year
    return result


def _as_user(verdict):
    """The user's own reference: checked like any other, reported, never blocking and never shown as verified."""
    verdict["checked"] = verdict["status"]
    verdict["status"] = USER
    verdict["notes"].insert(0, f"用户提供；自动核对结果：{verdict['checked']}（只报告，用不用由用户决定）")
    verdict.pop("unreachable", None)


def check_all(entries, this_year=None):
    this_year = this_year or dt.date.today().year
    results, offline = [], None
    seen = {}
    for index, entry in enumerate(entries, 1):
        if not isinstance(entry, dict):
            results.append({"text": str(entry), "status": DROP, "notes": ["这一项不是 {...} 对象"], "year": None, "per_year": None})
            continue
        key = clean_doi(entry.get("doi")).lower() or _norm(entry.get("title")) or _norm(entry.get("text"))
        if key and key in seen:
            results.append(
                {"text": str(entry.get("text") or ""), "status": DROP, "notes": [f"和第 {seen[key]} 条重复"], "year": None, "per_year": None}
            )
            continue
        if key:
            seen[key] = index
        if offline and clean_doi(entry.get("doi")):
            results.append(
                {"text": str(entry.get("text") or ""), "status": PENDING, "notes": [offline], "year": _int(entry.get("year")),
                 "per_year": None, "unreachable": True,
                 "support": UNCHECKED if str(entry.get("quote") or "").strip() else NO_QUOTE}
            )
            continue
        try:
            verdict = check_entry(entry, this_year)
            verdict.update(support(entry, verdict.pop("_source_text", "")))
            mismatch = text_problems(entry) if verdict["status"] in (PASS, PENDING) else []
            if mismatch:
                verdict["status"] = FIX
                verdict["notes"].append("；".join(mismatch) + "。text 是印在论文里的那一行，要和核对过的字段一致")
            if entry.get("from_user"):
                _as_user(verdict)
            results.append(verdict)
        except Offline as exc:
            note = f"连不上核对服务（{exc}），这条没核对"
            if exc.down:
                offline = note
            verdict = {"text": str(entry.get("text") or ""), "status": PENDING, "notes": [note], "year": _int(entry.get("year")),
                       "per_year": None, "unreachable": True,
                       "support": UNCHECKED if str(entry.get("quote") or "").strip() else NO_QUOTE}
            if entry.get("from_user"):
                _as_user(verdict)
            results.append(verdict)
    # Keep the excerpt and the claim beside each verdict so the citing sentence can be compared with them.
    for entry, result in zip(entries, results):
        if isinstance(entry, dict):
            for key in ("quote", "supports"):
                if str(entry.get(key) or "").strip():
                    result.setdefault(key, str(entry[key]).strip())
    return results


def plain_reason(result):
    """Why an entry is not 已验证, in words a user understands (no HTTP codes or proxy errors)."""
    if result.get("unreachable"):
        return "核对时连不上文献数据库，没有核对"
    note = "；".join(result.get("notes") or [])
    note = re.sub(r"（[^（）]*(?:HTTP|Errno|Tunnel|timed out|SSL|URLError|urlopen|Forbidden|proxy)[^（）]*）", "", note, flags=re.I)
    return note or result.get("status", "")


def report(results, entries, minimum=DEFAULT_MIN, this_year=None, report_only=False):
    this_year = this_year or dt.date.today().year
    lines = []
    if report_only:
        lines += ["只排版：下面的核对只报告给用户，不改原文的参考文献。", ""]
    for index, r in enumerate(results, 1):
        extra = f"，年均引用 {r['per_year']}" if r["per_year"] is not None else ""
        lines.append(f"第 {index} 条：{r['status']}{extra}  {r['text'][:60]}")
        lines.extend("    - " + note for note in r["notes"])
        if r.get("support") and r["status"] != DROP:
            lines.append(f"    - 观点支撑：{r['support']}" + (f"（{r['support_note']}）" if r.get("support_note") else ""))
    usable = [i for i, r in enumerate(results) if r["status"] != DROP]
    counted = [
        results[i]["year"]
        for i in usable
        if isinstance(entries[i], dict) and not entries[i].get("classic") and results[i]["year"]
    ]
    recent = sum(1 for y in counted if this_year - y < RECENT_YEARS)
    if report_only:
        lines.append("")
        lines.append("→ 下一步：")
        lines.append("  refs.json 的字段（authors/title/year/journal/doi）是脚本从原文猜的；结论是“缺 title”、作者或刊名对不上这类时，"
                     "先对照原文把字段补对（text 保持原样），再运行 refs。")
        lines.append("  不要删改用户原文里的文献。finish 会生成核对清单，原样附在给用户的回复里，由用户决定要不要改。")
        return "\n".join(lines)
    lines.append("")
    verified = [i for i in usable if results[i]["status"] == PASS]
    unchecked = [i for i in usable if results[i].get("unreachable")]
    lines.append(f"可用 {len(usable)} 条（不含“排除”；其中已验证 {len(verified)} 条"
                 + (f"，{len(unchecked)} 条因为连不上核对服务根本没核对" if unchecked else "")
                 + f"），本次至少要 {minimum} 条（用户另有要求按用户的）。")
    if counted:
        lines.append(
            f"近 {RECENT_YEARS} 年（{this_year - RECENT_YEARS + 1} 年以后）{recent}/{len(counted)} 条，"
            f"占 {round(100 * recent / len(counted))}%（建议不低于 60%，不是硬性要求；classic 为 true 的经典文献不算在内）。"
        )
    lines.append("")
    lines.append("→ 下一步：")
    lines.append("  排除：从 refs.json 删掉，另找一篇；不要把它改成查到的那篇。")
    lines.append("  信息不一致：确认是同一篇后，按提示同时改 refs.json 里的字段和 text，再运行 refs；不改不能用。")
    lines.append("  无法验证：先问用户是否接受，用户同意后在 job.json 写 \"accept_unverified_refs\": true，否则 run 会拦下；"
                 "能换成可核实的文献就换。finish 会把原因列进给用户的清单。")
    if any(r.get("unreachable") for r in results):
        lines.append("  连不上核对服务的条目：联网后重新运行 refs。确实没法联网时先问用户是否接受未核实的文献，"
                     "用户同意后在 job.json 写 \"accept_unverified_refs\": true，否则 run 会拦下。")
    lines.append("  观点支撑：给每条写 quote（支撑你那句话的原文原句，必写）；没找到或没对照的，run 会列出引用句和摘录让你逐条核对，清单里如实标出。")
    if len(usable) < minimum:
        lines.append(f"  可用的不够 {minimum} 条：继续检索；实在找不到就如实告诉用户还差 {minimum - len(usable)} 条，不要凑数。")
    lines.append("  都处理好后，把要用的条目的 text 按正文第一次引用的顺序写进 content.json 的 references。")
    return "\n".join(lines)
