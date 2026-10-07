"""Check reference entries against public DOI records (doi.org, Crossref, OpenAlex).

The agent writes WORK/refs.json; ``xaufe.py refs`` prints one verdict per entry
plus the totals that references/citations.md asks for.  Only metadata is
checked here; relevance and quality are judged by the agent.
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
RECENT_YEARS = 5
DEFAULT_MIN = 6

# Registration agencies whose records doi.org can return as CSL JSON.
NEGOTIABLE = {"crossref", "datacite", "medra", "jalc", "kisti"}

PASS, FIX, DROP, PENDING, USER = "通过", "需改正", "排除", "待核实", "用户提供"


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
        raise Offline(f"HTTP {exc.code}") from exc
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
    return html.unescape(re.sub(r"<[^>]+>", " ", text))


def _online():
    try:
        fetch(RA_URL + "10.1038")
        return True
    except (Offline, Unsupported):
        return False


def _page_check(entry, result, why=""):
    """Entries without a usable DOI record: the linked page must show the title."""
    url = str(entry.get("url") or "").strip()
    title = str(entry.get("title") or "").strip()
    if not url:
        result.update(status=FIX, notes=[why + "请加上原文页面的 url（能打开、页面上有题目）"])
    elif not title:
        result.update(status=FIX, notes=["缺 title：写上文献的题目，脚本要在链接页面上找它"])
    else:
        try:
            page = fetch_page(url)
        except Offline as exc:
            if not _online():
                raise Offline(str(exc), down=True) from exc
            result.update(status=DROP, notes=[f"链接打不开（{exc}）：换一个能打开的原文页面链接，或用 DOI"])
            return result
        if _norm(title) in _norm(page):
            result["notes"].append("链接页面上有这个题目")
        else:
            result.update(
                status=DROP,
                notes=["链接页面上找不到这个题目：链接要指向这篇文献本身（不是网站首页），不要编链接"],
            )
    return result


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
        result.update(status=DROP, notes=[f"ISBN {isbn} 查不到这本书；有书目页面链接就写进 url 再试"])
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
    return result


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
    for item in items if isinstance(items, list) else []:
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
    if entry.get("from_user"):
        result["status"] = USER
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
            result.update(status=DROP, notes=["没有 DOI，也没有可访问的链接"])
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
        result["notes"].append(f"记录里第一作者写作 {record_author}，和中文名没法自动比对，自己核对")
    elif record_author and _norm(record_author) not in _norm(given_first):
        result["status"] = FIX
        result["notes"].append(f"第一作者对不上：记录里是 {record_author}，你写的是 {given_first or '（空）'}")

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
    containers = [c for c in containers if c and _cjk(c) == _cjk(journal)]
    if journal and containers and not any(_similar(journal, c) for c in containers):
        result["status"] = FIX
        result["notes"].append(f"刊名/会议名和记录不一样：记录里是 {containers[0]}，确认后按记录写")

    if (
        doi.lower().startswith("10.48550/")
        or str(csl.get("publisher", "")).lower() == "arxiv"
        or openalex.get("type") == "preprint"
        or crossref.get("type") == "posted-content"
    ):
        result["notes"].append("预印本（质量分 1）：查有没有正式发表的版本，有就换成正式版")

    if "cited_by_count" in openalex and record_year:
        result["per_year"] = round(openalex["cited_by_count"] / max(1, this_year - record_year), 1)
    result["year"] = record_year
    return result


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
        if offline and clean_doi(entry.get("doi")) and not entry.get("from_user"):
            results.append(
                {"text": str(entry.get("text") or ""), "status": PENDING, "notes": [offline], "year": _int(entry.get("year")), "per_year": None}
            )
            continue
        try:
            results.append(check_entry(entry, this_year))
        except Offline as exc:
            note = f"连不上核对服务（{exc}），这条没核对"
            if exc.down:
                offline = note
            results.append(
                {"text": str(entry.get("text") or ""), "status": PENDING, "notes": [note], "year": _int(entry.get("year")), "per_year": None}
            )
    return results


def report(results, entries, minimum=DEFAULT_MIN, this_year=None):
    this_year = this_year or dt.date.today().year
    lines = []
    for index, r in enumerate(results, 1):
        extra = f"，年均引用 {r['per_year']}" if r["per_year"] is not None else ""
        lines.append(f"第 {index} 条：{r['status']}{extra}  {r['text'][:60]}")
        lines.extend("    - " + note for note in r["notes"])
    usable = [i for i, r in enumerate(results) if r["status"] != DROP]
    counted = [
        results[i]["year"]
        for i in usable
        if isinstance(entries[i], dict) and not entries[i].get("classic") and results[i]["year"]
    ]
    recent = sum(1 for y in counted if this_year - y < RECENT_YEARS)
    lines.append("")
    lines.append(f"可用 {len(usable)} 条（不含“排除”），要求至少 {minimum} 条。")
    if counted:
        lines.append(
            f"近 {RECENT_YEARS} 年（{this_year - RECENT_YEARS + 1} 年以后）{recent}/{len(counted)} 条，"
            f"占 {round(100 * recent / len(counted))}%，要求不低于 60%（classic 为 true 的经典文献不算在内）。"
        )
    lines.append("")
    lines.append("→ 下一步：")
    lines.append("  排除：从 refs.json 删掉，另找一篇；不要把它改成查到的那篇。")
    lines.append("  需改正：确认是同一篇后，按提示同时改 refs.json 里的字段和 text，再运行 refs。")
    lines.append("  待核实：可以照常用；finish 会把它列进给用户的清单。")
    if len(usable) < minimum:
        lines.append(f"  可用的不够 {minimum} 条：继续检索；实在找不到就如实告诉用户还差 {minimum - len(usable)} 条，不要凑数。")
    lines.append("  都处理好后，把要用的条目的 text 按正文第一次引用的顺序写进 content.json 的 references。")
    return "\n".join(lines)
