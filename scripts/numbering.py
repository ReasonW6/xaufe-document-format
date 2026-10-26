"""The label Word shows in front of an auto-numbered paragraph ("1.", "第一章", "(a)", "•").

python-docx's paragraph text leaves these labels out, but they are text the reader sees: a
heading numbered by its list reads "1. 引言", not "引言".  ``list_labels`` replays Word's
counters over the document so that callers can compare or keep that visible text."""

from __future__ import annotations

from docx.oxml.ns import qn

from style_model import style_chain

_CN_DIGITS = "零一二三四五六七八九"
_TIANGAN = "甲乙丙丁戊己庚辛壬癸"
_ROMAN = ((1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
          (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"))


def _chinese(n):
    if n < 10:
        return _CN_DIGITS[n]
    if n < 100:
        tens, ones = divmod(n, 10)
        return ("" if tens == 1 else _CN_DIGITS[tens]) + "十" + (_CN_DIGITS[ones] if ones else "")
    if n >= 1000:
        return str(n)
    hundreds, rest = divmod(n, 100)
    tail = "" if not rest else "零" + _CN_DIGITS[rest] if rest < 10 else ("一" if rest < 20 else "") + _chinese(rest)
    return _CN_DIGITS[hundreds] + "百" + tail


def _roman(n):
    out = ""
    for value, mark in _ROMAN:
        while n >= value:
            out, n = out + mark, n - value
    return out


def _letters(n):
    return chr(ord("A") + (n - 1) % 26) * ((n - 1) // 26 + 1)


FORMATS = {
    "decimal": str,
    "decimalZero": lambda n: f"{n:02d}",
    "decimalFullWidth": lambda n: "".join(chr(ord(c) + 0xFEE0) for c in str(n)),
    "decimalHalfWidth": str,
    "decimalEnclosedCircle": lambda n: chr(0x2460 + n - 1) if 1 <= n <= 20 else str(n),
    "decimalEnclosedParen": lambda n: chr(0x2474 + n - 1) if 1 <= n <= 20 else f"({n})",
    "decimalEnclosedFullstop": lambda n: chr(0x2488 + n - 1) if 1 <= n <= 20 else f"{n}.",
    "upperRoman": _roman,
    "lowerRoman": lambda n: _roman(n).lower(),
    "upperLetter": _letters,
    "lowerLetter": lambda n: _letters(n).lower(),
    "chineseCounting": _chinese,
    "chineseCountingThousand": _chinese,
    "chineseLegalSimplified": lambda n: _chinese(n).translate(str.maketrans("一二三四五六七八九十百", "壹贰叁肆伍陆柒捌玖拾佰")),
    "ideographDigital": lambda n: "".join(_CN_DIGITS[int(c)] if c != "0" else "〇" for c in str(n)),
    "ideographTraditional": lambda n: _TIANGAN[(n - 1) % 10],
    "japaneseCounting": _chinese,
    "taiwaneseCounting": _chinese,
    "none": lambda n: "",
}


def _val(el, tag, default=None):
    found = el.find(qn(tag)) if el is not None else None
    return found.get(qn("w:val")) if found is not None else default


def _levels(abstract):
    levels = {}
    for lvl in abstract.findall(qn("w:lvl")):
        levels[int(lvl.get(qn("w:ilvl"), "0"))] = lvl
    return levels


def _num_pr(doc, p):
    """(numId, ilvl) of a paragraph, directly or through its style; None when not numbered."""
    num = p.find(f"{qn('w:pPr')}/{qn('w:numPr')}")
    if num is None:
        style_id = p.xpath("./w:pPr/w:pStyle/@w:val")
        try:
            style = doc.styles.get_by_id(style_id[0] if style_id else None, 1)
        except Exception:
            style = None
        for st in reversed(style_chain(style) if style is not None else []):
            num = st.element.find(f"{qn('w:pPr')}/{qn('w:numPr')}")
            if num is not None:
                break
    if num is None:
        return None
    num_id = _val(num, "w:numId")
    ilvl = _val(num, "w:ilvl")
    if num_id in (None, "0"):
        return None
    return num_id, int(ilvl or 0)


def list_labels(doc):
    """{paragraph element: label} for every numbered paragraph, in document order.

    A label is the level text with its counters filled in plus what follows it (a tab or a space
    becomes one space).  A format this module cannot render gives ``None`` for that paragraph."""
    try:
        numbering = doc.part.numbering_part.element
    except (KeyError, NotImplementedError, AttributeError):
        return {}
    abstracts = {a.get(qn("w:abstractNumId")): a for a in numbering.findall(qn("w:abstractNum"))}
    nums = {}
    for num in numbering.findall(qn("w:num")):
        overrides = {}
        for over in num.findall(qn("w:lvlOverride")):
            start = _val(over, "w:startOverride")
            overrides[int(over.get(qn("w:ilvl"), "0"))] = (int(start) if start else None, over.find(qn("w:lvl")))
        nums[num.get(qn("w:numId"))] = (_val(num, "w:abstractNumId"), overrides)

    counters, started, labels = {}, set(), {}
    for p in doc.element.body.iter(qn("w:p")):
        found = _num_pr(doc, p)
        if found is None or found[0] not in nums:
            continue
        num_id, ilvl = found
        abstract_id, overrides = nums[num_id]
        levels = _levels(abstracts.get(abstract_id)) if abstract_id in abstracts else {}
        for level, (_, lvl) in overrides.items():
            if lvl is not None:
                levels[level] = lvl
        if ilvl not in levels:
            continue
        count = counters.setdefault(abstract_id, {})
        if num_id not in started:  # a lvlOverride/startOverride restarts the list where this num begins
            started.add(num_id)
            for level, (start, _) in overrides.items():
                if start is not None:
                    count[level] = start - 1
        start = int(_val(levels[ilvl], "w:start", "1"))
        count[ilvl] = count.get(ilvl, start - 1) + 1
        for deeper in [k for k in count if k > ilvl]:
            del count[deeper]
        lvl = levels[ilvl]
        text = _val(lvl, "w:lvlText", "")
        legal = lvl.find(qn("w:isLgl")) is not None
        rendered = True
        for level in range(9):
            mark = f"%{level + 1}"
            if mark not in text:
                continue
            fmt = "decimal" if legal and level < ilvl else _val(levels.get(level), "w:numFmt", "decimal")
            if fmt not in FORMATS or level not in levels:
                rendered = False
                break
            value = count.get(level, int(_val(levels[level], "w:start", "1")))
            text = text.replace(mark, FORMATS[fmt](value))
        if _val(lvl, "w:numFmt", "decimal") == "bullet":
            rendered = True
        if not rendered:
            labels[p] = None
            continue
        suffix = _val(lvl, "w:suff", "tab")
        labels[p] = text + ("" if suffix == "nothing" or not text else " ")
    return labels
