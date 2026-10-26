"""Heading numbers in the template's way: 一、 / （一） / 1. / （1）, counted again from the text.

Format-only included: a heading numbered another way ("1.1", "第一章", "1、") or not at all is
renumbered, and each change is reported to the user.  A heading whose number belongs to another
level of the same document ("（一）" set as h1 while the other h2s use "（一）") is a level mistake in
mapping.json / content.json, not a number to rewrite, so it is refused instead."""

from __future__ import annotations
import re

LEVELS = ("h1", "h2", "h3", "h4")
CN_DIGITS = "一二三四五六七八九"
CN = "一二三四五六七八九十百零〇"
# The typed number at the start of a heading, by kind; the first match wins.
SCHEMES = (
    ("第X章", re.compile(rf"^\s*第[{CN}\d]+\s*(?:章|部分|篇)\s*[、:：.．]?\s*")),
    ("第X节", re.compile(rf"^\s*第[{CN}\d]+\s*节\s*[、:：.．]?\s*")),
    ("X、", re.compile(rf"^\s*[{CN}]+\s*[、.．]\s*")),
    ("（X）", re.compile(rf"^\s*[（(]\s*[{CN}]+\s*[)）]\s*[、.．]?\s*")),
    ("（n）", re.compile(r"^\s*[（(]\s*\d+\s*[)）]\s*[、.．]?\s*")),
    ("n.n.n", re.compile(r"^\s*\d+\s*[.．]\s*\d+\s*[.．]\s*\d+(?:\s*[.．]\s*\d+)*\s*[.．、]?\s*")),
    ("n.n", re.compile(r"^\s*\d+\s*[.．]\s*\d+\s*[.．、]?\s*")),
    ("n.", re.compile(r"^\s*\d+\s*[.．]\s*")),
    ("n、", re.compile(r"^\s*\d+\s*、\s*")),
    ("n)", re.compile(r"^\s*\d+\s*[)）]\s*")),
    ("n ", re.compile(r"^\s*\d{1,2}\s+(?=\S)")),
    ("①", re.compile(r"^\s*[①-⑳]\s*")),
    ("Chapter n", re.compile(r"^\s*Chapter\s+\d+\s*[:：.．]?\s*", re.I)),
)
TEMPLATE = {"h1": "X、", "h2": "（X）", "h3": "n.", "h4": "（n）"}
# Headings that are never numbered: they follow the body, not part of its chapters.
UNNUMBERED = re.compile(r"^\s*(致\s*谢|谢\s*辞|后\s*记|附\s*录.*|参\s*考\s*文\s*献|结\s*束\s*语)\s*$")


def chinese(n):
    if n <= 10:
        return "十" if n == 10 else CN_DIGITS[n - 1]
    tens, ones = divmod(n, 10)
    return ("" if tens == 1 else CN_DIGITS[tens - 1]) + "十" + (CN_DIGITS[ones - 1] if ones else "")


def scheme(text):
    """(kind, rest): the kind of typed number at the start and the words after it."""
    for kind, rx in SCHEMES:
        m = rx.match(text)
        if m and text[m.end():].strip():
            return kind, text[m.end():].strip()
    return None, text.strip()


def label(level, n):
    return {"h1": f"{chinese(n)}、", "h2": f"（{chinese(n)}）", "h3": f"{n}.", "h4": f"（{n}）"}[level]


NAME = {"h1": "一级", "h2": "二级", "h3": "三级", "h4": "四级"}


def conflicts(items, kinds):
    """Headings whose number shows they sit at the wrong level: the template's number of another
    level ("（一）" set as h1, "一、" set as h2), or a number the document mostly uses at another level.
    "1." is the template's h3 number but also the top of a "1" / "1.1" / "1.1.1" scheme, so it only
    signals h3 when the document has no "1.1" numbers."""
    from collections import Counter

    decimal = any(kind in ("n.n", "n.n.n") for kind in kinds)
    signal = {TEMPLATE[level]: level for level in ("h1", "h2", "h3") if not (decimal and level == "h3")}
    used = {}
    for (level, _), kind in zip(items, kinds):
        if kind:
            used.setdefault(kind, Counter())[level] += 1
    problems = []
    for (level, text), kind in zip(items, kinds):
        other = signal.get(kind)
        if other and other != level:
            problems.append(f"标题“{text.strip()[:20]}”设成了 {level}，但它的编号是模板{NAME[other]}标题的写法，应为 {other}")
        elif kind and not other and len(used[kind]) > 1 and used[kind].most_common(1)[0][0] != level:
            problems.append(f"标题“{text.strip()[:20]}”设成了 {level}，但文档里同样编号的标题大多是 {used[kind].most_common(1)[0][0]}")
    return problems


def renumber(items):
    """items: [(level, text)] in document order (h1-h4 only).  Returns [(index, before, after)] for
    the headings whose text changes.  Raises ValueError naming every heading whose number shows it
    is set at the wrong level (the level is fixed in mapping.json / content.json, not the number)."""
    kinds = [scheme(text)[0] for _, text in items]
    problems = conflicts(items, kinds)
    if problems:
        raise ValueError("；".join(problems) + "：多半是级别填错了，先改级别（mapping.json 的 role 或 content.json 的 level）再 run")
    counters = dict.fromkeys(LEVELS, 0)
    changes = []
    for i, (level, text) in enumerate(items):
        if UNNUMBERED.match(text):
            continue
        counters[level] += 1
        for deeper in LEVELS[LEVELS.index(level) + 1:]:
            counters[deeper] = 0
        after = label(level, counters[level]) + scheme(text)[1]
        if after != text.strip():
            changes.append((i, text, after))
    return changes
