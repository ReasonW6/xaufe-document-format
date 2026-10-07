"""Fit long cover text on the first page without changing words or font sizes.

Long titles, course names and field values are wrapped into balanced lines that
stay inside their underlined areas. When the extra lines would push the
completion date off page 1, spare blank lines are collapsed step by step
(level 1-2), then the fixed line pitch of the title/field lines is reduced
(level 3-4). Font sizes never change. The real render remains the final judge:
xaufe.py retries with the next level if the date still leaves page 1.

The same plan is used by the writer and by check_format, so a fitted cover is
checked against the exact geometry it was given.
"""

from __future__ import annotations
import math
import unicodedata
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

MAX_LEVEL = 4
# Free height (pt) between the last default cover line and the bottom date area,
# measured in Word 16 for the untouched masters; kept slightly conservative.
SLACK_PT = {"A": 80, "B": 100}
LINE_PT = 36  # fixed pitch of title and field lines in both masters
COURSE_LINE_PT = 47  # single-spaced 36 pt heading line (B course name)
PITCH = {0: 720, 1: 720, 2: 720, 3: 560, 4: 480}  # twips; font size unchanged
NO_LINE_START = set("，。、；：！？）》」』”’,.;:!?)%")
NO_LINE_END = set("（《「『“‘(")
GOOD_BREAK_BEFORE = set("（(《")
GOOD_BREAK_AFTER = set("）)》、，,与和及")


def _wide(char):
    return unicodedata.east_asian_width(char) in ("W", "F")


def width(text, size_pt):
    """Approximate rendered width in twips (CJK = 1 em, other = 0.55 em)."""
    return sum(size_pt * 20 * (1 if _wide(c) else 0.55) for c in text)


def _tokens(text):
    """CJK characters break anywhere; runs of ASCII letters/digits stay together."""
    out, word = [], ""
    for char in text:
        if not _wide(char) and not char.isspace():
            word += char
            continue
        if word:
            out.append(word)
            word = ""
        out.append(char)
    if word:
        out.append(word)
    return out


def balanced_lines(text, size_pt, max_twips):
    """Split text into the fewest lines that fit, with similar line lengths.

    Prefers breaks around brackets and conjunctions; never starts a line with
    closing punctuation or ends one with opening punctuation.
    """
    text = text.strip()
    if not text or width(text, size_pt) <= max_twips:
        return [text]
    tokens = _tokens(text)
    count = max(2, math.ceil(width(text, size_pt) / max_twips))
    while count <= len(tokens):
        lines = _split(tokens, size_pt, count, max_twips)
        if lines:
            return lines
        count += 1
    return list(text)


def _break_penalty(tokens, i, em):
    """Cost of breaking before tokens[i]; None when the break is not allowed."""
    before, after = tokens[i - 1], tokens[i]
    if after[0] in NO_LINE_START or before[-1] in NO_LINE_END or after.isspace():
        return None
    if before[-1] in GOOD_BREAK_AFTER or after[0] in GOOD_BREAK_BEFORE:
        return 0.0
    return 9 * em * em  # same cost as two lines each off by ~2 characters


def _split(tokens, size_pt, count, max_twips):
    """Dynamic programming over break positions for exactly `count` lines."""
    em = size_pt * 20
    widths = [width(t, size_pt) for t in tokens]
    prefix = [0.0]
    for w in widths:
        prefix.append(prefix[-1] + w)
    target = prefix[-1] / count
    n = len(tokens)
    inf = float("inf")
    best = [[inf] * (n + 1) for _ in range(count + 1)]
    back = [[0] * (n + 1) for _ in range(count + 1)]
    best[0][0] = 0.0
    for k in range(1, count + 1):
        for i in range(k, n + 1):
            for j in range(k - 1, i):
                if best[k - 1][j] == inf:
                    continue
                line = prefix[i] - prefix[j]
                if line > max_twips:
                    continue
                penalty = 0.0
                if j > 0:
                    penalty = _break_penalty(tokens, j, em)
                    if penalty is None:
                        continue
                cost = best[k - 1][j] + (line - target) ** 2 + penalty
                if cost < best[k][i]:
                    best[k][i] = cost
                    back[k][i] = j
    if best[count][n] == inf:
        return None
    cuts, i = [], n
    for k in range(count, 0, -1):
        j = back[k][i]
        cuts.append((j, i))
        i = j
    return ["".join(tokens[a:b]).strip() for a, b in reversed(cuts)]


def _run_size(paragraph, index, default):
    from style_model import run_properties

    runs = paragraph.runs
    if index >= len(runs):
        return default
    return int(run_properties(paragraph, runs[index]._r).get("sz", str(default * 2))) / 2


def _ind(paragraph):
    nodes = paragraph._p.xpath("./w:pPr/w:ind")
    return {k.split("}")[1]: v for k, v in nodes[0].attrib.items()} if nodes else {}


def _tabs(paragraph):
    return [
        (t.get(qn("w:val")), int(t.get(qn("w:pos"))))
        for t in paragraph._p.xpath("./w:pPr/w:tabs/w:tab")
        if t.get(qn("w:val")) != "clear"
    ]


def plan(doc, spec, values, level=None, min_level=0):
    """Compute wrapped texts, field geometry and the vertical fit level.

    doc is a freshly loaded master (after customization.apply_cover) whose cover
    paragraphs still have their original layout. values hold the final cover
    strings ('' for blank fields).
    """
    cover = spec["cover"]
    page = spec["page_twips"]
    text_width = page["width"] - page["left"] - page["right"]
    ps = doc.paragraphs
    result = {"fields": {}, "field_tabs": None, "field_indent": {}, "extra_pt": 0.0}

    # Title: label + value on one paragraph.
    title = values.get("title_zh", "")
    tp = ps[cover["title_paragraph"]]
    label_size = _run_size(tp, 0, 18)
    value_size = _run_size(tp, 1, 16)
    first = int(_ind(tp).get("firstLine", "0"))
    label_w = width(tp.runs[0].text, label_size)
    available = max(value_size * 20, text_width - first - label_w)
    title_lines = balanced_lines(title, value_size, available) if title else [""]
    result["title"] = {
        "lines": title_lines,
        "indent": (
            {"left": str(round(first + label_w)), "right": "0", "hanging": str(round(label_w))}
            if len(title_lines) > 1 and spec["id"] == "A"
            else None
        ),
    }
    # A wrapped title always uses the spare blank line right under it (p6).
    result["extra_pt"] += LINE_PT * (len(title_lines) - 1) - (LINE_PT if len(title_lines) > 1 else 0)

    # Course name (B cover first line), centred 36 pt heading.
    if cover.get("course_paragraph") is not None:
        cp = ps[cover["course_paragraph"]]
        size = _run_size(cp, 0, 36)
        right = int(_ind(cp).get("right", "0"))
        course = values.get("course_name", "").strip().strip("《》")
        lines = balanced_lines("《" + course + "》", size, text_width - right - size * 6)
        result["course"] = {"lines": lines if course else ["《　　　　　》"]}
        result["extra_pt"] += COURSE_LINE_PT * (len(result["course"]["lines"]) - 1)

    # Field lines share one underlined zone: label end E .. right tab R.
    field_ps = {key: ps[i] for key, i in cover["fields"].items()}
    sample = ps[cover["fields"]["student_name"]]
    size = _run_size(sample, 1, 16)
    label_size = _run_size(sample, 0, 16)
    left = int(_ind(sample).get("left", "0"))
    label_end = left + width(sample.runs[0].text, label_size)
    right_tab = max(pos for kind, pos in _tabs(sample) if kind == "right")
    pad = size * 20
    default_zone = right_tab - label_end - pad
    max_zone = text_width - label_end - pad
    widths = {key: width(values.get(key, ""), size) for key in field_ps}
    widest = max(widths.values(), default=0)
    if widest > default_zone:
        new_right = text_width if widest > max_zone else min(text_width, label_end + widest + 2 * pad)
        result["field_tabs"] = {"center": round((label_end + new_right) / 2), "right": round(new_right)}
    for key in field_ps:
        value = values.get(key, "")
        lines = balanced_lines(value, size, max_zone) if value else [""]
        result["fields"][key] = lines
        if len(lines) > 1:
            # Continuation lines start just left of the field zone; the first line's
            # tab after the label is never captured by this hanging-indent stop.
            start = round(label_end) - 60
            result["field_indent"][key] = {"left": str(start), "hanging": str(start - left)}
        result["extra_pt"] += LINE_PT * (len(lines) - 1)

    multiline = len(title_lines) > 1
    if level is None:
        level = min_level
        while level < MAX_LEVEL and result["extra_pt"] > capacity(spec, level, multiline):
            level += 1
    result["level"] = max(level, min_level)
    result["fits"] = result["extra_pt"] <= capacity(spec, result["level"], multiline)
    return result


def _spacers(spec, level):
    """Blank cover paragraphs collapsed at each level (indices from the masters)."""
    supervisor_slot = "supervisor" in spec["cover"]["fields"]
    extra = [] if supervisor_slot else [13]
    if level <= 0:
        return []
    chosen = [7, 8] + extra
    if level >= 2:
        chosen += [6, 4]
    if level >= 3:
        chosen += [1, 2] if spec["id"] == "A" else [1]
    return sorted(set(chosen))


def _reclaimed_pt(spec, level, title_multiline=False):
    sizes = {1: 36 if spec["id"] == "A" else 27, 2: 36, 4: 47, 6: 36, 7: 36, 8: 36, 13: 36}
    if title_multiline:
        sizes[6] = 0  # already used by the wrapped title
    gained = sum(sizes[i] for i in _spacers(spec, level))
    lines = 1 + len(spec["cover"]["fields"])
    gained += lines * (720 - PITCH[level]) / 20
    return gained


def capacity(spec, level, title_multiline=False):
    return SLACK_PT[spec["id"]] + _reclaimed_pt(spec, level, title_multiline)


def apply(doc, spec, layout):
    """Write wrapped texts and adjusted geometry into the cover paragraphs."""
    cover = spec["cover"]
    ps = doc.paragraphs
    tp = ps[cover["title_paragraph"]]
    title_lines = layout["title"]["lines"]
    tp.runs[1].text = "\n".join(title_lines) if any(title_lines) else "　" * 9
    if layout["title"]["indent"]:
        _set(tp, "ind", layout["title"]["indent"])
        _set(tp, "jc", {"val": "left"})
    if len(title_lines) > 1:
        # The spare line under a wrapped title is used by the title itself.
        _collapse(ps[6])
    if "course" in layout:
        ps[cover["course_paragraph"]].runs[0].text = "\n".join(layout["course"]["lines"])
    for key, lines in layout["fields"].items():
        p = ps[cover["fields"][key]]
        p.runs[1].text = "\n".join("\t" + line + "\t" for line in lines)
        if layout["field_tabs"]:
            tabs = p._p.get_or_add_pPr().find(qn("w:tabs"))
            for tab in tabs.findall(qn("w:tab")):
                kind = tab.get(qn("w:val"))
                if kind in layout["field_tabs"]:
                    tab.set(qn("w:pos"), str(layout["field_tabs"][kind]))
        if key in layout["field_indent"]:
            ind = p._p.get_or_add_pPr().find(qn("w:ind"))
            ind.attrib.pop(qn("w:firstLine"), None)
            ind.attrib.pop(qn("w:firstLineChars"), None)
            for name, value in layout["field_indent"][key].items():
                ind.set(qn("w:" + name), value)
    level = layout["level"]
    for index in _spacers(spec, level):
        _collapse(ps[index])
    if PITCH[level] != 720:
        for index in [cover["title_paragraph"], *cover["fields"].values()]:
            spacing = ps[index]._p.get_or_add_pPr().find(qn("w:spacing"))
            spacing.set(qn("w:line"), str(PITCH[level]))
            spacing.set(qn("w:lineRule"), "exact")


def _collapse(paragraph):
    _set(paragraph, "spacing", {"line": "20", "lineRule": "exact", "before": "0", "after": "0"})


def _set(paragraph, tag, attributes):
    ppr = paragraph._p.get_or_add_pPr()
    old = ppr.find(qn("w:" + tag))
    node = OxmlElement("w:" + tag)
    for key, value in attributes.items():
        node.set(qn("w:" + key), value)
    if old is not None:
        old.addprevious(node)
        ppr.remove(old)
    else:
        ppr.append(node)
