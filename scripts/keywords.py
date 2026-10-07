"""Normalize only identified keyword separators, without rebuilding hyperlinks/runs."""

from __future__ import annotations
import re
from copy import deepcopy
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# A blank the author typed after the colon belongs to the label and is kept.
LABEL = re.compile(r"^(关键词\s*[:：]|Key\s*words\s*[:：])[ 　]*", re.I)


def visible(p):
    return "".join(p._p.xpath(".//w:t/text()"))


def replace_range(nodes, start, end, replacement):
    pos = 0
    inserted = False
    for node in nodes:
        text = node.text or ""
        a, b = pos, pos + len(text)
        pos = b
        if b <= start or a >= end:
            continue
        l = max(0, start - a)
        r = min(len(text), end - a)
        node.text = text[:l] + ("" if inserted else replacement) + text[r:]
        node.set(qn("xml:space"), "preserve")
        inserted = True
    if not inserted:
        raise ValueError("关键词空格替换范围无效。")


def normalize(p, words=None):
    original = visible(p)
    m = LABEL.match(original)
    if m is None:
        raise ValueError("关键词段须包含“关键词：”或“Key words：”标签，请核对段落角色。")
    # Fields, real tabs and objects are not flattened to infer semantic boundaries.
    if p._p.xpath(".//w:fldSimple|.//w:instrText|.//w:drawing|.//w:pict|.//w:tab|.//w:br"):
        raise ValueError("关键词含复杂域、制表或对象；先明确词项及无损处理方式，不自动拆词。")
    # Blanks at the end of the paragraph are layout, not part of a keyword.
    trimmed = None
    core = original.rstrip(" 　")
    if len(core) > m.end() and core != original:
        replace_range(p._p.xpath(".//w:t"), len(core), len(original), "")
        trimmed, original = original, visible(p)
    value = original[m.end() :]
    separators = []
    if words is not None:
        if not isinstance(words, list) or any(
            not isinstance(x, str) or not x or x.strip() != x for x in words
        ):
            raise ValueError("keyword_items 必须是非空字符串列表；空关键词用空列表。")
        if not words:
            if value.strip():
                raise ValueError("词项列表与原关键词文字不一致。")
        else:
            pattern = r"\A" + r"([ ]+)".join(re.escape(word) for word in words) + r"\Z"
            matched = re.match(pattern, value)
            if matched is None:
                raise ValueError("明确词项与原文不一致；禁止借关键词排版改写内容。")
            separators = [
                (m.end() + matched.start(i), m.end() + matched.end(i)) for i in range(1, len(words))
            ]
    else:
        separators = [
            (m.end() + x.start(), m.end() + x.end()) for x in re.finditer(r" {2,}", value)
        ]
        tokens = value.split(" ")
        if not separators and len(tokens) > 1 and all(re.search(r"[一-鿿]", t) for t in tokens):
            # Every word contains Chinese characters, so single spaces cannot be
            # inside an English phrase: they are the separators.
            separators = [(m.end() + x.start(), m.end() + x.end()) for x in re.finditer(r" ", value)]
    nodes = p._p.xpath(".//w:t")
    expected = original
    for a, b in reversed(separators):
        if a == m.end() or b == len(original):
            raise ValueError("关键词首尾存在空白，不能确定为词间分隔；请明确词项。")
        expected = expected[:a] + "   " + expected[b:]
        replace_range(nodes, a, b, "   ")
    if visible(p) != expected:
        raise ValueError("关键词空格修改未通过逐字保全检查。")
    before = trimmed or original
    return (
        None
        if expected == before
        else {
            "before": before,
            "after": expected,
            "allowed_change": "仅已识别的词间ASCII空格统一为3个，并去掉段末空白",
        }
    )


def split_label_runs(p):
    text = visible(p)
    m = LABEL.match(text)
    if m is None:
        raise ValueError("关键词标签缺失。")
    boundary = m.end()
    offset = 0
    out = []
    for r in list(p._p.xpath(".//w:r")):
        texts = r.findall(qn("w:t"))
        text = "".join(t.text or "" for t in texts)
        if not text:
            continue
        end = offset + len(text)
        if offset < boundary < end:
            if any(e.tag not in (qn("w:rPr"), qn("w:t")) for e in r):
                raise ValueError("关键词标签与复杂运行混排，不能安全分离。")
            clone = deepcopy(r)
            r.addnext(clone)
            for el in texts:
                r.remove(el)
            for el in clone.findall(qn("w:t")):
                clone.remove(el)
            for target, value in (
                (r, text[: boundary - offset]),
                (clone, text[boundary - offset :]),
            ):
                t = OxmlElement("w:t")
                t.set(qn("xml:space"), "preserve")
                t.text = value
                target.append(t)
            out.append((r, True))
            out.append((clone, False))
        else:
            out.append((r, end <= boundary))
        offset = end
    return out


def spacing_issue(p):
    text = visible(p)
    m = LABEL.match(text)
    if m is None:
        return "关键词标签缺失。"
    value = text[m.end() :]
    if value.strip() != value:
        return "关键词首尾存在未明确的空白。"
    if any(len(x.group()) != 3 for x in re.finditer(r" {2,}", value)):
        return "已识别的关键词分隔符不是3个ASCII空格。"
    if p._p.xpath(".//w:tab"):
        return "关键词含制表符；需要明确词项分隔。"
    return None
