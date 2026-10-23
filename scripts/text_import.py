"""Verbatim import of a .txt/.md original into content.json (format only: every word kept).

Paragraphs are cut where the original leaves a blank line (hard-wrapped lines are joined),
Markdown syntax is dropped the way ``source_compare.fragments`` drops it, and each paragraph
gets the role ``xaufe.guess_roles`` would give it in a Word original.  The result is a draft
for the agent to review: the words are the original's, the structure is a guess.
"""

from __future__ import annotations
import re
import shutil
from pathlib import Path
from textio import read_text_any

LINK = re.compile(r"\[([^\]]+)\]\(([^)\s]+)[^)]*\)")
IMAGE = re.compile(r"^!\[([^\]]*)\]\(([^)\s]+)[^)]*\)$")
CJK = re.compile(r"[　-〿㐀-鿿＀-￯]")
KEYWORD_SPLIT = re.compile(r"\s*[；;，,、]\s*|\s{2,}|\s+(?=[㐀-鿿])")


def _join(lines):
    """Hard-wrapped lines -> one paragraph: no space between Chinese, one space between words."""
    text = ""
    for line in lines:
        if text and not (CJK.search(text[-1]) or CJK.search(line[0])):
            text += " "
        text += line
    return text


def _pieces(text, markdown=False):
    """The original's lines -> [("para"|"heading"|"code"|"table"|"image", payload)].  In Markdown a
    blank line ends a table; a plain .txt often leaves blank lines between the rows of one table."""
    found, lines, fence, code = [], [], None, []

    def flush():
        if lines:
            found.append(("para", _join(lines)))
        lines.clear()

    rows = []
    for raw in text.splitlines() + [""]:
        stripped = raw.strip()
        if fence:
            if stripped.startswith(fence):
                found.append(("code", "\n".join(code)))
                fence, code = None, []
            else:
                code.append(raw.rstrip())
            continue
        if rows and (stripped or markdown) and not stripped.startswith("|"):
            found.append(("table", rows))
            rows = []
        if stripped.startswith(("```", "~~~")):
            flush()
            fence = stripped[:3]
            continue
        if not stripped or re.fullmatch(r"[-*_=\s]{3,}", stripped):
            flush()
            continue
        image = IMAGE.match(stripped)
        if image:
            flush()
            found.append(("image", (image.group(1), image.group(2))))
            continue
        if stripped.startswith("|"):
            flush()
            cells = [c.strip() for c in stripped.strip("|").split("|")]
            if not all(re.fullmatch(r":?-+:?", c) for c in cells if c):
                rows.append(cells)
            continue
        if stripped.startswith("#"):
            flush()
            found.append(("heading", (len(stripped) - len(stripped.lstrip("#")), stripped.lstrip("#").strip())))
            continue
        stripped = re.sub(r"(?<!\\)(\*\*|__|`)", "", stripped)
        stripped = re.sub(r"^[-*+]\s+", "", stripped)
        lines.append(stripped)
    if rows:
        found.append(("table", rows))
    if fence and code:
        found.append(("code", "\n".join(code)))
    return found


def _paragraph(text):
    """A paragraph block; Markdown links become runs so the link itself is kept."""
    if not LINK.search(text):
        return {"type": "paragraph", "text": text}
    runs, at = [], 0
    for m in LINK.finditer(text):
        if m.start() > at:
            runs.append({"text": text[at:m.start()]})
        runs.append({"text": m.group(1), "link": m.group(2)})
        at = m.end()
    if at < len(text):
        runs.append({"text": text[at:]})
    return {"type": "paragraph", "runs": runs}


def _keywords(text):
    body = re.sub(r"^\s*(关键词|key\s*words?)\s*[:：]\s*", "", text, flags=re.I)
    return [k for k in KEYWORD_SPLIT.split(body.strip().rstrip("。.")) if k]


def to_content(path, work):
    """(content, notes) for the original at ``path``; images found next to it are copied into ``work``."""
    from xaufe import guess_roles

    path = Path(path)
    raw = read_text_any(path)
    pieces = _pieces(raw, markdown=path.suffix.lower() == ".md")
    blocks, notes = [], []
    plain = re.sub(r"```.*?```|~~~.*?~~~", "", raw, flags=re.S)
    marked = len(re.findall(r"\*\*[^*\n]+\*\*|__[^_\n]+__|`[^`\n]+`", plain))
    if marked:
        notes.append(f"原稿有 {marked} 处 Markdown 加粗或行内代码标记（**、__、`），导入时只留了文字、没带格式："
                     "要保留某处加粗，就在 content.json 那一段改用 runs 写 {\"text\": \"…\", \"bold\": true}。")
    texts = [(i, kind, payload) for i, (kind, payload) in enumerate(pieces) if kind in ("para", "heading")]
    first_heading = next((i for i, kind, _ in texts if kind == "heading"), None)
    pseudo = []
    for i, kind, payload in texts:
        text = payload if kind == "para" else payload[1]
        style = None
        if kind == "heading" and not (i == first_heading and i == texts[0][0]):
            style = f"markdown heading {min(payload[0], 3)}"  # not a template style name: special words still win
        pseudo.append({"id": f"x{i}", "kind": "paragraph", "text": text, "style": style, "images": [],
                       "equations": [], "section_break": False, "block_index": i})
    roles, _ = guess_roles(pseudo)
    # Numbering decides a heading's level (as in Word originals); Markdown "#" only when unnumbered.
    from xaufe import _heading_level
    for b in pseudo:
        level = _heading_level(dict(b, style=None))
        if roles[b["id"]].startswith("h") and level:
            roles[b["id"]] = f"h{level}"
    # An abstract without its "摘要" label: the body paragraphs right before the keywords, ahead of any heading.
    seen_heading = False
    for at, b in enumerate(pseudo):
        role = roles[b["id"]]
        seen_heading = seen_heading or role in ("h1", "h2", "h3", "h4", "preface")
        if role in ("keywords_zh", "keywords_en") and not seen_heading:
            back = at - 1
            while back >= 0 and roles[pseudo[back]["id"]] == "body":
                roles[pseudo[back]["id"]] = "abstract_" + role[-2:]
                back -= 1
    role_of = {int(b["id"][1:]): roles[b["id"]] for b in pseudo}
    text_of = {int(b["id"][1:]): b["text"] for b in pseudo}

    content = {"mode": "document", "metadata": {}, "blocks": blocks}
    abstract = {"abstract_zh": [], "abstract_en": []}
    front = []  # abstract pieces, kept as paragraphs if this turns out not to be a full paper
    references = []
    pending_caption = None
    for i, (kind, payload) in enumerate(pieces):
        role = role_of.get(i)
        text = text_of.get(i)
        if kind == "code":
            blocks.append({"type": "code", "text": payload})
        elif kind == "table":
            width = max(len(r) for r in payload)
            rows = [r + [""] * (width - len(r)) for r in payload]
            table = {"type": "table", "header": rows[0], "rows": rows[1:] or [[""] * width]}
            if pending_caption is not None and blocks and blocks[-1] is pending_caption:
                table["caption"] = blocks.pop()["text"]
            blocks.append(table)
        elif kind == "image":
            alt, target = payload
            source = (path.parent / target).resolve()
            if source.is_file() and source.suffix.lower() in (".png", ".jpg", ".jpeg", ".gif", ".bmp"):
                shutil.copy2(source, Path(work) / source.name)
                blocks.append({"type": "image", "path": source.name})
            else:
                notes.append(f"图片 {target} 没找到（或不是常见图片格式），没有导入：请用户提供图片文件后补一个 image 块。")
        elif role in ("title_zh", "title_en"):
            content["metadata"].setdefault(role, text)
            if role == "title_en":
                front.append((role, text))  # only a full paper prints it (on the English abstract page)
        elif role in ("abstract_label_zh", "abstract_label_en", "reference_title", "toc_title", "preserve"):
            front.append((role, text))
            if role in ("toc_title", "preserve"):
                notes.append(f"“{text[:20]}”看起来是原稿的目录，没有导入：目录由模板按标题重新生成。")
        elif role in ("abstract_zh", "abstract_en"):
            abstract[role].append(text)
            front.append((role, text))
        elif role in ("keywords_zh", "keywords_en"):
            content[role] = _keywords(text)
            front.append((role, text))
        elif role == "reference":
            references.append(text)
        elif role == "preface":
            blocks.append({"type": "heading", "level": 1, "text": text, "preface": True})
        elif role in ("h1", "h2", "h3", "h4"):
            blocks.append({"type": "heading", "level": int(role[1]), "text": text})
        elif role == "caption" and text.startswith(("图", "Figure")) and blocks and blocks[-1]["type"] == "image":
            blocks[-1]["caption"] = text
        else:
            blocks.append({"type": "paragraph", "text": text} if role == "caption" else _paragraph(text))
            pending_caption = blocks[-1] if role == "caption" else None
            continue
        pending_caption = None
    if references:
        content["references"] = references
    paper = abstract["abstract_zh"] and content.get("keywords_zh") and references and any(
        b["type"] == "heading" for b in blocks)
    if paper:
        content["mode"] = "paper"
        for role, parts in abstract.items():
            if parts:
                content[role] = "\n".join(parts)
    else:
        content.pop("keywords_zh", None)
        content.pop("keywords_en", None)
        content["metadata"].pop("title_en", None)
        kept = [{"type": "paragraph", "text": t} for role, t in front
                if role not in ("toc_title", "preserve", "reference_title")]
        content["blocks"] = kept + blocks
        if kept:
            notes.append("原稿不是完整的论文结构（摘要、关键词、标题、参考文献缺一项），摘要和关键词按普通段落放在正文最前面；"
                         "要按论文结构排就先问用户，再把它们移到 abstract_zh / keywords_zh 并改 mode。")
    if not content["metadata"].get("title_zh"):
        notes.append("没认出全篇题目：把题目写进 metadata.title_zh（和 job.json 的 info.title_zh 一致）。")
    return content, notes
