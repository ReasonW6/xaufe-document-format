#!/usr/bin/env python3
"""The only command an agent needs. Run from the skill folder:

    python -X utf8 scripts/xaufe.py compare
    python -X utf8 scripts/xaufe.py read --docx 用户的文件.docx
    python -X utf8 scripts/xaufe.py new --out "成品.docx" [--source 原稿]
    python -X utf8 scripts/xaufe.py check --work WORK
    python -X utf8 scripts/xaufe.py inspect --work WORK --docx 原稿.docx
    python -X utf8 scripts/xaufe.py refs --work WORK [--min 篇数]
    python -X utf8 scripts/xaufe.py run --work WORK [--office none] [--office-note "原因"]
    python -X utf8 scripts/xaufe.py finish --work WORK
    python -X utf8 scripts/xaufe.py preview --work WORK --docx 文件.docx
    python -X utf8 scripts/xaufe.py cancel --work WORK

Output is plain Chinese text meant to be read and followed step by step.
"""

from __future__ import annotations
import argparse
import hashlib
import json
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
ROOT = HERE.parent

EXIT_ASK_USER = 3


def _read(path, what):
    path = Path(path)
    if not path.is_file():
        raise ValueError(f"找不到 {what}：{path}")
    from engine import read_text_any

    try:
        return json.loads(read_text_any(path))
    except json.JSONDecodeError as exc:
        raise ValueError(
            f"{what} 不是合法的 JSON（第 {exc.lineno} 行附近）：{exc.msg}。"
            "最常见的原因是文字里用了英文双引号 \"，改成中文引号“ ”即可（中文引号在 JSON 里没问题）。"
        ) from exc


def _write(path, data):
    Path(path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def _work(value):
    import delivery as D

    root, state = D.load(value)
    return root, state


# ---------------------------------------------------------------- compare
def cmd_compare(_args):
    print((ROOT / "references/comparison.md").read_text(encoding="utf-8").strip())


# ---------------------------------------------------------------- new
def cmd_new(args):
    import delivery as D

    job = D.create(args.out, sources=args.source)
    work = job["workspace"]
    print(f"WORK = {work}")
    print(f"最终 Word 将保存到：{job['destination']}")
    print("下一步：")
    print(f"  1. 写 {Path(work) / 'job.json'}（写法见 references/job.md）")
    print(f"  2. 运行 check --work \"{work}\" 看还缺什么封面信息")
    print(
        f"  3. 新写文档：写 {Path(work) / 'content.json'}；已有 Word：运行 inspect --work \"{work}\" --docx 原稿.docx"
    )


# ---------------------------------------------------------------- check
def _intake(root):
    import intake as I

    job = _read(root / "job.json", "job.json")
    if isinstance(job, dict) and job.get("task") not in ("write", "format"):
        raise ValueError(
            'job.json 缺 task：这次要写内容就写 "task": "write"，只排版就写 "task": "format"（判断方法见 SKILL.md 第 0 步）。'
        )
    return I.from_job(job, root / "intake.json")


def _task(root):
    return _read(root / "job.json", "job.json").get("task")


# ---------------------------------------------------------------- writing checks
NUMBER = re.compile(r"\d[\d,.]*\s*(?:%|％|‰|元|万|亿|人|名|倍|个百分点|户|家)|百分之")
CITE = re.compile(r"[\[［]\d+(?:\s*[,，\-–]\s*\d+)*[\]］]")
CN_DIGITS = "一二三四五六七八九"
HEADING_NUMBER = {
    1: re.compile(r"^([一二三四五六七八九十]+)\s*、"),
    2: re.compile(r"^[（(]([一二三四五六七八九十]+)[)）]"),
    3: re.compile(r"^(\d+)\s*[.．]"),
}


def _cn_number(text):
    if text.isdigit():
        return int(text)
    tens, ten, ones = text.partition("十")
    if not ten:
        return CN_DIGITS.index(text) + 1 if len(text) == 1 else -1
    return (CN_DIGITS.index(tens) + 1 if tens else 1) * 10 + (CN_DIGITS.index(ones) + 1 if ones else 0)


def _ref_key(text):
    return re.sub(r"\s+", "", re.sub(r"^\s*[\[［]\d+[\]］]", "", str(text)))


def _writing_problems(root, content):
    """What a written document must satisfy before it is laid out (task = write)."""
    import refcheck as R

    problems = []
    references = content.get("references") or []
    result_path = root / "refs-result.json"
    if references and not result_path.is_file():
        problems.append("参考文献还没有核对：按 references/citations.md 写 refs.json 并运行 refs。")
    elif references:
        saved = _read(result_path, "refs-result.json")
        refs_json = root / "refs.json"
        current = hashlib.sha256(refs_json.read_bytes()).hexdigest() if refs_json.is_file() else None
        if saved.get("refs_sha256") != current:
            problems.append("refs.json 改过以后还没有重新运行 refs。")
        usable = {_ref_key(r["text"]) for r in saved.get("results", []) if r["status"] in (R.PASS, R.PENDING, R.USER)}
        for index, text in enumerate(references, 1):
            if _ref_key(text) not in usable:
                problems.append(f"参考文献第 {index} 条不在 refs 核对可用的条目里（要用 refs.json 里的 text 原样写）：{str(text)[:40]}")
    for block in content.get("blocks", []):
        if block.get("type") == "paragraph":
            runs = block.get("runs") or []
            text = block.get("text") or "".join(r.get("text", "") for r in runs)
            cited = CITE.search(text) or any(r.get("superscript") and CITE.search(r.get("text", "")) for r in runs)
        elif block.get("type") == "table":
            cells = [str(c) for row in [block.get("header") or []] + (block.get("rows") or []) for c in row]
            text = " ".join(cells)
            cited = CITE.search(str(block.get("caption") or ""))
        else:
            continue
        if NUMBER.search(text) and not cited:
            problems.append(f"有具体数字但没有标出处（删掉数字，或在句末标上核对过的文献序号）：{text[:40]}")
    counters = {1: 0, 2: 0, 3: 0}
    for block in content.get("blocks", []):
        level = block.get("level")
        if block.get("type") != "heading" or block.get("preface") or level not in HEADING_NUMBER:
            continue
        match = HEADING_NUMBER[level].match(block.get("text", "").strip())
        if not match:
            continue
        number = _cn_number(match.group(1))
        if number != counters[level] + 1:
            problems.append(f"标题编号不连续：“{block['text'][:20]}”前面一个同级标题是第 {counters[level]} 个")
        counters[level] = number
        for deeper in range(level + 1, 4):
            counters[deeper] = 0
    return problems


def _reference_checklist(root):
    """Ready-made table for the reply to the user (task = write)."""
    import refcheck as R

    content_path = root / "content.json"
    if not content_path.is_file():
        return None
    references = _read(content_path, "content.json").get("references") or []
    saved = _read(root / "refs-result.json", "refs-result.json") if (root / "refs-result.json").is_file() else {}
    minimum = saved.get("min", R.DEFAULT_MIN)
    by_key = {_ref_key(r["text"]): r for r in saved.get("results", [])}
    lines = []
    if references:
        lines += ["| 序号 | 文献 | 核实情况 | 支撑的论点 |", "|---|---|---|---|"]
        for index, text in enumerate(references, 1):
            r = by_key.get(_ref_key(text), {})
            status = {R.PASS: "已核对", R.USER: "你提供的"}.get(r.get("status"))
            if status is None:
                status = "待核实：" + "；".join(r.get("notes", []))[:40]
            supports = r.get("supports") or ""
            if supports and r.get("relevance") == "中":
                supports += "（弱支撑）"
            lines.append(f"| [{index}] | {_ref_key(text)[:50]} | {status} | {supports} |")
    else:
        lines.append("这次没有找到能核实的参考文献，请提供你读过的文献。")
    if 0 < len(references) < minimum:
        lines.append(f"参考文献现在 {len(references)} 篇，比要求的 {minimum} 篇少 {minimum - len(references)} 篇，请补充你读过的文献。")
    lines.append("请在提交前再核对一遍全部参考文献。")
    return "\n".join(lines)


def cmd_check(args):
    import intake as I

    root, _ = _work(args.work)
    state = _intake(root)
    print(I.summary(state))
    if state["stage"] != "ready":
        print("\n→ 先问用户上面缺的项（可以让用户选“留空自己填”），把回答写进 job.json 后再运行 check。")
        return EXIT_ASK_USER
    print("\n→ 封面信息齐了，可以准备内容并运行 run。")
    return 0


# ---------------------------------------------------------------- inspect
H1 = re.compile(r"^(第[一二三四五六七八九十百\d]+章|[一二三四五六七八九十]+\s*[、.．])")
H2 = re.compile(r"^[（(][一二三四五六七八九十]+[)）]")
H3 = re.compile(r"^\d+\s*[.．、]\s*\S")
NUM2 = re.compile(r"^\d+\.\d+(?!\.\d)\s*\S")
NUM3 = re.compile(r"^\d+\.\d+\.\d+\s*\S")
COVER_WORDS = ("学号", "学生姓名", "指导教师", "班级", "完成日期", "学  号", "学　　号")


def _plain(text):
    return re.sub(r"[\s　]+", "", text)


def _heading_level(block):
    style = (block.get("style") or "").lower()
    match = re.search(r"(?:heading|标题)\s*(\d)", style)
    if match and int(match.group(1)) <= 3:
        return int(match.group(1))
    text = block["text"].strip()
    if len(text) > 40 or text.endswith(("。", "；", "，")):
        return None
    if H1.match(text):
        return 1
    if H2.match(text) or NUM2.match(text):
        return 2
    if NUM3.match(text) or (H3.match(text) and len(text) <= 30):
        return 3
    return None


def guess_roles(blocks):
    """Best-effort role for every top-level paragraph; the agent must review it.

    Paragraphs that already use this skill's styles (an earlier output) are
    mapped exactly; other documents are guessed from text patterns.
    """
    from customization import ROLE_STYLES

    known = {name: role for role, name in ROLE_STYLES.items() if role != "cover"}
    paragraphs = [b for b in blocks if b["kind"] == "paragraph"]
    cover_end = None
    for b in blocks[:60]:
        if b["kind"] == "paragraph" and b["section_break"]:
            before = "".join(x["text"] for x in blocks if x["block_index"] <= b["block_index"])
            if any(word in before for word in COVER_WORDS):
                cover_end = b["block_index"] + 1
            break
    roles, area, seen_heading = {}, None, False
    for b in paragraphs:
        text = b["text"].strip()
        plain = _plain(text)
        style = (b.get("style") or "").lower()
        if cover_end is not None and b["block_index"] < cover_end:
            role = "preserve"
        elif b.get("style") in known:
            role = known[b["style"]]
            seen_heading = seen_heading or role in ("h1", "h2", "h3", "h4", "preface")
        elif not text and b["images"]:
            role = "figure"
        elif not text and b["equations"]:
            role = "equation"
        elif not text:
            role = "spacer"
        elif plain in ("目录", "目錄"):
            role, area = "toc_title", "toc"
        elif style.startswith("toc") or (area == "toc" and re.search(r"\d+\s*$", text) and not seen_heading):
            role = "preserve"
        elif plain in ("摘要", "内容摘要", "中文摘要"):
            role, area = "abstract_label_zh", "abstract_zh"
        elif re.match(r"^关键词\s*[:：]", plain):
            role, area = "keywords_zh", None
        elif plain.lower() == "abstract":
            role, area = "abstract_label_en", "abstract_en"
        elif re.match(r"^key\s*words?\s*[:：]", text, re.I):
            role, area = "keywords_en", None
        elif plain in ("参考文献", "主要参考文献"):
            role, area = "reference_title", "references"
        elif area == "references":
            role = "reference"
        elif area in ("abstract_zh", "abstract_en"):
            role = area
        elif plain in ("序言", "引言", "前言"):
            role, area, seen_heading = "preface", None, True
        elif re.match(r"^(图|表|Figure|Table)\s*\d", text) and len(text) <= 60:
            role = "caption"
        elif _heading_level(b):
            role, area, seen_heading = "h" + str(_heading_level(b)), None, True
        elif (
            not seen_heading
            and area is None
            and len(text) <= 200
            and re.search(r"[A-Za-z]", text)
            and not re.search(r"[一-鿿]", text)
            and not text.endswith(".")
            and not roles_has(roles, "title_en")
        ):
            role = "title_en"
        elif not seen_heading and area is None and len(text) <= 40 and not text.endswith("。") and not roles_has(roles, "title_zh"):
            role = "title_zh"
        else:
            role, area = "body", (None if area == "toc" else area)
        roles[b["id"]] = role
    return roles, cover_end


def roles_has(roles, role):
    return role in roles.values()


def cmd_inspect(args):
    import delivery as D
    import engine as E
    import intake as I

    root, s = _work(args.work)
    source = Path(args.docx).expanduser().resolve()
    if source.suffix.lower() != ".docx":
        raise ValueError("只能读取 .docx。旧版 .doc 请先用 Word/WPS“另存为 .docx”。")
    state = _intake(root)
    inv = E.inventory(source)
    roles, cover_end = guess_roles(inv["blocks"])
    wants_cover = state.get("include_cover", True)
    own_output = any((b.get("style") or "").startswith("XAUFE ") for b in inv["blocks"])
    values = set(roles.values())
    # A full paper (abstract + headings) gets the template's structure: its own TOC,
    # abstract pages and page numbering. An earlier output of this skill already has it.
    paper = (
        not own_output
        and "abstract_zh" in values
        and bool(values & {"h1", "h2", "h3", "preface"})
        and not any(b.get("section_break") for b in inv["blocks"][(cover_end or 0) :] if b["kind"] == "paragraph")
    )
    if paper:
        for b in inv["blocks"]:
            if b["kind"] == "paragraph" and (roles.get(b["id"]) == "toc_title" or (
                roles.get(b["id"]) == "preserve" and (cover_end is None or b["block_index"] >= cover_end)
            )):
                roles[b["id"]] = "delete"
    mapping = {
        "source": str(source),
        "input_sha256": inv["input_sha256"],
        "mode": "paper" if paper else "document",
        "format_tables": True,
        "replace_cover_before": (cover_end if cover_end is not None else 0) if wants_cover else None,
        "paragraphs": [
            {"id": b["id"], "role": roles[b["id"]], "text": b["text"][:60]}
            for b in inv["blocks"]
            if b["kind"] == "paragraph"
        ],
    }
    _write(root / "mapping.json", mapping)
    if not any(x["path"] == str(source) for x in s["source_documents"]):
        s["source_documents"].append(
            {"path": str(source), "sha256": D.sha(source), "role": "original"}
        )
        D.write_json(root / D.MARKER, s)
    counts = {}
    for role in roles.values():
        counts[role] = counts.get(role, 0) + 1
    print(f"已生成 {root / 'mapping.json'}，共 {len(roles)} 段。自动判断结果：")
    print("  " + "，".join(f"{k} {v} 段" for k, v in sorted(counts.items())))
    if cover_end is not None and wants_cover:
        print(f"  检测到原稿开头有旧封面（到第 {cover_end} 个块为止），会换成所选模板的封面。")
    elif cover_end is not None:
        print("  注意：原稿开头有封面，但用户不要封面。先和用户确认；脚本不会自动删除封面。")
    elif wants_cover:
        print("  原稿没有封面，会在最前面加上所选模板的封面。")
    if paper:
        print("  原稿有摘要和标题：按完整论文结构排（mode=paper），会自动生成目录、摘要单独成页。")
        if "delete" in values or "delete" in roles.values():
            print("  原稿里的旧目录已标为 delete，会换成自动生成的目录。")
        from customization import profile_for

        if profile_for(state)["english_abstract"] and "abstract_en" not in values:
            print(
                "  注意：所选版式需要英文摘要，原稿没有。先问用户：要不要你把中文摘要翻译成英文补上"
                "（写进 mapping.json 的 add，见 references/content.md），还是这次不要英文摘要"
                "（job.json 的 format 写 \"english_abstract\": false）。"
            )
    elif not own_output:
        print("  原稿没有摘要或标题，按普通文档排（mode=document），不加目录。")
    print("→ 打开 mapping.json，逐段核对 role（尤其是标题级别和正文），改错的项；然后运行 run。")
    print("  可用 role 见 references/content.md 的“已有 Word”一节。")
    return 0


# ---------------------------------------------------------------- run
def _resolve_mapping(root):
    mapping = _read(root / "mapping.json", "mapping.json")
    rows = mapping.get("paragraphs")
    if not isinstance(rows, list):
        raise ValueError("mapping.json 缺少 paragraphs 列表；请重新运行 inspect。")
    roles = {}
    for row in rows:
        if not isinstance(row, dict) or not row.get("id") or not row.get("role"):
            raise ValueError("mapping.json 的每一段都要有 id 和 role。")
        roles[row["id"]] = row["role"]
    resolved = {
        "input_sha256": mapping.get("input_sha256"),
        "mode": mapping.get("mode", "document"),
        "format_tables": bool(mapping.get("format_tables", True)),
        "has_cover": False,
        "roles": roles,
    }
    if mapping.get("replace_cover_before") is not None:
        resolved["replace_cover_before"] = mapping["replace_cover_before"]
    if mapping.get("add"):
        resolved["add"] = mapping["add"]
    _write(root / "mapping.resolved.json", resolved)
    return Path(mapping["source"]), root / "mapping.resolved.json"


def _build(root, level):
    import engine as E

    content = root / "content.json"
    mapping = root / "mapping.json"
    draft = root / "draft.docx"
    if content.is_file() and mapping.is_file():
        raise ValueError("WORK 里同时有 content.json 和 mapping.json，只能保留一个：新写文档用 content.json，已有 Word 用 mapping.json。")
    if content.is_file():
        E.build(content, draft, force=True, intake_path=root / "intake.json", cover_min_level=level)
    elif mapping.is_file():
        source, resolved = _resolve_mapping(root)
        result = E.restyle(
            source, resolved, draft, force=True, intake_path=root / "intake.json", cover_min_level=level
        )
        if result.get("deleted_paragraphs") and level == 0:
            print(f"按 mapping 删除了 {len(result['deleted_paragraphs'])} 段（role 为 delete）：")
            for text in result["deleted_paragraphs"]:
                print("  - " + (text or "（空段）"))
    else:
        raise ValueError("还没有内容：新写文档请写 content.json；已有 Word 请先运行 inspect。")
    return draft


def _body_length(root):
    """Characters of body paragraphs in content.json (CJK characters plus Latin words, no punctuation)."""
    content = root / "content.json"
    if not content.is_file():
        return None
    data = _read(content, "content.json")
    total = 0
    for block in data.get("blocks", []):
        if block.get("type") != "paragraph":
            continue
        text = block.get("text") or "".join(r.get("text", "") for r in block.get("runs", []))
        total += len(re.findall(r"[\u4e00-\u9fff]", text)) + len(re.findall(r"[A-Za-z0-9]+", text))
    return total, len(data.get("references") or [])


def _review_template(result):
    warnings = {w["code"]: "请填写：这条警告怎么处理的——" + w["message"][:80] for w in result["warnings"]}
    static = result["validation_scope"] == "static-only"
    return {
        "vision": not static,
        "pages": {
            str(p["page"]): f"请填写：第{p['page']}页实际看到的情况（有无缺字、重叠、裁切、空白页，日期/目录/页眉是否正常）"
            for p in result["pages"]
        },
        "content_check": "请填写：需求与内容保全——对照用户要求和原文，正文、图表、公式、链接是否齐全，封面信息和留空是否正确",
        "layout_check": "请填写：格式与版面——字体字号、标题层级、目录与页码、分页、表格图片是否正常",
        "delivery_check": "请填写：最终交付——发现的问题是否都已改好，交付的是不是用户要的这一个 Word",
        "warnings": warnings,
    }


def cmd_run(args):
    import docx
    import delivery as D
    import engine as E
    import intake as I
    from cover_date_audit import CoverOverflow
    from cover_fit import MAX_LEVEL

    root, _ = _work(args.work)
    state = _intake(root)
    if state["stage"] != "ready":
        print(I.summary(state))
        print("\n→ 封面信息还不全。先问用户（可选“留空自己填”），写进 job.json 后重新运行 run。")
        return EXIT_ASK_USER
    if _task(root) == "write" and (root / "content.json").is_file():
        problems = _writing_problems(root, _read(root / "content.json", "content.json"))
        if problems:
            raise ValueError(
                "写作检查没通过（job.json 的 task 是 write）：\n  - "
                + "\n  - ".join(problems[:15])
                + "\n改好 content.json / refs.json 后重新 run。不要为了通过检查编造数字或出处。"
            )
    if args.office in ("word", "libreoffice", "wps", "wps-manual"):
        engine = "wps" if args.office.startswith("wps") else args.office
        D.configure_office(root, engine, "manual" if args.office == "wps-manual" else "auto")
    level = 0
    while True:
        draft = _build(root, level)
        try:
            result = D.prepare(
                root, draft, root / "intake.json", args.office_note, "none" if args.office == "none" else "auto"
            )
            break
        except CoverOverflow as exc:
            used = int(E.get_vars(docx.Document(draft)).get("XAUFE_COVER_FIT", "0"))
            if used >= MAX_LEVEL:
                raise ValueError(
                    "封面已经收紧到最紧，第一页仍放不下（"
                    + str(exc)
                    + "）。请问用户：缩短题目、课程名或过长的封面项，或同意缩小封面字号。"
                ) from exc
            level = used + 1
            print(f"封面太长，日期被挤出第一页；自动收紧封面空白（第 {level} 级）后重试……")
    _write(root / "review.json", _review_template(result))
    print(I.summary(state).splitlines()[0])
    stats = _body_length(root)
    if stats:
        print(
            f"正文约 {stats[0]} 字（不含标点、标题、摘要、参考文献），参考文献 {stats[1]} 条。"
            "写作时用户没提要求的，默认要正文不少于 5000 字、参考文献至少 6 篇；只排版时不用管。"
        )
    if result["validation_scope"] == "static-only":
        print("\n排版完成，但没有做实际分页和页面预览（原因见下）。")
        print("需要在最终回复里告诉用户：" + result["user_notice"])
    else:
        print(f"\n排版完成，共 {len(result['pages'])} 页。请逐页打开下面的图片查看：")
        for page in result["pages"]:
            print(f"  第{page['page']}页：{page['path']}")
    if result["warnings"]:
        print("\n需要处理的提示：")
        for w in result["warnings"]:
            print(f"  - {w['code']}：{w['message']}")
    review = root / "review.json"
    if result["validation_scope"] == "static-only":
        print(f"\n→ 下一步：这次没有页图。核对内容和格式后，把 {review} 里所有“请填写”换成你的实际检查结果，然后运行 finish。")
    else:
        print(f"\n→ 下一步：看完每一页后，把 {review} 里所有“请填写”换成你的实际检查结果，然后运行 finish。")
    print("  发现问题就改 job.json / content.json / mapping.json 后重新 run，不要直接改 WORK 里的 Word。")
    return 0


# ---------------------------------------------------------------- finish / cancel
def cmd_finish(args):
    import delivery as D

    root, _ = _work(args.work)
    D.attest(root, _read(root / "review.json", "review.json"))
    checklist = _reference_checklist(root) if _task(root) == "write" else None
    result = D.publish(root)
    print("已交付：" + result["output"])
    print("工作区已清理。")
    if result["user_notice"]:
        print("\n必须在给用户的回复里说明：" + result["user_notice"])
    else:
        print("\n已完成排版、分页和逐页检查。")
    if checklist:
        print("\n必须在给用户的回复里附上这张参考文献核实清单（原样照抄）：\n" + checklist)
    return 0


def cmd_refs(args):
    import refcheck as R

    root, _ = _work(args.work)
    entries = _read(root / "refs.json", "refs.json")
    if not isinstance(entries, list):
        raise ValueError("refs.json 应该是一个列表：[{...}, {...}]，写法见 references/citations.md")
    results = R.check_all(entries)
    for result, entry in zip(results, entries):
        if isinstance(entry, dict):
            result["supports"] = entry.get("supports", "")
            result["relevance"] = entry.get("relevance", "")
    _write(
        root / "refs-result.json",
        {"refs_sha256": hashlib.sha256((root / "refs.json").read_bytes()).hexdigest(), "min": args.min, "results": results},
    )
    print(R.report(results, entries, args.min))
    return 0


def cmd_read(args):
    """Print the text of a .docx so the agent can read the user's file."""
    import docx
    from docx.oxml.ns import qn

    source = Path(args.docx).expanduser().resolve()
    if source.suffix.lower() != ".docx":
        raise ValueError("只能读取 .docx。旧版 .doc 请先用 Word/WPS“另存为 .docx”。")
    doc = docx.Document(source)
    body = doc.element.body
    print(f"《{source.name}》的文字内容（按顺序）：")
    for element in body.iterchildren():
        if element.tag == qn("w:p"):
            text = "".join(element.xpath(".//w:t/text()")).strip()
            if text:
                print(text)
        elif element.tag == qn("w:tbl"):
            for row in element.xpath("./w:tr"):
                cells = ["".join(c.xpath(".//w:t/text()")).strip() for c in row.xpath("./w:tc")]
                print("[表格] " + " | ".join(cells))
    images = len(body.xpath(".//w:drawing"))
    if images:
        print(f"[另有 {images} 张图片]")
    return 0


def cmd_preview(args):
    """Page images of any .docx (e.g. one edited by hand), kept inside WORK."""
    import shutil
    import pymupdf as fitz
    from office_backend import export_pdf, OfficeUnavailable

    root, _ = _work(args.work)
    source = Path(args.docx).expanduser().resolve()
    out = root / "extra-preview" / source.stem
    if out.exists():
        shutil.rmtree(out)
    out.mkdir(parents=True)
    pdf = out / "preview.pdf"
    try:
        export_pdf(source, pdf)
    except OfficeUnavailable as exc:
        raise ValueError("没有可用的办公软件，生不成页图：" + str(exc)) from exc
    with fitz.open(pdf) as rendered:
        print(f"共 {len(rendered)} 页，请逐页打开查看：")
        for index, page in enumerate(rendered, 1):
            image = out / f"page-{index:03d}.png"
            page.get_pixmap(dpi=110).save(image)
            print(f"  第{index}页：{image}")
    try:
        from cover_date_audit import verify_pdf

        verify_pdf(source, pdf)
    except Exception as exc:
        print("注意（封面检查）：" + str(exc))
    return 0


def cmd_cancel(args):
    import delivery as D

    D.clean(args.work)
    print("已删除本次工作区，没有交付文件。")
    return 0


def main(argv=None):
    parser = argparse.ArgumentParser(description="西安财经大学文档格式：唯一入口")
    sub = parser.add_subparsers(dest="cmd", required=True)
    sub.add_parser("compare", help="打印 A/B 对比表（用户还没选版式时发给用户）")
    a = sub.add_parser("new", help="开始一个任务，创建工作区 WORK")
    a.add_argument("--out", required=True, help="最终 Word 的保存路径（.docx，不能已存在）")
    a.add_argument("--source", action="append", default=[], help="用户提供的原稿文件，可写多次")
    for name, text in (
        ("check", "检查 job.json，告诉你还缺什么"),
        ("run", "排版、分页、生成页图"),
        ("finish", "读取 review.json，交付 Word 并清理"),
        ("cancel", "放弃任务并清理"),
    ):
        a = sub.add_parser(name, help=text)
        a.add_argument("--work", required=True)
        if name == "run":
            a.add_argument(
                "--office",
                default="auto",
                choices=["auto", "word", "libreoffice", "wps", "wps-manual", "none"],
                help="none = 用户要求不启动办公软件",
            )
            a.add_argument("--office-note", help="办公软件找到了却无法调用时，写实际原因")
    a = sub.add_parser("read", help="打印用户 Word 文件里的文字（看原稿内容用）")
    a.add_argument("--docx", required=True)
    a = sub.add_parser("preview", help="给任意 .docx 生成页图（手工处理过的文档用）")
    a.add_argument("--work", required=True)
    a.add_argument("--docx", required=True)
    a = sub.add_parser("inspect", help="读取已有 Word，生成可核对的 mapping.json")
    a.add_argument("--work", required=True)
    a.add_argument("--docx", required=True)
    a = sub.add_parser("refs", help="核对 WORK/refs.json 里的参考文献（需要联网）")
    a.add_argument("--work", required=True)
    a.add_argument("--min", type=int, default=6, help="至少几篇；用户另有要求时填用户的数")
    args = parser.parse_args(argv)
    handler = {
        "compare": cmd_compare,
        "new": cmd_new,
        "check": cmd_check,
        "inspect": cmd_inspect,
        "preview": cmd_preview,
        "read": cmd_read,
        "refs": cmd_refs,
        "run": cmd_run,
        "finish": cmd_finish,
        "cancel": cmd_cancel,
    }[args.cmd]
    try:
        return handler(args) or 0
    except Exception as exc:  # every failure is reported as one readable line
        from wps_exchange import ManualExportRequired

        if isinstance(exc, ManualExportRequired):
            print("需要在 WPS 里手动导出 PDF，请求文件：" + str(exc.request), file=sys.stderr)
            print("做法见 references/special-cases.md 的 WPS 一节。", file=sys.stderr)
            return 4
        print("出错了：" + str(exc), file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
