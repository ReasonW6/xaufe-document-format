#!/usr/bin/env python3
"""Task record: chosen layout -> cover information -> ready.

xaufe.py rebuilds this record from the agent-written job.json on every run, so the
record always matches the latest user requirements. It never alters documents.
"""

from __future__ import annotations
import hashlib, json, uuid, os, tempfile
from datetime import datetime, date
from pathlib import Path
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError
from profiles import load_profile, TEMPLATES, LOGOS
from customization import profile_for, effective_profile, merge_request

VERSION = "1.0.0"
LABELS = {
    "course_name": "课程名称（B版封面第一行）",
    "title_zh": "论文／作业题目",
    "student_name": "学生姓名",
    "student_id": "学号",
    "major": "专业",
    "class_name": "班级",
    "supervisor": "指导教师",
    "date": "完成日期",
}
SOURCES = {"user": 3, "document": 2, "memory": 1}


def load(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def digest(value):
    return hashlib.sha256(
        json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode()
    ).hexdigest()


def write(path, state):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    state["version"] = VERSION
    state["integrity"] = digest({k: v for k, v in state.items() if k != "integrity"})
    fd, tmp = tempfile.mkstemp(prefix=path.name + ".", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(json.dumps(state, ensure_ascii=False, indent=2))
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.unlink(tmp)
    return state


def read(path):
    state = load(path)
    if state.get("version") != VERSION:
        raise ValueError(
            "交互记录格式不匹配，请重建本次记录并复用真实上下文的已确认要求；不要让用户重复选择或重新提供已知信息。"
        )
    if state.get("integrity") != digest({k: v for k, v in state.items() if k != "integrity"}):
        raise ValueError("任务记录被直接改动；请改 job.json 后重新运行 check 或 run。")
    return state


def task_timezone(name):
    return datetime.now().astimezone().tzinfo if name == "local" else ZoneInfo(name)


def start(path, job, timezone="Asia/Shanghai", job_id=None):
    if Path(path).exists():
        raise ValueError("交互记录已经存在；继续当前记录，或为新任务指定新路径。")
    if not str(job).strip():
        raise ValueError("必须描述本次任务，以免沿用其他人的资料。")
    try:
        task_timezone(timezone)
    except ZoneInfoNotFoundError as exc:
        raise ValueError("时区无效或缺少 tzdata；安装 requirements.txt 后重试。") from exc
    return write(
        path,
        {
            "version": VERSION,
            "job_id": job_id or str(uuid.uuid4()),
            "job": job,
            "stage": "awaiting_template",
            "timezone": timezone,
            "template": None,
            "selection_evidence": "",
            "candidates": {},
            "blank_fields": [],
            "metadata": {},
            "provenance": {},
            "missing": [],
            "conflicts": {},
            "events": [{"action": "task_started"}],
        },
    )


def select(path, template, user_choice):
    state = read(path)
    load_profile(template)
    if state["stage"] != "awaiting_template":
        raise ValueError(
            "此记录已有基础版式。用户明确换版时，在同一工作区新建intake记录，复用仍适用的已确认信息与格式要求，只补问新增缺项；不要再让用户确认已说清楚的换版。"
        )
    if not str(user_choice).strip():
        raise ValueError("缺少用户本次实际选择原话。不能由模型替用户选择。")
    state.update(
        {"template": template, "selection_evidence": user_choice, "stage": "awaiting_metadata"}
    )
    state["events"].append(
        {"action": "user_template_choice", "template": template, "evidence": user_choice}
    )
    return write(path, state)


def required_fields(state):
    """Question only fields actually used by the requested document scope."""
    spec = profile_for(state)
    fields = list(spec["cover"]["required_fields"])
    if state.get("include_cover", True):
        return fields
    # No cover does not implicitly remove an enabled header. Its default text
    # still uses title/major; explicit header strings need no personal fields.
    needed = []
    if spec["header"]["enabled"]:
        if "right" not in spec["header"]:
            needed.append("title_zh")
        if "left" not in spec["header"]:
            needed.append("major")
    return [key for key in fields if key in needed]


def document_scope(path, include_cover, user_request):
    state = read(path)
    if state["stage"] == "awaiting_template":
        raise ValueError("先记录基础版式，再记录是否使用封面。已有明确选择不必重复询问。")
    if not isinstance(include_cover, bool):
        raise ValueError("是否使用封面须为布尔值。")
    if not isinstance(user_request, str) or not user_request.strip():
        raise ValueError("记录用户是否使用封面的真实要求，不能为了跳过问题自行取消封面。")
    state["include_cover"] = include_cover
    state["events"].append(
        {"action": "document_scope", "include_cover": include_cover, "evidence": user_request}
    )
    write(path, state)
    return resolve(path, {})


def normalize_date(value):
    import re

    value = str(value).strip()
    m = re.fullmatch(r"(\d{4})\s*(?:年|[-/])\s*(\d{1,2})\s*(?:月|[-/])\s*(\d{1,2})\s*日?", value)
    if not m:
        raise ValueError("日期应为 YYYY-MM-DD 或 YYYY年M月D日，不猜测含糊日期。")
    return date(*map(int, m.groups())).isoformat()


def resolve(path, answers, now=None):
    state = read(path)
    if state["stage"] == "awaiting_template":
        raise ValueError("尚未由用户选定模板；不要先收集封面信息或处理文档。")
    spec = profile_for(state)
    fields = spec["cover"]["required_fields"]
    needed = set(required_fields(state))
    if not isinstance(answers, dict):
        raise ValueError("答案必须是JSON对象。")
    allowed = {
        "metadata",
        "blank_fields",
        "blank_all",
        "blank_instruction",
        "unblank_fields",
        "unblank_instruction",
    }
    for key in ("blank_fields", "unblank_fields"):
        val = answers.get(key, [])
        if not isinstance(val, list) or any(not isinstance(v, str) for v in val):
            raise ValueError(key + "必须是字段名数组。")
    if "blank_all" in answers and not isinstance(answers["blank_all"], bool):
        raise ValueError("blank_all必须为布尔值。")
    if set(answers) - allowed:
        raise ValueError("未知封面答案字段：" + ",".join(set(answers) - allowed))
    supplied = answers.get("metadata", {})
    if not isinstance(supplied, dict):
        raise ValueError("metadata 必须按字段提供可追溯信息。")
    for key, items in supplied.items():
        if key not in fields:
            raise ValueError(
                f"{key} 不在当前有效封面字段中；新增指导教师用 format 里的 cover_supervisor。"
            )
        if not isinstance(items, list):
            items = [items]
        if not items:
            raise ValueError("不能用空候选数组清除已知值；使用明确留空或更正。")
        clean = []
        for item in items:
            if not isinstance(item, dict) or set(item) - {
                "value",
                "source",
                "evidence",
                "target_confirmed",
                "corrects_previous",
            }:
                raise ValueError(
                    f"{key} 必须提供 value/source/evidence；记忆还需 target_confirmed。"
                )
            if item.get("source") not in SOURCES or not str(item.get("evidence", "")).strip():
                raise ValueError(f"{key} 来源不完整。模板占位、示例或推测不是有效来源。")
            if item["source"] == "memory" and item.get("target_confirmed") is not True:
                raise ValueError(f"{key} 的记忆未确认属于本次文档对象；不能套用别人的姓名学号。")
            if not isinstance(item.get("value"), str) or not item["value"].strip():
                raise ValueError(f"{key} 不能用空字符串伪装已知信息；明确留空请使用 blank_fields。")
            if "corrects_previous" in item and not isinstance(item["corrects_previous"], bool):
                raise ValueError("corrects_previous必须为布尔值。")
            if item.get("corrects_previous") and item["source"] != "user":
                raise ValueError("只有用户明确更正可撤销以前的用户值。")
            val = item["value"].strip()
            if key == "date":
                val = normalize_date(val)
            clean.append(dict(item, value=val))
        previous = list(state["candidates"].get(key, []))
        corrections = [x for x in clean if x.get("corrects_previous")]
        if corrections:
            if (
                len(corrections) != 1
                or len({x["value"] for x in clean if x["source"] == "user"}) != 1
            ):
                raise ValueError("同一字段的用户更正必须唯一，不能同时记录矛盾更正。")
            archived = [x for x in previous if x["source"] == "user"]
            previous = [x for x in previous if x["source"] != "user"]
            state["events"].append(
                {
                    "action": "user_value_corrected",
                    "field": key,
                    "previous_user_candidates": archived,
                    "replacement": corrections[0],
                }
            )
        for item in clean:
            if item not in previous:
                previous.append(item)
        state["candidates"][key] = previous
        state["events"].append({"action": "candidates_added", "field": key, "candidates": clean})
    blanks = set(state["blank_fields"])
    newblank = answers.get("blank_fields", [])
    if answers.get("blank_all"):
        newblank = fields
    if newblank:
        if not str(answers.get("blank_instruction", "")).strip():
            raise ValueError("留空必须记录用户实际指令，不能自行把未知字段全部留空后跳过询问。")
        if set(newblank) - set(fields):
            raise ValueError("留空字段不属于当前模板。")
        blanks.update(newblank)
        state["events"].append(
            {
                "action": "user_blank_choice",
                "fields": list(newblank),
                "evidence": answers["blank_instruction"],
            }
        )
    for k in answers.get("unblank_fields", []):
        if k not in fields:
            raise ValueError("unblank_fields 包含未知字段。")
        items = supplied.get(k, [])
        if isinstance(items, dict):
            items = [items]
        if (
            not any(x.get("source") == "user" for x in items)
            and not str(answers.get("unblank_instruction", "")).strip()
        ):
            raise ValueError("取消留空需用户明确指令，或本轮该字段的用户明确值。")
        blanks.discard(k)
        state["events"].append(
            {
                "action": "user_unblank_choice",
                "field": k,
                "evidence": answers.get("unblank_instruction")
                or [x["evidence"] for x in items if x.get("source") == "user"],
            }
        )
    metadata = {}
    provenance = {}
    missing = []
    conflicts = {}
    for key in fields:
        candidates = state["candidates"].get(key, [])
        if candidates:
            rank = max(SOURCES[x["source"]] for x in candidates)
            best = [x for x in candidates if SOURCES[x["source"]] == rank]
            vals = {x["value"] for x in best}
            if len(vals) == 1:
                metadata[key] = best[0]["value"]
                provenance[key] = best[0]
            elif key not in blanks and key in needed:
                conflicts[key] = sorted(vals)
        # Explicit blank always wins on cover; known title may still be used in abstract/header.
        if key in blanks:
            provenance["cover:" + key] = {"source": "user_blank"}
            continue
        if key in conflicts:
            continue
        if key not in metadata and key in needed:
            if key == "date":
                clock = now or datetime.now(task_timezone(state["timezone"]))
                if clock.tzinfo is None:
                    raise ValueError("内部时钟必须带时区，不能默认为 UTC。")
                value = clock.astimezone(task_timezone(state["timezone"])).date().isoformat()
                metadata[key] = value
                provenance[key] = {
                    "value": value,
                    "source": "runtime_today",
                    "timezone": state["timezone"],
                }
            else:
                missing.append(key)
    state.update(
        {
            "blank_fields": sorted(blanks),
            "metadata": metadata,
            "provenance": provenance,
            "missing": missing,
            "conflicts": conflicts,
            "stage": "ready" if not missing and not conflicts else "awaiting_metadata",
        }
    )
    state["events"].append(
        {
            "action": "cover_resolved",
            "missing": missing,
            "conflicts": conflicts,
            "date_source": provenance.get("date", {}).get("source"),
        }
    )
    return write(path, state)


def require(path, ready=True, now=None):
    if not path:
        raise ValueError(
            "缺少 --intake。需先记录真实的基础版式与本次要求；首次展示默认差异，已给的选择和信息直接复用，不重复提问。"
        )
    state = read(path)
    if state["stage"] == "awaiting_template" or not state.get("selection_evidence"):
        raise ValueError("模板尚未由用户选择；禁止开始文档处理。")
    if ready and state["stage"] != "ready":
        raise ValueError(
            "封面信息尚未解决："
            + ",".join(state.get("missing", []) + list(state.get("conflicts", {})))
        )
    # If the task spans midnight, an automatically chosen date tracks the processing day.
    if ready and state["provenance"].get("date", {}).get("source") == "runtime_today":
        clock = now or datetime.now(task_timezone(state["timezone"]))
        if clock.tzinfo is None:
            raise ValueError("内部时钟必须带时区。")
        today = clock.astimezone(task_timezone(state["timezone"])).date().isoformat()
        if state["metadata"].get("date") != today:
            state["metadata"]["date"] = today
            state["provenance"]["date"]["value"] = today
            state["events"].append({"action": "runtime_date_refreshed", "date": today})
            write(path, state)
    return state


def summary(state):
    """Short user-facing status: chosen layout/logo and what is still missing."""
    if state["stage"] == "awaiting_template":
        return "还没有记录用户选择的版式（A 或 B）。"
    spec = profile_for(state)
    display = {"A": "A（本科学年论文）", "B": "B（课程期末大作业）"}[state["template"]]
    color = {"green": "绿色", "red": "红色"}[spec["logo"]["color"]]
    lines = (
        [f"采用{display}，配{color}校徽。"]
        if state.get("include_cover", True)
        else [f"采用{display}的文档格式，本次不含封面。"]
    )
    conflicts = state.get("conflicts", {})
    blanks = set(state.get("blank_fields", []))
    missing = [
        k
        for k in required_fields(state)
        if k != "date"
        and k not in blanks
        and k not in conflicts
        and k not in state.get("metadata", {})
    ]
    if conflicts:
        lines.append(
            "需要核对："
            + "；".join(
                LABELS[k] + "（" + "／".join(map(str, values)) + "）"
                for k, values in conflicts.items()
            )
            + "。"
        )
    if missing:
        lines.append(
            "还缺：" + "、".join(LABELS[k] for k in missing) + "。这些项也可以留空，之后自己填写。"
        )
    if state["stage"] == "ready":
        used_blanks = [k for k in required_fields(state) if k in blanks]
        if used_blanks:
            lines.append("按要求留空：" + "、".join(LABELS[k] for k in used_blanks) + "。")
        if (
            state.get("include_cover", True)
            and "date" not in blanks
            and state.get("metadata", {}).get("date")
        ):
            date_value = state["metadata"]["date"]
            origin = state.get("provenance", {}).get("date", {}).get("source")
            lines.append(
                "日期采用"
                + ("执行当日：" if origin == "runtime_today" else "已提供的值：")
                + date_value
                + "。"
            )
        lines.append("信息已足够，直接继续排版，不需要再确认。")
    return "\n".join(lines)


def customize(path, request, user_request, *, replace=False, reset=False):
    state = read(path)
    if state["stage"] == "awaiting_template":
        raise ValueError("先记录基础版式选择；特别要求不等于另选模板。")
    if not isinstance(user_request, str) or not user_request.strip():
        raise ValueError("记录本次用户的实际特别要求，不另索强烈授权或二次确认。")
    if not isinstance(request, dict):
        raise ValueError("格式差异必须是JSON对象。")
    if reset and (replace or request):
        raise ValueError("恢复全部默认必须使用空请求且不与replace合用。")
    previous = state.get("format_request", {})
    patch = request
    request = {} if reset else (request if replace else merge_request(previous, request))
    effective_profile(state["template"], request)
    state["format_request"] = request
    state["format_request_evidence"] = user_request
    state["events"].append(
        {
            "action": "user_format_override",
            "request": request,
            "patch": patch,
            "previous": previous,
            "mode": "reset" if reset else "replace" if replace else "merge",
            "evidence": user_request,
        }
    )
    write(path, state)
    # Existing candidates are preserved; only genuinely new cover fields become missing.
    return resolve(path, {})


# "task" (write / format) is read by xaufe.py, not by the cover workflow.
JOB_KEYS = {"task", "template", "logo", "cover", "info", "blank", "format", "timezone"}
JOB_EVIDENCE = "job.json（执行者按用户在对话中的原话填写）"


def _drop_none(value):
    """null in job.json simply means "use the default" (same as leaving it out)."""
    if isinstance(value, dict):
        return {k: _drop_none(v) for k, v in value.items() if v is not None}
    return value


def from_job(job, path, now=None):
    """Build a fresh task record from job.json. Returns the resulting state.

    job.json is the single place where the agent writes the user's choices, so
    later rounds only edit the changed item and rerun.
    """
    if not isinstance(job, dict):
        raise ValueError("job.json 必须是一个 JSON 对象。")
    unknown = set(job) - JOB_KEYS
    if unknown:
        raise ValueError(
            "job.json 有不认识的项："
            + "、".join(sorted(unknown))
            + "。只能用 task、template、logo、cover、info、blank、format、timezone。"
        )
    template = job.get("template")
    if template not in TEMPLATES:
        raise ValueError('job.json 的 template 必须是 "A" 或 "B"。用户还没选就先问用户，不要替用户选。')
    logo = job.get("logo")
    if logo is not None and logo not in LOGOS:
        raise ValueError('job.json 的 logo 只能是 "red" 或 "green"；不写则 A 用绿色、B 用红色。')
    request = job.get("format", {})
    if not isinstance(request, dict):
        raise ValueError("job.json 的 format 必须是 JSON 对象（没有特别要求就写 {}）。")
    request = _drop_none(request)
    if "logo" in request:
        raise ValueError("校徽颜色写在 job.json 顶层的 logo，不要写进 format。")
    if logo is not None:
        request["logo"] = logo
    cover = job.get("cover", True)
    if not isinstance(cover, bool):
        raise ValueError("job.json 的 cover 只能是 true（要封面）或 false（不要封面）。")
    info = job.get("info", {})
    if not isinstance(info, dict) or any(not isinstance(v, str) for v in info.values()):
        raise ValueError("job.json 的 info 必须是“字段: 文字”的对象，值都用字符串。")
    blank = job.get("blank", [])
    if not isinstance(blank, list) or any(not isinstance(v, str) for v in blank):
        raise ValueError("job.json 的 blank 必须是字段名列表，例如 [\"student_id\"]。")
    timezone = job.get("timezone", "Asia/Shanghai")
    path = Path(path)
    if path.exists():
        path.unlink()
    # Same workspace -> same task id, so rebuilding identical input gives identical output.
    start(path, "job.json", timezone, job_id=str(uuid.uuid5(uuid.NAMESPACE_URL, str(path.resolve()))))
    select(path, template, JOB_EVIDENCE)
    if request:
        customize(path, request, JOB_EVIDENCE, replace=True)
    if not cover:
        document_scope(path, False, JOB_EVIDENCE)
    fields = profile_for(read(path))["cover"]["required_fields"]
    labels = "、".join(f"{k}（{LABELS[k]}）" for k in fields)
    for key in list(info) + blank:
        if key not in fields:
            hint = ""
            if key == "supervisor":
                hint = "B 版默认没有指导教师栏；用户要加就在 format 里写 \"cover_supervisor\": true。"
            raise ValueError(f"{key} 不是当前版式的封面字段。{hint}当前可用字段：{labels}。")
    both = [k for k in blank if info.get(k, "").strip()]
    if both:
        raise ValueError("同一项不能既填写又留空：" + "、".join(both))
    values = {k: v.strip() for k, v in info.items() if v.strip()}
    answers = {
        "metadata": {
            k: {"value": v, "source": "user", "evidence": JOB_EVIDENCE} for k, v in values.items()
        }
    }
    if blank:
        answers["blank_fields"] = blank
        answers["blank_instruction"] = JOB_EVIDENCE
    resolve(path, answers, now=now)
    return read(path)
