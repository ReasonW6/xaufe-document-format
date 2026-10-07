"""Read the bundled layouts (A/B) and school logos (red/green).

Layout and logo are independent: a layout only names its default logo colour.
There is no default layout; the user's choice must be recorded first.
"""

from __future__ import annotations
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TEMPLATES = ("A", "B")
LOGOS = ("red", "green")


def registry() -> dict:
    return json.loads((ROOT / "assets/templates.json").read_text(encoding="utf-8"))


def logo_info(color: str) -> dict:
    logos = registry()["logos"]
    if color not in logos:
        raise ValueError("校徽只能是 red（红）或 green（绿）。")
    return dict(logos[color], color=color)


def load_profile(template: str) -> dict:
    templates = registry()["templates"]
    if template not in templates:
        raise ValueError("版式只能是 A（本科学年论文）或 B（课程期末大作业）。校徽颜色另用 logo 设置。")
    spec = json.loads((ROOT / templates[template]["spec"]).read_text(encoding="utf-8"))
    if spec.get("id") != template:
        raise ValueError("模板配置标识不一致，请恢复完整技能包。")
    base = spec["logo"]
    spec["logo"] = dict(
        logo_info(base["default"]),
        width_emu=base["width_emu"],
        height_emu=base["height_emu"],
        alignment=base["alignment"],
    )
    return spec
