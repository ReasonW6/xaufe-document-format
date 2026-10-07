# xaufe-document-format

让 AI 助手按**西安财经大学**的模板写作、排版、检查 Word 文档的 Agent Skill。

支持两套版式：**A 本科学年论文**、**B 课程期末大作业**。成品带原校徽封面、目录、摘要和页眉，最后只交付一个 `.docx`。

> 非官方项目，与西安财经大学没有隶属关系。格式以学校、学院的最新通知为准。

## 能做什么

- **两套版式、两枚校徽，自由搭配**：A 默认绿校徽，B 默认红校徽，A 配红、B 配绿都可以。
- **写不写内容，按你说的来**：
  - 给了完整文档或说“只排版”，一个字都不改。
  - 只给了题目或方向，先问清主题和内容再写。
  - 给了提纲或草稿，就在你的材料上写完整。
- **已有 Word 一键套模板**：自动识别标题、摘要、关键词、参考文献，补上目录、摘要分页和分节页码。
- **长封面自动适配**：题目、课程名再长也自动换行对齐，完成日期始终留在第一页底部，字号不变。
- **参考文献不编造**：写作时先检索再用 DOI、ISBN 或原文页面逐条核对，交付时附一张核实清单。没有出处的具体数字不写，编号跳号会被拦下。
- **真实检查**：有 Word 或 WPS 时实际分页、刷新目录页码，并逐页看图检查；没有办公软件也能交付，但会说明哪些检查没做。
- **默认值都能改**：字体、字号、页码、页眉、英文摘要等；多轮修改只改你这次说的那一项。

## 两套版式

| | A：本科学年论文 | B：课程期末大作业 |
|---|---|---|
| 依据 | 《信息学院学年论文细则》 | 课程期末大作业五份模板 |
| 默认校徽 | 绿色 | 红色 |
| 封面填写项 | 题目、姓名、学号、专业、班级、指导教师、完成日期 | 课程名称、题目、姓名、学号、专业、班级、完成日期 |
| 文档顺序 | 封面 → 目录 → 中文摘要 → 英文摘要 → 正文 → 参考文献 | 封面 → 中文摘要 → 目录 → 正文 → 参考文献 |
| 页眉 | 左侧专业与文种，右侧题目 | 无 |
| 正文 | 五号宋体，固定 20 磅行距，首行缩进 2 字符 | 相同 |
| 页脚页码 | 默认不显示（目录仍有页码） | 默认不显示（目录仍有页码） |

完整的默认参数和每项的出处见 [references/format-spec.md](references/format-spec.md)。

## 安装

在 Claude Code 中使用，把仓库克隆到技能目录，再装依赖：

```bash
git clone https://github.com/ReasonW6/xaufe-document-format ~/.claude/skills/xaufe-document-format
```

```bash
python -m pip install -r ~/.claude/skills/xaufe-document-format/requirements.txt
```

- Windows 上 `~` 就是 `C:\Users\你的用户名`。
- 只想在某个项目里用，就克隆到该项目的 `.claude/skills/` 下。
- 其他支持 Agent Skills（`SKILL.md`）的工具，把整个文件夹放进它的技能目录即可。

## 怎么用

直接用自然语言说要求，助手会自己调用这个技能：

```text
用西财 B 版模板帮我写一篇《Python程序设计》期末大作业，方向是 Python 在数据分析中的应用。
```

```text
把 D:\论文\初稿.docx 按西财 A 版学年论文格式排一下，只排版，内容别动。
```

```text
封面换成红色校徽，正文改成小四号，其他不变。
```

大致流程：

1. 选版式。没说 A 还是 B，助手会发一张对比表让你选。
2. 补齐封面信息。缺的项一次问完，任何一项都可以选“留空自己填”。完成日期默认是今天。
3. 写作或排版。
4. 实际分页并逐页检查。
5. 交付一个 Word。

## 运行环境

- **Python 3.10+**，依赖见 [requirements.txt](requirements.txt)（python-docx、PyMuPDF、jsonschema、tzdata）。
- **Microsoft Word（推荐）**：用于实际分页、刷新目录页码和生成页图。脚本也支持 WPS 和 LibreOffice；都没有时照样交付 Word，并在回复里说明没做的检查。
- **联网（仅写作时）**：核对参考文献要访问 Crossref、OpenAlex、Open Library 和文献原文页面。

## 目录结构

```text
xaufe-document-format/
├── SKILL.md            给 AI 助手的分步说明（技能入口）
├── references/         对比表、job.json 写法、内容写法、参考文献规则、默认格式、特殊情况
├── scripts/xaufe.py    唯一命令入口：compare / read / new / check / inspect / refs / run / finish …
├── scripts/            排版、封面适配、格式检查、文献核对、办公软件调用、交付清理
├── assets/             两套清洁母版、两枚原校徽、检查基准
├── sources/            学校原始模板和细则（来源存档，不修改）
├── examples/           最小示例输入
└── tests/              自动测试和真实渲染验收脚本
```

## 验证情况

- **自动测试**：500 多项单元和集成测试。
- **真实 Word 验收**：8 个用例在 Word 中实际分页并逐页看图，覆盖：
  - A、B 默认版式；
  - 校徽互换；
  - 超长封面；
  - 多项格式定制；
  - 在上一版成品上再修改；
  - 不启动办公软件的静态交付；
  - 已有论文补齐结构。
- **弱模型验收**：用 Claude Haiku 4.5 当执行者，跑了从零写作、只排版、已有 Word、多轮修改、缺信息时提问等场景。每轮发现的问题都已改进技能，关键环节（文献核对、数字出处、编号连续）由脚本把关，不靠模型自觉。
- **只在 Windows + Word 上实测过**。macOS、Linux、WPS 自动导出、LibreOffice 理论上支持，但还没有实测。

维护时的检查命令：

```bash
python -X utf8 scripts/verify_bundle.py
```

```bash
python -X utf8 -B -m unittest discover -s tests -p "test_*.py"
```

```bash
python -X utf8 tests/run_acceptance.py --output-dir 技能目录外的空文件夹
```

改动任何文件后要重新生成 `SHA256SUMS`，`verify_bundle.py` 会指出不一致的文件。

## 注意

- AI 写的内容和找的文献，提交前请自己再读一遍、核对一遍。助手附的核实清单会标出哪些文献没能自动核实。
- 校徽图片、`sources/` 下的原始模板和细则归西安财经大学所有，这里只用于按学校格式排版。

## 许可证

代码和文档以 [MIT 许可证](LICENSE) 开源。校徽图片（`assets/logos/`）和学校原始模板、细则（`sources/`）不在许可范围内。
