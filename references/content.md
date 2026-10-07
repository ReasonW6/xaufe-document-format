# 内容怎么交给脚本

两种方式二选一：
- **content.json**：新写的文档，或用户给的是文字 / Markdown / TXT。
- **mapping.json**：用户给的是 Word（.docx）。由 `inspect` 自动生成，你只核对每段的角色。

WORK 里只能有其中一个文件。

## content.json

### 完整例子（B 版，完整结构）

```json
{
  "mode": "paper",
  "metadata": {"title_zh": "常见排序算法的比较"},
  "abstract_zh": "本文比较了冒泡排序、快速排序和归并排序……",
  "keywords_zh": ["排序算法", "时间复杂度", "稳定性"],
  "blocks": [
    {"type": "heading", "level": 1, "text": "引言", "preface": true},
    {"type": "paragraph", "text": "排序是计算机科学中最基本的操作之一……"},
    {"type": "heading", "level": 1, "text": "一、常见排序算法"},
    {"type": "heading", "level": 2, "text": "（一）冒泡排序"},
    {"type": "heading", "level": 3, "text": "1.基本思想"},
    {"type": "paragraph", "text": "……"},
    {"type": "table", "caption": "表1 三种算法的比较", "header": ["算法", "平均时间复杂度"], "rows": [["冒泡排序", "O(n²)"], ["快速排序", "O(n log n)"]]}
  ],
  "references": [
    "严蔚敏, 吴伟民. 数据结构(C语言版)[M]. 北京: 清华大学出版社, 2007."
  ]
}
```

### 各项

| 项 | 说明 |
|---|---|
| `mode` | `"paper"`：完整结构（封面、摘要、目录、正文、参考文献，顺序按版式）。`"document"`：只排 `blocks`，没有摘要和目录（普通文档、只给了正文时用） |
| `metadata.title_zh` | 文档题目，和 job.json 里的题目一致（封面题目留空时也要写真实题目） |
| `metadata.title_en` | 英文题目；有英文摘要时必须写 |
| `abstract_zh` / `keywords_zh` | 中文摘要（多段用 `\n` 分开）/ 关键词列表。`paper` 模式必须有 |
| `abstract_en` / `keywords_en` | 英文摘要 / Key words。A 版默认必须有；B 版只有 job.json 写了 `"english_abstract": true` 才写 |
| `blocks` | 正文，按顺序一块一块写，见下表 |
| `references` | 参考文献列表，每条一个字符串，按正文第一次引用的顺序；没写 `[1]` 会自动编号。写作时怎么找、怎么核对见 [citations.md](citations.md) |

`blocks` 的写法：

| 类型 | 写法 |
|---|---|
| 标题 | `{"type": "heading", "level": 1, "text": "一、引言"}`。编号写在文字里：一级“一、”，二级“（一）”，三级“1.”。默认最多三级 |
| 序言 | `{"type": "heading", "level": 1, "text": "序 言", "preface": true}` |
| 段落 | `{"type": "paragraph", "text": "……"}`。有上标、下标、加粗时用 `"runs": [{"text": "x"}, {"text": "2", "superscript": true}]` |
| 表格 | `{"type": "table", "caption": "表1 …", "header": [...], "rows": [[...], ...]}`，每行格数和表头相同 |
| 图片 | `{"type": "image", "path": "图片文件名.png", "caption": "图1 …"}`。图片文件放在 WORK 里，路径相对 content.json |
| 公式 | `{"type": "equation", "text": "E = mc²"}` |
| 代码 | `{"type": "code", "text": "print('hello')"}` |
| 分页 | `{"type": "page_break"}` |

- 全篇题目不要再写成正文的一级标题。
- 原文里的“摘要”“关键词”“参考文献”放进对应的项，不要塞进 blocks。
- **只排版时**：原文怎么写就怎么搬，一个字不改；原文没有摘要或参考文献，就用 `"mode": "document"`，不要补写。字数、文献篇数不够也不补，最多在回复里提醒。

## 写作要求（写作模式才看）

1. **先问后写。** 主题、方向、必须包含的内容没问清楚就不要动笔。
2. **结构按版式：**
   - A 版：中文题目、英文题目、中文摘要（150–300 字）和关键词（3–5 个）、英文摘要和 Key words、序言、正文（最多三级标题：一、/（一）/1.）、参考文献。
   - B 版：中文题目、中文摘要和关键词、正文、参考文献。用户要英文摘要时再加。
3. **篇幅：** 按用户要求。用户没说：A、B 都是正文不少于 5000 字、参考文献至少 6 篇（A 版细则原文：“正文：应不少于5000字”“参考文献至少6篇”；B 版沿用），并在回复里说明。`run` 会报出正文字数和参考文献条数。
4. **语言：** 规范的学术中文；摘要客观陈述，不写“本文认为”“我认为”。英文摘要与中文摘要内容一致。
5. **参考文献和数字（最容易出错）：** 按 [citations.md](citations.md) 检索、筛选，用 `refs` 命令核对。
   - 每条都必须是这次实际检索到的或用户给的，**绝不凭记忆写条目**。找不够就如实告诉用户还差几篇，不凑数。
   - 具体数字（多少人、百分比、几倍、金额）只能来自核对过的文献或用户给的资料，句末标引用序号；没有来源就不写数字，用“很多”“大多数”这类说法。
   - 正文引用处标上标序号（写法见 citations.md 第 8 节）。
   - 不编造实验数据、调查结果、样本（例如“本研究调查了某校 150 名学生”）。需要数据时请用户提供；只是演示分析方法时，说明这是方法示例，不写具体结果数字。
   - 字数不够就把论述、方法、代码讲解写充分，不要编数据凑字数。
   - 不加图片，除非用户给了图片文件。
6. **题目：** 一般不超过 25 字。标题编号要连续（一、二、三……，每个一级标题下从（一）重新数）。
7. **不要自己加分页**（`page_break`）：章节不需要另起一页，脚本会处理该分页的地方。用户要求时才加。
8. **JSON 里的引号：** 文字里需要引号时用中文引号“ ”，不要用英文双引号 `"`，否则 content.json 会坏。文件用写文件工具保存成 UTF-8；用 PowerShell 写时加 `-Encoding utf8`。

## mapping.json（已有 Word）

`inspect` 会生成：

```json
{
  "source": "D:\\原稿.docx",
  "input_sha256": "……",
  "mode": "document",
  "format_tables": true,
  "replace_cover_before": 0,
  "paragraphs": [
    {"id": "p0", "role": "title_zh", "text": "常见排序算法的比较"},
    {"id": "p1", "role": "h1", "text": "一、引言"},
    {"id": "p2", "role": "body", "text": "排序是……"}
  ]
}
```

**你只需要检查和修改每一段的 `role`。** `text` 只是开头文字，方便你看；不要改 `id`、`source`、`input_sha256`。

标题级别**按编号定**：“一、”是 `h1`，“（一）”是 `h2`，“1.”是 `h3`。原稿跳过了某一级（例如“三、”下面直接是“1.”）也照编号来，不要按上下关系改级别。

| role | 用在 |
|---|---|
| `body` | 普通正文段落 |
| `h1` / `h2` / `h3` | 一 / 二 / 三级标题（如“一、”“（一）”“1.”） |
| `preface` | “序言”“引言”这类不带编号的开篇标题 |
| `title_zh` | 文档大标题（文章开头的题目） |
| `abstract_label_zh` / `abstract_zh` / `keywords_zh` | “摘要”二字 / 摘要正文 / “关键词：……” |
| `abstract_label_en` / `abstract_en` / `keywords_en` | “Abstract” / 英文摘要正文 / “Key words: ……” |
| `toc_title` | “目录”二字 |
| `reference_title` / `reference` | “参考文献”标题 / 每条文献 |
| `caption` | 图题、表题（“图1 …”“表1 …”） |
| `figure` / `equation` / `code` | 只放图片 / 公式 / 代码的段落 |
| `spacer` | 空段落 |
| `preserve` | 原样保留不排版（旧封面） |
| `delete` | 删除这一段（只用于用户同意删的内容和被替换的旧目录） |

其他项：
- `mode`：inspect 自动判断。`"paper"`＝原稿有摘要和标题，按完整论文结构排：自动生成目录、摘要单独成页、分节设页码（原稿里的旧目录会标成 `delete`，由新目录代替）。`"document"`＝普通文档，只排样式，不加目录。
- `add`：只在 `mode` 为 `"paper"`、所选版式要英文摘要、原稿没有、**用户同意由你翻译**时使用，写法：
  ```json
  "add": {"title_en": "英文题目", "abstract_en": "英文摘要正文", "keywords_en": ["keyword one", "keyword two"]}
  ```
  用户不要英文摘要时，不写 `add`，改在 job.json 的 format 写 `"english_abstract": false`。
- `replace_cover_before`：`0` 表示在最前面加上所选版式的封面；大于 0 表示原稿前面有旧封面，会被替换成新封面（数字是旧封面后第一块内容的位置，由 inspect 自动判断）；`null` 表示不加封面（用户不要封面时）。旧封面判断错了就问用户。
- `format_tables`：`true` 表示表格也按模板调整宽度和字体。
- 原稿有英文摘要但 B 版默认没有：问用户保留还是去掉。保留就在 job.json 的 format 写 `"english_abstract": true`；用户同意去掉时，把这些段落的 role 改成 `delete`。
- `delete`：只用于用户同意删除的段落，以及 `mode: "paper"` 时被新目录代替的旧目录。run 会列出删了哪些段；回复里要告诉用户。
