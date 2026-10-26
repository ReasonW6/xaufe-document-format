# job.json 写法

`WORK/job.json` 记录这次的全部要求。每次 `check` / `run` 都会重新读取它，所以：
- **它永远写“当前完整的要求”**：上一轮的要求要保留，这一轮只改用户说的那一项。
- **恢复某项默认 = 把那一项从 job.json 里删掉。**

## 完整例子

```json
{
  "task": "format",
  "template": "B",
  "logo": "green",
  "info": {
    "course_name": "数据结构",
    "title_zh": "常见排序算法的比较",
    "student_name": "张三",
    "major": "软件工程",
    "class_name": "软工2301"
  },
  "blank": ["student_id"],
  "format": {
    "styles": {"body": {"cn": "仿宋", "size": 12}}
  }
}
```

## 各项含义

| 项 | 写法 | 说明 |
|---|---|---|
| `task` | `"write"` 或 `"format"` | 必填。这次要写内容写 `"write"`，只排版写 `"format"`（判断方法见 SKILL.md 的“需要用户决定的事”） |
| `template` | `"A"` 或 `"B"` | 必填。A=本科学年论文，B=课程期末大作业。用户没选就先问，不要替用户选 |
| `logo` | `"red"` 或 `"green"` | 可不写：A 默认绿，B 默认红 |
| `cover` | `true` / `false` | 可不写，默认有封面。用户说不要封面才写 `false` |
| `info` | 封面信息，值都是文字 | 只写用户说过或原稿里有的；不知道的不要写 |
| `blank` | 字段名列表 | 用户明确说“留空，我自己填”的项 |
| `format` | 特别要求，见下表 | 没有就写 `{}` 或不写 |
| `timezone` | 例如 `"Asia/Shanghai"` | 可不写，默认北京时间，用来算“今天” |
| `approved_changes` | `[{"text": "那段文字", "why": "用户怎么说的"}]` | 只排版时用户明确同意增删的文字；图片等对象写种类和数量（“删去图片 1 张”）。其他增删一律拦下；`add` 补的译文不用写 |
| `accept_unverified_refs` | `true` | 写作时连不上文献核对服务、且用户明确同意先用未核实的文献时才写；不写则 `run` 拦下 |
| `corrections` | 见下面“只排版时的纠正” | 只排版时为符合学校格式你补的、改的文字，交付时都会告诉用户 |
| `keep_as_is` | `"用户原话"` | 用户坚持题目、摘要等保持原样（或不要补）时写；超字数和缺的部分改为交付时提醒 |

`info` 和 `blank` 里能用的字段：

| 字段 | 含义 | A | B |
|---|---|---|---|
| `title_zh` | 题目 | ✓ | ✓ |
| `course_name` | 课程名称（只写课程名，不要书名号） | | ✓ |
| `student_name` | 学生姓名 | ✓ | ✓ |
| `student_id` | 学号 | ✓ | ✓ |
| `major` | 专业 | ✓ | ✓ |
| `class_name` | 班级 | ✓ | ✓ |
| `supervisor` | 指导教师 | ✓ | 需先在 format 写 `"cover_supervisor": true` |
| `date` | 完成日期，写成 `2026-10-07` | 不写=今天 | 不写=今天 |

## 只排版时的纠正（corrections）

换旧封面、去旧目录和图表目录、标题重编号、统一“内 容 摘 要”等标签、挪图题表题、表注和单位的写法、公式编号由脚本做并自己登记。下面这些由你做，每条写成 `{"part": …, "before": "原稿整段", "after": "改后整段", "why": "告诉用户的理由"}`：

| 情况 | part 和写法 |
|---|---|
| 缺中文摘要 / 关键词 | `abstract` / `keywords`，只写 after（按正文拟；关键词写成“关键词：词1；词2；词3”，3～5 个）和 why |
| A 版缺英文摘要 | `title_en`、`abstract_en`、`keywords_en`，只写 after（照中文翻译） |
| 摘要不在 150～300 字 | `abstract`，before 原文整段，after 改后整段，why 如“原摘要 139 字，学校要求 150～300 字，按正文补了一句” |
| 题目超过 24 字 | `title`，同时改 `info.title_zh` |
| 缺参考文献 | `references`，每条一项、只写 after；先 `refs --from-document` 核实通过，核实不了的不补 |
| 删一段 / 在某段后补一段 | after 写 `""` / 加 `"insert_after": "前一段整段原文"` |

before 要和原稿整段一致（只差空格可以）；before 为空表示补一段，题目摘要关键词放在正文前，文献放在最后。原稿是 Word 时脚本照 corrections 改；用 content.json 时你直接改 content.json，corrections 只登记，`run` 会核对改后的文字确实在里面。

## 用户说什么 → format 里怎么写

字号对照：三号=16，四号=14，小四=12，五号=10.5，小五=9（单位：磅）。

| 用户说 | format 里写 |
|---|---|
| A 版用红色校徽 / B 版用绿色校徽 | 不写在 format，写顶层 `"logo": "red"` / `"logo": "green"` |
| 正文改成仿宋小四 | `"styles": {"body": {"cn": "仿宋", "size": 12}}` |
| 正文英文改成 Arial | `"styles": {"body": {"latin": "Arial"}}` |
| 正文 1.5 倍行距 | `"styles": {"body": {"line": 360, "line_rule": "auto"}}` |
| 正文单倍 / 2 倍行距 | `"line": 240` / `"line": 480`，都配 `"line_rule": "auto"` |
| 正文固定 22 磅行距 | `"styles": {"body": {"line": 440, "line_rule": "exact"}}`（磅数×20） |
| 首行缩进 2 字符 | `"styles": {"body": {"chars": 200}}` |
| 一级标题改成三号 | `"styles": {"h1": {"size": 16}}` |
| 一级标题加粗 / 改蓝色 | `"styles": {"h1": {"bold": true}}` / `{"h1": {"color": "0070C0"}}` |
| 目录字号都改成小四 | `"styles": {"toc1": {"size": 12}, "toc2": {"size": 12}, "toc3": {"size": 12}}` |
| 要页码（页脚显示页码） | `"footer_page_numbers": true` |
| B 版也要英文摘要 | `"english_abstract": true`（content.json 要有英文摘要，见 content.md） |
| A 版不要英文摘要 | `"english_abstract": false` |
| B 版加指导教师栏 | `"cover_supervisor": true`，并在 info 写 `supervisor` |
| A 版去掉指导教师栏 | `"cover_supervisor": false` |
| 不要页眉 | `"header": {"enabled": false}` |
| B 版加页眉 | `"header": {"enabled": true, "left": "左侧文字", "right": "右侧文字"}` |
| 页眉左边改成某文字 | `"header": {"left": "某文字"}` |
| 去掉页眉横线 | `"header": {"border": false}` |
| 校徽居中 / 校徽宽 9 厘米 | `"logo_alignment": "center"` / `"logo_width_cm": 9` |
| 左右页边距 2.5 厘米 | `"page_twips": {"left": 1418, "right": 1418}`（1 厘米=567） |
| 需要四级标题 | `"heading_levels": 4` |
| 用户确认原稿标题就按自己的方式编号（如“（一）”是一级） | `"heading_numbering": "own"`（不写时，用了模板另一级编号的标题会被拦下） |
| 封面“期末大作业”改成“课程论文” | `"cover_label": "课程论文"` |
| 封面题目标签改成“论文题目：” | `"cover_title_label": "论文题目："` |
| 封面信息栏改楷体 | `"cover_fonts": {"fields": {"cn": "楷体"}}` |
| “关键词：”这几个字改小四 | `"keyword_labels": {"zh": {"size": 12}}` |
| A 版摘要放到目录前面 | `"parts": ["cover", "abstract_zh", "abstract_en", "toc", "body", "references"]` |

多个要求合在一个 format 里，例如：

```json
"format": {
  "footer_page_numbers": true,
  "styles": {
    "body": {"cn": "仿宋", "size": 12},
    "h1": {"size": 16}
  }
}
```

### “正文”和“全文”不一样

- “正文改成 X”只改 `body`。
- “全文改成 X 字体”要改所有文字：`body`、`h1`、`h2`、`h3`、`preface`、`title_zh`、`abstract_label_zh`、`abstract_zh`、`keywords_zh`、`reference_title`、`reference`、`toc_title`、`toc1`、`toc2`、`toc3`、`caption`、`table_text`、`header`，再加 `"cover_fonts": {"label": {...}, "title": {...}, "fields": {...}, "date": {...}}`（B 版还有 `"course"`）。

### styles 里能用的属性

| 属性 | 含义 | 例子 |
|---|---|---|
| `cn` / `latin` | 中文字体 / 英文字体 | `"仿宋"` / `"Times New Roman"` |
| `size` | 字号（磅，可以是 .5） | `12` |
| `bold` / `italic` | 加粗 / 斜体 | `true` |
| `color` | 颜色，6 位十六进制，不带 # | `"0070C0"` |
| `underline` | 下划线 | `"single"` / `"none"` |
| `align` | 对齐 | `"left"` / `"center"` / `"right"` / `"both"` |
| `chars` | 首行缩进（1 字符=100） | `200` |
| `first` / `left` / `hanging` | 首行 / 左 / 悬挂缩进（磅×20） | `420` |
| `line` + `line_rule` | 行距（见上表） | `360` + `"auto"` |
| `before` / `after` | 段前 / 段后（磅×20） | `120` |
| `before_lines` / `after_lines` | 段前 / 段后（1 行=100） | `50` |

可用角色（styles 的键）：`body` 正文、`h1`/`h2`/`h3`/`h4` 各级标题、`preface` 序言标题、`title_zh`/`title_en` 摘要页题目、`abstract_label_zh`/`abstract_label_en` “内容摘要”/“Abstract”、`abstract_zh`/`abstract_en` 摘要正文、`keywords_zh`/`keywords_en` 关键词、`toc_title` “目录”、`toc1`~`toc4` 目录条目、`reference_title` “参考文献”、`reference` 文献条目、`caption` 图表题、`figure` 图片段、`equation` 公式段、`table_text` 表格文字、`code` 代码、`header` 页眉。

## 多轮修改的例子

- 第 1 轮：用户说“正文仿宋小四”→ `"styles": {"body": {"cn": "仿宋", "size": 12}}`
- 第 2 轮：用户说“目录改小四”→ 保留第 1 轮的写法，再加 `"toc1"`、`"toc2"`、`"toc3"`。
- 第 3 轮：用户说“正文字号恢复默认”→ 只删掉 `body` 里的 `"size"`，保留 `"cn": "仿宋"` 和目录的设置。

写错了 `check` / `run` 会说哪里不对、怎么改；表里没有、脚本也不支持的要求见 [special-cases.md](special-cases.md) 的“脚本做不到的要求”。
