# 特殊情况

## 运行前

缺模块就 `python -m pip install -r requirements.txt` 后重试；没有 `python` 试 `python3` / `py -3`；输出文件已存在就换新文件名重新 `new`，不覆盖用户的文件。

## 办公软件（Word / WPS / LibreOffice）

`run` 会自动查找并使用电脑上能用的 Word、WPS 或 LibreOffice 做实际分页、更新目录页码、生成页图。

| 情况 | 怎么做 |
|---|---|
| 都找不到 | `run` 会自动改成“静态交付”：照样生成 Word 并检查格式，目录页码显示“待更新”。`finish` 会给出一段必须告诉用户的说明，照实写进回复 |
| Mac 上只有 Word | Mac 版 Word 不能自动调用，只会用 LibreOffice；没有就是静态交付，照上一行处理 |
| 报 `RENDER_ALTERNATIVE_REQUIRED`（装了但调不起来） | 环境支持“经用户批准后在本机执行命令”就用那种方式重跑同一条 `run`；做不到就 `run --work "WORK" --office-note "实际失败原因"`，进入静态交付 |
| 用户说不要打开 Word / 不要启动办公软件 | `run --work "WORK" --office none` |
| 用户指定用某个软件 | `--office word`、`--office wps` 或 `--office libreoffice` |
| 报“当前稿件导出失败” | 说明软件能用、是文档本身有问题；按错误提示修改内容后重试，不要改成静态交付 |

不要安装软件、改系统设置或关闭用户正在用的 Word。

### WPS 手动导出（只有 WPS、且自动调用失败时）

1. 运行 `run --work "WORK" --office wps-manual`，它会给出一个请求文件路径。
2. 用被授权的桌面操作工具（或请用户）在 WPS 打开请求里的 `source.docx`，导出整篇 PDF（勾选“标题书签”）到请求里写的位置。不要另存 Word。
3. 运行 `python -X utf8 scripts/wps_exchange.py accept --request 请求文件 --pdf 导出的PDF --exporter "WPS 版本" --evidence "怎么导出的"`。
4. 重新运行同一条 `run`。如果目录页码变了会再要一次导出，照做直到不再提示。

做不到手动导出时，用 `--office-note "原因"` 进入静态交付。

## 输入文件

| 情况 | 怎么做 |
|---|---|
| 旧版 .doc | 请用户（或你用 Word/WPS）另存为 .docx 再 `inspect` |
| PDF、扫描件、图片 | 不能直接排版。请用户提供 Word 或文字版；用户同意的话，你可以把文字整理进 content.json（这算写作/改写，要告诉用户） |
| 只排版、原文是 .txt/.md，或用户在对话里贴的文字 | 原样存成 .txt（贴的文字也一样），`new --source 那个文件` 登记为原稿，再 `inspect --text 那个文件` 逐字导入 content.json，只核对结构不改字。没登记原稿 `run` 不交付（SOURCE_REQUIRED） |
| 原稿有修订痕迹、批注 | `run` 会停下。请用户先在 Word 里“接受/拒绝所有修订”后再发；不要替用户决定 |
| 原稿有文本框、浮动图片、内容控件、嵌入对象 | `run` 会停下并说明是哪种对象。请用户改成普通文字/嵌入式图片后再发，或按下面“脚本做不到的要求”处理 |
| 原稿有英文摘要而 B 版默认没有 | 问用户保留还是去掉（见 content.md） |

## 验收没通过

`run` 列出的每个问题都带位置（第几页 / 哪一段 / 该改的输入项）、实际与应为、可能原因和改法（同类问题合并成一条，列出所有页码/段落），完整列表在 `WORK/issues.json`。改 job.json / content.json / mapping.json 后重新 `run`。问题没变就换一种改法；连续 3 次修改都没解决任何问题（或连续 8 次没通过）就停下，如实告诉用户哪些没解决。

常见代码（除了标“提示”的，都会拦下交付；每条报错自带原因和改法）：

| 代码 | 含义 |
|---|---|
| `BLANK_PAGE` / `WORD_BLANK_PAGE` | 多余空白页 / 分页符后紧跟分页、分节或段前分页（Word 里多一页空白，LibreOffice 可能看不出） |
| `TEXT_OUTSIDE_MARGINS` / `OFF_PAGE` / `IMAGE_*` / `TABLE_OUTSIDE_MARGINS` | 文字、图片、表格伸进页边距或出纸 |
| `TEXT_OVERLAP` / `LINE_TOO_TIGHT` / `TEXT_OVER_IMAGE` | 文字叠在一起 / 固定行距比字号小 / 文字压在图片上 |
| `INDENT_RENDERED` / `NOT_CENTERED` / `LINE_PITCH_RENDERED` / `BODY_IN_MARGIN` | 页图上缩进、居中、行距和模板不符，或正文压进上下边距 |
| `HEADER_OVERFLOW` | 页眉折成多行压进正文：问用户要短题目写进 `format.header.right` |
| `HEADER_*` / `PAGE_NUMBER_*` / `PAGE_NUMBERING` / `FOOTER_UNEXPECTED` | 页眉页脚、页码格式或连续性不对 |
| `COVER_DATE_TOO_HIGH` | 封面日期浮在保留区上面：page_twips 的下边距放不下页脚，不是封面太长 |
| `SOURCE_*` | 只排版时原文被删、改、多出，或没登记原稿；用户同意的增删写进 `approved_changes` |
| `INPUT_REJECTED` / `WRITING_CHECK` | 输入被拒（指出哪个 blocks 或哪一段）/ 写作硬性要求没满足 |
| `FONT_SUBSTITUTED`（提示） | 渲染机器缺宋体等字体，页图用了替代字体；照 `finish` 的说明告诉用户 |
| `RENDERER_DIFFERENCE`（提示） | 另装的 Word/WPS 排出的页数或目录页码和检查用的软件不同；告诉用户以它为准 |
| `ORIGINAL_HIGHLIGHT_REMOVED` / `PREVIOUS_MANUAL_EDIT`（提示） | 原稿的高亮被去掉（告诉用户）/ 原稿是以前 adopt 过的成品，那项手工处理要重做或写进 job.json |

`run` 打印的“需要处理的提示”和 `review.json` 的 `warnings` 一一对应（有页码的会带上页码），每条都要写怎么处理的。

## 封面太长

`run` 会自动换行、收紧封面空白和行距，必要时自动重试。仍放不下时会提示，请问用户：缩短题目/课程名/某一项，或允许缩小封面字号（允许的话在 format 里写 `cover_fonts`）。

## 脚本做不到的要求

用户的要求仍然有效，不能用“模板不允许”拒绝，也不能悄悄恢复默认。

1. 先确认是不是 [job.md](job.md) 的写法没找对（例如“全文字体”要改多个角色）。
2. 确实不支持时，告诉用户哪一项不能自动完成，其他部分照常完成。
3. 你有能力用 python-docx 或 Word 安全地编辑文档时：先让 `run` 通过，把 `WORK/final.docx` 复制一份，只改这一项格式，不动文字。然后运行

   ```bash
   python -X utf8 scripts/xaufe.py adopt --work "WORK" --docx "改好的文件" --note "改了什么"
   ```

   `adopt` 核对文字和成品完全一样，重做实际分页、渲染后检查、目录页码核对和原稿比对。只能改文字外观（字体、字号、颜色、加粗、对齐、缩进、间距、页眉文字、校徽大小），差异按角色合并成提示；带来的其他问题照样拦下。整类段落都改了的（job.json 的 `format.styles` 做得到）会被拒（ADOPT_USE_JOB）并给出要写的设置。手工处理记在成品里，以后改稿时会提醒。之后照常 review.json、`finish`。不要绕过 `adopt` 自己复制文件到交付位置。

## 常见报错

| 报错里说 | 意思和做法 |
|---|---|
| “还缺：……” | 封面信息不全，去问用户 |
| “不是当前版式的封面字段” / “不认识的项” | 报错会写出 JSON 路径（如 format.styles.h1）和最接近的正确写法，按提示改 |
| “mode: paper 需要 ……” | 完整结构缺了摘要/关键词/参考文献：写作时补写；只排版也要补，见 job.md 的“只排版时的纠正” |
| SOURCE_DIFFERENCES（写作模式的提示） | 原稿只是素材，没用上的文字只是提示：看 `WORK/source-comparison.json` 有没有用户要保留的，在 review.json 的 warnings 写明 |
| “检查之后 Word 又被改动” | 不要直接改 WORK 里的 Word，改输入文件后重新 `run` |
