# 特殊情况

## 运行前

- **提示缺少模块（ModuleNotFoundError）**：先运行 `python -m pip install -r requirements.txt`，再重试原命令。不要换别的方法绕开。
- **python 命令不存在**：试 `python3` 或 `py -3`。
- **输出文件已存在**：换一个新文件名重新 `new`，不要删除或覆盖用户已有的文件。

## 办公软件（Word / WPS / LibreOffice）

`run` 会自动查找并使用电脑上能用的 Word、WPS 或 LibreOffice 做实际分页、更新目录页码、生成页图。

| 情况 | 怎么做 |
|---|---|
| 都找不到 | `run` 会自动改成“静态交付”：照样生成 Word 并检查格式，目录页码显示“待更新”。`finish` 会给出一段必须告诉用户的说明，照实写进回复 |
| 报 `RENDER_ALTERNATIVE_REQUIRED`（装了但调不起来） | 如果你的运行环境支持“经用户批准后在本机会话里执行命令”，就用那种方式重跑同一条 `run`。做不到就运行 `run --work "WORK" --office-note "实际失败原因"`，进入静态交付 |
| 用户说不要打开 Word / 不要启动办公软件 | `run --work "WORK" --office none` |
| 用户指定用某个软件 | `--office word`、`--office wps` 或 `--office libreoffice` |
| 报“当前稿件导出失败” | 说明软件能用、是文档本身有问题；按错误提示修改内容后重试，不要改成静态交付 |

不要安装软件、改系统设置或关闭用户正在用的 Word。

### WPS 手动导出（只有 WPS、且自动调用失败时）

1. 运行 `run --work "WORK" --office wps-manual`，它会给出一个请求文件路径。
2. 用你被授权的桌面操作工具（或请用户）在 WPS 里打开请求里写的 `source.docx`，导出整篇 PDF，勾选“标题书签”，保存到请求里写的位置。不要另存 Word。
3. 运行 `python -X utf8 scripts/wps_exchange.py accept --request 请求文件 --pdf 导出的PDF --exporter "WPS 版本" --evidence "怎么导出的"`。
4. 重新运行同一条 `run`。如果目录页码变了会再要一次导出，照做直到不再提示。

做不到手动导出时，用 `--office-note "原因"` 进入静态交付。

## 输入文件

| 情况 | 怎么做 |
|---|---|
| 旧版 .doc | 请用户（或你用 Word/WPS）另存为 .docx 再 `inspect` |
| PDF、扫描件、图片 | 不能直接排版。请用户提供 Word 或文字版；用户同意的话，你可以把文字整理进 content.json（这算写作/改写，要告诉用户） |
| 原稿有修订痕迹、批注 | `run` 会停下。请用户先在 Word 里“接受/拒绝所有修订”后再发；不要替用户决定 |
| 原稿有文本框、浮动图片、内容控件、嵌入对象 | `run` 会停下并说明是哪种对象。告诉用户这一部分不能自动排版，可以：①用户把它改成普通文字/嵌入式图片后再发；②你有能力安全编辑 Word 时按下面“脚本做不到的要求”处理 |
| 原稿有英文摘要而 B 版默认没有 | 问用户保留还是去掉（见 content.md） |

## 封面太长

`run` 会自动换行、收紧封面空白和行距，必要时自动重试。仍放不下时会提示，请问用户：缩短题目/课程名/某一项，或允许缩小封面字号（允许的话在 format 里写 `cover_fonts`）。

## 脚本做不到的要求

用户的要求仍然有效，不能用“模板不允许”拒绝，也不能悄悄恢复默认。

1. 先确认是不是 [job.md](job.md) 的写法没找对（例如“全文字体”要改多个角色）。
2. 确实不支持时，告诉用户哪一项不能自动完成，其他部分照常完成。
3. 你有能力用 python-docx 或 Word 安全地编辑文档时，可以在 WORK 里复制一份成品，只改这一项，不动其他内容；改完运行 `python -X utf8 scripts/xaufe.py preview --work "WORK" --docx "改好的文件"` 生成页图逐页检查，确认无误后复制到交付位置，最后运行 `cancel` 清理 WORK。回复里说明这一项是手工处理的。

## 常见报错

| 报错里说 | 意思和做法 |
|---|---|
| “还缺：……” | 封面信息不全，去问用户 |
| “不是当前版式的封面字段” | job.json 的 info 写了这个版式没有的项，按提示改 |
| “job.json 有不认识的项” | 拼错了项名，对照 job.md |
| “mode: paper 需要 ……” | 完整结构缺了摘要/关键词/参考文献：写作时补写；只排版就改成 `"mode": "document"` |
| “原稿里有些文字在成品中没找到”（SOURCE_DIFFERENCES） | 打开 `WORK/source-comparison.json` 对照：没经同意漏掉的要补回；用户同意删改的在 review.json 的 warnings 写明 |
| “检查之后 Word 又被改动” | 不要直接改 WORK 里的 Word，改输入文件后重新 `run` |
