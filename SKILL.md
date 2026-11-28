---
name: paper-to-slides
description: >
  Turn a research paper or study material into an editable, evidence-mapped academic
  slide deck (PPTX) that preserves the organizer's template, with page-level source
  citations and non-repetitive speaker notes, plus real rendered PDF/PNG previews.
  Trigger when a task asks to build 学术 PPT / 汇报幻灯片 / an academic presentation from
  a paper, when /input/case.json and /input/task.md describe a paper-to-slides case, or
  when a deck must keep a template's master/layout/theme and page-level evidence mapping.
---

# paper-to-slides：论文 → 可编辑学术 PPT

本 Skill 负责把一篇论文（或研究材料）组织成**结构清楚、证据可追溯、可编辑、符合给定模板**的学术演示文稿。它不是"把段落切成行贴进页面"，而是**页级结构重写**：每页一个主张，主张背后挂论文里的图/表/公式/章节/数据，并在给定模板里排得下、讲得通、打开就能放映。

## 1. 运行时输入输出契约

这是 Skill authoring 任务。评测时本 Skill 被安装在 `/skills/weekly-skill`，由官方 DeepSeek Harness 在固定案例上运行。

- 输入：`/input/case.json`（案例任务、材料路径、输出契约 `output_contract`、观众/时长/页数上限/来源规则）与 `/input/task.md`（案例说明）。
- 材料：论文 PDF 与模板 PPTX 等，路径由 case.json/task.md 给出；PDF 只给原文，**没有预抽取答案层**，必须实际读文件。
- 工作区：`/app/work`（中间产物：抽取结果、outline、values）。
- 交付区：`/app/output`（最终产物：可编辑 PPTX、PDF/PNG 预览、evidence_map、result.json）。
- `result.json` 必须包含 case.json 中 `output_contract` 的**全部字段**，外加 `case_id`、`status`、`sources`、`uncertainty`、`change_record`、`artifacts`。

**事实依据是实际输入文件和原论文；不按错误的领域标签补造数据，不编造缺失的软件状态或科学数据。**

## 2. 目标与判断标准

只看三件事，不看页数多少：
1. **主线与论文自身贡献一致**，不是按段落顺序搬运。
2. **每页可回溯**到论文的图、表、章节、公式或数据（编号、单位、图注、条件不被改写）。
3. **成品可在指定模板与软件版本中打开、播放、编辑**，并能由他人按记录复跑出同一版式。

## 3. 六步方法

### 第 0 步 读输入，定契约
1. 读 `/input/case.json`，提取：`case_id`、论文/材料路径、模板路径、`output_contract` 字段、观众、时长、页数上限、来源/不确定性规则。
2. 读 `/input/task.md`，确认本案例的特殊要求。
3. 若 `output_contract` 字段未给出，以 case.json 与 task.md 中明确要求的交付物为准，并在 result.json 的 `sources` 中说明。

### 第 1 步 抽取与定位（脚本：`scripts/extract_pdf.py`）
```bash
python scripts/extract_pdf.py --pdf <paper.pdf> --out /app/work/extract [--dpi 120]
```
- 产出 `fulltext.txt`、逐页 `pages/page-*.txt`、`page_images/page-*.png` 与 `manifest.json`。
- 从 manifest 读取页数、检测到的图/表/公式编号及其所在页。
- 若 PyMuPDF 可用，`manifest.figure_images` 给出每页嵌入位图的 PNG 与 bbox；**按页就近把位图配到图/表**，矢量图回退到 `page_images/` 整页图。
- **逐页读正文**，不要只读摘要。把"示意图/原理图"与"实验数据图/表"分开标注，不把他人结果当本文结果。

### 第 2 步 提取主线（写进 outline 前先自问）
用一段话回答，再据此定页序：
1. 论文要解决什么问题（研究问题）？
2. 核心做法是什么（方法，方法名保持原文）？
3. 支撑结论的关键证据有哪些（图/表/公式/数据）？
4. 主结论是什么？限制与适用范围是什么？
主线 = 问题 → 方法 → 证据 → 结论 → 限制。**区分主结论与支撑性细节**，细节只作备注或后备证据。

### 第 3 步 设计页级结构（outline，schema 见 `schemas/outline.schema.json`）
- 每页一个主张；标题写成"主张"而不是"章节名"（如写成"方法 X 使误差下降 40%"而非"结果"）。
- 单页 3–5 个要点，每条是可讲的短短语，不是长句；数字、单位、方法名、结论表述与原文一致。
- 每页填写 `evidence`：`source_type`（figure/table/equation/section/paragraph/data/supplement）、`source_id`（如 `Fig. 3`、`Table 2`、`Eq. (7)`、`Sec. 3.2`）、`section`、`page`、`quote`（逐字数值/结论，不改写）、`confidence`。
- 每页填 `type`（`title/outline/section/content/two_content/image/thanks`）；build_deck 按类型选模板中匹配的版式，图文页引用抽出的 `figure_images` 位图。
- 页数不超过 case 给定的上限；超限时优先合并支撑性细节进备注。

### 第 4 步 生成讲稿/备注（与页面文字分离）
每页 `notes` 写：
- 过渡句（从上一页怎么接到本页）；
- 本页强调点（口头强调，而不是复述页面文字）；
- 被问到时能补充的证据（备用数字、对照组、限制）。

### 第 5 步 按模板生成可编辑文件（脚本：`scripts/build_deck.py`）
```bash
python scripts/build_deck.py --outline /app/work/outline.json \
  --template <template.pptx> --out /app/output/deck.pptx
```
- 有模板时**原位追加页面并复用模板的 slideMaster/slideLayout/theme/media**：逐页按 `type` 选择匹配版式，并复用该版式的 title/body/pic 占位符几何（标题、配色、页脚、字体规则不覆盖）。
- 无模板时用内置最小母版生成 16:9 可编辑 PPTX。
- 长标题/多图/大表格用"左侧要点 + 右侧图"或单图布局（`two_content` 双栏 / `image` 图文）；控制字号，防止溢出与自动缩放。
- 优先使用 python-pptx（若运行时可用）；否则用内置纯标准库 OOXML 写入器，二者输出等价可编辑。

### 第 6 步 渲染并检查（脚本：`scripts/render_deck.py`）
```bash
python scripts/render_deck.py --deck /app/output/deck.pptx --out /app/output [--dpi 120]
```
- 生成 `deck.pdf` 与逐页 `previews/preview-*.png`。
- **逐页看渲染图**，检查：文字是否溢出、字号是否过小、图是否错位、字体是否被替换、页脚/配色是否保留。
- 把渲染与版本记录写进 result.json 的 `change_record` / `sources`。

## 4. 汇总 result.json（脚本：`scripts/write_result.py`）
```bash
python scripts/write_result.py --case /input/case.json --out-dir /app/output \
  --values /app/work/values.json --status ok|partial|failed \
  --sources /app/work/sources.json --uncertainty /app/work/uncertainty.json \
  --change-record /app/work/change_record.json
```
- `values.json` 显式填写 output_contract 的每个字段（如 deck 路径、evidence_map、notes 等）。
- 脚本会补齐 `case_id`、`status`、`sources`、`uncertainty`、`change_record`、`artifacts`，并**列出未能填写的契约字段**。
- 状态只写真实情况：`ok`（全部完成）、`partial`（部分完成且已写明局限）、`failed`（关键步骤失败）。

## 5. 异常处理（必须如实记录，不假装成功）
| 情况 | 处置 |
|---|---|
| PDF 加密/不可读 | 记录 `status=partial`，在 `sources`/`warnings` 注明，改用可得的元数据/页面图继续或停止并说明。 |
| 公式为图片 / 无文本层 | 用 `ocr-tool`（若可用）或渲染页图人工转录；转录不确定处标 `confidence=low`，绝不臆造公式。 |
| 图表无法提取 | 保留图注与编号，引用"原论文 Fig. N"定位；不放错图、不把示意图当数据图。 |
| 模板占位符不足/无匹配版式 | 用最接近的版式；版式不足写入 `change_record`。 |
| 字体缺失 | 渲染后检查；若字体被替换，在 `change_record` 记录并尽量选择可用替代字体。 |
| 页数超限 | 合并支撑性细节到备注；仍超限则保留必要页并在 `uncertainty` 说明取舍。 |
| LibreOffice / pdftotext 缺失 | 脚本已内置探测；缺失时 `status=partial` 并写明"未能渲染"或"未能抽取文本"。 |

## 6. 依赖与环境
- Python 3（标准库即可；`numpy/matplotlib/Pillow` 可选）。
- `pdftotext`、`pdftoppm`（poppler）；`soffice`（LibreOffice，用于渲染）。
- `python-pptx` 可选：存在则更优，不存在自动退回纯标准库 OOXML。
- 脚本路径均相对本 Skill 目录：`scripts/build_deck.py` 等。

## 7. 最小调用示例（示意输出，非正式案例答案）
输入 outline（`examples/example_outline.json`）片段：
```json
{"title":"Example deck (synthetic)","slides":[
  {"id":"s01","type":"title","title":"A Sample Paper: Method and Results",
   "bullets":["Presenter","Date"],
   "notes":"Open with the one-line contribution.",
   "evidence":{"source_type":"title/abstract","source_id":"title","section":"abstract","page":1,"confidence":"high"}}
]}
```
调用：
```bash
python scripts/build_deck.py --outline examples/example_outline.json --out /app/work/example.pptx
python scripts/render_deck.py --deck /app/work/example.pptx --out /app/work
```
示意输出：`example.pptx`（可编辑）、`example.pdf`、`previews/preview-1.png`（真实渲染预览）。

## 8. 自查清单（交付前逐项过一遍）
1. 每页标题是否都是一个"主张"，而非段落/章节名？
2. 每页 `evidence` 是否都指向论文具体图/表/章节/页码，且编号与正文引用一致？
3. 数字、单位、方法名、结论是否与原文一致，未改写？
4. 备注是否补充了过渡与备用证据，而不是复述页面文字？
5. 模板母版/版式/主题/页脚是否保留？
6. 渲染 PDF/PNG 是否已生成且逐页检查过？
7. `result.json` 是否包含 output_contract 全部字段 + case_id + status + sources + uncertainty + change_record？
8. 任何 partial/failed 是否已写明具体局限与真实原因？
