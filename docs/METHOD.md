# paper-to-slides — 使用与方法说明（METHOD）

版本：1.0（与提交的 `SKILL.md`、`scripts/` 一致）

## 1. 用途与适用条件

把一篇论文或研究材料转换为可编辑、证据可追溯、符合给定模板的学术演示文稿（PPTX），
并附带逐页讲稿/备注与真实渲染预览（PDF + PNG）。适用于学术组会、开题/答辩、会议汇报、
论文精读讲解等"把连续叙述改写成页级可讲表达"的场景。

不负责：论文结论本身的重新推导、科研绘图的美化、学术写作润色。

## 2. 目录与依赖

```
SKILL.md                       Skill 本体（流程与规则）
scripts/build_deck.py          大纲 JSON → 可编辑 PPTX（模板原位追加 / 纯标准库 OOXML）
scripts/extract_pdf.py         PDF → 全文 + 逐页文本 + 图/表/公式定位 + 页面渲染
scripts/render_deck.py         PPTX → PDF + 逐页 PNG 预览（LibreOffice + poppler）
scripts/write_result.py        case 输出契约 + 状态 → result.json
schemas/outline.schema.json    页级大纲 JSON Schema
schemas/evidence_map.schema.json  页级证据映射 JSON Schema
examples/example_outline.json  最小调用示例（示意，非正式案例答案）
docs/METHOD.md                 本文档
```

依赖（脚本会自动探测并降级）：
- Python 3（标准库即可）；`python-pptx` 可选（有则更优，无则纯标准库 OOXML）。
- poppler `pdftotext`、`pdftoppm`（文本抽取与页面渲染）。
- LibreOffice `soffice`（Office 渲染为 PDF）。
- 无需 pip 安装；无网络依赖。

## 3. 输入输出

输入：论文 PDF、可选模板 PPTX、`/input/case.json`、`/input/task.md`。

输出（写入 `/app/output`）：
- `deck.pptx` —— 可编辑演示文件（页级结构、主张、图、备注、模板结果）。
- `deck.pdf`、`previews/preview-*.png` —— 真实渲染预览。
- `evidence_map.json` —— 每页内容对应论文图/表/章节/数据的映射。
- `outline.json` —— 页级大纲（含每页主张与证据）。
- `result.json` —— output_contract 全部字段 + case_id + status + sources + uncertainty + change_record + artifacts。

## 4. 关键方法

1. **主线提取**：把论文压缩为"问题 → 方法 → 证据 → 结论 → 限制"的可讲叙事骨架，区分主结论与支撑性细节。
2. **页结构与信息分层**：每页一个主张，3–5 个可讲短语，控制单页信息密度。
3. **图表引用与证据保留**：每页 `evidence` 记录 `source_type/source_id/section/page/quote/confidence`，编号、单位、图注、条件与原文一致。
4. **讲稿与页面分离**：备注写过渡、强调点、备用证据，不复述页面文字。
5. **模板约束与版面适配**：复用模板 slideMaster/slideLayout/theme/media；长标题、多图、大表格用"要点 + 图"布局。
6. **导出渲染与可编辑性**：LibreOffice 导出 PDF，逐页 PNG 预览，逐页检查溢出/字号/字体/版式。

## 5. 完整调用流程

```bash
# 1) 抽取论文
python scripts/extract_pdf.py --pdf paper.pdf --out work/extract

# 2) 读 case.json + task.md，写 work/outline.json（按 schemas/outline.schema.json）

# 3) 生成可编辑 PPTX（有模板则原位追加，保留模板设计）
python scripts/build_deck.py --outline work/outline.json \
  --template template.pptx --out output/deck.pptx

# 4) 渲染预览并检查
python scripts/render_deck.py --deck output/deck.pptx --out output

# 5) 写 result.json（values.json 显式填 output_contract 字段）
python scripts/write_result.py --case /input/case.json --out-dir output \
  --values work/values.json --status ok --sources work/sources.json \
  --uncertainty work/uncertainty.json --change-record work/change_record.json
```

## 6. 失败处理与自查

- 任何一步失败都要落到 `status=partial/failed` 并在 result.json 写明真实原因与影响范围；不假装成功。
- 逐项自查见 `SKILL.md` 第 8 节。
- 复跑性：同一 outline + 同一模板 + 同一脚本版本应产出同一版式；把模板路径、脚本版本、渲染工具写进 `sources`。

## 7. 限制

- 纯标准库 OOXML 路径不保证与 python-pptx 逐字节一致，但保证可编辑、可渲染、模板母版/版式/主题保留。
- 公式为图片、扫描件、加密 PDF 需 OCR/人工转录，转录不确定性必须标 `confidence=low`。
- 本 Skill 不生成论文结论，只负责"论文内容如何变成可讲、可核验、可编辑的 PPT"。
