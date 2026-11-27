# paper-to-slides

把一篇论文组织成**可编辑的学术 PPT**，同时保留**页级证据映射**、**讲稿/备注**与**真实渲染预览**。

参赛作品：玻尔产品周赛 · 北化分站赛第一期 · **论文转学术 PPT**
挑战 ID：`ppt-c7b80f0d`

## 它解决什么

| 维度 | 做法 |
|---|---|
| 对得上 | 按论文**自身的编号标题**切章节，不套通用模板 |
| 找得回 | 每条要点带 `page` + `section` + 逐字引文，写进页脚、备注与 `evidence.jsonl` |
| 留得住 | 论文卡片四字段：研究问题 / 数据与方法 / 主要发现 / 局限，另设"边界与不确定性"页 |
| 可编辑 | 输出原生文本框 PPTX（非位图），备注写进真正的 notesSlide |

## 一次调用

```bash
python3 /skills/weekly-skill/scripts/build_deck.py --outline outline.json --out deck.pptx
```

完整流程（PDF → PPTX + 预览 + result.json）由 `SKILL.md` 描述的六步方法驱动，四个纯标准库脚本：

- `scripts/extract_pdf.py` — PDF → 文本/分页/图表式引用/页 PNG
- `scripts/build_deck.py` — 大纲 → 可编辑 PPTX（python-pptx 优先，纯标准库 OOXML 兜底）
- `scripts/render_deck.py` — PPTX → PDF + PNG（LibreOffice + poppler）
- `scripts/write_result.py` — 输出契约字段 → result.json

## 目录

```
SKILL.md                     # 技能入口：方法、契约、异常、示例
scripts/                     # 四个脚本
schemas/                     # outline / evidence_map JSON Schema
docs/METHOD.md               # 章节切分与证据锚定方法论
examples/example_outline.json # 合成示例（非正式案例）
```

## 说明

本仓库仅含 Skill 本体与必要资源；正式案例、对照白板与评分由平台提供。
