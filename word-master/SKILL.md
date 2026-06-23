---
name: word-master
description: >
  AI-driven Word document creation and editing system. Use when user wants to
  create, edit, review, or format a .docx file. Covers "从零写文档" (creation)
  and "改已有文档" (editing) scenarios.
---

# Word Master Skill

> AI-driven Word document creation and editing system with dual-layer architecture:
> AI intent layer + deterministic execution layer.

## Core Scripts

| Script | Purpose |
|--------|---------|
| `${SKILL_DIR}/scripts/docx_reader.py` | Read DOCX, output rich-text Markdown with paragraph IDs |
| `${SKILL_DIR}/scripts/docx_editor.py` | Execute JSON DSL operations on DOCX |
| `${SKILL_DIR}/scripts/docx_diff.py` | Compare before/after DOCX, generate diff report |
| `${SKILL_DIR}/scripts/style_analyzer.py` | Extract format fingerprints from a template DOCX |
| `${SKILL_DIR}/scripts/style_transfer.py` | Apply template style profile to a target DOCX |
| `${SKILL_DIR}/scripts/template_analyzer.py` | Extract comprehensive format profile (paragraphs + tables) |
| `${SKILL_DIR}/scripts/table_analyzer.py` | Extract table style fingerprints (library) |

## Format Templates

| Type | Syntax | Example |
|------|--------|---------|
| JSON rules | `--template json:name` | `--template json:formal` |

Available JSON templates: `default`, `formal`, `casual`

## DSL Operations Reference

See `${SKILL_DIR}/references/edit-ops.md` for complete DSL specification.

---

## Workflow 1: Edit Mode (修改文档)

> For editing existing documents.

### Step 1: Read Document

```bash
python ${SKILL_DIR}/scripts/docx_reader.py document.docx --overview
```

This outputs rich-text Markdown with paragraph IDs (p0, p1, ...).

### Step 2: Generate Edit Operations

AI understands user intent and generates JSON DSL operations.
See `${SKILL_DIR}/references/edit-ops.md` for supported operations.

Example operations:
```json
[
  {"op": "replace_text", "find": "old", "replace": "new"},
  {"op": "rewrite_paragraph", "target": "p0", "content": "New content"},
  {"op": "set_font", "target": "p0", "name": "Arial", "bold": true},
  {"op": "set_paragraph_format", "target": "p0", "alignment": "center"}
]
```

### Step 3: Execute Edit

```bash
python ${SKILL_DIR}/scripts/docx_editor.py document.docx operations.json
```

Original file backed up as `document.docx.bak`.

### Step 4: Verify Changes

```bash
python ${SKILL_DIR}/scripts/docx_diff.py document.docx document_modified.docx
```

---

## Workflow 2: Style Transfer (样式迁移)
 
> 学习源模板文档的排版格式风格（包含段落格式与表格样式），应用到目标草稿文档。
 
### Step 1: 提取源模板格式指纹
 
```bash
python ${SKILL_DIR}/scripts/template_analyzer.py template.docx --output style_profile.json
```
 
可选参数：
- `--para-min-cluster-size N`：段落聚类最小样本数（默认 4）
- `--para-no-heading-aware`：禁用段落标题感知模式（使用普通的格式聚类）
 
这会直接生成一个包含 `roles` (段落格式角色) 和 `table_roles` (表格样式特征) 的综合 `style_profile.json` 配置文件。
 
### Step 2: AI 推断与微调样式角色
 
将生成的 `style_profile.json` 提供给 AI 审查，或由 AI 自动读取。AI 会推断这些段落和表格样式的应用意图，并对其进行调优。
 
**格式约束**：主结构必须包含：
- `roles`：段落格式角色列表
- `table_roles`：表格样式规则列表
 
**段落角色 (`roles`) 白名单值**：
| 类别 | 可用值 |
|------|--------|
| 标题 | `Heading 1`, `Heading 2`, `Heading 3`, `Heading 4`, `Heading 5`, `Heading 6`, `Heading 7`, `Heading 8`, `Heading 9` |
| 正文 | `Normal` |
| 列表 | `List Bullet`, `List Number` |
 
**表格样式 (`table_roles`) 结构说明**：
每个表格条目包含：
- `id`: 唯一标识符 (int)
- `structure`: 物理结构特征 (包含 `cols`, `rows`, `has_header_row`)
- `border`: 六方向边框（`top`, `bottom`, `left`, `right`, `insideH`, `insideV`）的 XML 格式定义
- `shading`: `header` 与 `body` 的底纹颜色 (HEX / None)
- `cell_margin`: 单元格四周的内边距大小 (如 `"1.0pt"`)
- `header_text` & `body_text`: 表头及表身文字的字体样式、粗细与对齐方式
 
**示例结构**：
```json
{
  "roles": [
    {"role": "Heading 1", "fingerprint": {"size": "14.0pt", "font": "黑体", "align": "center", "color": null}},
    {"role": "Normal", "fingerprint": {"size": "11.0pt", "font": "宋体", "align": "justify", "color": null}}
  ],
  "table_roles": [
    {
      "id": 0,
      "structure": {"cols": 3, "rows": 5, "has_header_row": true},
      "border": {
        "top": {"val": "single", "sz": 4, "color": "000000"}
      },
      "shading": {
        "header": "D9E1F2",
        "body": null
      },
      "header_text": {
        "font": "Arial",
        "size": "12.0pt",
        "bold": true,
        "align": "center"
      }
    }
  ]
}
```
 
> **注意**：`style_transfer.py` 会自动验证 profile 的强类型约束。
 
> **匹配规则**：样式转换采用精确名称和结构匹配。段落部分会套用对应 `role` 的 fingerprint；表格部分会基于最近邻打分算法（列数 + 表头行 + 行数相似度）匹配 `table_roles` 中的最近邻表格样式并自动生成应用操作。
 
### Step 3: 应用样式到目标草稿
 
```bash
python ${SKILL_DIR}/scripts/style_transfer.py --profile style_profile.json draft.docx output.docx
```
 
可选参数：
- `--review`：应用前暂停，供人工确认 profile
- `--skip-head N`：跳过前 N 个段落（排除封面）
- `--skip-tail N`：跳过后 N 个段落（排除签字页）
 
### Step 4: 验证结果
 
```bash
python ${SKILL_DIR}/scripts/docx_diff.py draft.docx output.docx
```

---

## Execution Discipline

1. **Read before write**: Always read existing code before modifying
2. **Test-driven**: Write failing test first, then implement
3. **Commit frequently**: Each task should be a separate commit
4. **Backup first**: docx_editor.py always creates .bak backup before editing