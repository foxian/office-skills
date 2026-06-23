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
| `${SKILL_DIR}/scripts/docx_structure.py` | 标题编号管理（添加/移除），支持灵活模板 |

## Format Templates

| Type | Syntax | Example |
|------|--------|---------|
| JSON rules | `--template json:name` | `--template json:formal` |

Available JSON templates: `default`, `formal`, `casual`

## DSL Operations Reference

See `${SKILL_DIR}/references/edit-ops.md` for complete DSL specification.

### Table Cell Operations

Read a document with tables → see cell IDs in the form `t{table_idx}r{row_idx}c{col_idx}`
(e.g. `t0r1c2` = table 0, row 1, column 2). Then target them in operations:

```json
[
  {"op": "replace_cell_text", "target": "t0r1c2", "find": "old", "replace": "new"},
  {"op": "rewrite_cell", "target": "t0r0c0", "content": "Brand new cell content"}
]
```

- `replace_cell_text`: in-cell targeted find/replace (preserves run-level formatting).
- `rewrite_cell`: full cell content replacement (single paragraph, plain text).
- Invalid target format → `ValueError`; out-of-range indices → `IndexError`.

---

## Workflow 1: Edit Mode (修改文档)

> For editing existing documents.

### Step 1: Read Document

```bash
python ${SKILL_DIR}/scripts/docx_reader.py document.docx --overview
```

This outputs rich-text Markdown with paragraph IDs (`p0`, `p1`, …) interleaved
with table cell IDs (`tNrNcN`, e.g. `t0r1c2`). Use the IDs in subsequent edit
operations. Pass `--overview` to skip tables (paragraphs only).

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

---

## Heading Numbering

标题编号管理：给 docx 的标题段落添加/移除编号前缀（文本前缀方案，非 Word 原生 numPr）。
与 markdown-master 的 `structure.py numbering` 对称，模板语法和 YAML 配置完全一致。

### CLI

```bash
# 添加编号
python ${SKILL_DIR}/scripts/docx_structure.py <file.docx> numbering add \
    [--h1 T] [--h2 T] [--h3 T] [--h4 T] [--h5 T] [--h6 T] \
    [--config FILE] [--start-from N] [--save-config FILE] \
    [-o output.docx]

# 移除编号
python ${SKILL_DIR}/scripts/docx_structure.py <file.docx> numbering remove [-o output.docx]
```

默认覆盖原文件并生成 `.bak` 备份；`-o` 输出到新文件时不生成 `.bak`。

### 编号模板语法

| 占位符 | 含义 | 示例 |
|---------|------|------|
| `{N}` / `{N:d}` | 十进制数字 | 3 |
| `{N:02d}` | 两位补零十进制数字 | 03 |
| `{N:R}` | 大写罗马数字 | III |
| `{N:r}` | 小写罗马数字 | iii |
| `{N:A}` | 大写字母 | C |
| `{N:a}` | 小写字母 | c |
| `{N:cn}` | 中文数字 | 三 |

N 范围是 1-6，对应标题级别 H1-H6。空模板字符串表示该级标题不添加编号。

### YAML 配置文件格式

```yaml
h1: "第{1}章 "
h2: "{1}.{2} "
h3: "（{3}）"
h4: ""
h5: ""
h6: ""
start_from: 1
```

### 常见编号用法示例

```bash
# 技术文档风格（1/1.1/1.1.1）
python ${SKILL_DIR}/scripts/docx_structure.py doc.docx numbering add \
    --h1 "{1} " --h2 "{1}.{2} " --h3 "{1}.{2}.{3} "

# 中文章节风格（第1章/1.1/（1））
python ${SKILL_DIR}/scripts/docx_structure.py doc.docx numbering add \
    --h1 "第{1}章 " --h2 "{1}.{2} " --h3 "（{3}）"

# 学术风格（I/A/1/a）
python ${SKILL_DIR}/scripts/docx_structure.py doc.docx numbering add \
    --h1 "{1:R} " --h2 "{2:A} " --h3 "{3} " --h4 "{4:a} "

# 保存当前配置为模板
python ${SKILL_DIR}/scripts/docx_structure.py doc.docx numbering add \
    --h1 "第{1}章 " --h2 "{1}.{2} " --save-config chinese_chapter.yaml

# 使用保存的配置
python ${SKILL_DIR}/scripts/docx_structure.py doc.docx numbering add \
    --config chinese_chapter.yaml
```