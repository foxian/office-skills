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
| `${SKILL_DIR}/scripts/md_to_docx.py` | Convert Markdown to DOCX with format templates |
| `${SKILL_DIR}/scripts/docx_reader.py` | Read DOCX, output rich-text Markdown with paragraph IDs |
| `${SKILL_DIR}/scripts/docx_editor.py` | Execute JSON DSL operations on DOCX |
| `${SKILL_DIR}/scripts/docx_diff.py` | Compare before/after DOCX, generate diff report |

## Format Templates

| Type | Syntax | Example |
|------|--------|---------|
| JSON rules | `--template json:name` | `--template json:formal` |
| DOCX styles | `--template docx:file.docx` | `--template docx:~/template.docx` |

Available JSON templates: `default`, `formal`, `casual`

## DSL Operations Reference

See `${SKILL_DIR}/references/edit-ops.md` for complete DSL specification.

---

## Workflow 1: Create Mode (从零创建)

> For creating new documents from scratch.

### Step 1: Outline Discussion

Discuss document structure and direction with user.

### Step 2: AI Writing

Generate `draft.md` with standard Markdown content.

### Step 3: Review Iteration

User and AI iterate on Markdown content directly.

### Step 4: Convert to DOCX

```bash
python ${SKILL_DIR}/scripts/md_to_docx.py draft.md output.docx --template json:formal
```

### Step 5: Micro-adjust (if needed)

Switch to Edit Mode for fine-tuning.

---

## Workflow 2: Edit Mode (修改文档)

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

## Execution Discipline

1. **Read before write**: Always read existing code before modifying
2. **Test-driven**: Write failing test first, then implement
3. **Commit frequently**: Each task should be a separate commit
4. **Backup first**: docx_editor.py always creates .bak backup before editing