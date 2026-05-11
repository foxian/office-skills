# Word Master Phase 1 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 完成 Word Master Phase 1，实现 set_font、set_paragraph_format 操作和双轨模板系统

**Architecture:** 在 docx_editor.py 中新增两个操作函数处理段落级别格式修改；md_to_docx.py 支持 json:xxx 和 docx:xxx 两种模板指定方式

**Tech Stack:** python-docx, JSON, pytest

---

## 文件结构

```
word-master/
├── skills/word-master/
│   ├── SKILL.md                              # 修改：更新工作流
│   ├── scripts/
│   │   ├── docx_editor.py                   # 修改：新增两个操作
│   │   └── md_to_docx.py                    # 修改：新增模板支持
│   └── references/
│       └── edit-ops.md                      # 修改：补充 DSL 规范
├── format_rules/                            # 新增目录
│   ├── default.json                         # 新增
│   ├── formal.json                          # 新增
│   └── casual.json                          # 新增
└── tests/
    └── test_docx_editor.py                  # 修改：新增 4 个测试
```

---

## 任务分解

### 任务 1: 补充 edit-ops.md DSL 规范

**Files:**
- Modify: `word-master/skills/word-master/references/edit-ops.md`

- [ ] **Step 1: 阅读现有 edit-ops.md 内容**

- [ ] **Step 2: 添加 set_font 操作规范**

在文件末尾添加：

```markdown
## set_font 操作

设置段落字体格式（段落级别）。

```json
{
  "op": "set_font",
  "target": "p0",
  "name": "Arial",
  "east_asia": "楷体",
  "size": "14pt",
  "bold": true,
  "italic": false
}
```

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| target | string | ✅ | 段落 ID，如 "p0" |
| name | string | 否 | 西文字体名 |
| east_asia | string | 否 | 中文字体名 |
| size | string | 否 | 字号，如 "14pt" 或 "28"（half-points） |
| bold | boolean | 否 | 是否加粗 |
| italic | boolean | 否 | 是否斜体 |

所有参数均作用到该段落的所有 Run。
```

- [ ] **Step 3: 添加 set_paragraph_format 操作规范**

在 set_font 规范后添加：

```markdown
## set_paragraph_format 操作

设置段落格式（对齐、行距、缩进、段前段后距）。

```json
{
  "op": "set_paragraph_format",
  "target": "p0",
  "alignment": "center",
  "line_spacing": 1.5,
  "first_line_indent": "2em",
  "space_before": "12pt",
  "space_after": "12pt"
}
```

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| target | string | ✅ | 段落 ID |
| alignment | string | 否 | left/center/right/justify |
| line_spacing | number | 否 | 行距倍数（1.5 表示 1.5 倍） |
| first_line_indent | string | 否 | 首行缩进，如 "2em" |
| space_before | string | 否 | 段前距，如 "12pt" |
| space_after | string | 否 | 段后距，如 "12pt" |

line_spacing 使用倍数制（1.0=单倍，1.5=1.5 倍）。
first_line_indent 单位支持 pt 和 em。
```

- [ ] **Step 4: 提交**

```bash
git add word-master/skills/word-master/references/edit-ops.md
git commit -m "docs: 补充 set_font 和 set_paragraph_format DSL 规范"
```

---

### 任务 2: 实现 set_font 操作

**Files:**
- Modify: `word-master/skills/word-master/scripts/docx_editor.py:1-67`
- Test: `word-master/tests/test_docx_editor.py`

- [ ] **Step 1: 阅读现有 docx_editor.py 代码**

确认现有 apply_operations 函数的结构和 _replace_text_in_runs 辅助函数位置。

- [ ] **Step 2: 编写 set_font 测试**

在 `test_docx_editor.py` 末尾添加：

```python
def test_set_font_name(tmp_path):
    """set_font should change font name for all runs in paragraph."""
    doc = docx.Document()
    p = doc.add_paragraph('Test text')
    in_path = tmp_path / "font_test.docx"
    out_path = tmp_path / "font_test_out.docx"
    doc.save(in_path)

    ops = [{"op": "set_font", "target": "p0", "name": "Arial"}]
    apply_operations(str(in_path), ops, str(out_path))

    doc_out = docx.Document(out_path)
    for run in doc_out.paragraphs[0].runs:
        assert run.font.name == "Arial"


def test_set_font_bold_italic(tmp_path):
    """set_font should toggle bold and italic."""
    doc = docx.Document()
    p = doc.add_paragraph('Bold and Italic')
    in_path = tmp_path / "bold_test.docx"
    out_path = tmp_path / "bold_test_out.docx"
    doc.save(in_path)

    ops = [{"op": "set_font", "target": "p0", "bold": True, "italic": True}]
    apply_operations(str(in_path), ops, str(out_path))

    doc_out = docx.Document(out_path)
    for run in doc_out.paragraphs[0].runs:
        assert run.bold == True
        assert run.italic == True


def test_set_font_size(tmp_path):
    """set_font should change font size."""
    doc = docx.Document()
    p = doc.add_paragraph('Size test')
    in_path = tmp_path / "size_test.docx"
    out_path = tmp_path / "size_test_out.docx"
    doc.save(in_path)

    ops = [{"op": "set_font", "target": "p0", "size": "14pt"}]
    apply_operations(str(in_path), ops, str(out_path))

    doc_out = docx.Document(out_path)
    for run in doc_out.paragraphs[0].runs:
        assert run.font.size.pt == 14.0
```

- [ ] **Step 3: 运行测试确认失败**

Run: `cd word-master/skills/word-master/scripts && python -m pytest ../../../tests/test_docx_editor.py::test_set_font_name -v`
Expected: FAIL (function not implemented)

- [ ] **Step 4: 实现 _apply_set_font 函数**

在 `_replace_text_in_runs` 函数后、`_backup_file` 函数前添加：

```python
def _apply_set_font(paragraph, params):
    """Apply font settings to all runs in a paragraph."""
    for run in paragraph.runs:
        if "name" in params:
            run.font.name = params["name"]
        if "east_asia" in params:
            run.font.east_asia = params["east_asia"]
        if "size" in params:
            size_str = params["size"]
            if size_str.endswith("pt"):
                from docx.shared import Pt
                run.font.size = Pt(float(size_str.rstrip("pt")))
            elif size_str.isdigit():
                from docx.shared import Pt
                run.font.size = Pt(int(size_str) / 2)
        if "bold" in params:
            run.bold = params["bold"]
        if "italic" in params:
            run.italic = params["italic"]
```

- [ ] **Step 5: 修改 apply_operations 添加 set_font 分支**

在 `apply_operations` 函数的 for 循环中添加：

```python
elif op['op'] == 'set_font':
    idx = int(op['target'].replace('p', ''))
    if idx < len(doc.paragraphs):
        _apply_set_font(doc.paragraphs[idx], op)
```

- [ ] **Step 6: 运行测试确认通过**

Run: `cd word-master/skills/word-master/scripts && python -m pytest ../../../tests/test_docx_editor.py::test_set_font_name ../../../tests/test_docx_editor.py::test_set_font_bold_italic ../../../tests/test_docx_editor.py::test_set_font_size -v`
Expected: PASS

- [ ] **Step 7: 提交**

```bash
git add word-master/skills/word-master/scripts/docx_editor.py word-master/tests/test_docx_editor.py
git commit -m "feat: 实现 set_font 操作"
```

---

### 任务 3: 实现 set_paragraph_format 操作

**Files:**
- Modify: `word-master/skills/word-master/scripts/docx_editor.py`
- Test: `word-master/tests/test_docx_editor.py`

- [ ] **Step 1: 编写 set_paragraph_format 测试**

在 `test_docx_editor.py` 末尾添加：

```python
def test_set_paragraph_alignment(tmp_path):
    """set_paragraph_format should change alignment."""
    doc = docx.Document()
    doc.add_paragraph('Left align')
    doc.add_paragraph('To center')
    in_path = tmp_path / "align_test.docx"
    out_path = tmp_path / "align_test_out.docx"
    doc.save(in_path)

    ops = [{"op": "set_paragraph_format", "target": "p1", "alignment": "center"}]
    apply_operations(str(in_path), ops, str(out_path))

    doc_out = docx.Document(out_path)
    assert doc_out.paragraphs[1].alignment == WD_ALIGN_PARAGRAPH.CENTER


def test_set_paragraph_line_spacing(tmp_path):
    """set_paragraph_format should change line spacing."""
    doc = docx.Document()
    p = doc.add_paragraph('Line spacing test')
    in_path = tmp_path / "linespacing_test.docx"
    out_path = tmp_path / "linespacing_test_out.docx"
    doc.save(in_path)

    ops = [{"op": "set_paragraph_format", "target": "p0", "line_spacing": 1.5}]
    apply_operations(str(in_path), ops, str(out_path))

    doc_out = docx.Document(out_path)
    assert doc_out.paragraphs[0].paragraph_format.line_spacing == 1.5
```

- [ ] **Step 2: 运行测试确认失败**

Run: `cd word-master/skills/word-master/scripts && python -m pytest ../../../tests/test_docx_editor.py::test_set_paragraph_alignment ../../../tests/test_docx_editor.py::test_set_paragraph_line_spacing -v`
Expected: FAIL

- [ ] **Step 3: 实现 _apply_set_paragraph_format 函数**

在 `_apply_set_font` 后添加：

```python
def _apply_set_paragraph_format(paragraph, params):
    """Apply paragraph format settings."""
    pf = paragraph.paragraph_format
    if "alignment" in params:
        align_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY
        }
        pf.alignment = align_map.get(params["alignment"], WD_ALIGN_PARAGRAPH.LEFT)
    if "line_spacing" in params:
        pf.line_spacing = params["line_spacing"]
    if "first_line_indent" in params:
        indent_str = params["first_line_indent"]
        if indent_str.endswith("em"):
            from docx.shared import Emu
            pf.first_line_indent = Emu(int(float(indent_str.rstrip("em")) * 914400 / 2))
        elif indent_str.endswith("pt"):
            from docx.shared import Pt
            pf.first_line_indent = Pt(float(indent_str.rstrip("pt")))
    if "space_before" in params:
        space_str = params["space_before"]
        if space_str.endswith("pt"):
            from docx.shared import Pt
            pf.space_before = Pt(float(space_str.rstrip("pt")))
    if "space_after" in params:
        space_str = params["space_after"]
        if space_str.endswith("pt"):
            from docx.shared import Pt
            pf.space_after = Pt(float(space_str.rstrip("pt")))
```

- [ ] **Step 4: 添加 WD_ALIGN_PARAGRAPH 导入**

在文件开头添加：

```python
from docx.enum.text import WD_ALIGN_PARAGRAPH
```

- [ ] **Step 5: 修改 apply_operations 添加 set_paragraph_format 分支**

在 set_font 分支后添加：

```python
elif op['op'] == 'set_paragraph_format':
    idx = int(op['target'].replace('p', ''))
    if idx < len(doc.paragraphs):
        _apply_set_paragraph_format(doc.paragraphs[idx], op)
```

- [ ] **Step 6: 运行测试确认通过**

Run: `cd word-master/skills/word-master/scripts && python -m pytest ../../../tests/test_docx_editor.py::test_set_paragraph_alignment ../../../tests/test_docx_editor.py::test_set_paragraph_line_spacing -v`
Expected: PASS

- [ ] **Step 7: 提交**

```bash
git add word-master/skills/word-master/scripts/docx_editor.py word-master/tests/test_docx_editor.py
git commit -m "feat: 实现 set_paragraph_format 操作"
```

---

### 任务 4: 实现 md_to_docx.py 模板支持

**Files:**
- Create: `word-master/skills/word-master/format_rules/default.json`
- Create: `word-master/skills/word-master/format_rules/formal.json`
- Create: `word-master/skills/word-master/format_rules/casual.json`
- Modify: `word-master/skills/word-master/scripts/md_to_docx.py`

- [ ] **Step 1: 创建 format_rules 目录**

```bash
mkdir -p word-master/skills/word-master/format_rules
```

- [ ] **Step 2: 创建 default.json**

```json
{
  "heading1": {
    "font": "方正小标宋简体",
    "size": "22pt",
    "bold": true,
    "alignment": "center",
    "space_before": "0pt",
    "space_after": "12pt"
  },
  "heading2": {
    "font": "方正黑体简体",
    "size": "16pt",
    "bold": true,
    "alignment": "left",
    "space_before": "12pt",
    "space_after": "6pt"
  },
  "body": {
    "font": "仿宋",
    "size": "16pt",
    "line_spacing": 1.5,
    "first_line_indent": "2em",
    "alignment": "justify",
    "space_before": "0pt",
    "space_after": "6pt"
  }
}
```

- [ ] **Step 3: 创建 formal.json**

```json
{
  "heading1": {
    "font": "方正小标宋简体",
    "size": "22pt",
    "bold": true,
    "alignment": "center",
    "space_before": "0pt",
    "space_after": "12pt"
  },
  "heading2": {
    "font": "方正黑体简体",
    "size": "16pt",
    "bold": true,
    "alignment": "left",
    "space_before": "12pt",
    "space_after": "6pt"
  },
  "body": {
    "font": "仿宋",
    "size": "16pt",
    "line_spacing": 1.5,
    "first_line_indent": "2em",
    "alignment": "justify",
    "space_before": "0pt",
    "space_after": "6pt"
  }
}
```

- [ ] **Step 4: 创建 casual.json**

```json
{
  "heading1": {
    "font": "微软雅黑",
    "size": "20pt",
    "bold": true,
    "alignment": "left",
    "space_before": "12pt",
    "space_after": "8pt"
  },
  "heading2": {
    "font": "微软雅黑",
    "size": "16pt",
    "bold": true,
    "alignment": "left",
    "space_before": "10pt",
    "space_after": "6pt"
  },
  "body": {
    "font": "微软雅黑",
    "size": "14pt",
    "line_spacing": 1.5,
    "first_line_indent": "2em",
    "alignment": "left",
    "space_before": "0pt",
    "space_after": "6pt"
  }
}
```

- [ ] **Step 5: 阅读现有 md_to_docx.py**

- [ ] **Step 6: 重写 md_to_docx.py 支持模板**

```python
import docx
from docx.shared import Pt, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
import sys
import os
import json


def _load_template(template_arg):
    """Load format rules from template.

    Args:
        template_arg: Either "json:xxx" (load format_rules/xxx.json) or
                     "docx:xxx" (load styles from xxx.docx)

    Returns:
        dict: Format rules for heading1, heading2, body
    """
    if template_arg.startswith("json:"):
        template_name = template_arg.split(":", 1)[1]
        skill_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        template_path = os.path.join(skill_dir, "format_rules", f"{template_name}.json")
        with open(template_path, "r", encoding="utf-8") as f:
            return json.load(f)
    elif template_arg.startswith("docx:"):
        docx_path = template_arg.split(":", 1)[1]
        return _load_styles_from_docx(docx_path)
    else:
        return None


def _load_styles_from_docx(docx_path):
    """Extract style rules from a DOCX file."""
    doc = docx.Document(docx_path)
    styles = {}
    for style in doc.styles:
        if style.type == 1:  # PARAGRAPH
            name = style.name.lower()
            if "heading 1" in name or "标题1" in name:
                styles["heading1"] = _style_to_dict(style)
            elif "heading 2" in name or "标题2" in name:
                styles["heading2"] = _style_to_dict(style)
            elif "normal" in name or "正文" in name:
                styles["body"] = _style_to_dict(style)
    if not styles:
        styles = {"heading1": {}, "heading2": {}, "body": {}}
    return styles


def _style_to_dict(style):
    """Convert a paragraph style to a dict format."""
    result = {}
    try:
        if style.font.name:
            result["font"] = style.font.name
        if style.font.size:
            result["size"] = f"{style.font.size.pt}pt"
        if style.font.bold is not None:
            result["bold"] = style.font.bold
        if style.alignment:
            align_map = {
                WD_ALIGN_PARAGRAPH.LEFT: "left",
                WD_ALIGN_PARAGRAPH.CENTER: "center",
                WD_ALIGN_PARAGRAPH.RIGHT: "right",
                WD_ALIGN_PARAGRAPH.JUSTIFY: "justify"
            }
            result["alignment"] = align_map.get(style.alignment, "left")
        if style.paragraph_format.line_spacing:
            result["line_spacing"] = style.paragraph_format.line_spacing
        if style.paragraph_format.first_line_indent:
            result["first_line_indent"] = f"{style.paragraph_format.first_line_indent.pt}pt"
    except:
        pass
    return result


def _apply_format_to_paragraph(paragraph, fmt):
    """Apply format dict to a paragraph."""
    if not fmt:
        return
    pf = paragraph.paragraph_format
    if "font" in fmt:
        for run in paragraph.runs:
            run.font.name = fmt["font"]
    if "size" in fmt:
        size_pt = float(fmt["size"].rstrip("pt"))
        for run in paragraph.runs:
            run.font.size = Pt(size_pt)
    if "bold" in fmt:
        for run in paragraph.runs:
            run.bold = fmt["bold"]
    if "alignment" in fmt:
        align_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY
        }
        paragraph.alignment = align_map.get(fmt["alignment"], WD_ALIGN_PARAGRAPH.LEFT)
    if "line_spacing" in fmt:
        pf.line_spacing = fmt["line_spacing"]
    if "first_line_indent" in fmt:
        indent_str = fmt["first_line_indent"]
        if indent_str.endswith("em"):
            pf.first_line_indent = Emu(int(float(indent_str.rstrip("em")) * 914400 / 2))
        elif indent_str.endswith("pt"):
            pf.first_line_indent = Pt(float(indent_str.rstrip("pt")))
    if "space_before" in fmt:
        pf.space_before = Pt(float(fmt["space_before"].rstrip("pt")))
    if "space_after" in fmt:
        pf.space_after = Pt(float(fmt["space_after"].rstrip("pt")))


def convert_markdown(md_path, out_path, template=None):
    doc = docx.Document()
    format_rules = None

    if template:
        format_rules = _load_template(template)

    with open(md_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    for line in lines:
        line = line.strip()
        if not line:
            continue
        if line.startswith("# "):
            p = doc.add_heading(line[2:], level=1)
            if format_rules and "heading1" in format_rules:
                _apply_format_to_paragraph(p, format_rules["heading1"])
        elif line.startswith("## "):
            p = doc.add_heading(line[3:], level=2)
            if format_rules and "heading2" in format_rules:
                _apply_format_to_paragraph(p, format_rules["heading2"])
        else:
            p = doc.add_paragraph(line)
            if format_rules and "body" in format_rules:
                _apply_format_to_paragraph(p, format_rules["body"])

    doc.save(out_path)


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: md_to_docx.py <input.md> <output.docx> [--template <template>]")
        sys.exit(1)

    md_path = sys.argv[1]
    out_path = sys.argv[2]
    template = None

    if "--template" in sys.argv:
        idx = sys.argv.index("--template")
        template = sys.argv[idx + 1]

    convert_markdown(md_path, out_path, template)
```

- [ ] **Step 7: 测试 JSON 模板**

创建测试文件 `test_template.md`:
```
# 测试标题

这是正文内容。

## 二级标题

更多正文。
```

Run:
```bash
cd word-master/skills/word-master/scripts
python md_to_docx.py test_template.md test_output.docx --template json:default
```

检查 test_output.docx 是否使用了 default.json 中的格式。

- [ ] **Step 8: 提交**

```bash
git add word-master/skills/word-master/format_rules/ word-master/skills/word-master/scripts/md_to_docx.py
git commit -m "feat: md_to_docx.py 支持双轨模板系统 (json:/docx:)"
```

---

### 任务 5: 更新 SKILL.md

**Files:**
- Modify: `word-master/skills/word-master/SKILL.md`

- [ ] **Step 1: 阅读现有 SKILL.md 和 ppt-master SKILL.md 对比**

- [ ] **Step 2: 重写 SKILL.md**

```markdown
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
```

- [ ] **Step 3: 提交**

```bash
git add word-master/skills/word-master/SKILL.md
git commit -m "docs: SKILL.md 对齐 ppt-master 风格"
```

---

### 任务 6: 验证全部测试通过

- [ ] **Step 1: 运行全部测试**

```bash
cd word-master/skills/word-master/scripts
python -m pytest ../../../tests/ -v
```

- [ ] **Step 2: 确认结果**

Expected: All tests PASS

- [ ] **Step 3: 最终提交**

```bash
git add -A
git commit -m "chore: Word Master Phase 1 完成"
```

---

## 验收清单

- [ ] edit-ops.md 补充完整
- [ ] set_font 操作测试通过
- [ ] set_paragraph_format 操作测试通过
- [ ] md_to_docx.py 模板支持工作正常
- [ ] SKILL.md 更新完成
- [ ] 全部测试通过