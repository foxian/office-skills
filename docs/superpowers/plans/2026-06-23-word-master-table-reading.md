# Word Master 表格读取增强 — 实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 让 `docx_reader.py` 按文档真实阅读顺序交织输出段落和表格内容（带单元格 ID），并将表格读取能力延伸到 `docx_editor.py`、`docx_diff.py` 和 `SKILL.md`。

**Architecture:** 将 `docx_reader.py` 的核心遍历逻辑从 `doc.paragraphs` 改为遍历 `doc.element.body` 子元素，区分 `<w:p>` 和 `<w:tbl>`，将表格渲染为带 `[tNrNcN]` ID 前缀的 Markdown 表格。`docx_editor.py` 新增解析 `tNrNcN` target 的两个 DSL 操作；`docx_diff.py` 新增表格逐单元格对比逻辑。

**Tech Stack:** Python 3.x, python-docx, pytest

## Global Constraints

- 工作目录：`word-master/`（所有路径相对此目录）
- 脚本目录：`scripts/`，测试目录：`tests/`
- 运行测试命令：`pytest tests/ -v`（在 `word-master/` 下执行）
- 段落 ID `p{i}` 的现有含义不变（`i` 为在 body 中的顺序位置）
- `--overview` 模式行为不变（只输出 Heading，不输出表格）
- 所有新 DSL 操作与现有操作向后兼容

---

### Task 1: 重构 docx_reader.py — 按文档顺序交织输出表格

**Files:**
- Modify: `scripts/docx_reader.py`
- Test: `tests/test_docx_reader.py`

**Interfaces:**
- Produces: `extract_rich_markdown(filepath, overview=False) -> str` — 签名不变，但输出现在包含表格块
- Produces: `_extract_table_markdown(table, table_idx: int) -> list[str]` — 新增，供 Task 3 的 diff 逻辑复用
- Produces: `_extract_cell_overrides(cell) -> list[str]` — 新增辅助函数

---

- [ ] **Step 1: 写失败测试（表格出现在输出中）**

在 `tests/test_docx_reader.py` 末尾追加：

```python
def test_extract_rich_markdown_with_table(tmp_path):
    """表格内容应出现在 extract_rich_markdown 输出中，带正确 ID"""
    doc = docx.Document()
    doc.add_paragraph('Before table.')
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = 'H1'
    table.cell(0, 1).text = 'H2'
    table.cell(1, 0).text = 'R1C1'
    table.cell(1, 1).text = 'R1C2'
    doc.add_paragraph('After table.')
    doc_path = tmp_path / "test_table.docx"
    doc.save(str(doc_path))

    md = extract_rich_markdown(str(doc_path))

    # 表格注释头
    assert '<!-- TABLE t0:' in md
    # 单元格 ID
    assert '[t0r0c0]' in md
    assert '[t0r1c1]' in md
    # 内容
    assert 'H1' in md
    assert 'R1C2' in md
    # 段落仍在
    assert 'Before table.' in md
    assert 'After table.' in md


def test_extract_rich_markdown_table_order(tmp_path):
    """表格应出现在 Before 段落之后、After 段落之前（文档阅读顺序）"""
    doc = docx.Document()
    doc.add_paragraph('BEFORE')
    table = doc.add_table(rows=1, cols=1)
    table.cell(0, 0).text = 'IN_TABLE'
    doc.add_paragraph('AFTER')
    doc_path = tmp_path / "order_test.docx"
    doc.save(str(doc_path))

    md = extract_rich_markdown(str(doc_path))

    before_pos = md.find('BEFORE')
    table_pos = md.find('IN_TABLE')
    after_pos = md.find('AFTER')
    assert before_pos < table_pos < after_pos


def test_overview_mode_excludes_tables(tmp_path):
    """overview 模式下表格不出现在输出中"""
    doc = docx.Document()
    doc.add_heading('Chapter 1', level=1)
    table = doc.add_table(rows=1, cols=1)
    table.cell(0, 0).text = 'TABLE_CONTENT'
    doc_path = tmp_path / "overview_test.docx"
    doc.save(str(doc_path))

    md = extract_rich_markdown(str(doc_path), overview=True)

    assert 'Chapter 1' in md
    assert 'TABLE_CONTENT' not in md
    assert '<!-- TABLE' not in md


def test_multi_paragraph_cell(tmp_path):
    """多段落单元格应以 ' / ' 连接"""
    import docx as _docx
    from docx.oxml import OxmlElement
    doc = _docx.Document()
    table = doc.add_table(rows=1, cols=1)
    cell = table.cell(0, 0)
    cell.paragraphs[0].text = 'Line1'
    cell.add_paragraph('Line2')
    doc_path = tmp_path / "multi_para.docx"
    doc.save(str(doc_path))

    md = extract_rich_markdown(str(doc_path))
    assert 'Line1 / Line2' in md
```

- [ ] **Step 2: 运行测试，验证失败**

```
pytest tests/test_docx_reader.py::test_extract_rich_markdown_with_table -v
```
预期：FAILED（`<!-- TABLE t0:` 不在输出中）

- [ ] **Step 3: 实现 `_extract_cell_overrides` 和 `_extract_table_markdown`，重构 `extract_rich_markdown`**

将 `scripts/docx_reader.py` 全部替换为以下内容：

```python
import docx
import sys
from docx.enum.text import WD_ALIGN_PARAGRAPH

def _extract_style_props(style):
    """
    提取 Word 样式的具体属性（字体、字号、对齐、间距等）。
    返回一个包含格式信息的字典。
    """
    props = {}
    if not hasattr(style, 'font'):
        return props
    
    # 提取字体属性
    if style.font.name: props['font'] = style.font.name
    if style.font.size: props['size'] = f"{style.font.size.pt}pt"
    if style.font.bold is not None: props['bold'] = style.font.bold
    if style.font.italic is not None: props['italic'] = style.font.italic
    
    # 提取段落格式属性
    if hasattr(style, 'paragraph_format'):
        pf = style.paragraph_format
        if pf.alignment is not None:
            align_map = {
                WD_ALIGN_PARAGRAPH.LEFT: 'LEFT',
                WD_ALIGN_PARAGRAPH.CENTER: 'CENTER',
                WD_ALIGN_PARAGRAPH.RIGHT: 'RIGHT',
                WD_ALIGN_PARAGRAPH.JUSTIFY: 'JUSTIFY'
            }
            props['align'] = align_map.get(pf.alignment, str(pf.alignment))
        if pf.line_spacing is not None: props['line_spacing'] = pf.line_spacing
        if pf.first_line_indent is not None:
            props['first_line_indent'] = f"{pf.first_line_indent.pt}pt"
        if pf.space_before is not None:
            props['space_before'] = f"{pf.space_before.pt}pt"
        if pf.space_after is not None:
            props['space_after'] = f"{pf.space_after.pt}pt"
            
    return props

def _extract_run_overrides(paragraph):
    """
    提取段落中各个 Run（文本块）相对于段落样式的手动格式覆盖。
    """
    overrides = []
    for run in paragraph.runs:
        text = run.text.strip()
        if not text:
            continue
        
        run_props = {}
        if run.font.name is not None: run_props['font'] = run.font.name
        if run.font.size is not None: run_props['size'] = f"{run.font.size.pt}pt"
        if run.font.bold is not None: run_props['bold'] = run.font.bold
        if run.font.italic is not None: run_props['italic'] = run.font.italic
        
        if run_props:
            prop_str = ", ".join(f"{k}={v}" for k, v in run_props.items())
            overrides.append(f"{prop_str} on \"{text}\"")
    return overrides

def _extract_cell_overrides(cell) -> list:
    """
    提取单元格内所有段落中 Run 的格式覆盖信息列表。
    """
    all_overrides = []
    for para in cell.paragraphs:
        all_overrides.extend(_extract_run_overrides(para))
    return all_overrides

def _extract_table_markdown(table, table_idx: int) -> list:
    """
    将单个 python-docx Table 对象渲染为带单元格 ID 标注的 Markdown 行列表。

    格式：
      <!-- TABLE t0: 3 cols × 5 rows -->
      | [t0r0c0] **内容** | ... |
      |---|---|---|
      | [t0r1c0] 内容 | ... |
    """
    rows = table.rows
    cols_count = len(table.columns)
    rows_count = len(rows)
    lines = []
    lines.append(f"<!-- TABLE t{table_idx}: {cols_count} cols × {rows_count} rows -->")

    for r_idx, row in enumerate(rows):
        cells = row.cells
        # 去除合并单元格导致的重复（python-docx 中合并单元格会重复出现）
        seen_tcs = set()
        unique_cells = []
        for cell in cells:
            tc_id = id(cell._tc)
            if tc_id not in seen_tcs:
                seen_tcs.add(tc_id)
                unique_cells.append(cell)

        cell_strs = []
        for c_idx, cell in enumerate(unique_cells):
            # 多段落单元格以 ' / ' 连接
            cell_text = ' / '.join(
                p.text for p in cell.paragraphs if p.text.strip()
            )
            overrides = _extract_cell_overrides(cell)
            cell_id = f"[t{table_idx}r{r_idx}c{c_idx}]"

            if r_idx == 0:
                # 首行视为表头，加粗
                display = f"{cell_id} **{cell_text}**"
            else:
                display = f"{cell_id} {cell_text}"

            if overrides:
                override_str = "; ".join(overrides)
                display += f" {{{override_str}}}"

            cell_strs.append(display)

        lines.append("| " + " | ".join(cell_strs) + " |")

        # 在表头行后插入分隔线
        if r_idx == 0:
            sep_cols = len(unique_cells)
            lines.append("|" + "|".join(["---"] * sep_cols) + "|")

    return lines

def extract_rich_markdown(filepath, overview=False):
    """
    读取 DOCX 文档，按文档阅读顺序交织输出段落和表格的 Markdown 表示。

    :param filepath: DOCX 文件路径
    :param overview: 如果为 True，则只输出标题大纲（Heading），表格不参与输出
    """
    doc = docx.Document(filepath)
    lines = []
    lines.append(f"<!-- DOCUMENT: {filepath} -->")

    if not overview:
        # 收集文档中实际使用到的所有样式名称（仅段落）
        used_style_names = set(p.style.name for p in doc.paragraphs if p.text.strip())
        if used_style_names:
            lines.append("<!-- STYLES:")
            for s_name in sorted(used_style_names):
                if s_name in doc.styles:
                    style = doc.styles[s_name]
                    props = _extract_style_props(style)
                    if props:
                        prop_str = ", ".join(f"{k}:{v}" for k, v in props.items())
                        lines.append(f"  {s_name} -> {prop_str}")
                    else:
                        lines.append(f"  {s_name} -> (default)")
                else:
                    lines.append(f"  {s_name} -> (not found in doc.styles)")
            lines.append("-->")

    lines.append("")  # 头部后的空行

    # 按文档顺序遍历 body 子元素，区分段落（w:p）和表格（w:tbl）
    para_idx = 0
    table_idx = 0

    # 预先建立 python-docx 的段落/表格对象列表，用于通过 XML element 反查
    _para_map = {id(p._element): p for p in doc.paragraphs}
    _table_map = {id(t._element): t for t in doc.tables}

    for child in doc.element.body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

        if tag == 'p':
            para = _para_map.get(id(child))
            if para is None:
                para_idx += 1
                continue
            text = para.text.strip()
            if not text:
                para_idx += 1
                continue
            style_name = para.style.name if para.style is not None else "Normal"

            if overview and not style_name.startswith('Heading'):
                para_idx += 1
                continue

            overrides = _extract_run_overrides(para) if not overview else []
            if overrides:
                override_str = " | override: " + "; ".join(overrides)
                lines.append(f"### [p{para_idx}] {text} {{{style_name}{override_str}}}")
            else:
                lines.append(f"### [p{para_idx}] {text} {{{style_name}}}")
            para_idx += 1

        elif tag == 'tbl':
            if overview:
                table_idx += 1
                continue
            table = _table_map.get(id(child))
            if table is None:
                table_idx += 1
                continue
            table_lines = _extract_table_markdown(table, table_idx)
            lines.extend(table_lines)
            lines.append("")  # 表格后空行
            table_idx += 1

    return "\n".join(lines)

if __name__ == "__main__":
    import sys
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    filepath = sys.argv[1]
    overview = "--overview" in sys.argv
    print(extract_rich_markdown(filepath, overview))

__all__ = [
    'extract_rich_markdown',
    '_extract_style_props',
    '_extract_run_overrides',
    '_extract_cell_overrides',
    '_extract_table_markdown',
]
```

- [ ] **Step 4: 运行所有 reader 测试，验证通过**

```
pytest tests/test_docx_reader.py -v
```
预期：所有测试 PASS

- [ ] **Step 5: 集成验证（人工）**

```
python scripts/docx_reader.py examples/投标书_技术部分.docx | head -100
```
预期：输出中可见 `<!-- TABLE t0:` 和 `[t0r0c0]` 格式的单元格 ID，表格前后有段落内容。

- [ ] **Step 6: Commit**

```
git add scripts/docx_reader.py tests/test_docx_reader.py
git commit -m "feat(reader): 按文档顺序交织输出表格内容，带单元格 ID 标注"
```

---

### Task 2: 扩展 docx_editor.py — 新增单元格内容编辑操作

**Files:**
- Modify: `scripts/docx_editor.py`
- Test: `tests/test_docx_editor_table.py`

**Interfaces:**
- Consumes: `doc.tables[n].rows[n].cells[n]` — python-docx 标准 API
- Produces: DSL op `replace_cell_text`：`{"op": "replace_cell_text", "target": "tNrNcN", "find": str, "replace": str}`
- Produces: DSL op `rewrite_cell`：`{"op": "rewrite_cell", "target": "tNrNcN", "content": str}`
- Produces: `_parse_cell_target(target: str) -> tuple[int, int, int]` — 解析 tNrNcN，失败抛 `ValueError`

---

- [ ] **Step 1: 写失败测试**

在 `tests/test_docx_editor_table.py` 末尾追加：

```python
def test_replace_cell_text(tmp_path):
    """replace_cell_text 应精确替换指定单元格内的文本"""
    doc = docx.Document()
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = 'Original'
    table.cell(0, 1).text = 'Keep'
    table.cell(1, 0).text = 'Also Keep'
    table.cell(1, 1).text = 'Target Old'
    doc_path = tmp_path / "cell_replace.docx"
    doc.save(str(doc_path))

    from docx_editor import apply_operations
    ops = [{"op": "replace_cell_text", "target": "t0r1c1", "find": "Old", "replace": "New"}]
    out_path = tmp_path / "cell_replace_out.docx"
    apply_operations(str(doc_path), ops, str(out_path))

    out_doc = docx.Document(str(out_path))
    assert out_doc.tables[0].cell(1, 1).text == 'Target New'
    assert out_doc.tables[0].cell(0, 0).text == 'Original'  # 其他单元格不变


def test_rewrite_cell(tmp_path):
    """rewrite_cell 应整体重写指定单元格内容"""
    doc = docx.Document()
    table = doc.add_table(rows=1, cols=2)
    table.cell(0, 0).text = 'Old Content'
    table.cell(0, 1).text = 'Unchanged'
    doc_path = tmp_path / "cell_rewrite.docx"
    doc.save(str(doc_path))

    from docx_editor import apply_operations
    ops = [{"op": "rewrite_cell", "target": "t0r0c0", "content": "Brand New"}]
    out_path = tmp_path / "cell_rewrite_out.docx"
    apply_operations(str(doc_path), ops, str(out_path))

    out_doc = docx.Document(str(out_path))
    assert out_doc.tables[0].cell(0, 0).text == 'Brand New'
    assert out_doc.tables[0].cell(0, 1).text == 'Unchanged'


def test_invalid_cell_target(tmp_path):
    """非法 target 格式应抛出 ValueError"""
    doc = docx.Document()
    doc.add_table(rows=1, cols=1)
    doc_path = tmp_path / "invalid_target.docx"
    doc.save(str(doc_path))

    from docx_editor import apply_operations
    import pytest
    ops = [{"op": "replace_cell_text", "target": "p0", "find": "x", "replace": "y"}]
    with pytest.raises(ValueError, match="Invalid cell target"):
        apply_operations(str(doc_path), ops, str(tmp_path / "out.docx"))


def test_cell_target_out_of_range(tmp_path):
    """越界索引应抛出 IndexError"""
    doc = docx.Document()
    doc.add_table(rows=1, cols=1)
    doc_path = tmp_path / "oob.docx"
    doc.save(str(doc_path))

    from docx_editor import apply_operations
    import pytest
    ops = [{"op": "rewrite_cell", "target": "t9r0c0", "content": "x"}]
    with pytest.raises(IndexError):
        apply_operations(str(doc_path), ops, str(tmp_path / "out.docx"))
```

- [ ] **Step 2: 运行测试，验证失败**

```
pytest tests/test_docx_editor_table.py::test_replace_cell_text tests/test_docx_editor_table.py::test_rewrite_cell tests/test_docx_editor_table.py::test_invalid_cell_target -v
```
预期：FAILED（操作未定义）

- [ ] **Step 3: 在 `scripts/docx_editor.py` 中添加辅助函数和两个新操作**

在文件顶部的 `import` 区域末尾添加（在现有 import 之后）：
```python
import re
```
（如已存在则跳过）

在 `_backup_file` 函数之前（约第845行之前），添加以下三个函数：

```python
def _parse_cell_target(target: str):
    """
    解析 'tNrNcN' 格式的单元格 target，返回 (table_idx, row_idx, col_idx)。
    格式非法时抛出 ValueError；索引越界时由调用方抛出 IndexError。
    """
    m = re.fullmatch(r't(\d+)r(\d+)c(\d+)', target)
    if not m:
        raise ValueError(
            f"Invalid cell target: {target!r}. Expected format: tNrNcN (e.g. t0r1c2)"
        )
    return int(m.group(1)), int(m.group(2)), int(m.group(3))


def _apply_replace_cell_text(doc, op):
    """
    replace_cell_text: 在指定单元格内执行文本替换。
    op 必须包含: target (tNrNcN), find (str), replace (str)
    """
    t_idx, r_idx, c_idx = _parse_cell_target(op['target'])
    tables = doc.tables
    if t_idx >= len(tables):
        raise IndexError(
            f"Table index {t_idx} out of range (document has {len(tables)} table(s))"
        )
    rows = tables[t_idx].rows
    if r_idx >= len(rows):
        raise IndexError(
            f"Row index {r_idx} out of range (table t{t_idx} has {len(rows)} row(s))"
        )
    cells = rows[r_idx].cells
    if c_idx >= len(cells):
        raise IndexError(
            f"Col index {c_idx} out of range (row t{t_idx}r{r_idx} has {len(cells)} cell(s))"
        )
    cell = cells[c_idx]
    find_txt = op['find']
    repl_txt = op['replace']
    for para in cell.paragraphs:
        _replace_text_in_runs(para, find_txt, repl_txt)


def _apply_rewrite_cell(doc, op):
    """
    rewrite_cell: 整体重写指定单元格的纯文本内容。
    op 必须包含: target (tNrNcN), content (str)
    """
    t_idx, r_idx, c_idx = _parse_cell_target(op['target'])
    tables = doc.tables
    if t_idx >= len(tables):
        raise IndexError(
            f"Table index {t_idx} out of range (document has {len(tables)} table(s))"
        )
    rows = tables[t_idx].rows
    if r_idx >= len(rows):
        raise IndexError(
            f"Row index {r_idx} out of range (table t{t_idx} has {len(rows)} row(s))"
        )
    cells = rows[r_idx].cells
    if c_idx >= len(cells):
        raise IndexError(
            f"Col index {c_idx} out of range (row t{t_idx}r{r_idx} has {len(cells)} cell(s))"
        )
    cell = cells[c_idx]
    # 清空所有段落后写入新内容
    for para in cell.paragraphs:
        para.text = ''
    if cell.paragraphs:
        cell.paragraphs[0].text = op['content']
    else:
        cell.add_paragraph(op['content'])
```

在 `apply_operations` 函数的 `elif op['op'] == 'update_style_definition':` 分支之后、`doc.save(outpath)` 之前，添加：

```python
        elif op['op'] == 'replace_cell_text':
            _apply_replace_cell_text(doc, op)
        elif op['op'] == 'rewrite_cell':
            _apply_rewrite_cell(doc, op)
```

- [ ] **Step 4: 运行所有表格编辑器测试，验证通过**

```
pytest tests/test_docx_editor_table.py -v
```
预期：所有测试 PASS

- [ ] **Step 5: Commit**

```
git add scripts/docx_editor.py tests/test_docx_editor_table.py
git commit -m "feat(editor): 新增 replace_cell_text 和 rewrite_cell 单元格编辑 DSL 操作"
```

---

### Task 3: 扩展 docx_diff.py — 新增表格内容差异对比

**Files:**
- Modify: `scripts/docx_diff.py`
- Test: `tests/test_docx_diff.py`

**Interfaces:**
- Consumes: `doc.tables` — python-docx 标准 API
- Produces: `compare_tables(doc1, doc2) -> str` — 返回 `## Tables` Markdown 章节字符串（空字符串表示无差异）

---

- [ ] **Step 1: 查看现有 test_docx_diff.py 的结构**

```
pytest tests/test_docx_diff.py -v --collect-only
```
了解现有测试用例名称，避免命名冲突。

- [ ] **Step 2: 写失败测试**

在 `tests/test_docx_diff.py` 末尾追加：

```python
def test_diff_detects_table_cell_change(tmp_path):
    """表格单元格内容变化应出现在 diff 报告的 ## Tables 章节"""
    import docx as _docx
    from docx_diff import generate_rich_diff

    def make_doc(path, cell_text):
        doc = _docx.Document()
        table = doc.add_table(rows=2, cols=2)
        table.cell(0, 0).text = 'H1'
        table.cell(0, 1).text = 'H2'
        table.cell(1, 0).text = 'Unchanged'
        table.cell(1, 1).text = cell_text
        doc.save(path)

    p1 = str(tmp_path / "doc1.docx")
    p2 = str(tmp_path / "doc2.docx")
    make_doc(p1, 'Old Value')
    make_doc(p2, 'New Value')

    report = generate_rich_diff(p1, p2)

    assert '## Tables' in report
    assert 't0r1c1' in report
    assert 'Old Value' in report
    assert 'New Value' in report


def test_diff_no_table_changes(tmp_path):
    """两文档表格相同时 Tables 章节应不包含变更"""
    import docx as _docx
    from docx_diff import generate_rich_diff

    def make_doc(path):
        doc = _docx.Document()
        table = doc.add_table(rows=1, cols=1)
        table.cell(0, 0).text = 'Same'
        doc.save(path)

    p1 = str(tmp_path / "same1.docx")
    p2 = str(tmp_path / "same2.docx")
    make_doc(p1)
    make_doc(p2)

    report = generate_rich_diff(p1, p2)

    # Tables 章节存在但无变更行
    if '## Tables' in report:
        # 不应有 → 箭头（表示内容变化）
        tables_section = report.split('## Tables')[-1]
        assert '→' not in tables_section


def test_diff_table_count_mismatch(tmp_path):
    """两文档表格数量不一致时应报告新增/删除"""
    import docx as _docx
    from docx_diff import generate_rich_diff

    def make_doc1(path):
        doc = _docx.Document()
        table = doc.add_table(rows=1, cols=1)
        table.cell(0, 0).text = 'Only in doc1'
        doc.save(path)

    def make_doc2(path):
        _docx.Document().save(path)  # 无表格

    p1 = str(tmp_path / "has_table.docx")
    p2 = str(tmp_path / "no_table.docx")
    make_doc1(p1)
    make_doc2(p2)

    report = generate_rich_diff(p1, p2)

    assert '## Tables' in report
    assert '删除' in report or 'deleted' in report.lower() or 't0' in report
```

- [ ] **Step 3: 运行测试，验证失败**

```
pytest tests/test_docx_diff.py::test_diff_detects_table_cell_change -v
```
预期：FAILED（`## Tables` 不在输出中）

- [ ] **Step 4: 在 `scripts/docx_diff.py` 中添加表格对比函数，并集成到 `generate_rich_diff`**

在 `generate_diff` 函数之后（约第56行之后）添加：

```python
def compare_tables(doc1, doc2) -> str:
    """
    对比两个文档的表格内容（按单元格文本）。
    返回 Markdown 格式的 '## Tables' 章节字符串；无差异时仍返回章节头。
    """
    tables1 = doc1.tables
    tables2 = doc2.tables
    max_tables = max(len(tables1), len(tables2))

    if max_tables == 0:
        return ""

    lines = ["## Tables", ""]

    for t_idx in range(max_tables):
        if t_idx >= len(tables1):
            # 文档2 新增的表格
            t2 = tables2[t_idx]
            lines.append(f"### t{t_idx} [新增, {len(t2.columns)}×{len(t2.rows)}]")
            lines.append("+ 整个表格为新增")
            lines.append("")
            continue
        if t_idx >= len(tables2):
            # 文档1 有但文档2 删除的表格
            t1 = tables1[t_idx]
            lines.append(f"### t{t_idx} [删除, {len(t1.columns)}×{len(t1.rows)}]")
            lines.append("- 整个表格被删除")
            lines.append("")
            continue

        t1 = tables1[t_idx]
        t2 = tables2[t_idx]
        rows1 = t1.rows
        rows2 = t2.rows
        max_rows = max(len(rows1), len(rows2))
        cell_diffs = []

        for r_idx in range(max_rows):
            if r_idx >= len(rows1) or r_idx >= len(rows2):
                cell_diffs.append(
                    f"- 行数不同：文档1 {len(rows1)} 行，文档2 {len(rows2)} 行"
                )
                break
            cells1 = rows1[r_idx].cells
            cells2 = rows2[r_idx].cells
            max_cols = max(len(cells1), len(cells2))
            for c_idx in range(max_cols):
                if c_idx >= len(cells1) or c_idx >= len(cells2):
                    cell_diffs.append(
                        f"- [t{t_idx}r{r_idx}] 列数不同：文档1 {len(cells1)} 列，文档2 {len(cells2)} 列"
                    )
                    break
                text1 = cells1[c_idx].text
                text2 = cells2[c_idx].text
                if text1 != text2:
                    cell_diffs.append(
                        f'- [t{t_idx}r{r_idx}c{c_idx}] "{text1}" → "{text2}"'
                    )

        lines.append(
            f"### t{t_idx} ({len(t1.columns)}×{len(t1.rows)})"
        )
        if cell_diffs:
            lines.extend(cell_diffs)
        else:
            lines.append("（无变更）")
        lines.append("")

    return "\n".join(lines)
```

在 `generate_rich_diff` 函数中，将 `return format_diff_report(...)` 一行替换为：

```python
    # Format paragraph report
    para_report = format_diff_report(
        diff_results, original=path1, modified=path2,
        context=context, include_same=context > 0
    )

    # Format table report
    d1 = docx.Document(path1)
    d2 = docx.Document(path2)
    table_report = compare_tables(d1, d2)

    if table_report:
        return para_report + "\n\n" + table_report
    return para_report
```

- [ ] **Step 5: 运行所有 diff 测试，验证通过**

```
pytest tests/test_docx_diff.py -v
```
预期：所有测试 PASS

- [ ] **Step 6: Commit**

```
git add scripts/docx_diff.py tests/test_docx_diff.py
git commit -m "feat(diff): 新增表格内容逐单元格差异对比，在报告中输出 Tables 章节"
```

---

### Task 4: 更新 SKILL.md — 补充表格读取和编辑说明

**Files:**
- Modify: `SKILL.md`

**Interfaces:**
- 无代码接口，仅文档更新

---

- [ ] **Step 1: 修改 SKILL.md 的 Workflow 1 Step 1 说明**

在 `SKILL.md` 第50行（`This outputs rich-text Markdown with paragraph IDs (p0, p1, ...).`）之后追加：

```markdown
Tables in the document are rendered inline in reading order as Markdown tables with cell IDs:
- `[t0r1c2]` = table 0, row 1, column 2
- Table header (row 0) is shown in **bold**
- Format overrides (bold, italic, font) are annotated inline as `{...}`
```

- [ ] **Step 2: 修改 SKILL.md 的 Workflow 1 Step 2 示例操作**

在现有的 JSON 示例中（约第58-65行）追加表格编辑操作示例：

```json
[
  {"op": "replace_text", "find": "old", "replace": "new"},
  {"op": "rewrite_paragraph", "target": "p0", "content": "New content"},
  {"op": "set_font", "target": "p0", "name": "Arial", "bold": true},
  {"op": "set_paragraph_format", "target": "p0", "alignment": "center"},
  {"op": "replace_cell_text", "target": "t0r1c1", "find": "旧值", "replace": "新值"},
  {"op": "rewrite_cell", "target": "t0r2c0", "content": "全新内容"}
]
```

- [ ] **Step 3: 更新 Core Scripts 表格，补充 `_extract_cell_overrides` 和 `_extract_table_markdown` 说明**

无需改动（这两个是内部辅助函数，不需要在 SKILL.md 中公开列出）。

- [ ] **Step 4: Commit**

```
git add SKILL.md
git commit -m "docs(skill): 更新 Workflow 1 说明，补充表格 ID 读取和单元格编辑操作示例"
```

---

### Task 5: 全量回归验证

- [ ] **Step 1: 运行全部测试套件**

```
pytest tests/ -v
```
预期：全部 PASS，无回归。

- [ ] **Step 2: 集成验证**

```
python scripts/docx_reader.py examples/投标书_技术部分.docx
```
人工检查输出：
- 段落和表格交织排列
- 每个表格前有 `<!-- TABLE t{n}: ... -->` 注释
- 单元格 ID 格式正确（如 `[t0r0c0]`）
- 表头行内容加粗

- [ ] **Step 3: 最终 commit（如有未提交内容）**

```
git status
```
确认所有更改已提交。
