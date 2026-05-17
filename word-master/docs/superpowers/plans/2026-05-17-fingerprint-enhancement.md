# Fingerprint Enhancement Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 扩展 `style_analyzer.py` 的格式指纹，新增 `font`、`space_before`、`space_after`、`line_spacing` 四个字段，并更新聚类 key 和评分函数，使样式迁移对中文文档的"假样式"场景有更高的区分度。

**Architecture:** 改动集中在两个脚本中的三个函数。`compute_effective_fingerprint()` 新增四个维度的提取逻辑；`_fingerprint_key()` 将 `font` 加入聚类 key；`style_transfer.py` 中的 `_score()` 按新的 7 维权重重写。所有变更向后兼容（`fingerprints.json` 和 `style_profile.json` 的顶层结构不变）。

**Tech Stack:** Python 3, python-docx, pytest

**Spec:** `docs/superpowers/specs/2026-05-17-fingerprint-enhancement-design.md`

---

## File Map

| 文件 | 操作 | 说明 |
|------|------|------|
| `skills/word-master/scripts/style_analyzer.py` | Modify | 扩展 `compute_effective_fingerprint()` 和 `_fingerprint_key()` |
| `skills/word-master/scripts/style_transfer.py` | Modify | 重写 `_score()` 评分函数 |
| `tests/test_style_analyzer.py` | Modify | 新增针对 font/spacing 字段的测试 |
| `tests/test_style_transfer.py` | Modify | 新增针对新 `_score()` 逻辑的测试 |

---

## Task 1: 扩展 `compute_effective_fingerprint()` — 新增 font 字段

**Files:**
- Modify: `skills/word-master/scripts/style_analyzer.py`
- Test: `tests/test_style_analyzer.py`

- [ ] **Step 1: 写失败测试 — run-level font name 应被提取**

在 `tests/test_style_analyzer.py` 末尾追加：

```python
def test_font_from_run_override():
    """Run-level font name should be captured in fingerprint."""
    doc = docx.Document()
    p = doc.add_paragraph()
    run = p.add_run("标题文字")
    run.font.name = "黑体"
    fp = compute_effective_fingerprint(p)
    assert fp["font"] == "黑体"

def test_font_null_when_not_set():
    """font should be None when neither run nor style sets a font."""
    doc = docx.Document()
    p = doc.add_paragraph()
    p.add_run("plain text")
    fp = compute_effective_fingerprint(p)
    assert "font" in fp  # key exists even if null
```

- [ ] **Step 2: 运行测试，确认失败**

```bash
cd word-master
python -m pytest tests/test_style_analyzer.py::test_font_from_run_override tests/test_style_analyzer.py::test_font_null_when_not_set -v
```

预期：`FAILED` — `KeyError: 'font'`

- [ ] **Step 3: 在 `compute_effective_fingerprint()` 末尾的 return 前加入 font 提取逻辑**

在 `style_analyzer.py` 的 `compute_effective_fingerprint()` 函数中，在 `eff_align` 之后、`return` 之前加入：

```python
# Font family
run_font = None
for run in paragraph.runs:
    if run.font.name is not None:
        run_font = run.font.name
        break
style_font_name = style_font.name if style_font and hasattr(style_font, 'name') else None
eff_font = run_font if run_font is not None else style_font_name
```

并在 `return` 字典中加入 `"font": eff_font`：

```python
return {
    "size": size_str,
    "bold": eff_bold,
    "italic": eff_italic,
    "align": eff_align,
    "font": eff_font,
}
```

- [ ] **Step 4: 运行测试，确认通过**

```bash
python -m pytest tests/test_style_analyzer.py::test_font_from_run_override tests/test_style_analyzer.py::test_font_null_when_not_set -v
```

预期：`PASSED`

- [ ] **Step 5: 确认旧测试仍然通过**

```bash
python -m pytest tests/test_style_analyzer.py -v
```

预期：全部 `PASSED`

- [ ] **Step 6: Commit**

```bash
git add tests/test_style_analyzer.py skills/word-master/scripts/style_analyzer.py
git commit -m "feat(fingerprint): add font field to compute_effective_fingerprint"
```

---

## Task 2: 扩展 `compute_effective_fingerprint()` — 新增 spacing 和 line_spacing 字段

**Files:**
- Modify: `skills/word-master/scripts/style_analyzer.py`
- Test: `tests/test_style_analyzer.py`

- [ ] **Step 1: 写失败测试 — 段落间距应被提取**

在 `tests/test_style_analyzer.py` 末尾追加：

```python
def test_space_before_from_paragraph():
    """space_before should be extracted from paragraph format."""
    from docx.shared import Pt
    doc = docx.Document()
    p = doc.add_paragraph()
    p.add_run("标题")
    p.paragraph_format.space_before = Pt(12)
    fp = compute_effective_fingerprint(p)
    assert fp["space_before"] == "12.0pt"

def test_space_before_null_when_not_set():
    """space_before should be None when not set."""
    doc = docx.Document()
    p = doc.add_paragraph()
    p.add_run("正文")
    fp = compute_effective_fingerprint(p)
    assert "space_before" in fp
    assert "space_after" in fp
    assert "line_spacing" in fp

def test_line_spacing_absolute_value_is_null():
    """Absolute line spacing (Pt value) should be normalized to None."""
    from docx.shared import Pt
    from docx.enum.text import WD_LINE_SPACING
    doc = docx.Document()
    p = doc.add_paragraph()
    p.add_run("文字")
    p.paragraph_format.line_spacing = Pt(24)  # absolute, not a multiplier
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    fp = compute_effective_fingerprint(p)
    assert fp["line_spacing"] is None  # absolute mode not supported
```

- [ ] **Step 2: 运行测试，确认失败**

```bash
python -m pytest tests/test_style_analyzer.py::test_space_before_from_paragraph tests/test_style_analyzer.py::test_space_before_null_when_not_set tests/test_style_analyzer.py::test_line_spacing_absolute_value_is_null -v
```

预期：`FAILED` — `KeyError: 'space_before'`

- [ ] **Step 3: 在 `compute_effective_fingerprint()` 的 return 前加入 spacing 提取逻辑**

在 font 提取逻辑之后、`return` 之前加入：

```python
# Spacing fields (paragraph-level)
pf = paragraph.paragraph_format
style_pf2 = style_pf  # already computed above

def _get_pt_str(val):
    """Convert Pt value to 'Xpt' string, or None if not set."""
    if val is None:
        return None
    try:
        return f"{val.pt}pt"
    except AttributeError:
        return None

sb_val = pf.space_before if pf else None
if sb_val is None and style_pf2:
    sb_val = style_pf2.space_before
eff_space_before = _get_pt_str(sb_val)

sa_val = pf.space_after if pf else None
if sa_val is None and style_pf2:
    sa_val = style_pf2.space_after
eff_space_after = _get_pt_str(sa_val)

# Line spacing: only support multiplier (float), not absolute Pt
ls_val = pf.line_spacing if pf else None
if ls_val is None and style_pf2:
    ls_val = style_pf2.line_spacing
# python-docx: if line_spacing is a Pt object, it's absolute mode — not supported
if ls_val is not None:
    try:
        ls_val.pt  # this only works if it's a Pt (absolute), which we don't support
        ls_val = None
    except AttributeError:
        pass  # it's a float (multiplier mode) — keep it
eff_line_spacing = ls_val
```

并更新 `return` 字典：

```python
return {
    "size": size_str,
    "bold": eff_bold,
    "italic": eff_italic,
    "align": eff_align,
    "font": eff_font,
    "space_before": eff_space_before,
    "space_after": eff_space_after,
    "line_spacing": eff_line_spacing,
}
```

- [ ] **Step 4: 运行测试，确认通过**

```bash
python -m pytest tests/test_style_analyzer.py::test_space_before_from_paragraph tests/test_style_analyzer.py::test_space_before_null_when_not_set tests/test_style_analyzer.py::test_line_spacing_absolute_value_is_null -v
```

预期：全部 `PASSED`

- [ ] **Step 5: 确认全量测试仍然通过**

```bash
python -m pytest tests/test_style_analyzer.py -v
```

预期：全部 `PASSED`

- [ ] **Step 6: Commit**

```bash
git add tests/test_style_analyzer.py skills/word-master/scripts/style_analyzer.py
git commit -m "feat(fingerprint): add space_before, space_after, line_spacing to fingerprint"
```

---

## Task 3: 更新聚类 Key — 加入 font

**Files:**
- Modify: `skills/word-master/scripts/style_analyzer.py`
- Test: `tests/test_style_analyzer.py`

- [ ] **Step 1: 写失败测试 — 不同字体的相同尺寸段落应被分为不同聚类**

```python
def test_different_fonts_form_different_clusters(tmp_path):
    """Paragraphs with same size but different fonts should be in separate clusters."""
    doc = docx.Document()
    # 5 paragraphs with 黑体
    for i in range(5):
        p = doc.add_paragraph()
        run = p.add_run(f"黑体段落{i}")
        run.font.size = Pt(14)
        run.font.name = "黑体"
    # 5 paragraphs with 宋体
    for i in range(5):
        p = doc.add_paragraph()
        run = p.add_run(f"宋体段落{i}")
        run.font.size = Pt(14)
        run.font.name = "宋体"
    path = tmp_path / "test_fonts.docx"
    doc.save(str(path))
    result = extract_fingerprints(str(path), min_cluster_size=4)
    # Should have 2 clusters: one for 黑体, one for 宋体
    assert len(result) == 2
    fonts = {r["fingerprint"]["font"] for r in result}
    assert "黑体" in fonts
    assert "宋体" in fonts
```

- [ ] **Step 2: 运行测试，确认失败（旧 key 不含 font，两组会被合并）**

```bash
python -m pytest tests/test_style_analyzer.py::test_different_fonts_form_different_clusters -v
```

预期：`FAILED` — `assert len(result) == 2` 失败（旧 key 只有 size/bold/italic/align，两组被合并成 1）

- [ ] **Step 3: 更新 `_fingerprint_key()`**

```python
def _fingerprint_key(fp):
    """Convert fingerprint dict to a hashable key for grouping."""
    return (fp.get("size"), fp.get("bold"), fp.get("italic"), fp.get("align"), fp.get("font"))
```

- [ ] **Step 4: 运行测试，确认通过**

```bash
python -m pytest tests/test_style_analyzer.py::test_different_fonts_form_different_clusters -v
```

预期：`PASSED`

- [ ] **Step 5: 确认全量测试**

```bash
python -m pytest tests/test_style_analyzer.py -v
```

预期：全部 `PASSED`

- [ ] **Step 6: Commit**

```bash
git add tests/test_style_analyzer.py skills/word-master/scripts/style_analyzer.py
git commit -m "feat(fingerprint): add font to clustering key in _fingerprint_key"
```

---

## Task 4: 更新 `_score()` — 新增维度权重

**Files:**
- Modify: `skills/word-master/scripts/style_transfer.py`
- Test: `tests/test_style_transfer.py`

- [ ] **Step 1: 了解现有 `_score()` 实现**

阅读 `skills/word-master/scripts/style_transfer.py` 中的 `_score()` 函数（当前约 L11-L38），理解其结构后再修改。

- [ ] **Step 2: 写失败测试 — font 匹配应影响评分**

在 `tests/test_style_transfer.py` 中找到 `_score` 相关测试，追加：

```python
from style_transfer import _score

def test_score_font_mismatch_reduces_score():
    """Mismatched font should reduce the score compared to matching font."""
    fp = {"size": "14.0pt", "bold": False, "italic": None, "align": "center",
          "font": "黑体", "space_before": None, "space_after": None}
    role_same_font = {"size": "14.0pt", "bold": False, "italic": None, "align": "center",
                      "font": "黑体", "space_before": None, "space_after": None}
    role_diff_font = {"size": "14.0pt", "bold": False, "italic": None, "align": "center",
                      "font": "宋体", "space_before": None, "space_after": None}
    score_same = _score(fp, role_same_font)
    score_diff = _score(fp, role_diff_font)
    assert score_same > score_diff

def test_score_total_weight_is_one():
    """When all fields match, score should be 1.0."""
    fp = {"size": "14.0pt", "bold": True, "italic": False, "align": "center",
          "font": "黑体", "space_before": "12.0pt", "space_after": "6.0pt"}
    assert _score(fp, fp) == pytest.approx(1.0, abs=0.01)
```

- [ ] **Step 3: 运行测试，确认失败**

```bash
python -m pytest tests/test_style_transfer.py::test_score_font_mismatch_reduces_score tests/test_style_transfer.py::test_score_total_weight_is_one -v
```

预期：`FAILED` — 旧 `_score()` 不认识 font/space_before/space_after 字段

- [ ] **Step 4: 重写 `_score()` 函数**

按照设计文档中的权重表重写（总计 = 1.0）：

```python
def _score(fp, role_fp):
    """
    Calculate similarity score between two fingerprints (0.0 ~ 1.0).
    Weights: size=2/8, bold=1/8, italic=1/8, align=1/8, font=2/8,
             space_before=0.5/8, space_after=0.5/8.
    line_spacing is recorded in fingerprint but not scored.
    Size uses linear decay: max(0, 1 - diff/10).
    """
    total = 0.0

    # size (weight=2/8=0.25): linear decay
    s1 = fp.get("size")
    s2 = role_fp.get("size")
    if s1 and s2:
        try:
            diff = abs(float(s1.rstrip("pt")) - float(s2.rstrip("pt")))
            total += max(0.0, 1.0 - diff / 10.0) * 0.25
        except ValueError:
            pass
    elif s1 == s2:
        total += 0.25

    # bold (weight=1/8=0.125)
    total += 0.125 if fp.get("bold") == role_fp.get("bold") else 0.0

    # italic (weight=1/8=0.125)
    total += 0.125 if fp.get("italic") == role_fp.get("italic") else 0.0

    # align (weight=1/8=0.125)
    total += 0.125 if fp.get("align") == role_fp.get("align") else 0.0

    # font (weight=2/8=0.25)
    f1 = fp.get("font")
    f2 = role_fp.get("font")
    if f1 is None and f2 is None:
        total += 0.25
    elif f1 == f2:
        total += 0.25
    # else: mismatch, 0

    # space_before (weight=0.5/8=0.0625)
    sb1 = fp.get("space_before")
    sb2 = role_fp.get("space_before")
    if sb1 is None and sb2 is None:
        total += 0.0625
    elif sb1 == sb2:
        total += 0.0625

    # space_after (weight=0.5/8=0.0625)
    sa1 = fp.get("space_after")
    sa2 = role_fp.get("space_after")
    if sa1 is None and sa2 is None:
        total += 0.0625
    elif sa1 == sa2:
        total += 0.0625

    return total
```

- [ ] **Step 5: 运行测试，确认通过**

```bash
python -m pytest tests/test_style_transfer.py::test_score_font_mismatch_reduces_score tests/test_style_transfer.py::test_score_total_weight_is_one -v
```

预期：全部 `PASSED`

- [ ] **Step 6: 确认全量测试（包括旧的 style_transfer 和 style_analyzer 测试）**

```bash
python -m pytest tests/ -v
```

预期：全部 `PASSED`

- [ ] **Step 7: Commit**

```bash
git add tests/test_style_transfer.py skills/word-master/scripts/style_transfer.py
git commit -m "feat(fingerprint): rewrite _score() with 7-dimension weights including font and spacing"
```

---

## Task 5: 端到端验证

**Files:**
- 使用已有的测试文档 `examples/招标书_技术文件_陕西公司店塔电厂MIS系统升级改造项目燃料与检修管理系统、备份等设备.docx`

- [ ] **Step 1: 重新运行 `style_analyzer.py`，观察新的指纹输出**

```bash
cd word-master
python skills/word-master/scripts/style_analyzer.py "examples/招标书_技术文件_陕西公司店塔电厂MIS系统升级改造项目燃料与检修管理系统、备份等设备.docx" --output examples/fingerprints_v2.json
```

- [ ] **Step 2: 验证指纹质量**

查看 `examples/fingerprints_v2.json`，人工检查：
- 各组的 `font` 字段是否有效填充（非全为 null）
- 指纹组的数量是否比旧版 `fingerprints.json`（5组）增加
- 不同语义角色（标题 vs 正文）是否被有效区分到不同组

- [ ] **Step 3: Commit 最终结果**

```bash
git add examples/fingerprints_v2.json
git commit -m "test: regenerate fingerprints with enhanced v2 extractor for validation"
```
