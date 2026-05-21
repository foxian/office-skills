# Heading-Aware 指纹提取 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 为 `style_analyzer.py` 的 `extract_fingerprints()` 新增 `heading_aware` 模式——以段落大纲级别（outlineLvl）而非样式名分组标题指纹，并对同级多段取平均。

**Architecture:** 新增工具函数 `_get_outline_level(paragraph)`（读取 outlineLvl，支持样式继承链回退）和 `_average_fingerprints(fps)`（多段聚合）；`extract_fingerprints()` 新增 `heading_aware=False` 参数；CLI 新增 `--heading-aware` flag。默认行为完全不变。

**Tech Stack:** Python 3.10+, python-docx, lxml（项目已有）

**Spec:** `docs/superpowers/specs/2026-05-21-heading-aware-fingerprint-design.md`

---

## 文件变更

| 操作 | 文件 |
|------|------|
| Modify | `word-master/skills/word-master/scripts/style_analyzer.py` |
| Modify | `word-master/tests/test_style_analyzer.py` |

---

## Task 1: 新增 `_get_outline_level` 函数

**Files:**
- Modify: `word-master/skills/word-master/scripts/style_analyzer.py`
- Test: `word-master/tests/test_style_analyzer.py`

- [ ] **Step 1: 写失败测试**

在 `word-master/tests/test_style_analyzer.py` 末尾追加：

```python
from docx.oxml.ns import qn as _qn

def test_get_outline_level_returns_none_for_normal_para():
    """Normal paragraph without outlineLvl returns None."""
    doc = docx.Document()
    p = doc.add_paragraph("普通段落")
    from style_analyzer import _get_outline_level
    assert _get_outline_level(p) is None

def test_get_outline_level_reads_direct_xml():
    """Paragraph with explicit outlineLvl=2 returns 2."""
    doc = docx.Document()
    p = doc.add_paragraph("三级标题")
    pPr = p._element.get_or_add_pPr()
    outlineLvl = pPr.get_or_add_outlineLvl()
    outlineLvl.set(_qn('w:val'), '2')
    from style_analyzer import _get_outline_level
    assert _get_outline_level(p) == 2

def test_get_outline_level_inherits_from_style():
    """Paragraph using Heading 3 style (no direct outlineLvl) returns 2 via style chain."""
    doc = docx.Document()
    p = doc.add_paragraph("继承样式标题", style="Heading 3")
    from style_analyzer import _get_outline_level
    # Heading 3 style carries outlineLvl=2 in its XML definition
    assert _get_outline_level(p) == 2
```

- [ ] **Step 2: 运行确认失败**

```
python -m pytest word-master/tests/test_style_analyzer.py::test_get_outline_level_returns_none_for_normal_para word-master/tests/test_style_analyzer.py::test_get_outline_level_reads_direct_xml word-master/tests/test_style_analyzer.py::test_get_outline_level_inherits_from_style -v
```

预期：`ImportError: cannot import name '_get_outline_level'`

- [ ] **Step 3: 实现 `_get_outline_level`**

在 `style_analyzer.py` 的 `_detect_list_type` 函数之前插入（imports 区已有 `from docx.oxml.ns import qn`）：

```python
def _get_outline_level(paragraph) -> int | None:
    """
    Read paragraph outline level (0=Heading 1, ..., 8=Heading 9).
    Strategy:
      1. Check paragraph-level XML directly.
      2. Walk style inheritance chain (para.style -> base_style -> ...),
         up to 10 hops to guard against circular references.
    Returns int 0-8 or None (not a heading).
    """
    def _read_outline_lvl_from_pPr(pPr_elem):
        if pPr_elem is None:
            return None
        ol = pPr_elem.find(qn('w:outlineLvl'))
        if ol is None:
            return None
        val = ol.get(qn('w:val'))
        if val is None:
            return None
        try:
            level = int(val)
            return level if 0 <= level <= 8 else None
        except ValueError:
            return None

    # 1. Paragraph-level direct attribute
    para_pPr = paragraph._element.find(qn('w:pPr'))
    level = _read_outline_lvl_from_pPr(para_pPr)
    if level is not None:
        return level

    # 2. Style inheritance chain
    style = paragraph.style
    for _ in range(10):
        if style is None:
            break
        style_pPr = style.element.find(qn('w:pPr'))
        level = _read_outline_lvl_from_pPr(style_pPr)
        if level is not None:
            return level
        style = style.base_style

    return None
```

- [ ] **Step 4: 运行确认通过**

```
python -m pytest word-master/tests/test_style_analyzer.py::test_get_outline_level_returns_none_for_normal_para word-master/tests/test_style_analyzer.py::test_get_outline_level_reads_direct_xml word-master/tests/test_style_analyzer.py::test_get_outline_level_inherits_from_style -v
```

预期：3 PASSED

- [ ] **Step 5: 全量回归**

```
python -m pytest word-master/tests/test_style_analyzer.py -v
```

预期：所有原有测试 + 3 个新测试 PASSED

- [ ] **Step 6: 提交**

```
git add word-master/skills/word-master/scripts/style_analyzer.py word-master/tests/test_style_analyzer.py
git commit -m "feat: add _get_outline_level - reads paragraph outline level with style chain fallback"
```

---

## Task 2: 新增 `_average_fingerprints` 聚合函数

**Files:**
- Modify: `word-master/skills/word-master/scripts/style_analyzer.py`
- Test: `word-master/tests/test_style_analyzer.py`

### 聚合规则（来自 Spec §3）

| 字段 | 规则 |
|------|------|
| `size` | 数值平均，`f"{val:.1f}pt"`；None 段跳过；全 None 结果为 None |
| `bold`, `italic` | 多数投票，None 参与；平局按 False > None > True |
| `align` | 多数投票，None 参与；平局取第一段值 |
| `font` | 最高频，None 跳过；全 None 结果为 None；平局取第一段值 |
| `color` | 最高频，None 参与投票；平局取第一段值 |
| `space_before`, `space_after`, `first_line_indent` | 数值平均，None 视为 0，`f"{val:.1f}pt"` |
| `line_spacing` | float 平均，None 跳过；全 None 结果为 None |

- [ ] **Step 1: 写失败测试**

在 `test_style_analyzer.py` 末尾追加：

```python
def test_average_fingerprints_size_averaging():
    """size field: numeric average, skip None, format to 1 decimal."""
    from style_analyzer import _average_fingerprints
    fps = [
        {"size": "12.0pt", "bold": True, "italic": False, "align": "justify",
         "font": "宋体", "color": None, "space_before": "0.0pt", "space_after": "0.0pt",
         "first_line_indent": "0.0pt", "line_spacing": None},
        {"size": "14.0pt", "bold": True, "italic": False, "align": "justify",
         "font": "宋体", "color": None, "space_before": "0.0pt", "space_after": "0.0pt",
         "first_line_indent": "0.0pt", "line_spacing": None},
    ]
    result = _average_fingerprints(fps)
    assert result["size"] == "13.0pt"

def test_average_fingerprints_size_skips_none():
    """size=None paragraphs are skipped; result is average of remaining."""
    from style_analyzer import _average_fingerprints
    fps = [
        {"size": None, "bold": False, "italic": False, "align": None,
         "font": None, "color": None, "space_before": "0.0pt", "space_after": "0.0pt",
         "first_line_indent": "0.0pt", "line_spacing": None},
        {"size": "12.0pt", "bold": False, "italic": False, "align": None,
         "font": None, "color": None, "space_before": "0.0pt", "space_after": "0.0pt",
         "first_line_indent": "0.0pt", "line_spacing": None},
    ]
    result = _average_fingerprints(fps)
    assert result["size"] == "12.0pt"

def test_average_fingerprints_bold_tiebreak():
    """bold tie: True:1 False:1 => False (False > None > True priority)."""
    from style_analyzer import _average_fingerprints
    base = {"size": None, "italic": False, "align": None, "font": None,
            "color": None, "space_before": "0.0pt", "space_after": "0.0pt",
            "first_line_indent": "0.0pt", "line_spacing": None}
    fps = [{**base, "bold": True}, {**base, "bold": False}]
    result = _average_fingerprints(fps)
    assert result["bold"] == False

def test_average_fingerprints_bold_none_wins():
    """bold: None:2 True:1 => None."""
    from style_analyzer import _average_fingerprints
    base = {"size": None, "italic": False, "align": None, "font": None,
            "color": None, "space_before": "0.0pt", "space_after": "0.0pt",
            "first_line_indent": "0.0pt", "line_spacing": None}
    fps = [{**base, "bold": None}, {**base, "bold": None}, {**base, "bold": True}]
    result = _average_fingerprints(fps)
    assert result["bold"] is None
```

- [ ] **Step 2: 运行确认失败**

```
python -m pytest word-master/tests/test_style_analyzer.py::test_average_fingerprints_size_averaging word-master/tests/test_style_analyzer.py::test_average_fingerprints_size_skips_none word-master/tests/test_style_analyzer.py::test_average_fingerprints_bold_tiebreak word-master/tests/test_style_analyzer.py::test_average_fingerprints_bold_none_wins -v
```

预期：`ImportError: cannot import name '_average_fingerprints'`

- [ ] **Step 3: 实现 `_average_fingerprints`**

在 `_get_outline_level` 之后、`_detect_list_type` 之前插入：

```python
def _average_fingerprints(fps: list[dict]) -> dict:
    """
    Aggregate a list of fingerprint dicts into one representative fingerprint.
    Aggregation rules per field: see Spec §3.
    """
    if not fps:
        return {}

    def _pt_to_float(s):
        if s is None:
            return 0.0
        try:
            return float(s.rstrip("pt"))
        except (ValueError, AttributeError):
            return 0.0

    def _vote(values, tiebreak_order=None, skip_none=False):
        from collections import Counter
        filtered = [v for v in values if v is not None] if skip_none else values
        if not filtered:
            return None
        counts = Counter(filtered)
        max_count = max(counts.values())
        winners = [k for k, v in counts.items() if v == max_count]
        if len(winners) == 1:
            return winners[0]
        if tiebreak_order is not None:
            for priority in tiebreak_order:
                if priority in winners:
                    return priority
        # Fallback: first occurrence in fps
        for v in values:
            if v in winners:
                return v
        return winners[0]

    bold_tiebreak = [False, None, True]

    size_vals = [_pt_to_float(fp.get("size")) for fp in fps if fp.get("size") is not None]
    ls_vals = [fp.get("line_spacing") for fp in fps if fp.get("line_spacing") is not None]

    def _avg_pt(field):
        vals = [_pt_to_float(fp.get(field)) for fp in fps]
        return f"{sum(vals) / len(vals):.1f}pt"

    return {
        "size": f"{sum(size_vals) / len(size_vals):.1f}pt" if size_vals else None,
        "bold": _vote([fp.get("bold") for fp in fps], tiebreak_order=bold_tiebreak),
        "italic": _vote([fp.get("italic") for fp in fps], tiebreak_order=bold_tiebreak),
        "align": _vote([fp.get("align") for fp in fps]),
        "font": _vote([fp.get("font") for fp in fps], skip_none=True),
        "color": _vote([fp.get("color") for fp in fps]),
        "space_before": _avg_pt("space_before"),
        "space_after": _avg_pt("space_after"),
        "first_line_indent": _avg_pt("first_line_indent"),
        "line_spacing": sum(ls_vals) / len(ls_vals) if ls_vals else None,
    }
```

- [ ] **Step 4: 运行确认通过**

```
python -m pytest word-master/tests/test_style_analyzer.py::test_average_fingerprints_size_averaging word-master/tests/test_style_analyzer.py::test_average_fingerprints_size_skips_none word-master/tests/test_style_analyzer.py::test_average_fingerprints_bold_tiebreak word-master/tests/test_style_analyzer.py::test_average_fingerprints_bold_none_wins -v
```

预期：4 PASSED

- [ ] **Step 5: 全量回归**

```
python -m pytest word-master/tests/test_style_analyzer.py -v
```

- [ ] **Step 6: 提交**

```
git add word-master/skills/word-master/scripts/style_analyzer.py word-master/tests/test_style_analyzer.py
git commit -m "feat: add _average_fingerprints - aggregate multiple paragraph fingerprints per heading level"
```

---

## Task 3: 修改 `extract_fingerprints` + CLI flag

**Files:**
- Modify: `word-master/skills/word-master/scripts/style_analyzer.py`
- Test: `word-master/tests/test_style_analyzer.py`

- [ ] **Step 1: 写失败测试**

在 `test_style_analyzer.py` 末尾追加：

```python
def test_heading_aware_mismatched_style_grouped_by_outline(tmp_path):
    """
    Integration: Heading 4 style + outlineLvl=2 grouped as Heading 3.
    Heading 3 style + outlineLvl=2 also grouped as Heading 3.
    Both merge into one Heading 3 group; no Heading 4 group.
    """
    from docx.oxml.ns import qn as _qn
    doc = docx.Document()
    p_wrong = doc.add_paragraph("三级标题(用了Heading4样式)", style="Heading 4")
    pPr = p_wrong._element.get_or_add_pPr()
    ol = pPr.get_or_add_outlineLvl()
    ol.set(_qn('w:val'), '2')
    doc.add_paragraph("三级标题(正确样式)", style="Heading 3")
    for i in range(5):
        doc.add_paragraph(f"正文段落{i}")
    path = tmp_path / "mismatch.docx"
    doc.save(str(path))
    result = extract_fingerprints(str(path), min_cluster_size=1, heading_aware=True)
    heading3 = [r for r in result if r.get("heading_role") == "Heading 3"]
    heading4 = [r for r in result if r.get("heading_role") == "Heading 4"]
    assert len(heading3) == 1, f"Expected 1 Heading 3, got {len(heading3)}"
    assert len(heading4) == 0, f"Expected 0 Heading 4, got {len(heading4)}"

def test_heading_aware_body_still_clustered(tmp_path):
    """Body paragraphs (no outlineLvl) still go through format clustering."""
    from docx.shared import Pt
    doc = docx.Document()
    doc.add_paragraph("一级标题", style="Heading 1")
    for i in range(5):
        p = doc.add_paragraph(f"正文{i}")
        p.runs[0].font.size = Pt(12)
    path = tmp_path / "body.docx"
    doc.save(str(path))
    result = extract_fingerprints(str(path), min_cluster_size=1, heading_aware=True)
    body_groups = [r for r in result if "heading_role" not in r]
    assert len(body_groups) >= 1

def test_heading_aware_false_unchanged(tmp_path):
    """Default heading_aware=False: existing behavior preserved, no heading_role field."""
    doc = docx.Document()
    for i in range(5):
        doc.add_paragraph(f"正文段落{i}")
    path = tmp_path / "default.docx"
    doc.save(str(path))
    result = extract_fingerprints(str(path), min_cluster_size=4)
    for r in result:
        assert "heading_role" not in r
```

- [ ] **Step 2: 运行确认失败**

```
python -m pytest word-master/tests/test_style_analyzer.py::test_heading_aware_mismatched_style_grouped_by_outline word-master/tests/test_style_analyzer.py::test_heading_aware_body_still_clustered word-master/tests/test_style_analyzer.py::test_heading_aware_false_unchanged -v
```

预期：`TypeError: extract_fingerprints() got an unexpected keyword argument 'heading_aware'`

- [ ] **Step 3: 修改 `extract_fingerprints`**

函数签名改为：
```python
def extract_fingerprints(filepath, min_cluster_size=4, heading_aware=False):
```

在 `doc = docx.Document(filepath)` 之后、原有 `clusters = {}` 循环之前，插入：

```python
    if heading_aware:
        heading_paras = {}    # "Heading N" -> list of fingerprint dicts
        heading_examples = {} # "Heading N" -> first paragraph text
        body_clusters = {}    # fingerprint_key -> list of (text, fp)

        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                continue
            outline_level = _get_outline_level(para)
            if outline_level is not None:
                role = f"Heading {outline_level + 1}"
                fp = compute_effective_fingerprint(para)
                if role not in heading_paras:
                    heading_paras[role] = []
                    heading_examples[role] = text
                heading_paras[role].append(fp)
            else:
                fp = compute_effective_fingerprint(para)
                key = _fingerprint_key(fp)
                if key not in body_clusters:
                    body_clusters[key] = []
                body_clusters[key].append((text, fp))

        result = []
        for role in sorted(heading_paras, key=lambda r: int(r.split()[-1])):
            result.append({
                "heading_role": role,
                "fingerprint": _average_fingerprints(heading_paras[role]),
                "example": heading_examples[role],
            })
        for i, (key, members) in enumerate(body_clusters.items()):
            has_chapter_pattern = any(CHAPTER_PATTERN.match(t) for t, _ in members)
            if len(members) < min_cluster_size and not has_chapter_pattern:
                continue
            representative_text, representative_fp = members[0]
            result.append({"id": i, "fingerprint": representative_fp, "example": representative_text})
        return result

    # --- Original mode (heading_aware=False) ---
```

原有 `clusters = {}` 开始的代码保持不动。

- [ ] **Step 4: 修改 CLI `__main__` 块**

将原有 `if __name__ == "__main__":` 块替换为：

```python
if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Extract style fingerprints from a DOCX template.")
    parser.add_argument("filepath", help="Source template DOCX path")
    parser.add_argument("--min-cluster-size", type=int, default=4)
    parser.add_argument("--output", default="fingerprints.json")
    parser.add_argument(
        "--heading-aware",
        action="store_true",
        help=(
            "Group heading paragraphs by outline level (navigation level) instead of style name. "
            "Use when the template has mismatched heading styles but correct navigation structure."
        ),
    )
    args = parser.parse_args()
    fingerprints = extract_fingerprints(
        args.filepath,
        args.min_cluster_size,
        heading_aware=args.heading_aware,
    )
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(fingerprints, f, ensure_ascii=False, indent=2)
    print(f"[INFO] Extracted {len(fingerprints)} fingerprint groups -> {args.output}")
```

- [ ] **Step 5: 运行新测试**

```
python -m pytest word-master/tests/test_style_analyzer.py::test_heading_aware_mismatched_style_grouped_by_outline word-master/tests/test_style_analyzer.py::test_heading_aware_body_still_clustered word-master/tests/test_style_analyzer.py::test_heading_aware_false_unchanged -v
```

预期：3 PASSED

- [ ] **Step 6: 全量回归**

```
python -m pytest word-master/tests/test_style_analyzer.py -v
```

预期：全部 PASSED（含所有原有测试）

- [ ] **Step 7: 验证 CLI**

```
python word-master/skills/word-master/scripts/style_analyzer.py --help
```

预期：输出包含 `--heading-aware` 说明

- [ ] **Step 8: 提交**

```
git add word-master/skills/word-master/scripts/style_analyzer.py word-master/tests/test_style_analyzer.py
git commit -m "feat: extract_fingerprints heading_aware mode - group headings by outlineLvl, average multi-para fingerprints"
```

---

## 验证计划

```
python -m pytest word-master/tests/test_style_analyzer.py -v
```

Task 1 新增 3 条 + Task 2 新增 4 条 + Task 3 新增 3 条 + 原有约 17 条 = 全部 PASSED。
