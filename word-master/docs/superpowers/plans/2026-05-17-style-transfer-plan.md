# Style Transfer Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 新建 `style_analyzer.py` 和 `style_transfer.py` 两个脚本，实现从源模板中提取正文格式规律（格式指纹），通过 LLM 推断语义角色，并将结果作为 `apply_style` DSL 指令委托 `docx_editor.py` 就地应用到目标草稿文档的正文段落上。

**Architecture:** `style_analyzer.py` 只读分析源模板，输出 `fingerprints.json`；LLM 将指纹转为 `style_profile.json`；`style_transfer.py` 作为规划器生成 DSL 操作列表后调用 `docx_editor.apply_operations()` 执行。两个新脚本与现有 `docx_reader.py`、`docx_editor.py` 完全独立，只通过 `apply_operations()` 接口耦合。

**Tech Stack:** python-docx, Python 3, pytest, JSON

---

## 文件结构

```
word-master/
├── skills/word-master/
│   └── scripts/
│       ├── style_analyzer.py     # 新建：格式指纹提取器
│       └── style_transfer.py     # 新建：样式转移总控脚本
└── tests/
    ├── test_style_analyzer.py    # 新建：分析器单元测试
    └── test_style_transfer.py    # 新建：转移器集成测试
```

---

## 任务分解

### 任务 1：实现 `style_analyzer.py` — 有效指纹计算

**Files:**
- Create: `word-master/skills/word-master/scripts/style_analyzer.py`
- Test: `word-master/tests/test_style_analyzer.py`

核心功能：对每个段落，合并 `paragraph.style` 的属性和 run-level 手动覆盖，计算出**有效格式指纹**（effective fingerprint）。

有效值计算规则（以字号为例）：
- 若 run 上有 `font.size`（非 None）→ 使用 run 值
- 否则使用 `paragraph.style.font.size`
- 若仍为 None → 标记为 `null`

- [ ] **Step 1: 编写有效指纹计算的测试**

在 `word-master/tests/test_style_analyzer.py` 中创建：

```python
import docx
from docx.shared import Pt
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'skills', 'word-master', 'scripts'))
from style_analyzer import compute_effective_fingerprint

def test_run_override_wins(tmp_path):
    """Run-level font size should override paragraph style."""
    doc = docx.Document()
    p = doc.add_paragraph()
    run = p.add_run("测试文字")
    run.font.size = Pt(16)
    fp = compute_effective_fingerprint(p)
    assert fp["size"] == "16.0pt"

def test_style_level_used_when_no_run_override(tmp_path):
    """Paragraph style font size used when run has no override."""
    doc = docx.Document()
    p = doc.add_paragraph("正文段落")
    # Normal style default - no explicit run override
    fp = compute_effective_fingerprint(p)
    # size may be None if style has none, but should not crash
    assert "size" in fp

def test_bold_run_override(tmp_path):
    """Bold set on run should be captured in fingerprint."""
    doc = docx.Document()
    p = doc.add_paragraph()
    run = p.add_run("加粗文字")
    run.bold = True
    fp = compute_effective_fingerprint(p)
    assert fp["bold"] == True
```

- [ ] **Step 2: 运行测试确认失败**

```bash
cd word-master
python -m pytest tests/test_style_analyzer.py -v
```
Expected: `ImportError: cannot import name 'compute_effective_fingerprint'`

- [ ] **Step 3: 实现 `compute_effective_fingerprint`**

创建 `word-master/skills/word-master/scripts/style_analyzer.py`：

```python
import docx
import json
import sys
from docx.enum.text import WD_ALIGN_PARAGRAPH


def _get_effective_value(run_val, style_val):
    """Return run value if explicitly set, else fall back to style value."""
    if run_val is not None:
        return run_val
    return style_val


def compute_effective_fingerprint(paragraph):
    """
    计算段落的有效格式指纹，合并 style 属性 和 run-level 手动覆盖。
    取第一个非空 run 的值作为代表（多 run 时取众数逻辑简化为取首个）。
    """
    style = paragraph.style
    style_font = style.font if hasattr(style, 'font') else None
    style_pf = style.paragraph_format if hasattr(style, 'paragraph_format') else None

    # 从 runs 中取有效覆盖值（取第一个有显式设置的 run）
    run_size = None
    run_bold = None
    run_italic = None
    for run in paragraph.runs:
        if run.font.size is not None and run_size is None:
            run_size = run.font.size
        if run.bold is not None and run_bold is None:
            run_bold = run.bold
        if run.italic is not None and run_italic is None:
            run_italic = run.italic

    # 字号
    style_size = style_font.size if style_font and style_font.size else None
    eff_size = run_size or style_size
    size_str = f"{eff_size.pt}pt" if eff_size else None

    # 加粗
    style_bold = style_font.bold if style_font else None
    eff_bold = _get_effective_value(run_bold, style_bold)

    # 斜体
    style_italic = style_font.italic if style_font else None
    eff_italic = _get_effective_value(run_italic, style_italic)

    # 对齐
    align_map = {
        WD_ALIGN_PARAGRAPH.LEFT: "left",
        WD_ALIGN_PARAGRAPH.CENTER: "center",
        WD_ALIGN_PARAGRAPH.RIGHT: "right",
        WD_ALIGN_PARAGRAPH.JUSTIFY: "justify",
    }
    pf_align = paragraph.paragraph_format.alignment
    style_align = style_pf.alignment if style_pf else None
    eff_align_raw = pf_align or style_align
    eff_align = align_map.get(eff_align_raw) if eff_align_raw else None

    return {
        "size": size_str,
        "bold": eff_bold,
        "italic": eff_italic,
        "align": eff_align,
    }
```

- [ ] **Step 4: 运行测试确认通过**

```bash
cd word-master
python -m pytest tests/test_style_analyzer.py -v
```
Expected: 3 tests PASS

- [ ] **Step 5: 提交**

```bash
git add word-master/skills/word-master/scripts/style_analyzer.py word-master/tests/test_style_analyzer.py
git commit -m "feat: style_analyzer - compute_effective_fingerprint"
```

---

### 任务 2：实现 `style_analyzer.py` — 正文过滤与聚类输出

**Files:**
- Modify: `word-master/skills/word-master/scripts/style_analyzer.py`
- Modify: `word-master/tests/test_style_analyzer.py`

- [ ] **Step 1: 编写正文过滤和聚类的测试**

在 `test_style_analyzer.py` 末尾追加：

```python
from style_analyzer import extract_fingerprints
import re

def _make_doc_with_paragraphs(tmp_path, paras):
    """Helper: create a docx with given (text, bold, size_pt) tuples."""
    doc = docx.Document()
    for text, bold, size_pt in paras:
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.bold = bold
        run.font.size = Pt(size_pt)
    path = tmp_path / "test.docx"
    doc.save(path)
    return str(path)

def test_small_cluster_filtered(tmp_path):
    """Paragraphs appearing only once (small cluster) should be filtered out."""
    # 1 big-font paragraph (cover-like), 10 body paragraphs
    paras = [("封面标题", True, 26)]
    paras += [("正文内容第{}段".format(i), False, 12) for i in range(10)]
    path = _make_doc_with_paragraphs(tmp_path, paras)
    result = extract_fingerprints(path, min_cluster_size=4)
    # Only one fingerprint group should survive (the body one)
    assert len(result) == 1
    assert result[0]["fingerprint"]["size"] == "12.0pt"

def test_chapter_pattern_signals_body(tmp_path):
    """Paragraphs matching Chinese chapter pattern should be included."""
    paras = [("第一章 项目概述", True, 14)] * 5
    path = _make_doc_with_paragraphs(tmp_path, paras)
    result = extract_fingerprints(path, min_cluster_size=4)
    assert len(result) == 1
    assert result[0]["example"].startswith("第一章")
```

- [ ] **Step 2: 运行测试确认失败**

```bash
cd word-master
python -m pytest tests/test_style_analyzer.py::test_small_cluster_filtered -v
```
Expected: `ImportError: cannot import name 'extract_fingerprints'`

- [ ] **Step 3: 实现 `extract_fingerprints` 函数**

在 `style_analyzer.py` 末尾追加：

```python
CHAPTER_PATTERN = re.compile(
    r'^(第[一二三四五六七八九十\d]+章|[一二三四五六七八九十]+、|（[一二三四五六七八九十\d]+）|\d+\.\s)'
)


def _fingerprint_key(fp):
    """Convert fingerprint dict to a hashable key for grouping."""
    return (fp.get("size"), fp.get("bold"), fp.get("italic"), fp.get("align"))


def extract_fingerprints(filepath, min_cluster_size=4):
    """
    分析 DOCX 文档，提取正文格式指纹聚类。
    过滤掉成员数 < min_cluster_size 的小聚类（封面等特殊页面格式）。
    每个聚类保留一条代表性示例文字。
    返回 list of {id, fingerprint, example}。
    """
    import re
    doc = docx.Document(filepath)
    
    clusters = {}  # key -> list of (paragraph, fingerprint)
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        fp = compute_effective_fingerprint(para)
        key = _fingerprint_key(fp)
        if key not in clusters:
            clusters[key] = []
        clusters[key].append((text, fp))

    # 策略 A: 过滤小聚类；策略 C: 章节模式命中的聚类豁免过滤
    result = []
    for i, (key, members) in enumerate(clusters.items()):
        has_chapter_pattern = any(CHAPTER_PATTERN.match(text) for text, _ in members)
        if len(members) < min_cluster_size and not has_chapter_pattern:
            continue
        representative_text, representative_fp = members[0]
        result.append({
            "id": i,
            "fingerprint": representative_fp,
            "example": representative_text,
        })
    return result
```

还需在文件顶部添加 `import re`（与现有 import 合并）。

- [ ] **Step 4: 运行测试确认通过**

```bash
cd word-master
python -m pytest tests/test_style_analyzer.py -v
```
Expected: 所有 5 个 tests PASS

- [ ] **Step 5: 添加 CLI 入口**

在 `style_analyzer.py` 末尾追加：

```python
if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Extract style fingerprints from a DOCX template.")
    parser.add_argument("filepath", help="Source template DOCX path")
    parser.add_argument("--min-cluster-size", type=int, default=4)
    parser.add_argument("--output", default="fingerprints.json")
    args = parser.parse_args()
    fingerprints = extract_fingerprints(args.filepath, args.min_cluster_size)
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(fingerprints, f, ensure_ascii=False, indent=2)
    print(f"[INFO] Extracted {len(fingerprints)} fingerprint groups → {args.output}")
```

- [ ] **Step 6: 提交**

```bash
git add word-master/skills/word-master/scripts/style_analyzer.py word-master/tests/test_style_analyzer.py
git commit -m "feat: style_analyzer - extract_fingerprints with A+C body filtering"
```

---

### 任务 3：实现 `style_transfer.py` — 应用阶段（指纹匹配 + DSL 生成 + 委托执行）

**Files:**
- Create: `word-master/skills/word-master/scripts/style_transfer.py`
- Test: `word-master/tests/test_style_transfer.py`

注意：LLM 推理阶段（fingerprints → style_profile）依赖外部调用，测试中用 mock profile 代替。

- [ ] **Step 1: 编写指纹匹配和 DSL 生成测试**

创建 `word-master/tests/test_style_transfer.py`：

```python
import docx
from docx.shared import Pt
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'skills', 'word-master', 'scripts'))
from style_transfer import match_fingerprint_to_role, generate_apply_ops

MOCK_PROFILE = {
    "roles": [
        {"role": "Heading 1", "fingerprint": {"size": "16.0pt", "bold": True, "align": "center"}},
        {"role": "Normal",    "fingerprint": {"size": "12.0pt", "bold": False, "align": "justify"}},
    ]
}

def test_exact_match_heading(tmp_path):
    """Exact fingerprint match should return correct role."""
    fp = {"size": "16.0pt", "bold": True, "align": "center"}
    role = match_fingerprint_to_role(fp, MOCK_PROFILE)
    assert role == "Heading 1"

def test_closest_match_normal(tmp_path):
    """Closest fingerprint should return Normal for body-like format."""
    fp = {"size": "12.0pt", "bold": False, "align": "justify"}
    role = match_fingerprint_to_role(fp, MOCK_PROFILE)
    assert role == "Normal"

def test_low_similarity_returns_none(tmp_path):
    """Fingerprint with very different size should return None (skip)."""
    fp = {"size": "36.0pt", "bold": True, "align": "center"}  # 封面大标题
    role = match_fingerprint_to_role(fp, MOCK_PROFILE, threshold=0.6)
    assert role is None

def test_generate_ops(tmp_path):
    """generate_apply_ops should produce apply_style ops for matched paragraphs."""
    doc = docx.Document()
    p1 = doc.add_paragraph()
    p1.add_run("第一章 概述").bold = True
    p1.runs[0].font.size = Pt(16)
    p1.runs[0].bold = True
    p2 = doc.add_paragraph()
    p2.add_run("这是正文内容。")
    p2.runs[0].font.size = Pt(12)
    path = tmp_path / "draft.docx"
    doc.save(path)
    
    ops = generate_apply_ops(str(path), MOCK_PROFILE)
    # At least one op should be generated
    assert len(ops) > 0
    assert all(op["op"] == "apply_style" for op in ops)
```

- [ ] **Step 2: 运行测试确认失败**

```bash
cd word-master
python -m pytest tests/test_style_transfer.py -v
```
Expected: `ImportError`

- [ ] **Step 3: 实现 `style_transfer.py` 核心逻辑**

创建 `word-master/skills/word-master/scripts/style_transfer.py`：

```python
import docx
import json
import sys
import os

# 复用现有的指纹计算函数
sys.path.insert(0, os.path.dirname(__file__))
from style_analyzer import compute_effective_fingerprint


def _score(fp, role_fp):
    """
    计算两个指纹的相似度分数（0.0 ~ 1.0）。
    比较 size / bold / align 三个维度，各占 1/3 权重。
    """
    score = 0.0
    total = 3.0

    # size 相似度：差值 <= 2pt 得满分，否则线性衰减
    s1 = fp.get("size")
    s2 = role_fp.get("size")
    if s1 and s2:
        try:
            diff = abs(float(s1.rstrip("pt")) - float(s2.rstrip("pt")))
            score += max(0.0, 1.0 - diff / 10.0) * (1.0 / total)
        except ValueError:
            pass
    elif s1 == s2:  # both None
        score += 1.0 / total

    # bold 精确匹配
    if fp.get("bold") == role_fp.get("bold"):
        score += 1.0 / total

    # align 精确匹配
    if fp.get("align") == role_fp.get("align"):
        score += 1.0 / total

    return score


def match_fingerprint_to_role(fp, profile, threshold=0.6):
    """
    将段落指纹与 style_profile 中的角色做最近邻匹配。
    返回最佳匹配的 Word 样式名，若最高分 < threshold 则返回 None（跳过该段落）。
    """
    best_role = None
    best_score = 0.0
    for entry in profile.get("roles", []):
        s = _score(fp, entry["fingerprint"])
        if s > best_score:
            best_score = s
            best_role = entry["role"]
    return best_role if best_score >= threshold else None


def generate_apply_ops(draft_path, profile, threshold=0.6, skip_head=0, skip_tail=0):
    """
    遍历 draft.docx 的段落，对正文段落做指纹匹配并生成 apply_style DSL 列表。
    skip_head/skip_tail: 跳过头尾 N 个段落（人工兜底）。
    """
    doc = docx.Document(draft_path)
    paras = doc.paragraphs
    n = len(paras)
    ops = []

    for i, para in enumerate(paras):
        if i < skip_head or i >= n - skip_tail:
            continue
        if not para.text.strip():
            continue
        fp = compute_effective_fingerprint(para)
        role = match_fingerprint_to_role(fp, profile, threshold)
        if role:
            ops.append({"op": "apply_style", "target": f"p{i}", "style": role})
    return ops
```

- [ ] **Step 4: 运行测试确认通过**

```bash
cd word-master
python -m pytest tests/test_style_transfer.py -v
```
Expected: 4 tests PASS

- [ ] **Step 5: 提交**

```bash
git add word-master/skills/word-master/scripts/style_transfer.py word-master/tests/test_style_transfer.py
git commit -m "feat: style_transfer - fingerprint matching and DSL generation"
```

---

### 任务 4：实现 `style_transfer.py` — CLI 入口与 `docx_editor` 集成

**Files:**
- Modify: `word-master/skills/word-master/scripts/style_transfer.py`

- [ ] **Step 1: 在 `style_transfer.py` 末尾添加 CLI 和 `docx_editor` 集成代码**

```python
def run(template_path, draft_path, output_path,
        profile_path=None, review=False, skip_head=0, skip_tail=0):
    """
    主入口：两阶段样式转移流程。
    profile_path: 若提供则跳过分析和 LLM 推理，直接复用。
    review: 若 True，在生成 profile 后暂停等待用户确认。
    """
    # --- 阶段 1：获取 style_profile ---
    if profile_path:
        print(f"[INFO] Loading existing profile: {profile_path}")
        with open(profile_path, "r", encoding="utf-8") as f:
            profile = json.load(f)
    else:
        from style_analyzer import extract_fingerprints
        print(f"[INFO] Analyzing template: {template_path}")
        fingerprints = extract_fingerprints(template_path)

        # ── LLM 推理占位 ──
        # 当前版本将 fingerprints 直接打印到控制台，由外部 LLM（AI 会话）生成
        # style_profile.json 后再以 --profile 传入执行第二阶段。
        # TODO: 未来可集成 API 调用自动完成此步骤。
        fingerprints_path = "fingerprints.json"
        with open(fingerprints_path, "w", encoding="utf-8") as f:
            json.dump(fingerprints, f, ensure_ascii=False, indent=2)
        print(f"[INFO] Fingerprints saved to {fingerprints_path}")
        print("[INFO] Please review fingerprints.json, ask LLM to generate style_profile.json,")
        print("       then re-run with: --profile style_profile.json")
        return

    # --- 审核模式：暂停等待用户确认 ---
    if review:
        print(f"[REVIEW] Profile loaded. Edit style_profile.json if needed, then press Enter to continue...")
        input()

    # --- 阶段 2：生成 DSL 并委托 docx_editor 执行 ---
    print(f"[INFO] Generating apply_style operations for: {draft_path}")
    ops = generate_apply_ops(draft_path, profile,
                              skip_head=skip_head, skip_tail=skip_tail)
    print(f"[INFO] Generated {len(ops)} operations")

    # 委托 docx_editor 执行（复用其备份和安全执行机制）
    editor_path = os.path.join(os.path.dirname(__file__), "docx_editor.py")
    sys.path.insert(0, os.path.dirname(__file__))
    from docx_editor import apply_operations
    apply_operations(draft_path, ops, output_path)
    print(f"[INFO] Done → {output_path}")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Word document style transfer tool.")
    parser.add_argument("template", nargs="?", help="Source template DOCX (skip if using --profile)")
    parser.add_argument("draft", help="Target draft DOCX to apply styles to")
    parser.add_argument("output", help="Output DOCX path")
    parser.add_argument("--profile", help="Reuse existing style_profile.json (skip analysis)")
    parser.add_argument("--review", action="store_true", help="Pause after profile load for manual review")
    parser.add_argument("--skip-head", type=int, default=0, metavar="N", help="Skip first N paragraphs")
    parser.add_argument("--skip-tail", type=int, default=0, metavar="N", help="Skip last N paragraphs")
    args = parser.parse_args()
    run(
        template_path=args.template,
        draft_path=args.draft,
        output_path=args.output,
        profile_path=args.profile,
        review=args.review,
        skip_head=args.skip_head,
        skip_tail=args.skip_tail,
    )
```

还需在文件顶部追加 `import json` 导入（若未添加）。

- [ ] **Step 2: 运行已有测试，确认不回退**

```bash
cd word-master
python -m pytest tests/test_style_transfer.py tests/test_style_analyzer.py -v
```
Expected: 所有已有测试 PASS

- [ ] **Step 3: 手动冒烟测试（可选但推荐）**

若手边有任意 .docx 文件，用以下命令验证 CLI 正常运行：
```bash
cd word-master/skills/word-master/scripts
python style_transfer.py path/to/template.docx path/to/draft.docx output.docx
```
Expected: 打印 `[INFO] Fingerprints saved to fingerprints.json` 并退出（LLM 阶段占位）

- [ ] **Step 4: 提交**

```bash
git add word-master/skills/word-master/scripts/style_transfer.py
git commit -m "feat: style_transfer CLI with docx_editor delegation"
```

---

### 任务 5：更新 `SKILL.md` 文档

**Files:**
- Modify: `word-master/skills/word-master/SKILL.md`

- [ ] **Step 1: 阅读现有 SKILL.md**

- [ ] **Step 2: 在 Core Scripts 表格中追加两行**

```markdown
| `${SKILL_DIR}/scripts/style_analyzer.py` | Extract format fingerprints from a template DOCX |
| `${SKILL_DIR}/scripts/style_transfer.py` | Apply template style profile to a target DOCX |
```

- [ ] **Step 3: 在 Workflow 2 后追加 Workflow 3**

```markdown
---

## Workflow 3: Style Transfer (样式迁移)

> 学习源模板文档的正文格式风格，应用到目标草稿文档。

### Step 1: 提取源模板格式指纹

```bash
python ${SKILL_DIR}/scripts/style_analyzer.py template.docx --output fingerprints.json
```

### Step 2: AI 推断样式角色

将 `fingerprints.json` 内容提供给 AI，请求生成 `style_profile.json`。
AI 应根据字号、加粗、对齐等特征推断 Word 内置样式（Heading 1, Normal 等）。

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
```

- [ ] **Step 4: 提交**

```bash
git add word-master/skills/word-master/SKILL.md
git commit -m "docs: SKILL.md - add style transfer workflow 3"
```

---

### 任务 6：运行全量测试验证

- [ ] **Step 1: 运行全量测试**

```bash
cd word-master
python -m pytest tests/ -v
```
Expected: 所有测试 PASS，无新失败

- [ ] **Step 2: 提交（若有修复）**

```bash
git commit -am "fix: resolve any test failures from style transfer integration"
```
