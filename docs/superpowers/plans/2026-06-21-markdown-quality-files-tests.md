# markdown-master quality.py 与 files.py 单元测试实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 为 `markdown-master/scripts/quality.py`（lint/zhlint/linkcheck）和 `markdown-master/scripts/files.py`（split/merge）补齐自动化单元测试，共 28 个用例（16+12），不修改任何源码。

**Architecture:** 纯函数级 pytest 测试，直接 import `run_lint` / `run_zhlint` / `run_linkcheck` / `cmd_split` / `cmd_merge`。文档内容用 Python 字符串内联；文件 I/O 用 `tmp_path` fixture。已知缺陷 `test_zhlint_inline_code_protected` 用 `@pytest.mark.xfail` 标记。

**Tech Stack:** Python 3、pytest（项目已用）、`markdown-master/scripts/_md_utils.py`、`markdown-master/conftest.py`（已配置 sys.path）。

**Spec:** `docs/superpowers/specs/2026-06-21-markdown-quality-files-tests-design.md`

---

## 文件结构

- **Create:** `markdown-master/tests/test_quality.py` — 16 个用例，覆盖 run_lint（7）+ run_zhlint（4）+ run_linkcheck（5）
- **Create:** `markdown-master/tests/test_files.py` — 12 个用例，覆盖 cmd_split（7）+ cmd_merge（5）
- **不修改：** `scripts/` 下任何源码、`conftest.py`、现有 6 个测试文件

被测函数签名（来自 `markdown-master/scripts/quality.py` 和 `files.py`，已确认）：
- `run_lint(content: str, fix: bool) -> tuple[str, list[tuple[int, str]]]` — 返回 `(修复后内容, [(行号, 描述), ...])`
- `run_zhlint(content: str, fix: bool) -> tuple[str, list[tuple[int, str]]]`
- `run_linkcheck(content: str, file_path: str, check_remote: bool) -> list[tuple[int, str]]`
- `cmd_split(input_path: str, by: str, output_dir: str | None) -> None` — 直接写文件到磁盘，print 到 stdout
- `cmd_merge(inputs: list[str], output: str, separator: str) -> None` — 写文件到 output，无 .md 时 `sys.exit(1)`

约定：
- `run_lint` / `run_zhlint` 的 issues 中，行号 `0` 表示文档级问题（如缺少 H1），否则从 1 开始
- `run_lint` 返回的 `issues` 列表在 `--fix` 路径下仍包含被修复的条目（"发现并修复"语义）
- `cmd_split` 的文件名格式：preamble 为 `00_preamble.md`；其余为 `XX_<sanitized_title>.md`，XX 是从 0 开始的两位补零序号；空 preamble 不输出
- `cmd_split` 的 `_sanitize_filename` 会先剥除编号前缀（如 `第1章 `、`1.1 `、`（1）`），再替换 `\/:*?"<>|` 为 `_`，截断到 80 字符

---

### Task 1: 创建 test_quality.py 并实现 run_lint 检测用例（spec §1-6）

**Files:**
- Create: `markdown-master/tests/test_quality.py`

- [ ] **Step 1: 创建 test_quality.py，写入文件头注释和 6 个 run_lint 检测用例**

```python
"""quality.py 的单元测试：覆盖 run_lint / run_zhlint / run_linkcheck（本地路径分支）。

远程 HTTP linkcheck 分支不写测试——核心断言逻辑已由本地路径分支覆盖，
HTTP 行为不是本工具核心，且 mock urllib 会让 CI 不稳定。详见
docs/superpowers/specs/2026-06-21-markdown-quality-files-tests-design.md。
"""
import pytest
from quality import run_lint, run_zhlint, run_linkcheck


def test_lint_clean_doc():
    """正常文档无任何格式问题，issues 应为空。"""
    content = "# 标题\n\n正文段落\n\n## 子标题\n\n更多正文\n"
    _result, issues = run_lint(content, fix=False)
    assert issues == []


def test_lint_missing_h1():
    """文档没有 H1 标题时，应报文档级问题，行号为 0。"""
    content = "## 子标题\n\n正文\n"
    _result, issues = run_lint(content, fix=False)
    assert (0, "缺少 H1 标题") in issues


def test_lint_heading_skip():
    """从 H1 直接跳到 H3 应报标题层级跳跃，行号指向 H3 所在行。"""
    content = "# A\n\n### C\n"
    _result, issues = run_lint(content, fix=False)
    skip_issues = [i for i in issues if "标题层级跳跃" in i[1]]
    assert len(skip_issues) == 1
    assert skip_issues[0][0] == 3  # ### C 在第 3 行


def test_lint_heading_skip_ignored_in_code_block():
    """代码块内的 ### 不应触发标题层级跳跃告警。"""
    content = "# A\n\n```\n### C\n```\n"
    _result, issues = run_lint(content, fix=False)
    skip_issues = [i for i in issues if "标题层级跳跃" in i[1]]
    assert skip_issues == []


def test_lint_excessive_blank_lines():
    """连续 4 行空行应报"连续空行过多"，行号指向起始空行。"""
    content = "# A\n\n\n\n\n正文\n"
    _result, issues = run_lint(content, fix=False)
    blank_issues = [i for i in issues if "连续空行过多" in i[1]]
    assert len(blank_issues) == 1
    # 第 2-5 行是空行，起始行号为 2
    assert blank_issues[0][0] == 2


def test_lint_code_block_no_language():
    """单独 ``` 报"代码块未指定语言"；```python 不报。"""
    content = "# A\n\n```\ncode\n```\n\n```python\ncode\n```\n"
    _result, issues = run_lint(content, fix=False)
    no_lang_issues = [i for i in issues if "代码块未指定语言" in i[1]]
    # 只有一处无语言代码块（第 3 行的 ```），第 9 行的 ```python 不报
    assert len(no_lang_issues) == 1
    assert no_lang_issues[0][0] == 3
```

- [ ] **Step 2: 运行这 6 个用例，确认全部通过**

Run: `cd markdown-master && python -m pytest tests/test_quality.py -v`
Expected: 6 passed

- [ ] **Step 3: 提交**

```bash
git add markdown-master/tests/test_quality.py
git commit -m "test(quality): add run_lint detection tests"
```

---

### Task 2: 添加 run_lint --fix 用例（spec §7）

**Files:**
- Modify: `markdown-master/tests/test_quality.py`

- [ ] **Step 1: 在 test_quality.py 末尾追加 fix 用例**

```python
def test_lint_fix_collapses_blank_lines():
    """--fix 时 4 行空行被压缩为 2 行；issues 仍包含该条（"发现并修复"而非"静默修复"）。"""
    content = "# A\n\n\n\n\n正文\n"
    result, issues = run_lint(content, fix=True)
    # 4 行空行压缩为 2 行
    assert "# A\n\n\n正文\n" == result
    # issues 仍报告这条（行号是压缩前的起始行号 2）
    blank_issues = [i for i in issues if "连续空行过多" in i[1]]
    assert len(blank_issues) == 1
    assert blank_issues[0][0] == 2
```

- [ ] **Step 2: 运行，确认通过**

Run: `cd markdown-master && python -m pytest tests/test_quality.py::test_lint_fix_collapses_blank_lines -v`
Expected: 1 passed

- [ ] **Step 3: 提交**

```bash
git add markdown-master/tests/test_quality.py
git commit -m "test(quality): add run_lint --fix test"
```

---

### Task 3: 添加 run_zhlint 用例（spec §8-11，含 1 个 xfail）

**Files:**
- Modify: `markdown-master/tests/test_quality.py`

- [ ] **Step 1: 在 test_quality.py 末尾追加 4 个 zhlint 用例**

```python
def test_zhlint_missing_space_zh_en():
    """中文紧接英文/数字应报缺空格；fix 后插入空格。"""
    content = "中文abc\n"
    result, issues = run_zhlint(content, fix=True)
    assert "中文 abc" in result
    assert any("缺少空格" in i[1] for i in issues)


def test_zhlint_missing_space_en_zh():
    """英文紧接中文同样应报缺空格；fix 后插入空格。"""
    content = "abc中文\n"
    result, issues = run_zhlint(content, fix=True)
    assert "abc 中文" in result
    assert any("缺少空格" in i[1] for i in issues)


def test_zhlint_half_punct():
    """中文语境中使用半角逗号应报；fix 后转为全角。"""
    content = "中文,英文\n"
    result, issues = run_zhlint(content, fix=True)
    assert "中文，英文" in result
    assert any("半角标点" in i[1] for i in issues)


@pytest.mark.xfail(reason="run_zhlint 未调用 _is_inline_code，行内代码段会被错误插入空格")
def test_zhlint_inline_code_protected():
    """行内代码段内的内容不应被插入空格。

    已知缺陷：run_zhlint 主流程未调用 _is_inline_code，
    行内代码 `code` 与相邻中文会被错误插入空格。详见 spec。
    """
    content = "中文`code`abc\n"
    result, _issues = run_zhlint(content, fix=True)
    # 期望：行内代码段不受影响，原样保留
    assert "中文`code`abc" in result
```

- [ ] **Step 2: 运行，确认 3 个 pass + 1 个 xfail**

Run: `cd markdown-master && python -m pytest tests/test_quality.py -k zhlint -v`
Expected: 3 passed, 1 xfailed（`test_zhlint_inline_code_protected`）

- [ ] **Step 3: 提交**

```bash
git add markdown-master/tests/test_quality.py
git commit -m "test(quality): add run_zhlint tests with xfail for inline code bug"
```

---

### Task 4: 添加 run_linkcheck 本地路径用例（spec §12-16）

**Files:**
- Modify: `markdown-master/tests/test_quality.py`

- [ ] **Step 1: 在 test_quality.py 末尾追加 5 个 linkcheck 用例**

```python
def test_linkcheck_local_image_missing(tmp_path):
    """引用不存在的本地图片应报"图片路径不存在"，行号正确。"""
    doc = tmp_path / "doc.md"
    doc.write_text("# A\n\n![alt](missing.png)\n", encoding="utf-8")
    content = doc.read_text(encoding="utf-8")
    issues = run_linkcheck(content, str(doc), check_remote=False)
    img_issues = [i for i in issues if "图片路径不存在" in i[1]]
    assert len(img_issues) == 1
    assert img_issues[0][0] == 3  # 第 3 行


def test_linkcheck_local_image_exists(tmp_path):
    """引用本地真实存在的图片不报错。"""
    (tmp_path / "img.png").write_bytes(b"\x89PNG\r\n\x1a\n")
    doc = tmp_path / "doc.md"
    doc.write_text("# A\n\n![alt](img.png)\n", encoding="utf-8")
    content = doc.read_text(encoding="utf-8")
    issues = run_linkcheck(content, str(doc), check_remote=False)
    assert issues == []


def test_linkcheck_local_link_missing(tmp_path):
    """本地链接目标不存在应报"本地链接目标不存在"。"""
    doc = tmp_path / "doc.md"
    doc.write_text("# A\n\n[text](missing.md)\n", encoding="utf-8")
    content = doc.read_text(encoding="utf-8")
    issues = run_linkcheck(content, str(doc), check_remote=False)
    link_issues = [i for i in issues if "本地链接目标不存在" in i[1]]
    assert len(link_issues) == 1
    assert link_issues[0][0] == 3


def test_linkcheck_anchor_skipped(tmp_path):
    """锚点链接（#anchor）不报错。"""
    doc = tmp_path / "doc.md"
    doc.write_text("# A\n\n[text](#section)\n", encoding="utf-8")
    content = doc.read_text(encoding="utf-8")
    issues = run_linkcheck(content, str(doc), check_remote=False)
    assert issues == []


def test_linkcheck_link_in_code_block_skipped(tmp_path):
    """代码块内的图片/链接不应被检查。"""
    doc = tmp_path / "doc.md"
    doc.write_text("# A\n\n```\n![alt](missing.png)\n```\n", encoding="utf-8")
    content = doc.read_text(encoding="utf-8")
    issues = run_linkcheck(content, str(doc), check_remote=False)
    assert issues == []
```

- [ ] **Step 2: 运行，确认全部通过**

Run: `cd markdown-master && python -m pytest tests/test_quality.py -k linkcheck -v`
Expected: 5 passed

- [ ] **Step 3: 提交**

```bash
git add markdown-master/tests/test_quality.py
git commit -m "test(quality): add run_linkcheck local-path tests"
```

---

### Task 5: 创建 test_files.py 并实现 cmd_split 用例（spec §1-7）

**Files:**
- Create: `markdown-master/tests/test_files.py`

- [ ] **Step 1: 创建 test_files.py，写入文件头注释和 7 个 split 用例**

```python
"""files.py 的单元测试：覆盖 cmd_split / cmd_merge。

cmd_split / cmd_merge 直接 print 到 stdout 且 cmd_merge 会 sys.exit，
不返回值。测试策略：调用后读产出文件做断言；merge 空输入用
pytest.raises(SystemExit) 捕获退出码。详见
docs/superpowers/specs/2026-06-21-markdown-quality-files-tests-design.md。
"""
import pytest
from files import cmd_split, cmd_merge


def test_split_by_h2_basic(tmp_path):
    """按 H2 拆分：preamble + 两个 H2 段，输出 3 个文件，序号和文件名正确。"""
    doc = tmp_path / "doc.md"
    doc.write_text(
        "# A\n\nintro\n\n## B\n\nb-body\n\n## C\n\nc-body\n",
        encoding="utf-8",
    )
    out_dir = tmp_path / "out"
    cmd_split(str(doc), "h2", str(out_dir))

    files = sorted(p.name for p in out_dir.iterdir())
    assert files == ["00_preamble.md", "01_B.md", "02_C.md"]
    assert "intro" in (out_dir / "00_preamble.md").read_text(encoding="utf-8")
    assert "## B" in (out_dir / "01_B.md").read_text(encoding="utf-8")
    assert "b-body" in (out_dir / "01_B.md").read_text(encoding="utf-8")
    assert "## C" in (out_dir / "02_C.md").read_text(encoding="utf-8")


def test_split_by_h2_no_preamble(tmp_path):
    """文档以 ## 开头时无 preamble，第一段文件名为 00_<标题>.md。"""
    doc = tmp_path / "doc.md"
    doc.write_text("## B\n\nb-body\n\n## C\n\nc-body\n", encoding="utf-8")
    out_dir = tmp_path / "out"
    cmd_split(str(doc), "h2", str(out_dir))

    files = sorted(p.name for p in out_dir.iterdir())
    # 没有 preamble，直接从 00 开始编号
    assert "00_B.md" in files
    assert "01_C.md" in files
    assert "00_preamble.md" not in files


def test_split_preamble_only_blank_skipped(tmp_path):
    """只有正文（空行）无标题时，preamble 全空，不输出任何文件。"""
    doc = tmp_path / "doc.md"
    doc.write_text("\n\n\n", encoding="utf-8")
    out_dir = tmp_path / "out"
    cmd_split(str(doc), "h2", str(out_dir))

    files = list(out_dir.iterdir()) if out_dir.exists() else []
    assert files == []


def test_split_heading_in_code_block_ignored(tmp_path):
    """代码块内的 ## xxx 不触发拆分。"""
    doc = tmp_path / "doc.md"
    doc.write_text(
        "# A\n\n```\n## fake\n```\n\n## real\n\nbody\n",
        encoding="utf-8",
    )
    out_dir = tmp_path / "out"
    cmd_split(str(doc), "h2", str(out_dir))

    names = sorted(p.name for p in out_dir.iterdir())
    # 代码块内的 ## fake 不拆，只有 ## real 一个段 + preamble
    assert "01_real.md" in names
    assert "00_preamble.md" in names
    # 不应该有 fake 命名的文件
    assert not any("fake" in n for n in names)


def test_split_filename_sanitization(tmp_path):
    """标题含非法字符被替换为 _，长度截断到 80 字符（不含序号前缀和扩展名）。"""
    doc = tmp_path / "doc.md"
    # 标题含所有非法字符
    doc.write_text("# A\n\n## a\\b/c:d*e?f\"g<h>i|j\n\nbody\n", encoding="utf-8")
    out_dir = tmp_path / "out"
    cmd_split(str(doc), "h2", str(out_dir))

    names = [p.name for p in out_dir.iterdir() if not p.name.startswith("00_")]
    assert len(names) == 1
    name = names[0]
    # 非法字符全部被替换为 _
    sanitized_body = name[3:-3]  # 去掉 "01_" 前缀和 ".md" 后缀
    for ch in '\\/:*?"<>|':
        assert ch not in sanitized_body
    # 长度截断到 80
    assert len(sanitized_body) <= 80


def test_split_custom_output_dir(tmp_path):
    """--output-dir 指向的子目录会被自动创建。"""
    doc = tmp_path / "doc.md"
    doc.write_text("# A\n\n## B\n\nbody\n", encoding="utf-8")
    custom = tmp_path / "deep" / "nested" / "out"
    assert not custom.exists()

    cmd_split(str(doc), "h2", str(custom))

    assert custom.exists()
    assert any(p.name == "01_B.md" for p in custom.iterdir())


def test_split_by_h1(tmp_path):
    """--by h1 时按 H1 拆，H2 不拆。"""
    doc = tmp_path / "doc.md"
    doc.write_text(
        "# A\n\n## sub\n\nbody-a\n\n# B\n\nbody-b\n",
        encoding="utf-8",
    )
    out_dir = tmp_path / "out"
    cmd_split(str(doc), "h1", str(out_dir))

    names = sorted(p.name for p in out_dir.iterdir())
    # H1 拆出 A、B 两段，无 preamble（文档以 # 开头）
    assert "00_A.md" in names
    assert "01_B.md" in names
    # H2 sub 不应单独成文件
    assert not any("sub" in n for n in names)
    # sub 标题应留在 A 段内部
    assert "## sub" in (out_dir / "00_A.md").read_text(encoding="utf-8")
```

- [ ] **Step 2: 运行，确认全部通过**

Run: `cd markdown-master && python -m pytest tests/test_files.py -k split -v`
Expected: 7 passed

- [ ] **Step 3: 提交**

```bash
git add markdown-master/tests/test_files.py
git commit -m "test(files): add cmd_split tests"
```

---

### Task 6: 添加 cmd_merge 用例（spec §8-12）

**Files:**
- Modify: `markdown-master/tests/test_files.py`

- [ ] **Step 1: 在 test_files.py 末尾追加 5 个 merge 用例**

```python
def test_merge_two_files(tmp_path):
    """合并两个文件，默认分隔符 \n\n---\n\n 连接。"""
    a = tmp_path / "a.md"
    a.write_text("# A\n\nbody-a\n", encoding="utf-8")
    b = tmp_path / "b.md"
    b.write_text("# B\n\nbody-b\n", encoding="utf-8")
    out = tmp_path / "out.md"

    cmd_merge([str(a), str(b)], str(out), "\n\n---\n\n")

    content = out.read_text(encoding="utf-8")
    assert "body-a" in content
    assert "body-b" in content
    assert "\n\n---\n\n" in content
    # a 在前 b 在后
    assert content.index("body-a") < content.index("body-b")


def test_merge_custom_separator(tmp_path):
    """--separator 生效，自定义分隔符出现在合并结果中。"""
    a = tmp_path / "a.md"
    a.write_text("AAA", encoding="utf-8")
    b = tmp_path / "b.md"
    b.write_text("BBB", encoding="utf-8")
    out = tmp_path / "out.md"

    cmd_merge([str(a), str(b)], str(out), "\n===\n")

    content = out.read_text(encoding="utf-8")
    assert content == "AAA\n===\nBBB"


def test_merge_directory_input(tmp_path):
    """目录输入时按文件名字母序合并所有 .md。"""
    src_dir = tmp_path / "src"
    src_dir.mkdir()
    (src_dir / "b.md").write_text("B-content", encoding="utf-8")
    (src_dir / "a.md").write_text("A-content", encoding="utf-8")
    (src_dir / "c.md").write_text("C-content", encoding="utf-8")
    out = tmp_path / "out.md"

    cmd_merge([str(src_dir)], str(out), "\n---\n")

    content = out.read_text(encoding="utf-8")
    # 按 a, b, c 字母序
    assert content.index("A-content") < content.index("B-content")
    assert content.index("B-content") < content.index("C-content")


def test_merge_skips_non_md(tmp_path):
    """目录里的 .txt / .png 被跳过，只合并 .md。"""
    src_dir = tmp_path / "src"
    src_dir.mkdir()
    (src_dir / "a.md").write_text("MD-content", encoding="utf-8")
    (src_dir / "notes.txt").write_text("TXT-content", encoding="utf-8")
    (src_dir / "pic.png").write_bytes(b"\x89PNG")
    out = tmp_path / "out.md"

    cmd_merge([str(src_dir)], str(out), "\n---\n")

    content = out.read_text(encoding="utf-8")
    assert "MD-content" in content
    assert "TXT-content" not in content


def test_merge_empty_inputs_exits_nonzero(tmp_path):
    """没有任何 .md 文件时 sys.exit(1)。"""
    src_dir = tmp_path / "src"
    src_dir.mkdir()
    (src_dir / "notes.txt").write_text("not md", encoding="utf-8")
    out = tmp_path / "out.md"

    with pytest.raises(SystemExit) as exc_info:
        cmd_merge([str(src_dir)], str(out), "\n---\n")
    assert exc_info.value.code == 1
```

- [ ] **Step 2: 运行，确认全部通过**

Run: `cd markdown-master && python -m pytest tests/test_files.py -k merge -v`
Expected: 5 passed

- [ ] **Step 3: 提交**

```bash
git add markdown-master/tests/test_files.py
git commit -m "test(files): add cmd_merge tests"
```

---

### Task 7: 全量回归验证

**Files:** 无（仅运行）

- [ ] **Step 1: 运行新增的两个测试文件，确认 28 个用例全绿（含 1 个 xfail）**

Run: `cd markdown-master && python -m pytest tests/test_quality.py tests/test_files.py -v`
Expected: 27 passed, 1 xfailed

- [ ] **Step 2: 运行整个 tests/ 目录，确认不破坏现有 6 个测试文件**

Run: `cd markdown-master && python -m pytest tests/ -v`
Expected: 全部 passed 或 xfailed，无 failed

- [ ] **Step 3: 确认源码未被修改**

Run: `git status markdown-master/scripts/`
Expected: `nothing to commit, working tree clean`（scripts/ 目录无任何改动）

- [ ] **Step 4: 最终提交（如有补漏）**

```bash
git status
# 若有未提交内容，补提交；若无，跳过本步
```

---

## Self-Review

**1. Spec coverage：**
- §test_quality.py §1-6（run_lint 检测）→ Task 1 ✓
- §test_quality.py §7（run_lint --fix）→ Task 2 ✓
- §test_quality.py §8-11（run_zhlint 含 xfail）→ Task 3 ✓
- §test_quality.py §12-16（run_linkcheck 本地）→ Task 4 ✓
- §test_files.py §1-7（cmd_split）→ Task 5 ✓
- §test_files.py §8-12（cmd_merge）→ Task 6 ✓
- 验收标准 §1（全绿含 xfail）→ Task 7 Step 1 ✓
- 验收标准 §2（不破坏现有测试）→ Task 7 Step 2 ✓
- 验收标准 §3（不改源码）→ Task 7 Step 3 ✓
- 验收标准 §4（覆盖全部编号）→ 16+12=28 个用例，与 spec 一致 ✓
- 验收标准 §5（中文 docstring）→ 每个测试函数都有 ✓
- 非目标（不修 _is_inline_code / 不写 HTTP / 不写 CLI / 不补其他测试）→ 计划未涉及 ✓

**2. Placeholder scan：** 无 TODO / TBD / "类似 Task N"，每个测试函数代码完整。

**3. Type consistency：**
- `run_lint(content, fix) -> (result_str, issues_list)` — Task 1/2 解构为 `_result, issues`，一致
- `run_zhlint(content, fix) -> (result_str, issues_list)` — Task 3 一致
- `run_linkcheck(content, file_path, check_remote) -> issues_list` — Task 4 一致
- `cmd_split(input_path, by, output_dir)` — Task 5 调用 `cmd_split(str(doc), "h2", str(out_dir))`，一致
- `cmd_merge(inputs, output, separator)` — Task 6 调用 `cmd_merge([str(a), str(b)], str(out), sep)`，一致
- 行号语义：spec 用例 5 要求"行号指向起始空行"，Task 1 `test_lint_excessive_blank_lines` 断言 `== 2`（content 第 2 行起是空行）✓

**潜在风险（执行时注意）：**
- Task 1 `test_lint_heading_skip`：`# A\n\n### C\n` 中 `### C` 在第 3 行（第 1 行 `# A`，第 2 行空行，第 3 行 `### C`）。断言 `== 3` 正确。
- Task 1 `test_lint_excessive_blank_lines`：`# A\n\n\n\n\n正文\n` 各行为 `['# A', '', '', '', '', '正文', '']`，4 行空行是第 2-5 行，起始行号 2。断言 `== 2` 正确。
- Task 3 `test_zhlint_inline_code_protected`：xfail 必须用 `@pytest.mark.xfail` 装饰器，不能写成 `pytest.xfail()` 调用，否则会被算作失败。计划中已用装饰器形式。
- Task 5 `test_split_by_h2_no_preamble`：文档 `## B\n\nb-body\n\n## C\n\nc-body\n` 第一行就是 `## B`，sections 第一个就是 `(B, [## B 行])`，文件名 `00_B.md`。断言正确。
- Task 5 `test_split_filename_sanitization`：标题 `a\b/c:d*e?f"g<h>i|j` 共 19 字符，远小于 80，截断不会触发，但断言 `<= 80` 仍成立。
- Task 6 `test_merge_empty_inputs_exits_nonzero`：目录里只有 `.txt`，`md_files` 为空，触发 `sys.exit(1)`。`exc_info.value.code == 1` 正确。

无问题。
