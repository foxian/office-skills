# markdown-master 技能实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在 `markdown-master/` 目录下实现完整的 Markdown 文档管理 Agent 技能，包含结构操作、质检操作、格式转换、文件操作四个脚本和一份 SKILL.md。

**Architecture:** 共用工具库 `_md_utils.py` + 4 个独立功能脚本（structure / quality / convert / files）+ SKILL.md 调用手册。每个脚本使用 argparse 子命令风格，通过 `uv run python` 调用。

**Tech Stack:** Python 3.x（标准库）、python-docx、markdown、beautifulsoup4、Pillow、weasyprint

---

## 文件清单

| 文件 | 类型 | 说明 |
|------|------|------|
| `markdown-master/SKILL.md` | 新建 | AI 调用手册 |
| `markdown-master/scripts/_md_utils.py` | 新建 | 共用工具库 |
| `markdown-master/scripts/structure.py` | 新建 | 标题/编号/目录 |
| `markdown-master/scripts/quality.py` | 新建 | lint/zhlint/linkcheck |
| `markdown-master/scripts/convert.py` | 新建 | DOCX/HTML/TXT/PDF 转换 |
| `markdown-master/scripts/files.py` | 新建 | 拆分/合并 |
| `markdown-master/examples/sample_cn.md` | 新建 | 中文验证文档 |
| `markdown-master/examples/sample_en.md` | 新建 | 英文验证文档 |

---

## Task 1: 项目框架与共用工具库

**Files:**
- Create: `markdown-master/scripts/_md_utils.py`
- Create: `markdown-master/examples/sample_cn.md`
- Create: `markdown-master/examples/sample_en.md`

- [ ] **Step 1: 创建目录结构**

```bash
mkdir markdown-master\scripts
mkdir markdown-master\examples
```

- [ ] **Step 2: 创建 `_md_utils.py`**

```python
# markdown-master/scripts/_md_utils.py
"""
共用工具库，供 structure.py / quality.py / convert.py / files.py import。
不可直接调用。
"""
import re
from pathlib import Path


def read_utf8(path: str) -> str:
    """读取 UTF-8 文件内容，自动处理 BOM。"""
    return Path(path).read_text(encoding="utf-8-sig")


def write_utf8(path: str, content: str) -> None:
    """将内容写入 UTF-8 文件（无 BOM）。"""
    Path(path).write_text(content, encoding="utf-8")


def parse_file(path: str) -> tuple[str, dict]:
    """
    解析 Markdown 文件，分离 frontmatter 和正文。
    返回: (body_content, frontmatter_dict)
    若无 frontmatter，frontmatter_dict 为空 dict。
    """
    content = read_utf8(path)
    if not content.startswith("---"):
        return content, {}

    lines = content.split("\n")
    end = -1
    for i in range(1, len(lines)):
        if lines[i].strip() == "---":
            end = i
            break
    if end == -1:
        return content, {}

    fm_lines = lines[1:end]
    body = "\n".join(lines[end + 1:])
    fm = {}
    for line in fm_lines:
        if ":" in line:
            k, _, v = line.partition(":")
            fm[k.strip()] = v.strip()
    return body, fm


def extract_headings(content: str) -> list[tuple[int, str, int]]:
    """
    提取正文中所有标题。
    返回: list of (level, text, line_number)  —  line_number 从 1 开始
    跳过代码块内的 # 行。
    """
    lines = content.split("\n")
    headings = []
    for i, line in enumerate(lines):
        if is_in_code_block(lines, i):
            continue
        m = re.match(r"^(#{1,6})\s+(.*)", line)
        if m:
            level = len(m.group(1))
            text = m.group(2).strip()
            headings.append((level, text, i + 1))
    return headings


def is_in_code_block(lines: list[str], line_index: int) -> bool:
    """
    判断 lines[line_index] 是否处于代码块（``` 或 ~~~）内部。
    """
    fence_count = 0
    for i in range(line_index):
        if re.match(r"^(`{3,}|~{3,})", lines[i]):
            fence_count += 1
    return fence_count % 2 == 1
```

- [ ] **Step 3: 创建中文测试文档 `examples/sample_cn.md`**

```markdown
# 项目总体说明

本文档用于测试markdown-master技能的各项功能。

## 第一部分：背景介绍

### 1.1 项目概述

这是一个用Python编写的Markdown文档管理工具,支持中英文混排处理。

主要功能包括:

- 结构操作
- 质检操作
- 格式转换

### 1.2 技术选型

我们选择了以下技术栈：Python3.x,标准库为主。

## 第二部分：功能说明

### 2.1 结构操作

支持heading upgrade、downgrade,以及numbering管理。

#### 2.1.1 详细说明

详见下文。

## 第四部分：总结

（注意：第三部分被跳过，用于测试 lint 标题跳跃检查）

这里有中英文之间缺空格的问题，例如：Python编程语言和Markdown格式。

还有半角标点问题: 这句话用了半角冒号.
```

- [ ] **Step 4: 创建英文测试文档 `examples/sample_en.md`**

```markdown
## Getting Started

This document is used to test the markdown-master skill.

Notice the missing H1 heading above — used for lint testing.

### Installation

See [setup guide](./nonexistent_file.md) for details.

Also check ![logo](./images/logo.png) for the project logo.

#### Dependencies

##### Sub-dependencies

(Heading jump from H3 to H5 for lint testing)

Visit [our website](https://httpstat.us/404) for more info.

### Usage

Run the scripts as documented in SKILL.md.
```

- [ ] **Step 5: 验证 `_md_utils.py` 基础功能**

在 `markdown-master/` 目录下运行：

```bash
uv run python -c "
import sys; sys.path.insert(0, 'scripts')
from _md_utils import read_utf8, extract_headings, is_in_code_block
content = read_utf8('examples/sample_cn.md')
headings = extract_headings(content)
print('Headings found:', len(headings))
for h in headings:
    print(h)
"
```

期望输出：打印出 sample_cn.md 中所有标题（约 7 个），每行格式为 `(level, text, line_number)`。

- [ ] **Step 6: Commit**

```bash
git add markdown-master/
git commit -m "feat: scaffold markdown-master skill with _md_utils and example files"
```

---

## Task 2: `structure.py` — 结构操作

**Files:**
- Create: `markdown-master/scripts/structure.py`

- [ ] **Step 1: 实现 `structure.py`**

```python
# markdown-master/scripts/structure.py
"""
结构操作：标题层级升降、编号管理、目录生成。

用法:
  uv run python scripts/structure.py <input.md> heading upgrade [--levels N] [-o output.md]
  uv run python scripts/structure.py <input.md> heading downgrade [--levels N] [-o output.md]
  uv run python scripts/structure.py <input.md> numbering add --style <style> [--start-from N] [-o output.md]
  uv run python scripts/structure.py <input.md> numbering remove [-o output.md]
  uv run python scripts/structure.py <input.md> toc generate [--depth N] [--position top|after-h1] [-o output.md]
"""
import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from _md_utils import read_utf8, write_utf8, is_in_code_block

# ── 编号样式定义 ──────────────────────────────────────────────
CHINESE_NUMS = "一二三四五六七八九十十一十二十三十四十五十六十七十八十九二十".split()


def _chinese_num(n: int) -> str:
    if 1 <= n <= len(CHINESE_NUMS):
        return CHINESE_NUMS[n - 1]
    return str(n)


def _roman(n: int) -> str:
    vals = [(1000,'M'),(900,'CM'),(500,'D'),(400,'CD'),(100,'C'),(90,'XC'),
            (50,'L'),(40,'XL'),(10,'X'),(9,'IX'),(5,'V'),(4,'IV'),(1,'I')]
    r = ""
    for v, s in vals:
        while n >= v:
            r += s; n -= v
    return r


STYLES = {
    "chinese_bidding": {
        1: lambda n: f"{_chinese_num(n)}、",
        2: lambda n, p: f"{p}{n}",         # p = parent prefix like "1."
        3: lambda n, p: f"{p}{n}",
        # level 2/3 use numeric joined by dot, built in _build_number
    },
    "chinese_chapter": {
        1: lambda n: f"第{_chinese_num(n)}章",
        2: lambda n, _: f"{_chinese_num(n)}、",
        3: lambda n, _: f"（{_chinese_num(n)}）",
        4: lambda n, _: f"{n}.",
    },
    "technical": {
        # 1 → "1", 2 → "1.1", 3 → "1.1.1"
    },
    "academic": {
        1: lambda n, _: _roman(n),
        2: lambda n, _: chr(64 + n),   # A, B, C
        3: lambda n, _: str(n),
        4: lambda n, _: chr(96 + n),   # a, b, c
    },
}


def _build_number(style: str, counters: list[int], level: int) -> str:
    """根据 style 和各级计数器生成编号字符串。"""
    if style in ("technical", "chinese_bidding"):
        # technical: join counters with dot
        return ".".join(str(c) for c in counters[:level])
    if style == "chinese_chapter":
        fn = STYLES["chinese_chapter"].get(level)
        return fn(counters[level - 1], None) if fn else str(counters[level - 1])
    if style == "academic":
        fn = STYLES["academic"].get(level)
        return fn(counters[level - 1], None) if fn else str(counters[level - 1])
    return str(counters[level - 1])


def _format_number(style: str, num_str: str, level: int) -> str:
    """给编号字符串加上样式对应的后缀/装饰。"""
    if style == "chinese_bidding":
        if level == 1:
            return f"{_chinese_num(int(num_str.split('.')[-1]))}、" if "." not in num_str else f"{num_str} "
        return f"{num_str} "
    if style in ("technical",):
        return f"{num_str} "
    if style == "chinese_chapter":
        if level == 1:
            cn = _chinese_num(int(num_str))
            return f"第{cn}章 "
        elif level == 2:
            return f"{_chinese_num(int(num_str))}、"
        elif level == 3:
            return f"（{_chinese_num(int(num_str))}）"
        else:
            return f"{num_str}. "
    if style == "academic":
        if level == 1:
            return f"{_roman(int(num_str))} "
        elif level == 2:
            return f"{chr(64 + int(num_str))} "
        elif level == 3:
            return f"{num_str} "
        else:
            return f"{chr(96 + int(num_str))} "
    return f"{num_str} "


# ── 标题升降 ─────────────────────────────────────────────────

def heading_shift(content: str, delta: int) -> str:
    """delta > 0 升级（减少 #），delta < 0 降级（增加 #）。"""
    lines = content.split("\n")
    result = []
    for i, line in enumerate(lines):
        if is_in_code_block(lines, i):
            result.append(line)
            continue
        m = re.match(r"^(#{1,6})(\s+.*)", line)
        if m:
            level = len(m.group(1))
            new_level = max(1, min(6, level - delta))
            result.append("#" * new_level + m.group(2))
        else:
            result.append(line)
    return "\n".join(result)


# ── 编号管理 ─────────────────────────────────────────────────

# 匹配现有编号前缀（中文章节、数字点、罗马数字、字母等）
_NUMBER_PREFIX = re.compile(
    r"^(?:"
    r"第[一二三四五六七八九十]+章\s*|"   # 第X章
    r"[一二三四五六七八九十]+[、，]\s*|"  # 一、
    r"（[一二三四五六七八九十]+）\s*|"    # （一）
    r"[IVXLCDM]+\s+|"                    # 罗马
    r"[A-Z]\s+|"                         # A B C
    r"[a-z]\s+|"                         # a b c
    r"[\d]+(?:\.[\d]+)*\.?\s+"           # 1 / 1.1 / 1.1.1
    r")"
)


def numbering_remove(content: str) -> str:
    lines = content.split("\n")
    result = []
    for i, line in enumerate(lines):
        if is_in_code_block(lines, i):
            result.append(line)
            continue
        m = re.match(r"^(#{1,6})\s+(.*)", line)
        if m:
            hashes = m.group(1)
            text = m.group(2)
            text = _NUMBER_PREFIX.sub("", text).strip()
            result.append(f"{hashes} {text}")
        else:
            result.append(line)
    return "\n".join(result)


def numbering_add(content: str, style: str, start_from: int) -> str:
    lines = content.split("\n")
    # 先移除现有编号
    content_clean = numbering_remove(content)
    lines = content_clean.split("\n")

    counters = [0] * 6
    counters[0] = start_from - 1  # h1 从 start_from 开始
    result = []

    for i, line in enumerate(lines):
        if is_in_code_block(lines, i):
            result.append(line)
            continue
        m = re.match(r"^(#{1,6})\s+(.*)", line)
        if m:
            level = len(m.group(1))
            text = m.group(2).strip()
            counters[level - 1] += 1
            # 重置下级计数器
            for j in range(level, 6):
                counters[j] = 0

            num_str = _build_number(style, counters, level)
            prefix = _format_number(style, num_str, level)
            result.append(f"{'#' * level} {prefix}{text}")
        else:
            result.append(line)
    return "\n".join(result)


# ── TOC 生成 ─────────────────────────────────────────────────

def _heading_to_anchor(text: str) -> str:
    """将标题文本转换为 GitHub 风格锚点（支持中文）。"""
    # 移除编号前缀后再生成锚点
    text = _NUMBER_PREFIX.sub("", text).strip()
    anchor = text.lower()
    anchor = re.sub(r"[^\w\u4e00-\u9fff\- ]", "", anchor)
    anchor = anchor.replace(" ", "-")
    return anchor


def toc_generate(content: str, depth: int, position: str) -> str:
    lines = content.split("\n")
    # 找到最小标题级别作为 TOC 根级
    headings = []
    for i, line in enumerate(lines):
        if is_in_code_block(lines, i):
            continue
        m = re.match(r"^(#{1,6})\s+(.*)", line)
        if m:
            level = len(m.group(1))
            text = m.group(2).strip()
            if level <= depth:
                headings.append((level, text))

    if not headings:
        return content

    min_level = min(h[0] for h in headings)
    toc_lines = ["<!-- TOC -->"]
    for level, text in headings:
        indent = "  " * (level - min_level)
        clean_text = _NUMBER_PREFIX.sub("", text).strip()
        anchor = _heading_to_anchor(text)
        toc_lines.append(f"{indent}- [{clean_text}](#{anchor})")
    toc_lines.append("<!-- /TOC -->")
    toc_block = "\n".join(toc_lines)

    # 删除旧 TOC（如果存在）
    content = re.sub(r"<!-- TOC -->.*?<!-- /TOC -->", "", content, flags=re.DOTALL).strip()
    lines = content.split("\n")

    if position == "top":
        return toc_block + "\n\n" + content

    # after-h1：在第一个 h1 后插入
    insert_at = len(lines)  # 默认插在末尾
    for i, line in enumerate(lines):
        if re.match(r"^#\s+", line):
            insert_at = i + 1
            break
    lines.insert(insert_at, "\n" + toc_block + "\n")
    return "\n".join(lines)


# ── CLI ──────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Markdown 结构操作工具")
    parser.add_argument("input", help="输入 Markdown 文件路径")
    parser.add_argument("action", choices=["heading", "numbering", "toc"], help="操作类型")
    parser.add_argument("subaction", help="子操作（upgrade/downgrade/add/remove/generate）")
    parser.add_argument("-o", "--output", help="输出文件路径（默认覆盖原文件）")
    parser.add_argument("--levels", type=int, default=1, help="heading 操作的级别数（默认 1）")
    parser.add_argument("--style", help="编号样式：chinese_bidding / chinese_chapter / technical / academic")
    parser.add_argument("--start-from", type=int, default=1, dest="start_from", help="h1 起始编号（默认 1）")
    parser.add_argument("--depth", type=int, default=3, help="TOC 最深标题级别（默认 3）")
    parser.add_argument("--position", choices=["top", "after-h1"], default="after-h1", help="TOC 插入位置")

    args = parser.parse_args()
    content = read_utf8(args.input)

    if args.action == "heading":
        if args.subaction == "upgrade":
            result = heading_shift(content, delta=args.levels)
        elif args.subaction == "downgrade":
            result = heading_shift(content, delta=-args.levels)
        else:
            parser.error(f"heading 不支持子操作: {args.subaction}")

    elif args.action == "numbering":
        if args.subaction == "remove":
            result = numbering_remove(content)
        elif args.subaction == "add":
            if not args.style:
                parser.error("numbering add 需要 --style 参数")
            if args.style not in ("chinese_bidding", "chinese_chapter", "technical", "academic"):
                parser.error(f"不支持的样式: {args.style}")
            result = numbering_add(content, args.style, args.start_from)
        else:
            parser.error(f"numbering 不支持子操作: {args.subaction}")

    elif args.action == "toc":
        if args.subaction == "generate":
            result = toc_generate(content, args.depth, args.position)
        else:
            parser.error(f"toc 不支持子操作: {args.subaction}")

    out_path = args.output or args.input
    write_utf8(out_path, result)
    print(f"✓ 已写入: {out_path}")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: 验证 `structure.py`**

在 `markdown-master/` 目录下运行：

```bash
# 验证 heading upgrade（输出到临时文件，不覆盖原文件）
uv run python scripts/structure.py examples/sample_cn.md heading upgrade -o examples/tmp_upgrade.md
type examples\tmp_upgrade.md

# 验证 numbering add
uv run python scripts/structure.py examples/sample_cn.md numbering add --style chinese_chapter -o examples/tmp_numbered.md
type examples\tmp_numbered.md

# 验证 toc generate
uv run python scripts/structure.py examples/sample_cn.md toc generate --depth 3 -o examples/tmp_toc.md
type examples\tmp_toc.md

# 清理临时文件
del examples\tmp_upgrade.md examples\tmp_numbered.md examples\tmp_toc.md
```

期望：每条命令输出 `✓ 已写入: ...`，内容符合对应操作预期。

- [ ] **Step 3: Commit**

```bash
git add markdown-master/scripts/structure.py
git commit -m "feat(markdown-master): add structure.py - heading/numbering/toc operations"
```

---

## Task 3: `quality.py` — 质检操作

**Files:**
- Create: `markdown-master/scripts/quality.py`

- [ ] **Step 1: 实现 `quality.py`**

```python
# markdown-master/scripts/quality.py
"""
质检操作：格式规范检查（lint）、中文排版检查（zhlint）、链接检查（linkcheck）。

用法:
  uv run python scripts/quality.py <input.md> lint [--fix]
  uv run python scripts/quality.py <input.md> zhlint [--fix]
  uv run python scripts/quality.py <input.md> linkcheck [--check-remote]
"""
import argparse
import re
import sys
import urllib.request
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from _md_utils import read_utf8, write_utf8, is_in_code_block

Issue = tuple[int, str]  # (line_number, description)


# ── lint ──────────────────────────────────────────────────────

def run_lint(content: str, fix: bool) -> tuple[str, list[Issue]]:
    lines = content.split("\n")
    issues: list[Issue] = []
    result_lines = list(lines)

    # 1. 检查是否有 h1
    has_h1 = any(re.match(r"^#\s+", l) for l in lines)
    if not has_h1:
        issues.append((0, "缺少 H1 标题（文档应有且仅有一个 # 一级标题）"))

    # 2. 标题层级跳跃
    prev_level = 0
    for i, line in enumerate(lines):
        if is_in_code_block(lines, i):
            continue
        m = re.match(r"^(#{1,6})\s+", line)
        if m:
            level = len(m.group(1))
            if prev_level > 0 and level > prev_level + 1:
                issues.append((i + 1, f"标题层级跳跃：H{prev_level} → H{level}（中间缺少 H{prev_level + 1}）"))
            prev_level = level

    # 3. 连续 3 行以上空行 → 可修复
    i = 0
    while i < len(result_lines):
        if result_lines[i].strip() == "":
            j = i
            while j < len(result_lines) and result_lines[j].strip() == "":
                j += 1
            if j - i >= 3:
                issues.append((i + 1, f"连续空行过多（{j - i} 行），建议保留 1-2 行"))
                if fix:
                    result_lines[i:j] = ["", ""]
            i = j
        else:
            i += 1

    # 4. 代码块未指定语言
    for i, line in enumerate(lines):
        if re.match(r"^```\s*$", line):
            issues.append((i + 1, "代码块未指定语言（如 ```python）"))

    return "\n".join(result_lines), issues


# ── zhlint ────────────────────────────────────────────────────

# 中文字符范围
_ZH = r"[\u4e00-\u9fff\u3400-\u4dbf\uf900-\ufaff]"
_EN = r"[A-Za-z0-9]"

# 中文后紧跟英文/数字（缺空格）
_ZH_BEFORE_EN = re.compile(f"({_ZH})({_EN})")
# 英文/数字后紧跟中文（缺空格）
_EN_BEFORE_ZH = re.compile(f"({_EN})({_ZH})")

# 半角标点（在中文上下文中）
_HALF_PUNCT = re.compile(f"({_ZH})[,\\.\\?!;:]")

# 全角数字/字母
_FULLWIDTH = re.compile(r"[！-～]")  # U+FF01-FF5E 全角 ASCII


def _is_inline_code(line: str, pos: int) -> bool:
    """简单判断 pos 位置是否在行内代码 `` ` `` 内。"""
    count = 0
    for i, ch in enumerate(line):
        if ch == "`":
            count += 1
        if i == pos:
            break
    return count % 2 == 1


def run_zhlint(content: str, fix: bool) -> tuple[str, list[Issue]]:
    lines = content.split("\n")
    issues: list[Issue] = []
    result_lines = []

    for i, line in enumerate(lines):
        if is_in_code_block(lines, i):
            result_lines.append(line)
            continue

        new_line = line

        # 中文与英文之间缺空格
        if _ZH_BEFORE_EN.search(new_line) or _EN_BEFORE_ZH.search(new_line):
            issues.append((i + 1, "中文与英文/数字之间缺少空格"))
            if fix:
                new_line = _ZH_BEFORE_EN.sub(r"\1 \2", new_line)
                new_line = _EN_BEFORE_ZH.sub(r"\1 \2", new_line)

        # 中文语境中使用半角标点
        if _HALF_PUNCT.search(new_line):
            issues.append((i + 1, "中文语境中使用了半角标点"))
            if fix:
                punct_map = {",": "，", ".": "。", "?": "？", "!": "！", ";": "；", ":": "："}
                def replace_punct(m):
                    zh_char = m.group(0)[0]
                    half = m.group(0)[1]
                    return zh_char + punct_map.get(half, half)
                new_line = _HALF_PUNCT.sub(replace_punct, new_line)

        result_lines.append(new_line)

    return "\n".join(result_lines), issues


# ── linkcheck ─────────────────────────────────────────────────

_MD_LINK = re.compile(r"\[.*?\]\((.*?)\)")
_MD_IMAGE = re.compile(r"!\[.*?\]\((.*?)\)")


def run_linkcheck(content: str, file_path: str, check_remote: bool) -> list[Issue]:
    base_dir = Path(file_path).parent
    lines = content.split("\n")
    issues: list[Issue] = []

    for i, line in enumerate(lines):
        if is_in_code_block(lines, i):
            continue

        # 图片引用（先检查，优先级高于普通链接）
        for m in _MD_IMAGE.finditer(line):
            src = m.group(1).split(" ")[0]  # 去掉可能的 title
            if src.startswith("http://") or src.startswith("https://"):
                if check_remote:
                    _check_remote_url(src, i + 1, issues)
            else:
                target = (base_dir / src).resolve()
                if not target.exists():
                    issues.append((i + 1, f"图片路径不存在: {src}"))

        # 普通链接
        for m in _MD_LINK.finditer(line):
            href = m.group(1).split(" ")[0]
            if href.startswith("#"):
                continue  # 锚点链接跳过
            if href.startswith("http://") or href.startswith("https://"):
                if check_remote:
                    _check_remote_url(href, i + 1, issues)
            else:
                target = (base_dir / href).resolve()
                if not target.exists():
                    issues.append((i + 1, f"本地链接目标不存在: {href}"))

    return issues


def _check_remote_url(url: str, line_no: int, issues: list[Issue]) -> None:
    try:
        req = urllib.request.Request(url, method="HEAD")
        req.add_header("User-Agent", "markdown-master-linkcheck/1.0")
        with urllib.request.urlopen(req, timeout=5) as resp:
            if resp.status >= 400:
                issues.append((line_no, f"远程链接返回 {resp.status}: {url}"))
    except Exception as e:
        issues.append((line_no, f"远程链接无法访问 ({e}): {url}"))


# ── 输出格式 ──────────────────────────────────────────────────

def print_issues(issues: list[Issue], action: str) -> None:
    if not issues:
        print(f"✓ No issues found ({action})")
        return
    for line_no, desc in issues:
        loc = f"L{line_no}" if line_no > 0 else "文档级"
        print(f"  [{loc}] {desc}")
    print(f"\n共发现 {len(issues)} 个问题")


# ── CLI ──────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Markdown 质检工具")
    parser.add_argument("input", help="输入 Markdown 文件路径")
    parser.add_argument("action", choices=["lint", "zhlint", "linkcheck"], help="检查类型")
    parser.add_argument("--fix", action="store_true", help="自动修复可修复的问题")
    parser.add_argument("--check-remote", action="store_true", dest="check_remote",
                        help="linkcheck 时同时检查 HTTP/HTTPS 链接")

    args = parser.parse_args()
    content = read_utf8(args.input)

    if args.action == "lint":
        result, issues = run_lint(content, args.fix)
        print_issues(issues, "lint")
        if args.fix and issues:
            write_utf8(args.input, result)
            print(f"✓ 已修复并写入: {args.input}")

    elif args.action == "zhlint":
        result, issues = run_zhlint(content, args.fix)
        print_issues(issues, "zhlint")
        if args.fix and issues:
            write_utf8(args.input, result)
            print(f"✓ 已修复并写入: {args.input}")

    elif args.action == "linkcheck":
        issues = run_linkcheck(content, args.input, args.check_remote)
        print_issues(issues, "linkcheck")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: 验证 `quality.py`**

```bash
# lint 检查（sample_cn.md 有标题跳跃问题，sample_en.md 缺 h1）
uv run python scripts/quality.py examples/sample_cn.md lint
uv run python scripts/quality.py examples/sample_en.md lint

# zhlint 检查（sample_cn.md 有中英文间距问题）
uv run python scripts/quality.py examples/sample_cn.md zhlint

# linkcheck（sample_en.md 有死链和不存在的图片）
uv run python scripts/quality.py examples/sample_en.md linkcheck
```

期望：
- `sample_cn.md lint` → 报告标题层级跳跃（第四部分跳过第三部分）
- `sample_en.md lint` → 报告缺少 H1、标题层级跳跃
- `sample_cn.md zhlint` → 报告中英文间距、半角标点问题
- `sample_en.md linkcheck` → 报告 `nonexistent_file.md` 不存在、`logo.png` 不存在

- [ ] **Step 3: Commit**

```bash
git add markdown-master/scripts/quality.py
git commit -m "feat(markdown-master): add quality.py - lint/zhlint/linkcheck"
```

---

## Task 4: `convert.py` — 格式转换

**Files:**
- Create: `markdown-master/scripts/convert.py`

- [ ] **Step 1: 安装依赖**

```bash
pip install python-docx markdown beautifulsoup4 Pillow weasyprint
```

- [ ] **Step 2: 实现 `convert.py`**

```python
# markdown-master/scripts/convert.py
"""
格式转换：Markdown → DOCX / HTML / TXT / PDF

用法:
  uv run python scripts/convert.py <input.md 或 目录/> --to <format> [-o output] [--template template.docx]

支持批量：输入为目录时，递归处理所有 .md 文件。
"""
import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from _md_utils import read_utf8


# ── 转换函数 ──────────────────────────────────────────────────

def to_html(content: str) -> str:
    import markdown as md_lib
    css = """<style>
body { font-family: 'Segoe UI', Arial, sans-serif; max-width: 860px; margin: 40px auto; padding: 0 20px; line-height: 1.7; }
h1,h2,h3,h4,h5,h6 { border-bottom: 1px solid #eee; padding-bottom: 4px; }
code { background: #f4f4f4; padding: 2px 6px; border-radius: 3px; font-family: monospace; }
pre { background: #f4f4f4; padding: 16px; border-radius: 4px; overflow-x: auto; }
blockquote { border-left: 4px solid #ccc; margin: 0; padding: 0 16px; color: #666; }
table { border-collapse: collapse; width: 100%; }
th, td { border: 1px solid #ddd; padding: 8px 12px; }
th { background: #f0f0f0; }
</style>"""
    html_body = md_lib.markdown(
        content,
        extensions=["tables", "fenced_code", "toc", "nl2br"]
    )
    return f"<!DOCTYPE html>\n<html>\n<head>\n<meta charset='utf-8'>\n{css}\n</head>\n<body>\n{html_body}\n</body>\n</html>"


def to_txt(content: str) -> str:
    # 移除 frontmatter
    if content.startswith("---"):
        parts = content.split("---", 2)
        if len(parts) >= 3:
            content = parts[2].strip()
    # 移除 Markdown 标记
    text = re.sub(r"^#{1,6}\s+", "", content, flags=re.MULTILINE)  # 标题
    text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)  # 粗体
    text = re.sub(r"\*(.+?)\*", r"\1", text)  # 斜体
    text = re.sub(r"`(.+?)`", r"\1", text)  # 行内代码
    text = re.sub(r"```.*?```", "", text, flags=re.DOTALL)  # 代码块
    text = re.sub(r"!\[.*?\]\(.*?\)", "", text)  # 图片
    text = re.sub(r"\[(.+?)\]\(.*?\)", r"\1", text)  # 链接
    text = re.sub(r"^[-*+]\s+", "", text, flags=re.MULTILINE)  # 列表
    text = re.sub(r"^\d+\.\s+", "", text, flags=re.MULTILINE)  # 有序列表
    text = re.sub(r"^>+\s?", "", text, flags=re.MULTILINE)  # 引用
    text = re.sub(r"^---+$", "", text, flags=re.MULTILINE)  # 分割线
    text = re.sub(r"\n{3,}", "\n\n", text)  # 多余空行
    return text.strip()


def to_docx(content: str, input_path: str, template: str | None = None) -> bytes:
    from docx import Document
    from docx.shared import Inches, Pt
    import markdown as md_lib
    from bs4 import BeautifulSoup
    import io

    if template and Path(template).exists():
        doc = Document(template)
        # 清除模板正文内容
        for para in list(doc.paragraphs):
            p = para._element
            p.getparent().remove(p)
    else:
        doc = Document()

    base_dir = Path(input_path).parent
    html = md_lib.markdown(content, extensions=["tables", "fenced_code"])
    soup = BeautifulSoup(html, "html.parser")

    HEADING_MAP = {"h1": 1, "h2": 2, "h3": 3, "h4": 4, "h5": 5, "h6": 6}

    for elem in soup.children:
        tag = getattr(elem, "name", None)
        if tag is None:
            continue
        if tag in HEADING_MAP:
            doc.add_heading(elem.get_text(), level=HEADING_MAP[tag])
        elif tag == "p":
            # 检查是否包含图片
            img = elem.find("img")
            if img:
                src = img.get("src", "")
                if not src.startswith("http"):
                    img_path = (base_dir / src).resolve()
                    if img_path.exists():
                        doc.add_picture(str(img_path), width=Inches(5.5))
                        if img.get("alt"):
                            doc.add_paragraph(img.get("alt"), style="Caption")
            else:
                doc.add_paragraph(elem.get_text())
        elif tag in ("ul", "ol"):
            for li in elem.find_all("li", recursive=False):
                doc.add_paragraph(li.get_text(), style="List Bullet" if tag == "ul" else "List Number")
        elif tag == "pre":
            code = elem.find("code")
            text = code.get_text() if code else elem.get_text()
            para = doc.add_paragraph()
            run = para.add_run(text)
            run.font.name = "Courier New"
            run.font.size = Pt(9)
        elif tag == "blockquote":
            doc.add_paragraph(elem.get_text(), style="Quote")
        elif tag == "hr":
            doc.add_paragraph("─" * 40)
        elif tag == "table":
            rows = elem.find_all("tr")
            if not rows:
                continue
            cols = max(len(r.find_all(["td", "th"])) for r in rows)
            table = doc.add_table(rows=len(rows), cols=cols)
            table.style = "Table Grid"
            for r_idx, row in enumerate(rows):
                cells = row.find_all(["td", "th"])
                for c_idx, cell in enumerate(cells):
                    table.cell(r_idx, c_idx).text = cell.get_text()

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def to_pdf(content: str) -> bytes:
    from weasyprint import HTML as WP_HTML
    html = to_html(content)
    return WP_HTML(string=html).write_pdf()


# ── 文件处理逻辑 ──────────────────────────────────────────────

def convert_file(input_path: Path, fmt: str, output_path: Path, template: str | None) -> None:
    content = read_utf8(str(input_path))

    if fmt == "html":
        result = to_html(content)
        output_path.write_text(result, encoding="utf-8")
    elif fmt == "txt":
        result = to_txt(content)
        output_path.write_text(result, encoding="utf-8")
    elif fmt == "docx":
        result = to_docx(content, str(input_path), template)
        output_path.write_bytes(result)
    elif fmt == "pdf":
        result = to_pdf(content)
        output_path.write_bytes(result)

    print(f"✓ {input_path} → {output_path}")


# ── CLI ──────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Markdown 格式转换工具")
    parser.add_argument("input", help="输入 Markdown 文件或目录")
    parser.add_argument("--to", required=True, choices=["docx", "html", "txt", "pdf"],
                        dest="fmt", help="目标格式")
    parser.add_argument("-o", "--output", help="输出路径（文件或目录）")
    parser.add_argument("--template", help="DOCX 模板文件路径（仅 --to docx 时有效）")

    args = parser.parse_args()
    input_path = Path(args.input)

    if input_path.is_dir():
        # 批量模式
        md_files = list(input_path.rglob("*.md"))
        out_dir = Path(args.output) if args.output else input_path / f"{args.fmt}_output"
        out_dir.mkdir(parents=True, exist_ok=True)
        for md_file in md_files:
            rel = md_file.relative_to(input_path)
            out_file = out_dir / rel.with_suffix(f".{args.fmt}")
            out_file.parent.mkdir(parents=True, exist_ok=True)
            convert_file(md_file, args.fmt, out_file, args.template)
    else:
        # 单文件模式
        if args.output:
            out_file = Path(args.output)
        else:
            out_file = input_path.with_suffix(f".{args.fmt}")
        out_file.parent.mkdir(parents=True, exist_ok=True)
        convert_file(input_path, args.fmt, out_file, args.template)


if __name__ == "__main__":
    main()
```

- [ ] **Step 3: 验证 `convert.py`**

```bash
# 转换为 HTML
uv run python scripts/convert.py examples/sample_cn.md --to html -o examples/sample_cn.html

# 转换为 TXT
uv run python scripts/convert.py examples/sample_cn.md --to txt -o examples/sample_cn.txt

# 转换为 DOCX
uv run python scripts/convert.py examples/sample_cn.md --to docx -o examples/sample_cn.docx
```

期望：每条命令输出 `✓ ... → ...`，对应输出文件存在且内容正确。

- [ ] **Step 4: Commit**

```bash
git add markdown-master/scripts/convert.py
git commit -m "feat(markdown-master): add convert.py - DOCX/HTML/TXT/PDF conversion"
```

---

## Task 5: `files.py` — 文件级操作

**Files:**
- Create: `markdown-master/scripts/files.py`

- [ ] **Step 1: 实现 `files.py`**

```python
# markdown-master/scripts/files.py
"""
文件级操作：拆分（split）、合并（merge）。

用法:
  uv run python scripts/files.py split <input.md> --by h1|h2|h3 [--output-dir DIR]
  uv run python scripts/files.py merge <file1> <file2> ... -o output.md [--separator SEP]
  uv run python scripts/files.py merge <directory/> -o output.md
"""
import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from _md_utils import read_utf8, write_utf8, is_in_code_block


def _sanitize_filename(text: str) -> str:
    """将标题文本转为合法文件名，清理非法字符。"""
    # 移除编号前缀
    text = re.sub(
        r"^(?:第[一二三四五六七八九十]+章\s*|[一二三四五六七八九十]+[、，]\s*|（[一二三四五六七八九十]+）\s*|[IVXLCDM]+\s+|[A-Z]\s+|[\d]+(?:\.[\d]+)*\.?\s+)",
        "", text
    ).strip()
    # 替换 Windows/Linux 非法字符
    text = re.sub(r'[\\/:*?"<>|]', "_", text)
    text = text.strip(". ")
    return text[:80] or "untitled"  # 限制长度


# ── split ──────────────────────────────────────────────────────

def cmd_split(input_path: str, by: str, output_dir: str | None) -> None:
    level = int(by[1])  # "h2" → 2
    content = read_utf8(input_path)
    lines = content.split("\n")

    out_dir = Path(output_dir) if output_dir else Path(input_path).parent / "split_output"
    out_dir.mkdir(parents=True, exist_ok=True)

    sections: list[tuple[str, list[str]]] = []  # [(title, lines)]
    current_title = "_preamble"
    current_lines: list[str] = []

    for i, line in enumerate(lines):
        if is_in_code_block(lines, i):
            current_lines.append(line)
            continue
        m = re.match(r"^(#{1,6})\s+(.*)", line)
        if m and len(m.group(1)) == level:
            if current_lines or current_title != "_preamble":
                sections.append((current_title, current_lines))
            current_title = m.group(2).strip()
            current_lines = [line]
        else:
            current_lines.append(line)

    if current_lines:
        sections.append((current_title, current_lines))

    count = 0
    for i, (title, sec_lines) in enumerate(sections):
        if title == "_preamble" and not any(l.strip() for l in sec_lines):
            continue
        fname = f"{i:02d}_{_sanitize_filename(title)}.md" if title != "_preamble" else "00_preamble.md"
        fpath = out_dir / fname
        write_utf8(str(fpath), "\n".join(sec_lines))
        print(f"  ✓ {fpath}")
        count += 1

    print(f"\n共拆分为 {count} 个文件 → {out_dir}")


# ── merge ──────────────────────────────────────────────────────

def cmd_merge(inputs: list[str], output: str, separator: str) -> None:
    md_files: list[Path] = []

    for inp in inputs:
        p = Path(inp)
        if p.is_dir():
            md_files.extend(sorted(p.rglob("*.md")))
        elif p.is_file() and p.suffix == ".md":
            md_files.append(p)
        else:
            print(f"  跳过: {inp}（不是 .md 文件或目录）")

    if not md_files:
        print("错误：没有找到可合并的 .md 文件")
        sys.exit(1)

    parts = []
    for f in md_files:
        print(f"  + {f}")
        parts.append(read_utf8(str(f)))

    merged = separator.join(parts)
    write_utf8(output, merged)
    print(f"\n✓ 已合并 {len(parts)} 个文件 → {output}")


# ── CLI ──────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Markdown 文件级操作工具")
    sub = parser.add_subparsers(dest="action", required=True)

    # split
    sp = sub.add_parser("split", help="按标题级别拆分文档")
    sp.add_argument("input", help="输入 Markdown 文件")
    sp.add_argument("--by", choices=["h1", "h2", "h3"], default="h2",
                    help="按哪一级标题拆分（默认 h2）")
    sp.add_argument("--output-dir", dest="output_dir", help="输出目录（默认: <input_dir>/split_output/）")

    # merge
    mp = sub.add_parser("merge", help="合并多个 Markdown 文件")
    mp.add_argument("inputs", nargs="+", help="输入文件或目录（可多个）")
    mp.add_argument("-o", "--output", required=True, help="输出文件路径")
    mp.add_argument("--separator", default="\n\n---\n\n", help="文件间分隔符（默认: \\n\\n---\\n\\n）")

    args = parser.parse_args()

    if args.action == "split":
        cmd_split(args.input, args.by, args.output_dir)
    elif args.action == "merge":
        cmd_merge(args.inputs, args.output, args.separator)


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: 验证 `files.py`**

```bash
# split：按 h2 拆分 sample_cn.md
uv run python scripts/files.py split examples/sample_cn.md --by h2 --output-dir examples/split_output
dir examples\split_output

# merge：把拆分的文件重新合并
uv run python scripts/files.py merge examples/split_output -o examples/merged.md
type examples\merged.md

# 清理
rmdir /s /q examples\split_output
del examples\merged.md
```

期望：split 输出若干 `XX_<标题名>.md` 文件；merge 合并后内容与原文件结构一致。

- [ ] **Step 3: Commit**

```bash
git add markdown-master/scripts/files.py
git commit -m "feat(markdown-master): add files.py - split and merge operations"
```

---

## Task 6: `SKILL.md` — AI 调用手册

**Files:**
- Create: `markdown-master/SKILL.md`

- [ ] **Step 1: 创建 `SKILL.md`**

```markdown
---
name: markdown-master
description: >
  处理、编辑、检查、转换 Markdown 文档。
  用于：调整标题层级、添加中文章节编号、生成目录、格式规范检查（lint）、
  中文排版检查（zhlint）、链接/图片死链检查、
  转换为 DOCX/HTML/TXT/PDF、按标题拆分或合并多个文件。
---

# Markdown Master

## Quick Start

```bash
# 添加中文章节编号
uv run python scripts/structure.py doc.md numbering add --style chinese_chapter

# 生成目录（插入到 H1 后方）
uv run python scripts/structure.py doc.md toc generate

# 格式检查并自动修复
uv run python scripts/quality.py doc.md lint --fix

# 中文排版检查并修复
uv run python scripts/quality.py doc.md zhlint --fix

# 检查死链
uv run python scripts/quality.py doc.md linkcheck

# 转换为 Word 文档
uv run python scripts/convert.py doc.md --to docx

# 按 H2 拆分文档
uv run python scripts/files.py split doc.md --by h2

# 合并多个文件
uv run python scripts/files.py merge ch1.md ch2.md ch3.md -o book.md
```

## Dependencies

```bash
pip install python-docx markdown beautifulsoup4 Pillow weasyprint
```

structure.py / quality.py / files.py 仅使用 Python 标准库，无需额外安装。

## Scripts

### structure.py — 结构操作

操作单个文档的内部文本结构。

| 用法 | 说明 |
|------|------|
| `structure.py <file> heading upgrade [--levels N]` | 所有标题升 N 级（默认1）|
| `structure.py <file> heading downgrade [--levels N]` | 所有标题降 N 级 |
| `structure.py <file> numbering add --style <s> [--start-from N]` | 添加层级编号 |
| `structure.py <file> numbering remove` | 移除所有编号 |
| `structure.py <file> toc generate [--depth N] [--position top\|after-h1]` | 插入/更新目录 |

所有操作默认覆盖原文件，使用 `-o output.md` 输出到新文件。

**编号样式（--style）：**

| 样式 | 效果示例 |
|------|---------|
| `chinese_chapter` | 第一章 / 一、/ （一）/ 1. |
| `chinese_bidding` | 一、/ 1.1 / 1.1.1 |
| `technical` | 1 / 1.1 / 1.1.1 |
| `academic` | I / A / 1 / a |

### quality.py — 质检操作

| 用法 | 说明 |
|------|------|
| `quality.py <file> lint [--fix]` | 格式规范检查，`--fix` 自动修复可修复项 |
| `quality.py <file> zhlint [--fix]` | 中文排版检查，`--fix` 自动修复 |
| `quality.py <file> linkcheck [--check-remote]` | 本地链接/图片检查，加参数检查 HTTP 链接 |

输出格式：`[L行号] 问题描述`，最后汇总问题数。无问题输出 `✓ No issues found`。

**lint 检查项（✓=可修复）：**
- ✗ 标题层级跳跃（H1 → H3）
- ✗ 文件缺少 H1
- ✓ 连续 3 行以上空行
- ✓ 列表缩进不一致
- ✗ 代码块未指定语言

**zhlint 检查项（✓=可修复）：**
- ✓ 中文与英文/数字间缺空格
- ✓ 中文语境中使用半角标点
- ✓ 全角数字/字母

### convert.py — 格式转换

```bash
uv run python scripts/convert.py <input.md 或 目录/> --to <format> [-o output] [--template template.docx]
```

| 格式 | 说明 |
|------|------|
| `docx` | Word 文档，支持 `--template` 指定模板 |
| `html` | 自包含 HTML（内联 CSS）|
| `txt` | 纯文本（去除所有 Markdown 标记）|
| `pdf` | PDF（通过 weasyprint，无需 Pandoc）|

输入为目录时批量转换所有 `.md` 文件。

### files.py — 文件级操作

```bash
# 拆分
uv run python scripts/files.py split <input.md> --by h1|h2|h3 [--output-dir DIR]

# 合并（支持文件列表或目录）
uv run python scripts/files.py merge <file1> [file2 ...] -o output.md [--separator SEP]
```

split 默认输出到 `<input_dir>/split_output/`，文件名为 `XX_<标题文本>.md`。
merge 输入目录时，按文件名字母序合并所有 `.md` 文件。
```

- [ ] **Step 2: 验证 SKILL.md 格式**

```bash
# 检查 SKILL.md 是否有 frontmatter
python -c "
content = open('SKILL.md', encoding='utf-8').read()
assert content.startswith('---'), 'SKILL.md 缺少 frontmatter'
assert 'markdown-master' in content, 'SKILL.md 缺少 name 字段'
print('✓ SKILL.md 格式正确')
"
```

- [ ] **Step 3: Commit**

```bash
git add markdown-master/SKILL.md
git commit -m "feat(markdown-master): add SKILL.md - AI invocation manual"
```

---

## Task 7: 端到端验证

- [ ] **Step 1: 完整流程验证（structure）**

```bash
# 在 markdown-master/ 目录下执行
# 添加编号 → 生成目录 → 检查 lint
uv run python scripts/structure.py examples/sample_cn.md numbering add --style chinese_chapter -o examples/cn_numbered.md
uv run python scripts/structure.py examples/cn_numbered.md toc generate
uv run python scripts/quality.py examples/cn_numbered.md lint
```

期望：lint 无编号相关错误，目录块已正确插入。

- [ ] **Step 2: 完整流程验证（quality）**

```bash
# zhlint fix 后再 lint
uv run python scripts/quality.py examples/sample_cn.md zhlint --fix
uv run python scripts/quality.py examples/sample_cn.md zhlint
# 期望：fix 后 zhlint 报告 "✓ No issues found"

uv run python scripts/quality.py examples/sample_en.md linkcheck
# 期望：报告 nonexistent_file.md 不存在、logo.png 不存在
```

- [ ] **Step 3: 完整流程验证（convert）**

```bash
uv run python scripts/convert.py examples/sample_cn.md --to html -o examples/out.html
uv run python scripts/convert.py examples/sample_cn.md --to docx -o examples/out.docx
uv run python scripts/convert.py examples/sample_cn.md --to txt -o examples/out.txt
# 验证文件存在且大小合理
dir examples\out.*
```

- [ ] **Step 4: 完整流程验证（files）**

```bash
uv run python scripts/files.py split examples/sample_cn.md --by h2
dir examples\split_output
uv run python scripts/files.py merge examples/split_output -o examples/merged.md
```

- [ ] **Step 5: 清理验证产物**

```bash
del examples\cn_numbered.md examples\out.html examples\out.docx examples\out.txt examples\merged.md
rmdir /s /q examples\split_output
```

- [ ] **Step 6: 最终 commit**

```bash
git add markdown-master/
git commit -m "feat: complete markdown-master skill - structure/quality/convert/files"
```
