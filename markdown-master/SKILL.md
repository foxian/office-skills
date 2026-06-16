
---
name: markdown-master
description: &gt;
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
| `structure.py &lt;file&gt; heading upgrade [--levels N]` | 所有标题升 N 级（默认1）|
| `structure.py &lt;file&gt; heading downgrade [--levels N]` | 所有标题降 N 级 |
| `structure.py &lt;file&gt; numbering add --style &lt;s&gt; [--start-from N]` | 添加层级编号 |
| `structure.py &lt;file&gt; numbering remove` | 移除所有编号 |
| `structure.py &lt;file&gt; toc generate [--depth N] [--position top\|after-h1]` | 插入/更新目录 |

所有操作默认覆盖原文件，使用 `-o output.md` 输出到新文件。

**编号样式（--style）：**

| 样式 | 效果示例 |
|------|---------|
| `chinese_chapter` | 第1章 / 1、 / （1） / 1. |
| `chinese_bidding` | 1、 / 1.1 / 1.1.1 |
| `technical` | 1 / 1.1 / 1.1.1 |
| `academic` | I / A / 1 / a |

### quality.py — 质检操作

| 用法 | 说明 |
|------|------|
| `quality.py &lt;file&gt; lint [--fix]` | 格式规范检查，`--fix` 自动修复可修复项 |
| `quality.py &lt;file&gt; zhlint [--fix]` | 中文排版检查，`--fix` 自动修复 |
| `quality.py &lt;file&gt; linkcheck [--check-remote]` | 本地链接/图片检查，加参数检查 HTTP 链接 |

输出格式：`[L行号] 问题描述`，最后汇总问题数。无问题输出 `✓ No issues found`。

**lint 检查项：**
- ✗ 标题层级跳跃（H1 -&gt; H3）
- ✗ 文件缺少 H1
- ✓ 连续 3 行以上空行
- ✗ 代码块未指定语言

**zhlint 检查项：**
- ✓ 中文与英文/数字间缺空格
- ✓ 中文语境中使用半角标点

### convert.py — 格式转换

```bash
uv run python scripts/convert.py &lt;input.md 或 目录/&gt; --to &lt;format&gt; [-o output] [--template template.docx]
```

| 格式 | 说明 |
|------|------|
| `docx` | Word 文档，支持 `--template` 指定模板 |
| `html` | 自包含 HTML（内联 CSS）|
| `txt` | 纯文本（去除所有 Markdown 标记）|
| `pdf` | PDF（通过 weasyprint）|

输入为目录时批量转换所有 `.md` 文件。

### files.py — 文件级操作

```bash
# 拆分
uv run python scripts/files.py split &lt;input.md&gt; --by h1|h2|h3 [--output-dir DIR]

# 合并（支持文件列表或目录）
uv run python scripts/files.py merge &lt;file1&gt; [file2 ...] -o output.md [--separator SEP]
```

split 默认输出到 `&lt;input_dir&gt;/split_output/`，文件名为 `XX_&lt;标题文本&gt;.md`。
merge 输入目录时，按文件名字母序合并所有 `.md` 文件。

