# markdown-master 技能设计文档

## 背景与目标

`markdown-master` 是一个 AI Agent 技能（Skill），定位是对标 `word-master` 和 `docx` skill，为 AI Agent 提供完整的 Markdown 文档处理能力。

Agent 读取 `SKILL.md` 后，能够调用对应的 Python 脚本完成以下三类核心任务：
- **结构操作**：调整标题层级、管理编号、生成目录
- **质检操作**：格式规范检查、中文排版检查、链接检查
- **格式转换**：将 Markdown 转换为 DOCX、HTML、TXT、PDF

功能以中文场景为主、中英文通用。脚本独立运行，通过 `uv run python` 调用。

---

## 文件结构

```
markdown-master/
├── SKILL.md                   ← AI 调用手册（触发词、子命令速查、示例）
├── examples/
│   ├── sample_cn.md           ← 中文测试文档（含多级标题、中英混排、图片引用）
│   └── sample_en.md           ← 英文测试文档
└── scripts/
    ├── _md_utils.py           ← 共用解析库（非直接调用）
    ├── structure.py           ← 标题层级、编号、目录
    ├── quality.py             ← lint、中文排版、链接检查
    ├── convert.py             ← 格式转换（DOCX/HTML/TXT/PDF）
    └── files.py               ← 文件级操作（拆分、合并）
```

**设计原则：**
- 每个脚本职责单一：`structure.py` 只处理单文档内部变换；`files.py` 只处理跨文件操作
- `_md_utils.py` 提供所有脚本共用的解析和 IO 基础能力，不直接调用
- 统一子命令风格：`python scripts/<script>.py <input> <action> [options]`

---

## 各脚本详细设计

### `_md_utils.py` — 共用工具库

供其他脚本 import，不直接调用。提供：

- `parse_file(path)` → 返回 `(content, frontmatter_dict)`
- `extract_headings(content)` → 标题列表，每项含 `(level, text, line_number)`
- `read_utf8(path)` / `write_utf8(path, content)` → 统一 UTF-8 文件读写
- `is_in_code_block(lines, line_index)` → 判断某行是否在代码块内（防止误处理）

---

### `structure.py` — 结构操作

操作对象：单个 Markdown 文档的内部文本。

**调用格式：**
```bash
uv run python scripts/structure.py <input.md> <action> [options] [-o output.md]
```

| action | 必选参数 | 可选参数 | 说明 |
|--------|----------|----------|------|
| `heading upgrade` | — | `--levels N`（默认1） | 所有标题升 N 级（h2→h1） |
| `heading downgrade` | — | `--levels N`（默认1） | 所有标题降 N 级（h1→h2） |
| `numbering add` | `--style <样式>` | `--start-from N`（默认1） | 为标题添加层级编号 |
| `numbering remove` | — | — | 移除所有标题编号前缀 |
| `toc generate` | — | `--depth N`（默认3）、`--position top/after-h1`（默认after-h1） | 在文档内插入或更新目录块 |

**编号样式（`--style`）：**

| 样式名 | 示例 |
|--------|------|
| `chinese_bidding` | 一、/ 1.1 / 1.1.1 |
| `chinese_chapter` | 第一章 / 一、/ （一）/ 1. |
| `technical` | 1 / 1.1 / 1.1.1 |
| `academic` | I / A / 1 / a |

**默认输出：** 覆盖原文件。使用 `-o output.md` 输出到新文件。

**示例：**
```bash
# 升级标题
uv run python scripts/structure.py doc.md heading upgrade

# 添加招标文档编号，从第 2 章开始
uv run python scripts/structure.py doc.md numbering add --style chinese_bidding --start-from 2 -o numbered.md

# 在 h1 后面生成目录，最深到 h3
uv run python scripts/structure.py doc.md toc generate --depth 3 --position after-h1
```

---

### `quality.py` — 质检操作

**调用格式：**
```bash
uv run python scripts/quality.py <input.md> <action> [options]
```

| action | 可选参数 | 说明 |
|--------|----------|------|
| `lint` | `--fix` | 格式规范检查；`--fix` 自动修复可修复项 |
| `zhlint` | `--fix` | 中文排版检查；`--fix` 自动修复 |
| `linkcheck` | `--check-remote` | 本地链接/图片路径检查；加参数则顺带检查 HTTP/HTTPS 链接 |

**lint 检查项：**

| 检查项 | 可自动修复 |
|--------|-----------|
| 标题层级跳跃（如 h1 直接到 h3） | ✗ 仅报告 |
| 文件缺少 h1 | ✗ 仅报告 |
| 连续 3 行以上空行 | ✓ 修复 |
| 列表缩进不一致 | ✓ 修复 |
| 代码块未指定语言 | ✗ 仅报告 |

**zhlint 检查项：**

| 检查项 | 可自动修复 |
|--------|-----------|
| 中文与英文/数字之间缺少空格 | ✓ 修复 |
| 中文语境中使用半角标点（`,` `.` `?` `!`） | ✓ 修复 |
| 全角数字或全角字母 | ✓ 修复 |

**linkcheck 检查项：**
- 本地文件链接（`[text](./path.md)`）：检查文件是否存在
- 图片引用（`![alt](./images/img.png)`）：检查图片文件是否存在
- HTTP/HTTPS 链接（仅在 `--check-remote` 时检查）：发送 HEAD 请求，报告 4xx/5xx 状态

linkcheck 不支持 `--fix`（断链需人工判断如何修复）。

**输出格式：** 所有 action 以行号 + 问题描述的方式输出，最后打印问题总数。无问题时输出 `✓ No issues found`。

---

### `convert.py` — 格式转换

**调用格式：**
```bash
uv run python scripts/convert.py <input.md 或 目录/> --to <format> [-o output] [--template template.docx]
```

| 参数 | 说明 |
|------|------|
| `--to docx/html/txt/pdf` | 目标格式（必选） |
| `-o OUTPUT` | 输出路径。输入为目录时，OUTPUT 也应为目录 |
| `--template FILE` | 仅 `--to docx` 时有效，指定 Word 模板文件 |

**格式支持：**

| 格式 | 依赖 | 说明 |
|------|------|------|
| `docx` | `python-docx` | 支持标题、列表、代码块、表格、图片嵌入 |
| `html` | `markdown`、`beautifulsoup4` | 输出带内联 CSS 的自包含 HTML |
| `txt` | 标准库 | 去除所有 Markdown 标记，输出纯文本 |
| `pdf` | `weasyprint` | Markdown → HTML → PDF，纯 Python，无需 Pandoc |

**批量转换：** 输入路径为目录时，递归处理目录下所有 `.md` 文件。

**示例：**
```bash
# 单文件转 DOCX，使用自定义模板
uv run python scripts/convert.py report.md --to docx --template assets/template.docx

# 单文件转 PDF
uv run python scripts/convert.py report.md --to pdf -o output/report.pdf

# 批量转换目录下所有 md 为 HTML
uv run python scripts/convert.py docs/ --to html -o html_output/
```

---

### `files.py` — 文件级操作

**调用格式：**
```bash
uv run python scripts/files.py <action> [inputs...] [options]
```

| action | 参数 | 说明 |
|--------|------|------|
| `split` | `<input.md>` `--by h1/h2/h3` `--output-dir DIR` | 按指定标题级别拆成多个文件 |
| `merge` | `<file1> <file2> ...` `-o output.md` `--separator SEP` | 合并多个 Markdown 文件 |

**split 说明：**
- 拆分后每个文件以对应标题文本作为文件名（自动清理非法字符）
- `--output-dir` 默认为与输入文件同目录的 `split_output/` 子目录
- 保留各子文档中的完整内容，包括子标题

**merge 说明：**
- `--separator` 默认为 `\n\n---\n\n`
- 支持传入目录路径，自动递归找出所有 `.md` 文件按文件名字母序合并

**示例：**
```bash
# 按 h2 拆分，输出到 chapters/ 目录
uv run python scripts/files.py split book.md --by h2 --output-dir ./chapters/

# 合并三个文件
uv run python scripts/files.py merge ch1.md ch2.md ch3.md -o book.md

# 合并整个目录
uv run python scripts/files.py merge docs/ -o merged.md --separator "\n\n***\n\n"
```

---

## SKILL.md 结构

```
---
name: markdown-master
description: 处理、编辑、检查、转换 Markdown 文档。
  用于：调整标题层级、添加章节编号、生成目录、格式检查、
  中文排版检查、链接检查、转换为 DOCX/HTML/PDF、拆分或合并文件。
---

## Quick Start       ← 最常用命令速查（各脚本各一个示例）
## Dependencies      ← pip install 清单
## Scripts
### structure.py     ← 子命令表 + 完整示例
### quality.py
### convert.py
### files.py
```

---

## 依赖清单

```bash
# 核心（structure / quality / files）
# 无第三方依赖，仅使用标准库

# convert.py
pip install python-docx markdown beautifulsoup4 Pillow weasyprint
```

---

## 验证方案

准备 `examples/` 目录：
- `sample_cn.md`：中文文档，含多级标题、中英混排、本地图片引用、中文排版问题
- `sample_en.md`：英文文档，含标题层级跳跃、死链

逐脚本验证：

| 脚本 | 验证命令 |
|------|----------|
| `structure.py` | `heading upgrade` / `numbering add --style chinese_chapter` / `toc generate` |
| `quality.py` | `lint --fix` / `zhlint --fix` / `linkcheck` |
| `convert.py` | `--to docx` / `--to html` / `--to pdf` |
| `files.py` | `split --by h2` / `merge` |

所有命令在 `examples/` 目录的示例文件上运行，确认输出符合预期。
