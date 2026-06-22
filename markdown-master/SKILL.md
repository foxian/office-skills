---
name: markdown-master
description: |
  Markdown 文档全能处理工具。当用户需要对 Markdown 文档进行以下操作时，务必使用此技能：
  - 调整标题层级（升级/降级）、添加章节编号（支持灵活配置每级格式）、生成/更新目录
  - 格式规范检查（lint）：检查标题层级跳跃、缺少 H1、连续空行过多、代码块未指定语言等，支持自动修复
  - 中文排版检查（zhlint）：中文与英文/数字间缺空格、中文语境使用半角标点等，支持自动修复
  - 链接/图片死链检查（支持检查本地文件和远程 HTTP 链接）
  - 格式转换：Markdown 转 DOCX/HTML/TXT/PDF（支持批量转换目录下所有文件）
  - 文件操作：按标题级别拆分文档、合并多个 Markdown 文件
  即使没有明确提到"markdown-master"，只要是 Markdown 相关的结构化编辑、质检、转换需求，都应该使用此技能。
---

# Markdown Master

## 使用前准备

首先检查并安装依赖：

```bash
# 检查 Python 是否可用
python --version

# 安装可选依赖（convert.py 需要）
pip install python-docx markdown beautifulsoup4 weasyprint pyyaml
```

注意：`structure.py` / `quality.py` / `files.py` 仅使用 Python 标准库，无需额外安装即可使用。只有 `convert.py` 需要上述依赖。

## 核心工作流程

### 场景 1：准备正式发布的中文文档

当用户需要把一份 Markdown 文档整理成正式版本时：

1. **先做格式检查**：`python scripts/quality.py doc.md lint --fix`
2. **再做中文排版优化**：`python scripts/quality.py doc.md zhlint --fix`
3. **添加章节编号**：`python scripts/structure.py doc.md numbering add --h1 "第{1}章 " --h2 "{1}.{2} " --h3 "（{3}）"`
4. **生成目录**：`python scripts/structure.py doc.md toc generate`
5. **检查链接**：`python scripts/quality.py doc.md linkcheck`
6. **导出为 Word/PDF**：`python scripts/convert.py doc.md --to docx`

### 场景 2：合并多篇文档

当用户有多份 Markdown 文件需要合并成一份时：

1. **合并文件**：`python scripts/files.py merge file1.md file2.md -o combined.md`
2. **统一调整标题层级**（如果需要）：`python scripts/structure.py combined.md heading upgrade --levels 1`
3. **后续同场景 1**（lint → zhlint → 编号 → 目录 → 导出）

### 场景 3：拆分长文档

当用户有一份很长的文档需要按章节拆分成多个文件时：

1. **拆分**：`python scripts/files.py split doc.md --by h2`
2. **检查拆分结果**

## 命令速查

### structure.py — 结构操作

| 用法 | 说明 |
|------|------|
| `python scripts/structure.py <file> heading upgrade [--levels N]` | 所有标题升 N 级（默认 1）|
| `python scripts/structure.py <file> heading downgrade [--levels N]` | 所有标题降 N 级 |
| `python scripts/structure.py <file> numbering add [--h1 T] [--h2 T] ... [--config FILE] [--start-from N] [--save-config FILE]` | 添加灵活编号，每级可独立配置格式，支持从 YAML 加载配置或保存当前配置 |
| `python scripts/structure.py <file> numbering remove` | 移除所有编号 |
| `python scripts/structure.py <file> toc generate [--depth N] [--position top\|after-h1]` | 插入/更新目录 |

所有操作默认覆盖原文件，使用 `-o output.md` 输出到新文件。

### 编号模板语法

| 占位符 | 含义 | 示例 |
|---------|------|------|
| `{N}` / `{N:d}` | 十进制数字 | 3 |
| `{N:02d}` | 两位补零十进制数字 | 03 |
| `{N:R}` | 大写罗马数字 | III |
| `{N:r}` | 小写罗马数字 | iii |
| `{N:A}` | 大写字母 | C |
| `{N:a}` | 小写字母 | c |
| `{N:cn}` | 中文数字 | 三 |

N 范围是 1-6，对应标题级别 H1-H6。空模板字符串表示该级标题不添加编号。

### YAML 配置文件格式

```yaml
h1: "第{1}章 "
h2: "{1}.{2} "
h3: "（{3}）"
h4: ""
h5: ""
h6: ""
start_from: 1
```

### 常见编号用法示例

```bash
# 技术文档风格（1/1.1/1.1.1）
python scripts/structure.py doc.md numbering add --h1 "{1} " --h2 "{1}.{2} " --h3 "{1}.{2}.{3} "

# 中文章节风格（第1章/1.1/（1））
python scripts/structure.py doc.md numbering add --h1 "第{1}章 " --h2 "{1}.{2} " --h3 "（{3}）"

# 学术风格（I/A/1/a）
python scripts/structure.py doc.md numbering add --h1 "{1:R} " --h2 "{2:A} " --h3 "{3} " --h4 "{4:a} "

# 保存当前配置为模板
python scripts/structure.py doc.md numbering add --h1 "第{1}章 " --h2 "{1}.{2} " --save-config chinese_chapter.yaml

# 使用保存的配置
python scripts/structure.py doc.md numbering add --config chinese_chapter.yaml
```

### quality.py — 质检操作

| 用法 | 说明 |
|------|------|
| `python scripts/quality.py <file> lint [--fix]` | 格式规范检查，`--fix` 自动修复可修复项 |
| `python scripts/quality.py <file> zhlint [--fix]` | 中文排版检查，`--fix` 自动修复 |
| `python scripts/quality.py <file> linkcheck [--check-remote]` | 本地链接/图片检查，加参数检查 HTTP 链接 |

输出格式：`[L行号] 问题描述`，最后汇总问题数。无问题输出 `✓ No issues found`。

**lint 检查项：**
- ✗ 标题层级跳跃（H1 -> H3）
- ✗ 文件缺少 H1
- ✓ 连续 3 行以上空行
- ✗ 代码块未指定语言

**zhlint 检查项：**
- ✓ 中文与英文/数字间缺空格
- ✓ 中文语境中使用半角标点

### convert.py — 格式转换

```bash
python scripts/convert.py <input.md 或 目录/ --to <format> [-o output] [--template template.docx]
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
python scripts/files.py split <input.md> --by h1|h2|h3 [--output-dir DIR]

# 合并（支持文件列表或目录）
python scripts/files.py merge <file1> [file2 ...] -o output.md [--separator SEP]
```

split 默认输出到 `<input_dir>/split_output/`，文件名为 `XX_<标题文本>.md`。
merge 输入目录时，按文件名字母序合并所有 `.md` 文件。
