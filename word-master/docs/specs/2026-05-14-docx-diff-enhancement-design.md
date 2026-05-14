# DOCX Diff Enhancement Design

## 1. 概述

**目标**：将 `docx_diff.py` 从"纯文本比对"升级为"富文本语义比对"，复用 Priority 1 中增强的 `docx_reader.py`，实现格式与内容双重验证能力。

**核心价值**：确保哪怕只是字号改变、加粗等微小格式调整，都能在 Diff 中显现，实现"所见即所得"的格式校验。

## 2. 技术架构

### 2.1 实现机制

```
docx_diff.py
    ├── 调用 docx_reader.py (两份文档各自生成带格式的 Markdown)
    ├── 对两份 Markdown 进行结构化比对
    └── 输出详细 Markdown diff
```

### 2.2 复用组件

- **docx_reader.py**：`extract_rich_markdown(filepath)` 输出带格式属性的 Markdown
  - STYLES 区块：文档使用的样式定义
  - 段落格式：`{StyleName}` 或 `{StyleName | override: bold on "text"}`
  - Run 级覆盖：局部格式变化标注

## 3. 输出格式

### 3.1 整体结构

```markdown
# DOCX Diff Report
<!-- ORIGINAL: original.docx -->
<!-- MODIFIED: modified.docx -->

## Styles Changes
[仅当两份文档的 STYLES 区块存在差异时显示]

## Paragraphs

### [p0] {Heading 1}
[段落级分组]

#### Fields
| Field | Original | Modified |
|-------|----------|----------|
| font  | 宋体     | 楷体     |
| size  | 12pt     | 14pt     |
| bold  | false    | true     |

### [新增] [p5] {Normal}
这是新增段落的内容，包含完整的段落文本和格式信息。

### [删除] [p3] {Normal}
~~这是被删除的段落内容。~~

### [p2] {Normal}
#### Fields
| Field | Original | Modified |
|-------|----------|----------|
| align | LEFT     | CENTER   |
```

### 3.2 变更类型处理

| 变更类型 | 标记方式 |
|----------|----------|
| 内容变化 | 段落级变更，Fields 表格列出具体字段差异 |
| 格式变化 | 段落级变更，Fields 表格列出样式/override 差异 |
| 新增段落 | `[新增] [pN]` + 完整内容 |
| 删除段落 | `[删除] [pN]` + ~~strikethrough~~ 格式 |
| 段落位置变化 | 在变更段落中标注 "moved from pN" / "moved to pN" |

### 3.3 混合粒度模式

- **段落级分组**：以 `[pN]` 为单位组织变更，便于快速定位
- **字段级差异**：每个变更段落下通过 Fields 表格展示具体字段变化
- **上下文保留**：未变更的相邻段落可选择性保留 1-2 段作为上下文

## 4. 命令行接口

### 4.1 基本用法

```bash
# 标准比对
python docx_diff.py original.docx modified.docx

# 输出到文件
python docx_diff.py original.docx modified.docx --output diff_report.md

# 同时输出到 stdout 和文件
python docx_diff.py original.docx modified.docx --output diff_report.md --verbose
```

### 4.2 参数说明

| 参数 | 说明 |
|------|------|
| `path1` | 修改前的 DOCX 文件路径 |
| `path2` | 修改后的 DOCX 文件路径 |
| `--output`, `-o` | 可选，输出文件路径（默认输出到 stdout） |
| `--verbose`, `-v` | 可选，同时输出到 stdout 和文件 |
| `--context`, `-c` | 可选，保留上下文段落数量（默认 1） |

### 4.3 错误处理

- **文件不存在**：输出错误信息到 stderr，exit(1)
- **文件损坏/非 DOCX**：输出错误信息到 stderr，exit(1)
- **解析失败**：输出错误信息到 stderr，exit(1)

```bash
# 错误示例
$ python docx_diff.py not_exist.docx modified.docx
Error: File not found: not_exist.docx
$ echo $?
1
```

## 5. 实现细节

### 5.1 核心函数

```python
def generate_rich_diff(path1: str, path2: str, context: int = 1) -> str:
    """
    生成富文本语义 diff。

    :param path1: 原文档路径
    :param path2: 新文档路径
    :param context: 保留上下文段落数量
    :return: Markdown 格式的 diff 报告
    """
    # 1. 解析两份文档为带格式的 Markdown
    md1 = extract_rich_markdown(path1)
    md2 = extract_rich_markdown(path2)

    # 2. 解析 Markdown 为结构化段落
    blocks1 = parse_markdown_blocks(md1)
    blocks2 = parse_markdown_blocks(md2)

    # 3. 对齐段落（基于索引和内容相似度）
    aligned = align_paragraphs(blocks1, blocks2)

    # 4. 生成字段级差异
    diff_blocks = []
    for orig, modif in aligned:
        if is_same(orig, modif):
            continue
        diff_blocks.append(compute_field_diff(orig, modif))

    # 5. 输出为 Markdown
    return format_diff_report(diff_blocks, context=context)
```

### 5.2 段落对齐算法

- **精确匹配**：索引相同且内容相同 → 无变化
- **内容匹配**：索引不同但内容相似度 > 阈值 → 移动/修改
- **模糊匹配**：使用 difflib 的 SequenceMatcher 找最佳匹配

### 5.3 字段差异计算

```python
def compute_field_diff(block1, block2):
    """计算两个段落块的字段级差异"""
    diff = {
        'type': determine_change_type(block1, block2),  # 'content', 'format', 'new', 'deleted'
        'original': block1,
        'modified': block2,
        'fields': []
    }

    # 比对样式
    if block1.style != block2.style:
        diff['fields'].append({
            'name': 'style',
            'orig': block1.style,
            'modif': block2.style
        })

    # 比对 override 信息
    if block1.overrides != block2.overrides:
        diff['fields'].append({
            'name': 'overrides',
            'orig': block1.overrides,
            'modif': block2.overrides
        })

    # 比对内容
    if block1.text != block2.text:
        diff['fields'].append({
            'name': 'content',
            'orig': block1.text,
            'modif': block2.text
        })

    return diff
```

## 6. 测试策略

### 6.1 单元测试

| 测试用例 | 说明 |
|----------|------|
| `test_diff_content_change` | 内容变化检测 |
| `test_diff_format_change` | 格式变化检测（字号、加粗、字体） |
| `test_diff_new_paragraph` | 新增段落检测 |
| `test_diff_deleted_paragraph` | 删除段落检测 |
| `test_diff_moved_paragraph` | 段落移动检测 |
| `test_diff_style_override` | Run 级 override 检测 |
| `test_diff_missing_file` | 文件不存在错误处理 |
| `test_diff_corrupt_file` | 文件损坏错误处理 |

### 6.2 集成测试

- 使用 `tests/test_input.docx` 和 `tests/test_output.docx` 进行端到端测试
- 验证 diff 输出包含预期的所有变更标记

## 7. 兼容性

- **向后兼容**：`generate_diff()` 函数保留，但内部实现调用 `generate_rich_diff()`
- **接口兼容**：CLI 参数保持一致，新增 `--output` 和 `--verbose` 选项
