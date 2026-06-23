# Word Master 表格读取增强设计

## 背景

`docx_reader.py` 目前只遍历 `doc.paragraphs`，完全跳过 Word 文档中的表格内容。这意味着 AI 在读取含有表格的文档时，会丢失表格中的所有信息，无法对其进行理解和编辑。本设计旨在补全这一功能缺口，并将表格读取能力延伸到下游工具链。

## 目标

1. `docx_reader.py` 按文档真实阅读顺序输出段落与表格交织内容
2. `docx_editor.py` 支持通过单元格 ID 精确替换/重写表格内容
3. `docx_diff.py` 支持对比两个文档的表格内容差异
4. `SKILL.md` Workflow 1 更新说明，指导 AI 使用表格 ID 进行编辑

---

## 一、docx_reader.py 核心重构

### 1.1 遍历方式变更

将现有的：

```python
for i, para in enumerate(doc.paragraphs):
```

改为遍历 `doc.element.body` 的子元素，按文档顺序区分段落（`<w:p>`）和表格（`<w:tbl>`）：

```python
from docx.oxml.ns import qn

para_idx = 0
table_idx = 0
for child in doc.element.body:
    tag = child.tag.split('}')[-1]
    if tag == 'p':
        # 处理段落
        ...
        para_idx += 1
    elif tag == 'tbl':
        # 处理表格
        ...
        table_idx += 1
```

### 1.2 ID 命名规范

| 元素 | ID 格式 | 示例 |
|------|---------|------|
| 段落 | `p{i}` | `p0`, `p3` |
| 表格单元格 | `t{table_idx}r{row_idx}c{col_idx}` | `t0r1c2` |

### 1.3 表格 Markdown 渲染格式

表格以标准 Markdown 格式渲染，含单元格 ID 前缀和格式覆盖标注：

```
<!-- TABLE t0: 3 cols × 5 rows -->
| [t0r0c0] **序号** | [t0r0c1] **设备名称** | [t0r0c2] **金额** |
|---|---|---|
| [t0r1c0] 1 | [t0r1c1] 防火墙 | [t0r1c2] 10万 |
| [t0r2c0] 2 | [t0r2c1] 入侵检测 {bold=True on "入侵检测"} | [t0r2c2] 8万 |
```

规则：
- 表格前输出 `<!-- TABLE t{n}: {cols} cols × {rows} rows -->` 注释行
- 首行（第0行）视为表头，内容加粗（`**...**`）
- 每个单元格内容以 `[t{n}r{n}c{n}]` 为前缀
- 若单元格内有格式覆盖（粗体、斜体、字号等），以 `{...}` 形式内联标注，与现有段落 override 格式一致
- 多段落单元格：各段落文本以 ` / ` 连接

### 1.4 新增函数

```python
def _extract_cell_overrides(cell) -> list[str]:
    """提取单元格内所有段落的格式覆盖信息"""

def _extract_table_markdown(table, table_idx: int, overview: bool = False) -> list[str]:
    """将单个 table 对象渲染为带 ID 标注的 Markdown 行列表"""

def extract_rich_markdown(filepath, overview=False) -> str:
    """
    读取 DOCX 文档，按文档阅读顺序交织输出段落和表格的 Markdown 表示。
    --overview 模式下仅输出 Heading 标题，表格不参与大纲。
    """
```

### 1.5 overview 模式行为

`--overview` 模式下表格**不输出**，行为与现有一致（只展示 Heading 层级大纲）。

---

## 二、docx_editor.py — 新增单元格编辑操作

### 2.1 新增 DSL 操作

| op | 必填参数 | 说明 |
|----|---------|------|
| `replace_cell_text` | `target`, `find`, `replace` | 在指定单元格内执行文本替换 |
| `rewrite_cell` | `target`, `content` | 整体重写单元格纯文本内容 |

`target` 格式：`"t{n}r{n}c{n}"`，如 `"t0r1c2"`。

### 2.2 target 解析逻辑

```python
import re

def _parse_cell_target(target: str):
    """解析 'tNrNcN' 格式，返回 (table_idx, row_idx, col_idx)，失败抛 ValueError"""
    m = re.fullmatch(r't(\d+)r(\d+)c(\d+)', target)
    if not m:
        raise ValueError(f"Invalid cell target: {target!r}. Expected format: tNrNcN")
    return int(m.group(1)), int(m.group(2)), int(m.group(3))
```

### 2.3 错误处理

- target 格式非法 → 抛出 `ValueError`，提示正确格式
- 表格/行/列索引越界 → 抛出 `IndexError`，提示实际范围

---

## 三、docx_diff.py — 表格内容差异对比

### 3.1 对比策略

- 按表格索引（`t0`, `t1`, ...）逐一对齐两份文档的表格
- 对比粒度：单元格文本内容（不对比格式）
- 两文档表格数量不一致时，多余的表格报告为"新增"或"删除"

### 3.2 输出格式

在现有 diff 报告末尾新增 `## Tables` 章节：

```markdown
## Tables

### t0 (3×5)
- [t0r1c1] "旧内容" → "新内容"
- [t0r2c0] "原文" → "修改后"

### t1 (新增, 2×3)
+ 整个表格为新增
```

---

## 四、SKILL.md 更新

在 Workflow 1 的 Step 1 说明中补充：

> 输出中 `[t0r1c2]` 格式标识第 0 个表格、第 1 行、第 2 列单元格。

在 Step 2 的示例操作中补充表格编辑示例：

```json
[
  {"op": "replace_cell_text", "target": "t0r1c1", "find": "旧值", "replace": "新值"},
  {"op": "rewrite_cell", "target": "t0r2c0", "content": "全新内容"}
]
```

---

## 五、测试计划

### 5.1 test_docx_reader.py

- [ ] 表格内容出现在输出中
- [ ] 单元格 ID 格式正确（`t0r0c0` 等）
- [ ] 表格与段落按文档顺序交织排列（通过包含表格的测试 DOCX 验证）
- [ ] `--overview` 模式下表格不出现
- [ ] 多段落单元格以 ` / ` 连接

### 5.2 test_docx_editor_table.py

- [ ] `replace_cell_text` 精确替换单元格内文本
- [ ] `rewrite_cell` 整体重写单元格内容
- [ ] 非法 target 格式抛出 `ValueError`
- [ ] 越界索引抛出 `IndexError`

### 5.3 test_docx_diff.py

- [ ] 表格单元格变化被捕获并出现在 `## Tables` 章节
- [ ] 两文档表格数量不一致时正确报告

### 5.4 集成验证

使用 `examples/投标书_技术部分.docx`（含大量表格）运行 `docx_reader.py`，人工核查渲染输出。

---

## 六、影响范围与兼容性

| 项目 | 影响 |
|------|------|
| 现有段落 ID（`p{i}`）| ✅ 不变，索引逻辑保持一致 |
| `--overview` 模式 | ✅ 行为不变 |
| STYLES 注释块 | ✅ 不变 |
| `docx_editor.py` 现有操作 | ✅ 向后兼容，新操作为增量添加 |
| `docx_diff.py` 现有输出 | ✅ 新增 Tables 章节，不影响现有内容 |
