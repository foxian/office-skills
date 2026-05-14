# Word Editor DSL Enhancement Design

## 1. Overview
本设计旨在补充和增强 Word Master 编辑器的 DSL (`edit-ops.md` 和 `docx_editor.py`)，提供 9 个新的操作指令，覆盖内容、表格、结构和格式四大维度。这是对已有指令（`replace_text`, `rewrite_paragraph`, `set_font`, `set_paragraph_format`）的重要补充，使得 AI 具备全面的文档修改能力。

**通用规则 — 容错与索引越界**：所有操作的辅助函数均遵循"失败安全"原则。若段落索引、表格索引、行/列索引、节索引等超出范围，或必填参数缺失/冲突，则打印 `[WARNING] op=<op_name>: <reason>, skipping.` 并跳过该操作（不抛出异常），以防止单条错误指令中断批量操作。

## 2. 架构选择
采用 **就地扩展** (In-place extension) 策略：
在 `docx_editor.py` 现有的 `apply_operations` 的 `if/elif` 路由链中，直接追加新指令的分支。每个新指令对应一个独立的私有辅助函数（如 `_apply_insert_paragraph`）。

**理由**：现有架构简单直接，增加 9 个操作函数不会导致代码急剧膨胀，且能保持与之前指令一致的错误处理和执行模式（"手术刀原则"）。

## 3. DSL 参数规范设计

所有操作指令均支持批量传入。操作对象定位统一风格：
- **段落**：沿用 `"pN"`（如 `"p5"`）的索引形式。
- **表格**：采用 `table: N`（整数索引，从0开始）定位。

### 3.1 内容操作

#### `insert_paragraph`（插入段落）
```json
{
  "op": "insert_paragraph",
  "after": "p5",
  "content": "新段落文字",
  "style": "Normal"
}
```
*   **`after`** / **`before`** (必填一): 定位锚点。`before: "p0"` 表示在文档首段前插入。互斥参数，若都不提供或都提供则打印警告并跳过。
*   **`content`** (必填): 新段落文字内容。
*   **`style`** (可选): 段落样式名。省略时默认继承**锚点段落**的样式（`after` 模式继承锚点段落样式，`before` 模式继承锚点段落样式；若锚点是 `p0` 且使用 `before` 无前段，则使用 `Normal`）。

#### `delete_paragraph`（删除段落）
```json
{"op": "delete_paragraph", "target": "p5"}
```
*   **`target`** (必填): 目标段落 ID。

### 3.2 表格操作

#### `modify_cell`（修改单元格）
```json
{"op": "modify_cell", "table": 0, "row": 1, "col": 2, "content": "新内容"}
```
*   **`table`** (必填): 表格索引 (0 开始)。
*   **`row`** / **`col`** (必填): 行列索引 (0 开始)。
*   **`content`** (必填): 单元格新文本内容。

#### `insert_row`（插入表格行）
```json
{"op": "insert_row", "table": 0, "after": 2}
```
*   **`table`** (必填): 表格索引。
*   **`after`** / **`before`** (必填一): 行号锚点。在某行之前或之后插入空行。互斥，若都不提供或都提供则跳过。

#### `insert_table`（插入新表格）
```json
{"op": "insert_table", "after": "p3", "rows": 3, "cols": 4}
```
*   **`after`** / **`before`** (必填一): 段落锚点，支持在段落前或后插入表格，与 `insert_paragraph` 保持对称。互斥，若都不提供或都提供则跳过。
*   **`rows`** / **`cols`** (必填): 初始行列数。

### 3.3 结构操作

#### `set_header_footer`（配置页眉页脚）
```json
{
  "op": "set_header_footer",
  "header": "页眉文字",
  "footer": "页脚文字",
  "section": 0,
  "even_page": false
}
```
*   **`header`** / **`footer`** (可选): 对应内容，省略则不修改。
*   **`section`** (可选): 节索引，默认 `0`（第一节）。
*   **`even_page`** (可选): `true` 时修改偶数页页眉/页脚，默认 `false`。

#### `insert_page_break`（插入分页符）
```json
{"op": "insert_page_break", "after": "p10"}
```
*   **`after`** / **`before`** (必填一): 段落锚点，支持在段落前或后插入分页符，与其他插入类操作保持对称。互斥，若都不提供或都提供则跳过。

### 3.4 格式操作

#### `apply_style`（强制应用样式）
```json
{"op": "apply_style", "target": "p2", "style": "Heading 1"}
```
*   **`target`** (必填): 目标段落 ID。
*   **`style`** (必填): 样式名。若文档中无此样式则打印警告并跳过（不抛异常）。

#### `set_page_setup`（设置页面格式）
```json
{
  "op": "set_page_setup",
  "margin_top": "2.54cm",
  "margin_bottom": "2.54cm",
  "margin_left": "3.17cm",
  "margin_right": "3.17cm",
  "orientation": "portrait",
  "section": 0
}
```
*   **`margin_top`** / **`margin_bottom`** / **`margin_left`** / **`margin_right`** (可选): 四个方向的边距，均支持 `"Xcm"` 或 `"Xpt"` 两种单位（如 `"2.54cm"` 或 `"72pt"`），省略则不修改对应边距。
*   **`orientation`** (可选): `"portrait"`（纵向）或 `"landscape"`（横向）。
*   **`section`** (可选): 节索引，默认 `0`。

## 4. 实施策略

所有写入操作遵循"失败安全"（与第 1 节通用规则对应）：
1. 辅助函数负责各自校验，遇到无效索引或互斥参数冲突，直接 `return` 跳过并输出 `[WARNING] op=<op_name>: <reason>, skipping.`。
2. 现有机制已包含 `_backup_file`（写入前备份），继续保持。

## 5. 测试要求
需要为 `docx_editor.py` 增加新指令的单元测试，确保：
- 新段落的 `before` / `after` 插入正确，且段落总数符合预期。
- `before: "p0"` 时样式回退到 `Normal`。
- 表格和单元格寻址越界时不崩溃，仅输出警告。
- 单位解析 (`cm` / `pt`) 换算结果符合预期（如 `2.54cm` ≈ `72pt`）。
- `insert_table` / `insert_page_break` 的 `before` 参数能正确插入。
