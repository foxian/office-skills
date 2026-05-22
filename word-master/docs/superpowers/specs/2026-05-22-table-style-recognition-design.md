# 表格样式识别与迁移功能设计

## 背景与目标

Word Master 现有的样式迁移功能（Workflow 3）只支持**段落样式**（字体、字号、对齐、行距等）的识别与转换，不支持**表格样式**。

本功能扩展目标：从模板 Word 文件中识别表格的完整视觉样式（边框、底纹、单元格内边距、文字格式），形成"表格样式库"，在需要转换的目标文档中自动匹配并应用。

---

## 功能范围

### 识别内容
- **Word 内置表格样式名称**（`<w:tblStyle>`，如 `Table Grid`、`Light Shading`）
- **边框属性**：六方向（top/bottom/left/right/insideH/insideV）的线型、粗细、颜色
- **底纹**：表头行底纹色、正文行底纹色
- **单元格内边距**：四方向（top/bottom/left/right）
- **文字格式**：表头行（字体、字号、加粗、对齐）+ 正文行（字体、字号、加粗、对齐）

### 迁移策略
- 优先引用 `tbl_style` 命名样式；若目标文档无该样式则直接写入视觉属性
- 所有目标表格都会被处理（无跳过阈值），多个样式库条目时取相似度最高者
- 文字格式全面覆盖（表头行 + 正文行均迁移）

---

## 架构设计

### 文件职责

| 文件 | 职责 | 变更类型 |
|------|------|----------|
| `style_analyzer.py` | 段落指纹提取（函数库） | **不动** |
| `table_analyzer.py` | 表格指纹提取（函数库） | ✅ 新建 |
| `template_analyzer.py` | 协调段落 + 表格分析，输出统一 profile | ✅ 新建 |
| `style_transfer.py` | 读 profile，生成段落 + 表格 ops | ✅ 扩展 |
| `docx_editor.py` | 执行 `apply_table_style` op | ✅ 扩展 |
| `validate_style_profile.py` | 验证 `table_roles` 字段格式 | ✅ 扩展 |
| `SKILL.md` | 更新 Workflow 3 Step 1 入口命令 | ✅ 更新 |

### 架构图

```
                    ┌─────────────────────────┐
                    │   template_analyzer.py   │  ← 新 CLI 入口
                    └────────────┬────────────┘
                                 │
              ┌──────────────────┴──────────────────┐
              │                                     │
    ┌─────────▼─────────┐               ┌──────────▼──────────┐
    │  style_analyzer.py │               │  table_analyzer.py  │
    │  （段落指纹提取）   │               │  （表格指纹提取）    │
    └───────────────────┘               └─────────────────────┘
              │                                     │
              └──────────────────┬──────────────────┘
                                 │ 合并
                    ┌────────────▼────────────┐
                    │    style_profile.json    │
                    │  roles + table_roles     │
                    └─────────────────────────┘
```

---

## 数据结构：`style_profile.json`

`roles` 字段保持不变（向后兼容），新增 `table_roles` 字段。

```json
{
  "roles": [
    {
      "role": "Heading 1",
      "fingerprint": {
        "size": "22.0pt", "bold": true, "italic": false,
        "align": "center", "font": "方正小标宋简体",
        "color": null, "space_before": "0.0pt",
        "space_after": "12.0pt", "first_line_indent": "0.0pt",
        "line_spacing": null
      }
    }
  ],
  "table_roles": [
    {
      "id": 0,
      "tbl_style": "Table Grid",
      "structure": { "cols": 3, "rows": 5, "has_header_row": true },
      "border": {
        "top":     { "val": "single", "sz": 4, "color": "000000" },
        "bottom":  { "val": "single", "sz": 4, "color": "000000" },
        "left":    { "val": "single", "sz": 4, "color": "000000" },
        "right":   { "val": "single", "sz": 4, "color": "000000" },
        "insideH": { "val": "single", "sz": 2, "color": "AAAAAA" },
        "insideV": { "val": "none",   "sz": 0, "color": "auto"   }
      },
      "shading": {
        "header": "D9E1F2",
        "body": null
      },
      "cell_margin": {
        "top": "0.0pt", "bottom": "0.0pt",
        "left": "5.4pt", "right": "5.4pt"
      },
      "header_text": {
        "font": "黑体", "size": "12.0pt",
        "bold": true, "align": "center"
      },
      "body_text": {
        "font": "宋体", "size": "10.5pt",
        "bold": false, "align": "left"
      },
      "example": "列一 | 列二 | 列三"
    }
  ]
}
```

### `table_roles` 字段说明

| 字段 | 类型 | 说明 |
|------|------|------|
| `id` | int | 样式库条目编号 |
| `tbl_style` | string \| null | Word 内置样式名，null 表示无命名样式 |
| `structure.cols` | int | 列数，通过 `<w:gridCol>` XML 节点数量计算（兼容合并单元格） |
| `structure.rows` | int | 行数（用于相似度匹配） |
| `structure.has_header_row` | bool | 是否有表头行 |
| `border.<dir>.val` | string | OOXML 线型（`single`/`none`/`dashed` 等） |
| `border.<dir>.sz` | int | 线宽（1/8 pt 单位，如 4 = 0.5pt） |
| `border.<dir>.color` | string | 十六进制颜色或 `"auto"` |
| `shading.header` | string \| null | 表头行底纹十六进制色（`w:val="clear"` 且 `w:fill` 不为 `auto`/`FFFFFF` 时有效），null 为无底纹 |
| `shading.body` | string \| null | 正文行底纹十六进制色，null 为无底纹 |
| `cell_margin.<dir>` | string | 内边距，`"Xpt"` 格式 |
| `header_text` | object \| null | 首行文字格式（无表头行时为 null） |
| `body_text` | object | 正文行文字格式 |
| `example` | string | 代表性表格首行文字摘要 |

---

## 各脚本核心逻辑

### `table_analyzer.py`（新建）

```
extract_table_fingerprints(filepath) → list[dict]

对每个表格：
  1. 读取 tbl_style：优先用 table.style.name API；若为 None 则回退读取
     <w:tblStyle w:val="..."> XML 节点
  2. 计算列数：通过 tbl._tbl.findall(qn('w:tblGrid/w:gridCol')) 数量，
     兼容合并单元格（不用 len(table.columns)）
  3. 计算行数：len(table.rows)
  4. 判断 has_header_row（三级优先级）：
     a. 首行存在 <w:tblHeader> XML 节点 → True（最可靠）
     b. 首行所有单元格 <w:shd w:val="clear" w:fill≠auto,FFFFFF> → True
     c. 首行所有单元格段落均加粗 → True
     否则 → False
  5. 读取六方向边框（<w:tblBorders>），每方向读取 val/sz/color 属性
  6. 读取底纹色（<w:shd>）：
     - 仅当 w:val="clear" 且 w:fill 不为 "auto" 或 "FFFFFF" 时，
       取 w:fill 作为有效底纹色，否则记为 null（无底纹）
     - 表头行底纹：首行第一个单元格的 <w:shd>
     - 正文行底纹：第二行（若存在）第一个单元格的 <w:shd>
  7. 读取单元格内边距（<w:tblCellMar>），取 top/bottom/left/right
  8. 提取首行文字格式（font/size/bold/align）
  9. 提取其余行文字格式
  10. 计算视觉指纹哈希 → 去重合并

合并单元格处理：
  - 底纹应用时通过比较 cell._tc 对象标识跳过重复的合并单元格

去重键（纯视觉属性，不含列数/行数）：
  (tbl_style, border_top.val, border_top.sz, border_top.color,
   insideH.val, insideH.sz, shading.header, shading.body,
   header_text.font, header_text.size, header_text.bold)

返回去重后列表，典型 1～4 条
```

### `template_analyzer.py`（新建 CLI 入口）

```bash
# 用法
python template_analyzer.py template.docx --output style_profile.json
python template_analyzer.py template.docx --output style_profile.json --para-no-heading-aware
python template_analyzer.py template.docx --output style_profile.json --para-min-cluster-size 2
```

参数命名规范：
- 段落相关参数统一加 `--para-` 前缀，透传给 `style_analyzer`
- 表格相关参数统一加 `--table-` 前缀，透传给 `table_analyzer`（未来扩展）

参数透传方式：**函数参数调用**（非 subprocess），映射关系：
- `--para-no-heading-aware` → `extract_fingerprints(..., heading_aware=False)`
- `--para-min-cluster-size N` → `extract_fingerprints(..., min_cluster_size=N)`

内部流程：
1. 调用 `style_analyzer.extract_fingerprints(heading_aware, min_cluster_size)` → 原始段落指纹
2. **字段归一化**：将输出中的 `heading_role` 字段重命名为 `role`；将无 `role` 字段的条目（body 聚类，含 `id` 字段）保留原样供 AI 推断角色名
3. 调用 `table_analyzer.extract_table_fingerprints()` → 表格样式库（已去重）
4. 合并输出 `style_profile.json`：`{"roles": [...], "table_roles": [...]}`

### `style_transfer.py`（扩展）

扩展位置：在 `generate_apply_ops()` 函数的 `return ops` 之前追加表格 ops（即段落 ops 生成完毕后）：

```
表格匹配逻辑：
  对目标文档每个表格 t_i：
    计算列数：通过 <w:gridCol> 数量（与 table_analyzer 一致）
    计算行数：len(table.rows)
    计算与 table_roles 每条的相似度：
      cols 相同         → +0.5
      has_header_row 一致 → +0.3
      行数差值 ≤ 2       → +0.2（依赖 structure.rows 字段）
    取最高分的 table_role（无截断阈值，总取最近的一条）
    生成：{"op": "apply_table_style", "target": "t{i}", "table_role": {...}}

  若 table_roles 为空或不存在 → 跳过所有表格，输出 [INFO] 日志
```

### `docx_editor.py`（扩展）

新增操作处理器 `_apply_table_style(doc, params)`：

```
1. 定位：table = doc.tables[t_idx]

2. 应用 tbl_style（若不为 null）：
   - 检查目标文档是否存在该样式且类型为表格样式：
     doc.styles[style_name].type == WD_STYLE_TYPE.TABLE
     （需导入 from docx.enum.style import WD_STYLE_TYPE）
   - 存在则：table.style = style_name
   - 不存在则跳过命名样式，继续用视觉属性覆盖

3. 写入六方向边框（先清除旧节点再新建）：
   tblPr = table._tbl.find(qn('w:tblPr'))  # 或 get_or_add_tblPr()
   旧 tblBorders = tblPr.find(qn('w:tblBorders'))
   若存在则 tblPr.remove(旧节点)
   新建 tblBorders，逐方向追加子节点：
     for dir in [top, left, bottom, right, insideH, insideV]:
       elem = OxmlElement(f'w:{dir}')
       elem.set(qn('w:val'), border[dir]['val'])
       elem.set(qn('w:sz'),  str(border[dir]['sz']))
       elem.set(qn('w:color'), border[dir]['color'])
       tblBorders.append(elem)
   tblPr.append(tblBorders)

4. 逐行应用底纹：
   - 维护已处理 _tc 集合，跳过合并单元格的重复引用
   - 首行所有单元格 → header shading（w:shd w:val="clear" w:fill=...）
   - 其余行所有单元格 → body shading（null 时写 w:fill="auto" w:val="clear"）

5. 写入单元格内边距（<w:tblCellMar>，方式同边框：先删旧再新建）

6. 逐单元格逐段落应用文字格式（复用现有 _apply_set_font 逻辑）：
   - 首行单元格段落 → header_text 格式
   - 其余行单元格段落 → body_text 格式
   - 跳过合并单元格重复引用（通过 _tc 集合）
```

### `validate_style_profile.py`（扩展）

新增 `table_roles` 校验（可选字段，缺失不报错）：
- 存在时必须为列表
- 每条需有 `id`（int）、`structure`（dict with cols/has_header_row）
- `border` 各方向若存在：`val`/`sz`/`color` 类型检查
- `shading.header`/`body`：hex 字符串或 null
- `header_text`/`body_text`：若存在，font/size/bold/align 类型检查

---

## CLI 参数规范

### `template_analyzer.py` 完整参数

| 参数 | 默认值 | 说明 |
|------|--------|------|
| `filepath` | 必填 | 源模板 docx 路径 |
| `--output` | `style_profile.json` | 输出文件路径 |
| `--para-heading-aware` | True | 按大纲级别识别标题段落 |
| `--para-no-heading-aware` | — | 禁用标题感知，按格式聚类 |
| `--para-min-cluster-size N` | 4 | 段落聚类最小成员数 |

> `style_analyzer.py` 原有 CLI（`--heading-aware`/`--min-cluster-size`）保留不动，向后兼容。

---

## SKILL.md 变更

Workflow 3 Step 1 更新：

```diff
- python ${SKILL_DIR}/scripts/style_analyzer.py template.docx --output fingerprints.json
+ python ${SKILL_DIR}/scripts/template_analyzer.py template.docx --output style_profile.json
```

Core Scripts 表格新增一行：

```diff
+ | `${SKILL_DIR}/scripts/template_analyzer.py` | 从模板提取完整格式 profile（段落 + 表格） |
+ | `${SKILL_DIR}/scripts/table_analyzer.py`    | 提取表格样式指纹（函数库）                |
```

---

## 验证计划

### 自动化测试

新增四个测试文件：

**`tests/test_table_analyzer.py`**
- 含表格 docx → 指纹字段完整性
- 2个视觉相同表格 → 去重为1条
- has_header_row 判断准确性
- 无表格文档 → 空列表，不报错

**`tests/test_template_analyzer.py`**
- 输出 profile 同时包含 `roles` 和 `table_roles`
- `--para-no-heading-aware` 参数透传有效
- `--para-min-cluster-size` 参数透传有效

**`tests/test_style_transfer_table.py`**
- 列数相同优先匹配
- 无完全匹配时取最高分（不跳过）
- 生成 ops 数量 = 目标文档表格数
- `table_roles` 为空时跳过且不报错

**`tests/test_docx_editor_table.py`**
- `apply_table_style` op 成功执行
- 边框写入 XML 正确
- 表头底纹正确应用于首行
- 正文行字体覆盖正确

### 运行命令

```bash
$env:PYTHONPATH="skills/word-master/scripts"
python -m pytest tests/test_table_analyzer.py tests/test_template_analyzer.py tests/test_style_transfer_table.py tests/test_docx_editor_table.py -v
```
