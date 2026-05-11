# Word Master Phase 1 实现设计

## 1. 项目定位

Word Master 是 AI 驱动的 Word 文档智能创作与编辑系统，属于 `office-master` 套件。核心采用"AI 意图层 + 脚本执行层"双层架构。

**当前阶段**：Phase 1 补全，将 MVP 推进到完整可用状态。

## 2. 现状分析

| 模块 | 状态 | 说明 |
|------|------|------|
| `docx_reader.py` | ✅ 完成 | 支持 `--overview` 全局概览，输出带段落 ID |
| `docx_editor.py` | ⚠️ 部分 | 支持 `rewrite_paragraph`、`replace_text`；缺失 `set_font`、`set_paragraph_format` |
| `docx_diff.py` | ✅ 完成 | 使用 `difflib.unified_diff` 比对 |
| `md_to_docx.py` | ⚠️ 基础 | 仅裸转换，无模板支持 |
| `edit-ops.md` | ⚠️ 不完整 | 仅记录 4 个基础操作 |
| `SKILL.md` | ⚠️ 简单 | 需要对齐 ppt-master 风格 |

## 3. 实现目标

1. 补全 `set_font` 和 `set_paragraph_format` 操作
2. 实现双轨模板系统（JSON 规则 + DOCX 样式）
3. 完善 DSL 规范文档
4. 更新 SKILL.md 对齐 ppt-master 风格
5. 补充测试用例确保质量

## 4. DSL 操作规范

### 4.1 现有操作

**`replace_text`**：查找替换
```json
{"op": "replace_text", "find": "A", "replace": "B", "scope": "all"}
```

**`rewrite_paragraph`**：重写段落
```json
{"op": "rewrite_paragraph", "target": "p0", "content": "新内容"}
```

### 4.2 新增操作

**`set_font`**：设置字体格式
```json
{
  "op": "set_font",
  "target": "p0",
  "name": "Arial",
  "east_asia": "楷体",
  "size": "14pt",
  "bold": true,
  "italic": false
}
```
| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| target | string | ✅ | 段落 ID，如 "p0" |
| name | string | 否 | 西文字体名 |
| east_asia | string | 否 | 中文字体名 |
| size | string | 否 | 字号，如 "14pt" 或 "28"（half-points） |
| bold | boolean | 否 | 是否加粗 |
| italic | boolean | 否 | 是否斜体 |

**`set_paragraph_format`**：设置段落格式
```json
{
  "op": "set_paragraph_format",
  "target": "p0",
  "alignment": "center",
  "line_spacing": 1.5,
  "first_line_indent": "2em",
  "space_before": "12pt",
  "space_after": "12pt"
}
```
| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| target | string | ✅ | 段落 ID |
| alignment | string | 否 | left/center/right/justify |
| line_spacing | number | 否 | 行距倍数 |
| first_line_indent | string | 否 | 首行缩进，如 "2em" |
| space_before | string | 否 | 段前距 |
| space_after | string | 否 | 段后距 |

## 5. 模板系统设计

### 5.1 双轨模板制

| 模板类型 | 指定方式 | 用途 |
|----------|----------|------|
| JSON 规则 | `--template json:formal` | 纯文本 Markdown 转换时的排版规则 |
| DOCX 样式 | `--template docx:参考.docx` | 复刻已有文档的样式体系 |

### 5.2 JSON 模板格式

**文件位置**：`${SKILL_DIR}/format_rules/{name}.json`

```json
{
  "heading1": {
    "font": "方正小标宋简体",
    "size": "22pt",
    "bold": true,
    "alignment": "center",
    "space_before": "0pt",
    "space_after": "12pt"
  },
  "heading2": {
    "font": "方正黑体简体",
    "size": "16pt",
    "bold": true,
    "alignment": "left",
    "space_before": "12pt",
    "space_after": "6pt"
  },
  "body": {
    "font": "仿宋",
    "size": "16pt",
    "line_spacing": 1.5,
    "first_line_indent": "2em",
    "alignment": "justify",
    "space_before": "0pt",
    "space_after": "6pt"
  }
}
```

### 5.3 Markdown 标签映射

| Markdown 语法 | 对应格式键 |
|---------------|------------|
| `# 标题` | heading1 |
| `## 标题` | heading2 |
| 普通段落 | body |

## 6. 文件改动清单

| 文件 | 操作 | 说明 |
|------|------|------|
| `scripts/docx_editor.py` | 修改 | 新增 `_apply_set_font()`、`_apply_set_paragraph_format()` |
| `scripts/md_to_docx.py` | 修改 | 支持 `--template json:xxx` 和 `--template docx:xxx` |
| `references/edit-ops.md` | 修改 | 补充 `set_font`、`set_paragraph_format` 规范 |
| `skills/word-master/SKILL.md` | 修改 | 对齐 ppt-master 风格工作流 |
| `format_rules/default.json` | 新增 | 默认公文格式 |
| `format_rules/formal.json` | 新增 | 正式文书格式 |
| `format_rules/casual.json` | 新增 | 日常文档格式 |
| `tests/test_docx_editor.py` | 修改 | 新增 4 个测试用例 |

## 7. 错误处理

| 场景 | 处理方式 |
|------|----------|
| 目标段落 ID 无效 | 跳过该操作，记录 warning |
| 字体名称不存在 | python-docx 忽略，使用原字体 |
| 数值格式错误 | 使用默认值 |
| 模板文件不存在 | 报错退出 |

## 8. 测试用例

| 测试函数 | 验证内容 |
|----------|----------|
| `test_set_font_name` | 设置字体名称生效 |
| `test_set_font_bold_italic` | 加粗/斜体切换 |
| `test_set_paragraph_alignment` | 居中/左对齐/右对齐 |
| `test_set_paragraph_line_spacing` | 1.5倍行距 |
| `test_mixed_operations` | 混合多种操作执行 |

## 9. SKILL.md 更新要点

参考 ppt-master 风格：
- 明确的执行纪律（Mandatory）
- 完整的脚本说明表格
- 清晰的工作流步骤（Step 1, Step 2...）
- 路径使用 `${SKILL_DIR}` 变量

## 10. 验收标准

- [ ] 6 个现有测试全部通过
- [ ] 新增 4 个测试全部通过
- [ ] `set_font` 能正确修改字体、字号、加粗、斜体
- [ ] `set_paragraph_format` 能正确修改对齐、行距、缩进
- [ ] `md_to_docx.py --template json:formal` 能按规则生成格式化的 DOCX
- [ ] `edit-ops.md` 完整记录所有 DSL 操作
- [ ] `SKILL.md` 风格对齐 ppt-master