# Word Master 设计规格说明书

## 1. 项目定位与核心哲学

`word-master` 是 AI 驱动的 Word 文档智能创作与编辑系统，归属于 `office-master` 套件，与 `ppt-master` 并列。

**核心哲学：双层架构（AI意图层 + 脚本执行层）**
- AI 负责“说要做什么”：通过阅读人类友好的 Markdown，产出结构化的意图或普通文本。
- 确定性引擎负责“怎么做”：通过精确解析 DSL 和调用底层库，安全、无损地执行修改或格式化，不让 AI 直接生成执行代码。

## 2. 两大核心工作流

Word Master 覆盖“从零写文档”和“改已有文档”两大核心场景。

### 2.1 创建模式 (从零创建)
解决 AI 写出 Markdown 极佳，但转换为高质量 Word 格式困难的痛点。

1. **大纲探讨**：对话确定结构与方向。
2. **AI 写作**：AI 产出标准纯文本 Markdown (`draft.md`)。
3. **审阅迭代**：用户与 AI 基于 Markdown 内容直接反复修改，敲定文字。
4. **格式转换**：调用 `md_to_docx.py --template <风格>`，将 Markdown 结合预设格式模板，一键渲染为排版规范的 DOCX。
5. **微调**：如需微小调整，无缝转入“编辑模式”。

### 2.2 编辑模式 (修改已有文档)
解决 DOCX 格式极其复杂，AI 直接修改容易破坏原格式的痛点。采用轻量的“读-改-验”对话循环。

1. **阅读层**：执行 `docx_reader.py`，将 DOCX 解析为带有段落 ID 和格式标注的富文本 Markdown，供 AI 快速“看懂”全貌或局部细节。
2. **意图层**：AI 理解用户指令，生成精准的编辑操作 (JSON DSL)。
3. **执行层**：执行 `docx_editor.py`（基于 `python-docx` 为主，`lxml` 兜底），在备份原文件的基础上进行精确修改。
4. **验证层**：执行 `docx_diff.py`，生成易读的变更报告供用户核验。

## 3. 核心模块职责拆分

| 模块 / 脚本 | 职责描述 |
|-------------|----------|
| `docx_reader.py` | 【阅读】读取 DOCX，输出带结构树、段落 ID 和基础样式的富文本 MD。支持全局概览 (`--overview`) 和局部读取 (`--section`)。 |
| `docx_editor.py` | 【执行】解析 JSON DSL，安全地在原文档上执行替换、插入、删除、调格式等操作。 |
| `docx_diff.py` | 【校验】对比修改前后的 DOCX 内容差异，生成变更摘要。 |
| `md_to_docx.py` | 【创建】按 `format_rules` 中的格式规范，将纯 Markdown 转换为排版精美的 DOCX。 |
| `docx_style_extractor.py`| 【逆向复刻】从参考 DOCX 中提取样式特征。分为“脚本提取原始特征”和“AI 推断样式体系”两步。 |
| `docx_template.py` | 【排版】为已有 DOCX 统一应用一套预制的格式规则库。 |

## 4. 编辑操作 DSL (JSON) 规范

`edit-ops.md` 将规范 AI 能够输出的操作集，保证修改过程的安全可控。

**内容操作**
- `replace_text`: 全局或局部查找替换。
- `rewrite_paragraph`: 重写指定 ID 的段落，仅改变内容，保留原有样式。
- `insert_paragraph`: 在指定 ID 后插入新段落。
- `delete_paragraph`: 删除目标段落。

**表格操作**
- `modify_cell`: 定位行列，修改单元格内容。
- `insert_row`: 在表格某行前后插入新行。
- `insert_table`: 插入具有特定样式的全新表格。

**格式操作**
- `set_font`: 修改目标范围内的字体（中英文字体、字号、加粗等）。
- `set_paragraph_format`: 设置对齐方式、行距、首行缩进、段前段后距。
- `apply_style`: 强制套用 Word 内部的预定义样式（如“标题 1”、“正文”）。
- `set_page_setup`: 页面调整（页边距、纸张方向）。

**结构操作**
- `insert_page_break`: 插入分页符。
- `set_header`: 设置页眉页脚。
- `insert_image`: 插入图片并设定宽度对齐。

## 5. 格式复刻机制 (Style Extraction)

针对不规范手动排版的源文档，采用“双步提取法”进行复刻：
1. **脚本层提取**：脚本如实输出所有段落的真实渲染特征（例如：“方正小标宋 22pt 加粗 居中”出现 5 次）。
2. **AI 模式识别**：AI 介入分析这些底层视觉数据，推断出“这代表一级标题”，从而生成抽象的规范化 `format_rules.json`。
3. 该规则可直接供给 `md_to_docx.py`，实现依葫芦画瓢的精准复刻。

## 6. MVP 实施计划

**Phase 1: 第一期（跑通核心循环与创建框架）**
- **创建模式**：实现 AI Markdown 写作 + `md_to_docx.py` 基础转换引擎。
- **编辑模式**：实现 `docx_reader.py` (输出带 ID 结构)，实现 `docx_editor.py` 支撑核心操作 (`replace_text`, `rewrite_paragraph`, `set_font`, `set_paragraph_format`)。
- **验证机制**：实现 `docx_diff.py` 基本文字比对。
- **目标**：能完整执行"从零写公文"和"修改公文字号/错别字"。

**Phase 2: 第二期（高级排版与增强能力）**
- **表格与结构**：扩展所有表格操作和分页、图片、页眉操作。
- **AI 审阅与推断**：增加 `reviewer.md` 指导 AI 审阅，实现 `docx_style_extractor.py` 的智能格式复刻。
- **健壮性兜底**：实现 `lxml` 对复杂 OOXML 结构的兜底修改方案。
