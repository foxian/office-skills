# Word Document Style Transfer Design

## 1. 概述 (Overview)
本项目旨在扩展 Word Master 的能力，实现"学习一个 Word 文档的样式与格式风格，并将其应用到另一个 Word 文档中"。

基于真实业务场景（如招投标书、合同等），两份文档通常都包含：
- **特殊结构页**：封面、签字承诺函、合同等（有复杂排版、表格、大面积留白）
- **正文段落**：标题、小节标题、正文、列表等连续文字段落
- **假样式**：文档未使用 Word 命名样式，而是靠手动加粗/调整字号来伪装格式

**本次设计（Phase 1）目标**：在两份文档都含有混合结构的前提下，从源文档中提取正文区的格式风格规律，并将其**就地**应用到目标文档中对应的正文段落上，不改变文档内容和结构，只改变样式属性。

## 2. 核心架构：两阶段就地样式替换

整体流程分为两个独立阶段：

### 阶段 1：提取源模板的正文样式规律 → `style_profile.json`

```
source_template.docx
  → [style_analyzer.py] 提取所有段落的有效格式指纹
  → [正文识别策略] 过滤掉封面/签字页等特殊段落
  → [LLM 推理] 将正文指纹映射为 Word 标准样式角色
  → style_profile.json
```

`style_profile.json` 是一份描述"什么视觉格式 = 什么语义角色"的映射表：
```json
{
  "roles": [
    {"role": "Heading 1", "fingerprint": {"size": "16pt", "bold": true, "align": "center"}},
    {"role": "Heading 2", "fingerprint": {"size": "14pt", "bold": true, "align": "left"}},
    {"role": "Normal",    "fingerprint": {"size": "12pt", "bold": false, "align": "justify"}}
  ]
}
```

### 阶段 2：将 profile 就地应用到目标草稿 → `output.docx`

```
draft.docx
  → [正文识别策略] 过滤出正文段落（跳过封面/签字页）
  → [指纹匹配] 每个正文段落的当前格式 ↔ profile 中的指纹做相似度匹配
  → 找到最近的角色（如 Heading 1）
  → 调用 paragraph.style = "Heading 1" 应用真样式，清除手动覆盖
  → 保存为 output.docx
```

整个过程**不改变文档段落数量和内容**，只修改段落的 `style` 属性（以及清除旧的手动 run-level 格式覆盖）。

## 3. 正文段落识别策略

这是整个系统的核心难点，对源模板和目标草稿均需执行。采用 **A+C 组合法**：

### 策略 A：聚类数量过滤（多数原则）
对全文所有段落做格式指纹聚类：
- 正文段落数量多 → **大聚类**（通常 10 个以上成员）
- 封面/签字页段落极少 → **小聚类**（1~3 个成员）

过滤掉成员数 < 阈值（默认 = 4）的聚类，这些小聚类视为"特殊页面格式"，不参与正文样式推断。

### 策略 C：内容模式匹配
检测段落文本是否符合中文文档章节标志（正则匹配）：
```
第[一二三四五六七八九十\d]+章 | 一、 | （一） | \d+\. 等
```
命中的段落及其后续段落视为"正文区起始"的强信号。

### 策略 D：指纹相似度兜底（应用阶段专用）
在对目标草稿应用样式时，将每个段落的格式指纹与 `style_profile.json` 中已知的正文指纹集合做相似度比较：
- 相似度超过阈值（如 0.6）→ 视为正文，执行样式应用
- 相似度过低 → 该段落格式和模板正文差异太大，跳过（可能是封面元素）

同时提供 `--skip-head N` 和 `--skip-tail N` 参数，让用户手动排除头部 N 个段落和尾部 N 个段落，作为识别不准时的人工兜底。

## 4. 具体实现模块 (Modules)

### 4.1 格式分析器 (`style_analyzer.py`)

独立、只读的分析脚本。

- **输入**：`source_template.docx`
- **处理**：
  - 遍历所有段落，将 `paragraph.style` 定义属性与 run-level 手动覆盖**合并**，得到每段的**有效格式指纹**（有效字号、有效加粗、有效对齐等）。
  - 应用策略 A+C 过滤出正文候选段落。
  - 去重聚类，每个指纹类型只保留一条代表性示例文字，生成紧凑的指纹摘要。
- **输出**：`fingerprints.json`，结构如下：
  ```json
  [
    {"id": 0, "fingerprint": {"size": "16pt", "bold": true, "align": "center"}, "example": "第一章 项目概述"},
    {"id": 1, "fingerprint": {"size": "12pt", "bold": false, "align": "justify"}, "example": "本项目旨在…"}
  ]
  ```

### 4.2 样式转移总控 (`style_transfer.py`)

用户入口脚本，统筹整个两阶段工作流。

**基础用法（每次从模板重新分析）：**
```bash
python style_transfer.py template.docx draft.docx output.docx
```

**复用已有配置（跳过分析和 LLM 推理）：**
```bash
python style_transfer.py --profile style_profile.json draft.docx output.docx
```

内部执行逻辑：
1. **获取 style profile**：
   - 有 `--profile`：直接加载，跳过分析阶段（0 token 消耗）。
   - 无 `--profile`：调用 `style_analyzer.py` → 调用 LLM 推理 → 生成 `style_profile.json`，同时保存到磁盘供下次复用。
2. **生成 DSL 操作指令**：遍历 `draft.docx` 段落，识别正文段落，对每个正文段落做指纹匹配，生成 `apply_style` 操作列表，例如：
   ```json
   [
     {"op": "apply_style", "target": "p1", "style": "Heading 1"},
     {"op": "apply_style", "target": "p3", "style": "Normal"}
   ]
   ```
3. **委托 `docx_editor.apply_operations()` 执行**：`style_transfer.py` 不自己操作 `python-docx`，而是直接调用现有的 `docx_editor` 执行引擎，复用其备份机制和安全执行逻辑。`style_transfer.py` 是纯粹的**规划器（大脑）**，`docx_editor.py` 是唯一的**执行器（手）**。

### 4.3 审核模式 (Review Mode)

提供 `--review` 标志，在生成 `style_profile.json` 后暂停：
```bash
python style_transfer.py template.docx draft.docx output.docx --review
```
暂停后，用户可直接编辑 `style_profile.json` 修正 LLM 的推断（如把误判的角色从 Heading 2 改为 Normal），确认后按回车继续执行应用阶段。

## 5. Phase 1 范围限定

**本阶段只处理正文段落的样式替换**。以下内容暂不在范围内：
- 封面、签字承诺函、合同表格页的格式提取与迁移（未来单独设计）
- 段落内容的增删改（只改样式，不碰文字）
- 页眉、页脚、页面边距等页面级设置的迁移

## 6. 后续实施步骤
1. **实现 `style_analyzer.py`**：段落遍历 + 有效指纹计算 + A+C 正文过滤 + 聚类去重。
2. **设计 LLM 推理 Prompt**：将指纹摘要转化为 `style_profile.json` 的 Prompt 模板。
3. **实现应用逻辑**：读取 `draft.docx`，识别正文段落，指纹匹配，生成 `apply_style` DSL 列表，调用 `docx_editor.apply_operations()` 执行。
4. **集成与测试**：用真实的含假样式、混合结构（封面+正文+签字页）的文档验证端到端效果。
