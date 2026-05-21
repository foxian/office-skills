# Heading-Aware 指纹提取 设计规格

## 背景

`style_analyzer.py` 的 `extract_fingerprints()` 函数用于从模板文档提取段落格式指纹，供 AI 推断各段落的样式角色（role），再由 `style_transfer.py` 应用到目标草稿。

当前实现按**视觉格式**聚类段落——相同格式的段落归入同一个 fingerprint cluster。这在模板规范时工作良好，但有一类常见异常：

**问题：** 源模板中，某些标题段落的**样式名与实际大纲级别不一致**。例如：
- 逻辑上是三级标题，导航目录正确显示为三级
- 但段落实际套用的是 `Heading 4` 样式（视觉格式是四级标题的格式）

Word 的导航目录（大纲）由 XML 属性 `<w:outlineLvl w:val="N"/>` 驱动，与样式名无关。因此导航正确，但样式名错误。

**后果：** 当前聚类逻辑按格式分组，无法区分"outlineLvl=2 的段落（三级）"与"outlineLvl=3 的段落（四级）"——如果它们恰好视觉格式相同，会被错误地归入同一 cluster，导致 AI 无法正确推断层级对应关系。

## 目标

在 `style_analyzer.py` 中增加 **heading-aware 提取模式**：提取标题指纹时，以段落的 `outlineLvl` 大纲级别（导航目录中的层级）为准，而非样式名，从而使 AI 获得正确的层级→格式对应关系。

## 设计决策

### 1. 新增参数，默认不改变现有行为

为 `extract_fingerprints()` 新增 `heading_aware=False` 参数。默认值 `False` 保持原有聚类逻辑不变，不破坏任何现有用法。

---

### 2. outlineLvl 读取策略：段落直接属性优先，回退样式继承链

启用 `heading_aware=True` 时，`_get_outline_level(paragraph)` 按以下顺序查找 `outlineLvl`：

1. **段落直接 XML 属性**：`para._element.pPr.outlineLvl`（若存在，直接返回）
2. **段落样式继承链**：依次沿 `para.style → para.style.base_style → ...` 查找，找到第一个有 `outlineLvl` 的样式 XML 节点

若两者均无，返回 `None`（视为非标题，走正文逻辑）。

**理由：** Word 中，`Heading 3` 样式本身在样式定义 XML 里携带 `outlineLvl=2`；用了正确样式名的标题段落通常不会在段落级别再显式设置 `outlineLvl`，而是继承自样式。仅读段落直接属性会遗漏这类正确的标题段落，导致它们错误地被归入正文聚类。

**实现要点：**
- 读取 `para.style` 的 XML：通过 `para.style.element.pPr.outlineLvl`（注意样式 element 与段落 element 不同）
- 沿继承链上溯时，以第一个有值的节点为准（不做数值合并）
- 样式继承链最多上溯 10 层（防止循环引用）

---

### 3. 同级多段取平均指纹

同一大纲级别可能有多个段落，取**所有该级别段落的各字段聚合值**：

| 字段 | 聚合方式 | None 处理 |
|------|---------|----------|
| `size` | 数值平均，格式化为 `"X.Xpt"`（1位小数） | 跳过 None 段，只平均有值的；若全为 None 则结果为 None |
| `bold` | 多数投票（True / False / None 各计数，取最多的） | None 作为独立选项参与投票；若 None 票最多，结果为 None（保留"继承自样式"语义） |
| `italic` | 同 bold | 同 bold |
| `align` | 多数投票 | None 作为独立选项参与投票 |
| `font` | 出现频率最高的字符串值 | None 跳过，不参与投票；若全为 None 则结果为 None |
| `color` | 出现频率最高的字符串值 | None 作为独立选项参与投票 |
| `space_before` | 数值平均，格式化为 `"X.Xpt"` | 视为 0.0pt（与 compute_effective_fingerprint 行为一致） |
| `space_after` | 同 space_before | 同上 |
| `first_line_indent` | 同 space_before | 同上 |
| `line_spacing` | 数值平均（仅 float 类型） | 跳过 None；若全为 None 则结果为 None |

**理由：** 多段平均比取第一段更稳健，避免因模板中的个别格式异常影响结果。

---

### 4. 正文段落走原有逻辑

无 `outlineLvl` 属性（含继承链均无）的段落，仍走原有格式聚类逻辑，并应用 `min_cluster_size` 过滤。

---

### 5. 输出格式与 id 字段规则

heading-aware 模式输出列表，条目顺序：**标题分组在前（按 Heading 1 → 9 升序），正文分组在后**。

- **标题分组**：含 `heading_role`、`fingerprint`、`example` 字段，**不含 `id` 字段**
- **正文分组**：含 `id`、`fingerprint`、`example` 字段，**不含 `heading_role` 字段**，`id` 从 `0` 开始对正文分组顺序编号

正文分组条目格式与原有 `heading_aware=False` 模式的输出完全一致。

---

### 6. 新增工具函数 `_get_outline_level`

单独封装 outlineLvl 读取逻辑（含继承链回退），便于测试和复用：

```python
def _get_outline_level(paragraph) -> int | None:
    """
    读取段落的大纲级别（0=Heading 1, 1=Heading 2, ...）。
    优先读段落直接 XML 属性，其次沿样式继承链查找。
    返回 0~8 整数，或 None（无大纲级别，视为正文）。
    """
```

---

### 7. CLI 增加 `--heading-aware` flag

```bash
python style_analyzer.py template.docx --heading-aware --output fingerprints.json
```

---

## 输出示例

```json
[
  {
    "heading_role": "Heading 1",
    "fingerprint": {"size": "16.0pt", "bold": true, "align": "center", "font": "黑体", "color": null, "space_before": "12.0pt", "space_after": "6.0pt", "first_line_indent": "0.0pt", "line_spacing": null},
    "example": "第一章 概述"
  },
  {
    "heading_role": "Heading 3",
    "fingerprint": {"size": "13.0pt", "bold": true, "align": "justify", "font": "黑体", "color": null, "space_before": "0.0pt", "space_after": "0.0pt", "first_line_indent": "0.0pt", "line_spacing": null},
    "example": "三级标题示例（此段实际用了Heading 4样式，但outlineLvl=2）"
  },
  {
    "id": 0,
    "fingerprint": {"size": "12.0pt", "bold": false, "align": "justify", "font": "宋体", "color": null, "space_before": "0.0pt", "space_after": "0.0pt", "first_line_indent": "24.0pt", "line_spacing": 1.5},
    "example": "正文内容"
  }
]
```

---

## 不在本次范围内

- `style_transfer.py` 不做修改（heading-aware 只影响指纹提取阶段，不影响应用阶段）
- 不自动修复模板文档中的样式名错误
- 不处理 `outlineLvl >= 9` 的情况（超出 Word 正常范围，返回 None 视为正文）

---

## 验证标准

1. **单元测试**：`_get_outline_level` 在段落无直接属性时返回 `None`（普通段落）
2. **单元测试**：`_get_outline_level` 在段落直接设置 `outlineLvl=2` 时返回 `2`
3. **单元测试**：`_get_outline_level` 在段落无直接属性但使用 `Heading 3` 样式时，通过继承链返回 `2`
4. **单元测试**：heading-aware 模式下，`Heading 4` 样式但 `outlineLvl=2` 的段落，归入 `Heading 3` 分组
5. **单元测试**：同一大纲级别多段，`size` 字段取数值平均；`size=None` 的段跳过不参与平均
6. **单元测试**：同一大纲级别多段，`bold` 多数投票，平局时 `None` 选项的处理（None 票最多则结果为 None）
7. **集成测试**：文档含样式名正确标题 + 样式名错误标题 + 正文，heading-aware 模式输出中标题按 outlineLvl 正确归组，正文走聚类逻辑
8. **回归测试**：`heading_aware=False`（默认）下，所有原有测试全部通过
