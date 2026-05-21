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

### 2. 标题段落判定：以 outlineLvl 为准

启用 `heading_aware=True` 时：
- 对每个非空段落，读取 `<w:pPr><w:outlineLvl w:val="N"/>` 属性
- 若该属性存在（值 0~8），将段落视为标题，归入 `Heading {N+1}` 分组
- 完全忽略段落的样式名（`para.style.name`）

**理由：** 用户确认模板文档的导航目录是规范的，因此 `outlineLvl` 比样式名更可信。

### 3. 同级多段取平均指纹

同一大纲级别可能有多个段落（且视觉格式可能略有差异）。取**所有该级别段落的各字段平均值**：
- `size`：数值平均后格式化为 `"Xpt"`（保留 1 位小数）
- `bold`、`italic`：多数投票（True/False/None 中取最多数的）
- `align`：多数投票
- `font`：取出现最频繁的值
- `color`：取出现最频繁的值（None 视为一种值参与投票）
- `space_before`、`space_after`：数值平均后格式化
- `first_line_indent`：数值平均后格式化
- `line_spacing`：数值平均（None 段忽略）

**理由：** 多段平均比取第一段更稳健，避免因模板中的个别格式异常影响结果。

### 4. 正文段落走原有逻辑

无 `outlineLvl` 属性的段落（即导航目录不认为其是标题的段落）仍走原有格式聚类逻辑，并应用 `min_cluster_size` 过滤。

### 5. 输出带 `heading_role` 字段

heading-aware 模式的标题分组结果中，每个条目增加 `"heading_role": "Heading N"` 字段，使下游 AI 可以直接识别该 fingerprint 对应的 Word 内置样式名，无需再猜测。

正文分组条目不含此字段，格式与原有一致。

### 6. 新增工具函数 `_get_outline_level`

单独封装 XML 读取逻辑，便于测试和复用：

```python
def _get_outline_level(paragraph) -> int | None:
    # 读取 <w:pPr><w:outlineLvl w:val="N"/>
    # 返回 0~8 整数，或 None（未设置）
```

### 7. CLI 增加 `--heading-aware` flag

```bash
python style_analyzer.py template.docx --heading-aware --output fingerprints.json
```

## 输出示例

```json
[
  {
    "heading_role": "Heading 1",
    "fingerprint": {"size": "16.0pt", "bold": true, "align": "center", "font": "黑体", ...},
    "example": "第一章 概述"
  },
  {
    "heading_role": "Heading 3",
    "fingerprint": {"size": "13.0pt", "bold": true, "align": "justify", "font": "黑体", ...},
    "example": "三级标题示例（此段实际用了Heading 4样式，但outlineLvl=2）"
  },
  {
    "id": 0,
    "fingerprint": {"size": "12.0pt", "bold": false, "align": "justify", "font": "宋体", ...},
    "example": "正文内容"
  }
]
```

## 不在本次范围内

- `style_transfer.py` 不做修改（heading-aware 只影响指纹提取阶段，不影响应用阶段）
- 不自动修复模板文档中的样式名错误
- 不处理 `outlineLvl >= 9` 的情况（超出 Word 正常范围，忽略即可）

## 验证标准

1. 单元测试：`_get_outline_level` 在无属性时返回 `None`，在 `outlineLvl=2` 时返回 `2`
2. 单元测试：heading-aware 模式下，`Heading 4` 样式但 `outlineLvl=2` 的段落，归入 `Heading 3` 分组
3. 单元测试：同一大纲级别多段，指纹取平均值
4. 回归测试：`heading_aware=False`（默认）下，所有原有测试全部通过
