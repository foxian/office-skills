# 格式指纹增强设计 (Fingerprint Enhancement Design)

## 1. 背景与问题

### 现状

`style_analyzer.py` 中的 `compute_effective_fingerprint()` 当前提取的指纹字段如下：

```json
{
  "size": "14.0pt",
  "bold": false,
  "italic": null,
  "align": "center"
}
```

### 核心问题：区分度严重不足

在"假样式"文档（作者使用 Word 命名样式但随后用 run-level 手动覆盖了视觉属性）中，由于大多数段落的 run-level 覆盖都将字号统一为相同值（如本次测试文档中所有段落均为 14.0pt），仅有 `size/bold/italic/align` 四个维度无法有效区分不同语义角色的段落，导致：

- 多个不同角色的段落被聚合成相同或相似的指纹组
- AI 在推断 style_profile 时缺乏足够的区分信号
- 样式迁移映射精度下降

### 缺失的最重要维度

在中文文档中，**字体家族（Font Family）** 是极为重要的语义信号：

| 字体 | 常见语义角色 |
|------|------------|
| 黑体 | 标题、重点 |
| 宋体 | 正文 |
| 仿宋_GB2312 | 引用、正文变体 |
| 楷体 | 特殊强调 |

此外，**段落间距**（space_before/space_after）是区分标题与正文的重要视觉信号（标题通常段前段后留白更大）。**行间距**（line_spacing）提供辅助参考。

---

## 2. 设计方案：扩展格式指纹

**总原则**：视觉维度为主，保持直接解析 `.docx` 原文件的方式（`python-docx`），不改为解析中间 MD 文件。

### 2.1 新的指纹结构

```json
{
  "id": 7,
  "fingerprint": {
    "size": "14.0pt",
    "bold": false,
    "italic": null,
    "align": "center",
    "font": "黑体",
    "space_before": "17.0pt",
    "space_after": "16.5pt",
    "line_spacing": 1.5
  },
  "example": "总体要求"
}
```

新增字段说明：

| 字段 | 来源 | 说明 |
|------|------|------|
| `font` | run.font.name → style.font.name（后备） | 中文主字体，优先取首个有显式字体覆盖的 Run；run.font 为 None 时 fallback 到 style.font.name；仍为 None 时标记为 `null` |
| `space_before` | paragraph.paragraph_format.space_before → style | 段前间距，单位 pt，格式化为字符串；paragraph_format 为 None 时 fallback 到 0pt |
| `space_after` | paragraph.paragraph_format.space_after → style | 段后间距，单位 pt，格式化为字符串；paragraph_format 为 None 时 fallback 到 0pt |
| `line_spacing` | paragraph.paragraph_format.line_spacing → style | 行间距，**仅支持倍数模式**（float，如 1.5）；若为绝对 Pt 值则规范化为 `null`（暂不支持）；若为 None 则为 `null` |

---

### 2.2 聚类 Key 的变更（`_fingerprint_key()`）

**规则：font 加入聚类 key，spacing 不加入。**

```python
# 旧
def _fingerprint_key(fp):
    return (fp.get("size"), fp.get("bold"), fp.get("italic"), fp.get("align"))

# 新
def _fingerprint_key(fp):
    return (fp.get("size"), fp.get("bold"), fp.get("italic"), fp.get("align"), fp.get("font"))
```

**理由**：
- `font` 加入：字体完全不同的段落不应被归为同一格式类，加入聚类 key 可提升分组精度。
- `space_before/after/line_spacing` 不加入：浮点精度问题（如 `16.9999pt` vs `17.0pt`）会导致同类段落被分裂成大量微小差异组，聚类失效。这三个字段保留在指纹数据中供 AI 参考，但不参与聚类分组。

---

### 2.3 匹配评分的权重变更（`_score()` in `style_transfer.py`）

总维度从 4 个扩展到 6 个参与评分（line_spacing 仅记录不评分），权重重新分配（总计 = 1.0）：

| 维度 | 权重 | 理由 |
|------|------|------|
| `size` | 2/8 = 0.25 | 字号是最强区分信号，双倍权重 |
| `bold` | 1/8 = 0.125 | |
| `italic` | 1/8 = 0.125 | |
| `align` | 1/8 = 0.125 | |
| `font` | 2/8 = 0.25 | 字体家族是第二强信号，双倍权重 |
| `space_before` | 0.5/8 = 0.0625 | 辅助参考 |
| `space_after` | 0.5/8 = 0.0625 | 辅助参考 |

`line_spacing` 只写入指纹供 AI 查看，**不参与数值评分**，原因是正文和标题的行间距差异规律在不同文档中不稳定。

---

## 3. 变更范围（改动文件）

| 文件 | 函数 | 变更描述 |
|------|------|---------|
| `style_analyzer.py` | `compute_effective_fingerprint()` | 新增字段：`font`、`space_before`、`space_after`、`line_spacing`；fallback 行为：font→默认字体，spacing→0/null，line_spacing→null（绝对Pt值亦规范化为null） |
| `style_analyzer.py` | `_fingerprint_key()` | 加入 font 参与聚类 key |
| `style_transfer.py` | `_score()` | 按新权重表重写评分逻辑（6 个维度：size 2/8, bold/italic/align 各 1/8, font 2/8, space_before/space_after 各 0.5/8） |

---

## 4. 不变范围

- 不改变解析数据源（继续直接解析 `.docx` 原文件，不改为解析中间 MD 文件）
- 不改变 `fingerprints.json` 和 `style_profile.json` 的顶层结构（向后兼容）
- 不改变 `style_transfer.py` 的调用方式和 CLI 参数
- `line_spacing` 仅追加到指纹 JSON，不加入聚类 key，不加入 `_score()` 评分

---

## 5. 成功标准

- 对《招标书》重新运行 `style_analyzer.py`，生成的 `fingerprints.json` 中各组的 `font` 字段能有效区分出不同角色的段落（如黑体组 vs 宋体组）
- 指纹组数量应有所增加（之前仅 5 组，增强后预期按字体区分会有更多组）
- 重新生成 `style_profile.json` 后，样式迁移的准确率相比之前有明显提升
- 量化指标（可选）：样式迁移后 `apply_style` 操作与预期角色匹配率 ≥ 80%
