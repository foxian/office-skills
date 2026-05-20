# Paragraph Role Detection Design

## 1. 概述 (Overview)

当前的 `style_transfer.py` 在处理草稿段落时，主要依靠视觉指纹相似度评分（`_score()`）来推断段落角色。这在 **Case A**（草稿段落已带有正确 Word 样式名）的场景下是冗余且低效的。

本次设计目标：

1. **引入多级确定性通道**：对于标题（Heading）和列表（List Bullet），直接通过样式名和 XML 底层节点作出 100% 确定的判定，不依赖指纹相似度推断。
2. **扩展可识别角色**：在原有 Heading 系列之外，增加对 `Normal`（标准正文）和 `List Bullet`（项目符号列表）的准确识别和样式应用。
3. **保持可扩展性**：为未来添加更多角色（如 `List Number`、`Callout Box` 等）预留清晰的挂载点。

---

## 2. 核心架构：四级角色判定流水线

对草稿文档中的每一个段落，按以下优先级依次判定其角色，一旦命中则立即返回，不再继续向下判断。

```
para
│
├─ 【第 0 级】Heading 样式名直接命中（确定性，优先级最高）
│   条件：para.style.name.startswith("Heading ")
│         且 style_name in available_roles
│   返回：role = para.style.name  （如 "Heading 1"）
│
├─ 【第 1 级】XML numPr 列表节点检测（确定性）
│   条件：段落 pPr 下存在 <w:numPr> 节点
│   返回：role = "List Bullet" 或 "List Number"
│   降级兜底（numPr 缺失时）：
│     → 样式名含 "bullet" / "list" 关键字
│     → 文本前缀为 •、·、-、*、o、▪ 等字符
│
├─ 【第 2 级】样式名为标准正文类（确定性）
│   条件：para.style.name in ("Normal", "Body Text",
│                             "Default Paragraph Style")
│   返回：role = "Normal"
│
├─ 【第 3 级】文本正则兜底（现有逻辑，启发式）
│   条件：_infer_role_from_text(text, profile) 返回非 None
│   返回：推断角色
│
└─ 【第 4 级】默认（兜底）
    返回：role = "Normal"
```

---

## 3. 具体修改模块 (Modules)

### 3.1 新增函数：`_detect_list_type(para, doc)` → `style_analyzer.py`

读取段落底层 XML 中的 `<w:numPr>` 节点，判定是否为 Word 原生列表，并区分项目符号（bullet）与数字编号（decimal/ordinal）类型。

**判定逻辑：**

1. 读取 `para._element.get_or_add_pPr().find(qn('w:numPr'))`。
2. 若 `numPr` 存在：
   - 取 `numId` 和 `ilvl`，在 `doc.part.numbering_part` 中追溯对应的 `abstractNum`，读取 `w:numFmt` 的值：
     - `numFmt == "bullet"` → 返回 `"List Bullet"`
     - 其他（`decimal`、`lowerLetter` 等）→ 返回 `"List Number"`
3. 若 `numPr` 不存在，则按以下顺序降级：
   - 样式名含 `"bullet"` → 返回 `"List Bullet"`
   - 样式名含 `"list"` 且文本以数字/字母序号开头 → 返回 `"List Number"`
   - 样式名含 `"list"` → 返回 `"List Bullet"`
   - 文本以常见项目符号字符开头（`•·-*o▪■◆▲`）→ 返回 `"List Bullet"`
4. 以上均未命中 → 返回 `None`。

**函数签名：**

```python
def _detect_list_type(paragraph, doc) -> str | None:
    """
    Detect if a paragraph is a Word native list.
    Returns "List Bullet", "List Number", or None.
    Uses <w:numPr> XML node as primary signal; falls back to
    style name keywords and text prefix characters.
    """
```

### 3.2 修改函数：`generate_apply_ops()` → `style_transfer.py`

在现有逻辑的 `for i, para in enumerate(paras):` 循环中，将判定逻辑重构为四级流水线：

```python
style_name = para.style.name

# 第 0 级
if style_name.startswith("Heading ") and style_name in available_roles:
    role = style_name

# 第 1 级
elif _detect_list_type(para, doc) is not None:
    detected = _detect_list_type(para, doc)
    role = detected if detected in available_roles else "Normal"

# 第 2 级
elif style_name in ("Normal", "Body Text", "Default Paragraph Style"):
    role = "Normal" if "Normal" in available_roles else None

# 第 3 级
else:
    role = _infer_role_from_text(text, profile)

# 第 4 级
if role is None:
    role = "Normal"
```

> **注意**：`_detect_list_type` 在每次循环中调用两次（判断 + 取值），后续可改为只调用一次并缓存结果。

### 3.3 `style_profile.json` 格式扩展

需在 `roles` 数组中补充 `Normal` 和 `List Bullet` 两个角色的条目。`role` 字段的值**直接使用 Word 内置样式名**，与 `apply_style` op 的 `style` 字段保持一致，零转换。

```json
{
  "roles": [
    {
      "role": "Heading 1",
      "fingerprint": { "size": "16.0pt", "font": "黑体", "bold": true, "align": "center" }
    },
    {
      "role": "Normal",
      "fingerprint": { "size": "11.0pt", "font": "宋体", "bold": false, "align": "justify" }
    },
    {
      "role": "List Bullet",
      "fingerprint": { "size": "11.0pt", "font": "宋体", "bold": false, "align": "left" }
    }
  ]
}
```

### 3.4 样式定义提取来源的升级（连锁决策）

与本次设计配套，**`extract_fingerprints()` 的提取来源应从"段落实例"升级为"样式定义"**：

- 对于有明确 Word 样式名的角色（Heading X、Normal、List Bullet），直接从 `doc.styles[role_name].font` 和 `doc.styles[role_name].paragraph_format` 中提取指纹，避免段落内 run-level 手动覆盖的污染。
- 此升级在下一个独立 Spec 中细化，本次设计不包含此改动。

---

## 4. `style_profile.json` 中 `role` 与 Word 样式名的约定

| 语义角色 | `role` 字段值 | Word 内置样式名 | 说明 |
|---------|-------------|---------------|------|
| 一级标题 | `"Heading 1"` | `Heading 1` | 直接使用内置名 |
| 二级标题 | `"Heading 2"` | `Heading 2` | 直接使用内置名 |
| 标准正文 | `"Normal"` | `Normal` | 直接使用内置名 |
| 项目符号列表 | `"List Bullet"` | `List Bullet` | 直接使用内置名 |

> **原则**：`role` 字段值 = Word 内置样式名，`apply_style` op 的 `style` 字段直接使用 `role` 值，不需要额外映射层。

---

## 5. 优势 (Advantages)

- **准确率提升**：Heading 和 List Bullet 的识别不再依赖浮点相似度比较，命中率从"概率性"提升到"确定性"。
- **性能提升**：对于 Case A 文档，跳过了 `_score()` 计算和指纹向量比较。
- **语义清晰**：`role` 字段直接对应 Word 样式名，整个管道的数据契约更简洁，调试也更直观。
- **可扩展**：四级流水线结构清晰，未来新增 `Callout Box`、`List Number`、`Caption` 等角色，只需在第 1～2 级之间插入新的检测函数即可。

---

## 6. 涉及文件

| 文件 | 操作 | 说明 |
|------|------|------|
| `skills/word-master/scripts/style_analyzer.py` | Modify | 新增 `_detect_list_type(para, doc)` 函数 |
| `skills/word-master/scripts/style_transfer.py` | Modify | 重构 `generate_apply_ops()` 中的角色判定为四级流水线 |
| `skills/word-master/scripts/validate_style_profile.py` | Modify | 补充对 `Normal` 和 `List Bullet` role 的校验白名单 |
| `tests/test_style_transfer.py` | Modify | 更新测试：验证列表段落被正确识别为 `List Bullet` |
| `tests/test_style_analyzer.py` | Modify | 新增测试：验证 `_detect_list_type` 在有/无 `numPr` 场景下的返回值 |

---

## 7. 后续实施步骤

1. 在 `style_analyzer.py` 中实现 `_detect_list_type()`，并编写单元测试覆盖 numPr 存在、不存在、文本前缀三种路径。
2. 重构 `style_transfer.py` 的 `generate_apply_ops()`，替换现有的单级指纹判定为四级流水线。
3. 更新 `validate_style_profile.py`，在角色白名单中加入 `"Normal"` 和 `"List Bullet"`。
4. 端到端测试：用一份包含标题、正文、列表的真实文档运行完整样式迁移，验证三类角色均被正确识别和应用。
