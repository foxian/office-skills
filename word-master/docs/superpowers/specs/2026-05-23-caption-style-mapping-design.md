# 图表题注高精度识别与样式绑定设计规格 (Caption Style Mapping Design Spec)

## 1. 概述 (Overview)

在文档样式迁移（`style_transfer.py`）中，**图表题注（Captions）** 是区分“普通正文”和“结构化元素”的关键排版特征。如果将题注误判为普通正文，不仅会导致字号、字体错乱，最糟糕的是极易造成**图表与题注发生跨页截断**（即图表在上一页，题注在下一页）的严重排版事故。

本设计的目标是：
1. **多级判定流水线升级**：在原有的段落角色检测流水线中引入 【第 1.5 级】题注专属检测通道，结合多级正则与物理位置上下文（Table / Drawing 节点），实现 100% 确定性的题注判定。
2. **样式级别规范套用**：将识别到的题注段落的 Word 样式名设为内置的 `"Caption"` 样式。
3. **安全创建与回退**：若目标文档或源模板缺少该样式，系统将在应用阶段动态创建具有高鲁棒性（区分中西文字体）的 `"Caption"` 样式。
4. **防断页版式保护**：精确应用 `keep_with_next`，物理隔绝跨页断开的问题。

---

## 2. 核心架构：多级判定流水线的升级

我们将对 `style_transfer.py` 及 `style_analyzer.py` 中的角色判定逻辑进行升级，在 **【第 1 级】** 和 **【第 2 级】** 之间插入 **【第 1.5 级】题注通道**：

```
para (段落)
│
├─ 【第 0 级】Heading 样式名直接命中 (Heading 1-9)
│
├─ 【第 1 级】XML numPr 列表节点检测 (List Bullet / List Number)
│
├─ 💡【第 1.5 级】题注段落检测 (Caption)  <-- [本次新增通道]
│   ├─ 判定 A：段落样式名直接为 "Caption" (或含 "caption" 关键字)
│   └─ 判定 B：物理布局上下文 + 文本正则匹配组合判定
│
├─ 【第 2 级】样式名为标准正文类 (Normal)
│
├─ 【第 3 级】文本正则分类 (基于视觉指纹相似度评分/启发式分类)
│
└─ 【第 4 级】默认兜底 (Normal)
```

---

## 3. 具体修改模块与算法实现

### 3.1 新增题注检测函数：`_detect_caption_type` -> `style_analyzer.py`

此函数用于检测当前段落是否属于题注。

**判定逻辑：**

1. **直接样式名判定**：
   若 `paragraph.style.name` 为 `"Caption"`，或不区分大小写地包含 `"caption"` 关键字，直接返回 `True`。
2. **文本正则初筛（支持多级编号）**：
   匹配以下正则表达式，重点优化对三级及以上编号（如 `图 1.2.3` 或 `Figure 2-1-3`）的捕获支持：
   ```python
   CAPTION_REGEX = re.compile(
       r"^(图|表|附图|附表|图表|Fig|Figure|Table|Chart)\s*(\d+([.-]\d+)*)\s*[:：]?\s*.*", 
       re.IGNORECASE
   )
   ```
   若当前段落文本不匹配此正则，直接返回 `False`。
3. **物理上下文判定（Look Back & Look Ahead - 容忍空段落）**：
   当文本通过正则初筛后，我们利用段落底层的 XML 兄弟节点关系检测其是否紧邻图表。为了容忍排版中常见的空行（空段落），我们支持对前/后最多两个节点的容错检测：
   * **向上看 (Look Back)**：检查当前段落的**前一个或前两个**节点：
     - 若为 `Table` 节点，或者段落的 XML 中包含图片元素（`<w:drawing>` 或 `<w:pict>`），说明是位于图表**下方**的题注。返回 `True`。
   * **向下看 (Look Ahead)**：检查当前段落的**后一个或后两个**节点：
     - 若为 `Table` 节点，说明是位于表格**上方**的题注。返回 `True`。
4. **格式降级兜底**：
   若没有检测到物理关联（例如图表为浮动框），且文本匹配正则，同时满足以下特征：
   - 段落字号（`size`）小于或等于正文默认字号（Normal）。
   - 段落对齐方式为居中（`center`）且无中文惯用的首行缩进（`first_line_indent == 0`）。
   满足条件则返回 `True`，否则返回 `False`。

**函数签名：**
```python
def _detect_caption_type(paragraph, doc) -> bool:
    """
    Detect if a paragraph is a document caption (for tables or images).
    Uses a combination of style name, regular expressions, and 
    physical layout adjacency to adjacent tables or drawings (supporting 1-2 paragraph gap).
    """
```

---

### 3.2 样式的动态提取与生成 -> `style_analyzer.py`

在提取源模板格式指纹（`extract_fingerprints`）时：
* 补充对 `"Caption"` 角色的支持。若模板文档的 `styles` 列表中定义了 `"Caption"` 样式，则提取其字体（font）、段落格式（paragraph_format）作为 `"Caption"` 角色的特征指纹。

---

### 3.3 样式应用阶段的处理与动态补齐 -> `style_transfer.py`

在将学习到的样式应用（`generate_apply_ops`）到目标文档时：

1. **白名单拓展**：
   在 `style_profile.json` 角色的合法白名单中新增 `"Caption"`。
2. **样式名覆盖**：
   一旦段落被判定为 `"Caption"`，样式转移操作将其直接绑定为样式名 `"Caption"`。
3. **目标文档样式补齐 (区分中西文字体)**：
   在应用样式前，先扫描目标草稿文档：
   - 若草稿中不包含 `"Caption"` 样式，在 `doc.styles` 中**动态创建**一个名为 `"Caption"` 的样式，其基样式（`base_style`）继承自 `"Normal"`。
   - 读取源模板中提取的 `"Caption"` 格式指纹：
     - 若有，将这些格式属性赋给新建的 `"Caption"` 样式；
     - 若没有指纹，则动态创建安全推荐值（9.5pt，居中对齐，段前 6pt，段后 6pt，行距单倍，无缩进）。
     - **中西文字体精细化处理**：为了解决 `python-docx` 中 `font.name` 仅对西文生效的缺陷，我们在动态创建样式时，将通过底层 XML 手动指定中文字体（`w:rFonts` 的 `w:eastAsia` 属性）为“仿宋”（或与源模板一致的中文字体），西文字体指定为 `"Times New Roman"`。

4. **物理断页防护应用 (物理关联保护策略)**：
   在 Word 的分页渲染机制中，`keep_with_next`（`w:keepNext`）的效果是**“当前段落与其下一段落保持在同一页”**。因此，直接在“下方题注”上应用 `keep_with_next` 是**无效的**。我们必须根据题注的相对位置应用以下物理保护策略：

   * **情景 A：题注在图表/图片下方（最常见）**
     - **动作**：必须在**包含该图表/图片的宿主段落**上，强制将段落属性 `paragraph_format.keep_with_next` 设为 `True`。
   * **情景 B：题注在表格下方**
     - **动作**：表格的所有行（`Table.rows`）除最后一行外，其底层 XML 行格式属性 `w:cantSplit`（禁止跨页断行）必须设为 `True`。此外，如果表格上一段是普通文字，在上一段上设置 `keep_with_next = True` 保证表格不与前文剥离。
   * **情景 C：题注在图表/表格上方**
     - **动作**：由于题注在上方，此时可以直接在**题注段落自身**，强制将其段落属性 `paragraph_format.keep_with_next` 设为 `True`，保证其与下方的图表在同一页。

---

## 4. `style_profile.json` 角色约定

在 `style_profile.json` 的 `roles` 列表中新增成员：

| 语义角色 | `role` 字段值 | Word 内置样式名 | 说明 |
|---------|-------------|---------------|------|
| 图表题注 | `"Caption"` | `Caption` | 内置样式，强制居中，防断页 |

**配置示例**：
```json
{
  "roles": [
    {
      "role": "Caption",
      "fingerprint": { "size": "9.5pt", "font": "仿宋", "bold": false, "align": "center" }
    }
  ]
}
```

---

## 5. 涉及修改文件

| 目录/文件 | 操作 | 说明 |
|------|------|------|
| `skills/word-master/scripts/style_analyzer.py` | Modify | 实现 `_detect_caption_type(para, doc)`，并在提取中支持 `Caption` 指纹 |
| `skills/word-master/scripts/style_transfer.py` | Modify | 1. 流水线中引入 【第 1.5 级】题注判定<br/>2. 新增目标文档样式补齐逻辑（支持中西文字体区分）<br/>3. 实现图文空间关联物理防断页保护 |
| `skills/word-master/scripts/validate_style_profile.py` | Modify | 补充对 `"Caption"` 语义角色的强类型验证白名单 |
| `tests/test_style_transfer.py` | Modify | 更新测试：验证在没有题注样式的目标文档中，能够成功创建该样式并应用图文防断页 |
| `tests/test_style_analyzer.py` | Modify | 新增测试：验证在物理位置紧邻 Table 或 Image（含 1~2 段空行容错）的各种上下文路径下能够准确识别 `Caption` |

---

## 6. 验证方案 (Verification Plan)

### 6.1 自动化测试
1. **测试用例 1**：物理紧邻 Table 上下方的段落能 100% 判定为 `Caption`。
2. **测试用例 2**：题注与图表之间存在 1~2 个空行段落时，识别算法依然能精准捕获 `Caption` 角色。
3. **测试用例 3**：对于没有 `"Caption"` 样式的文档，自动补齐该样式并分别设置中西文字体（仿宋 + Times New Roman）。
4. **测试用例 4**：验证下方题注时，宿主图片段落的 `<w:keepNext>` 为 `True`；上方题注时，题注段落自身的 `<w:keepNext>` 为 `True`。

### 6.2 手动集成测试
* 运行全流程样式迁移，应用在一个包含大量“表格和下方题注”的复杂文档中，将生成的最终文档与源文档进行 Diff，并在 Word 中验证页面底部边界处没有发生“图表与题注在换页时被生硬拆开”的排版错误。
