# 特殊页面识别设计规格

## 背景

当前 `template_analyzer.py` 提取的 `style_profile.json` 包含 `roles`（段落格式角色）和 `table_roles`（表格样式），但没有对文档中「特殊页面」（封面、签名页、目录页等）的识别和格式提取能力。

`style_transfer.py` 目前只提供 `--skip-head N` / `--skip-tail N` 参数粗略跳过首尾段落，无法精确处理特殊页面，也无法复用其格式规范做精确迁移。

## 目标

1. **识别**：自动识别文档中的特殊页面（封面、封底、目录、签名页、承诺函、版权声明、附录封面），位置推断与内容关键词两者结合。
2. **提取格式**：将特殊页面的段落格式摘要写入 `style_profile.json` 的新字段 `special_pages`，供未来精确迁移使用。
3. **精确跳过**：`style_transfer.py` 执行前，对草稿文档重新运行识别，精确确定哪些段落属于特殊页面，避免将正文样式错误应用到封面或签名页。
4. **可扩展**：通过注册表模式（Registry Pattern）支持无痛新增页面类型，不修改核心识别逻辑。

---

## 设计决策

### 1. 数据结构：新增 `special_pages` 顶层字段

`style_profile.json` 新增与 `roles`、`table_roles` 并列的 `special_pages` 字段，而非将特殊页面段落混入现有 `roles`。

**理由**：
- 封面大标题与正文 Heading 1 视觉格式可能完全不同，混入 `roles` 会造成角色冲突
- `style_transfer.py` 按段落逐一查找 role，没有「页面」概念；引入页面级 context 会大幅增加核心逻辑复杂度
- `special_pages` 字段天然容纳页面级元数据（置信度、识别依据、段落范围等）

### 2. 识别策略：草稿重新识别（非 range 静态记录）

Profile 只存模板的格式信息，不存段落 range。`style_transfer.py` 执行时在**草稿文档**上重新运行 `PageClassifier`，实时计算段落归属。

**理由**：草稿与模板段落数不同，将模板的 `para_range` 直接用于草稿会完全错位。

### 3. 扩展性：注册表模式

每种页面类型为独立的 `PageDetector` 子类，集中注册到 `DETECTORS` 列表。新增类型只需新建子类并追加到列表，不修改核心调度逻辑。

### 4. 触发方式：双入口

- `template_analyzer.py`：默认集成，提取模板格式时顺带识别，结果写入 `special_pages`
- `page_classifier.py`：独立脚本，可对任意文档（包括草稿）单独运行

---

## 架构设计

### 新增/修改文件

```
word-master/skills/word-master/scripts/
├── page_classifier.py        # [NEW] PageDetector 注册表 + PageClassifier + CLI
├── template_analyzer.py      # [MODIFY] 集成 PageClassifier，输出 special_pages
└── style_transfer.py         # [MODIFY] 执行前对草稿运行 PageClassifier，精确跳过
```

### 数据流

```
                    ┌─────────────────────────┐
  template.docx ──▶│  template_analyzer.py   │──▶ style_profile.json
                    │  (roles + table_roles   │    含 special_pages[格式信息]
                    │   + PageClassifier)     │
                    └─────────────────────────┘

                    ┌─────────────────────────┐
  draft.docx    ──▶│  PageClassifier         │──▶ {cover:[0~6], signature:[44~50]}
                    │  (草稿重新识别)          │    (段落 range 映射)
                    └─────────────────────────┘
                              │
                    ┌─────────▼───────────────┐
  style_profile ──▶│  style_transfer.py      │──▶ output.docx
  (skip_style_     │  跳过 special_pages 中   │
   transfer=true)  │  skip=true 的段落        │
                    └─────────────────────────┘
```

---

## PageDetector 接口

### 基类

```python
class PageDetector:
    page_type: str              # 唯一类型标识，如 "cover"
    search_zone: str            # "head" / "tail" / "any"
    keywords: list[str]         # 字面量关键词列表（in 匹配）
    keyword_patterns: list[str] # 正则关键词列表（re.search 匹配）；默认空列表
    min_score: float = 0.5      # 实例级置信度阈值；可被 CLI --min-score 全局覆盖

    def score(self, paras: list, page_index: int, total_pages: int) -> float:
        """
        计算这批段落属于该页面类型的置信度（0~1）。
        page_index: 当前页组在文档所有页组中的从 0 开始的索引
        total_pages: 文档总页组数
        默认实现：position_score × 0.4 + keyword_score × 0.6
        子类可 override 实现特殊逻辑。
        """

    def extract_format(self, paras: list, theme_fonts: dict) -> list[dict]:
        """
        提取该页面内各段落的格式摘要，写入 special_pages.paragraphs。
        默认实现：对每段调用 compute_effective_fingerprint()，
        并按下方「role 名推断规则」推断 role 名。
        """
```

### 分页切割策略

`PageClassifier.classify()` 在打分前先将文档段落列表切割为「页组」列表。**所有检测器共用同一份切割结果**，`page_index`／`total_pages` 在全局语义一致。

**切割规则（按优先级）**：

1. **显式分页符**：段落中任意 run 含 `<w:br w:type="page"/>` 节点，则在该段落**之后**切断（该段落属于当前页组，下一段落开始新页组）。
2. **`pageBreakBefore`**：段落 pPr 含 `<w:pageBreakBefore/>`，则在该段落**之前**切断（该段落为新页组的第一个段落）。
3. **分节符**：段落 pPr 含 `<w:sectPr>`（非文档末尾的 sectPr），同视为页面边界，在该段落**之后**切断。
4. **回退策略**（文档无任何分页符和分节符时）：以每 **15 段**为一个虚拟页组切割全文，产生统一的虚拟页组列表。所有检测器共用这一列表，`search_zone` 属性仅影响 `position_score` 的计算公式选择，不影响页组列表本身。

切割结果为有序的 `List[List[Paragraph]]`，`page_index` 即在此列表中的下标，`total_pages` 即列表长度。

### 默认打分逻辑

```
position_score（利用 page_index 和 total_pages 计算）:
  search_zone == "head" → 1.0 - page_index / total_pages
  search_zone == "tail" → page_index / (total_pages - 1) if total_pages > 1 else 1.0
  search_zone == "any"  → 0.5（不加分也不减分）

keyword_score（字面量 + 正则两类分别匹配，仅分母共用去重后的总数）:
  分母：len(list(dict.fromkeys(keywords + keyword_patterns)))（去重后的全部关键词总数）
  分子：分别遍历 keywords（in 匹配）和 keyword_patterns（re.search 匹配）统计命中数
        同一关键词同时出现在两表且均命中时，去重后仅计一次
  keyword_score = 命中的不重复关键词数 / 分母
  匹配范围：页组内所有段落文本拼接后一起搜索
  若 all_keywords 为空（如 back_cover），keyword_score = 0，final_score 仅靠 position_score

final_score = position_score × 0.4 + keyword_score × 0.6
（若 all_keywords 为空，权重重分配：position_score × 1.0）
```

### `min_score` 优先级规则

CLI `--min-score` / `--special-page-min-score` 参数为**全局覆盖**：当该参数存在时，忽略所有检测器实例的 `min_score` 字段，统一使用 CLI 传入值。未传入时各检测器使用自身 `min_score`（默认 0.5）。

### 注册表（7 种初始类型）

| 类型 | `page_type` | `search_zone` | `keywords`（字面量） | `keyword_patterns`（正则） |
|------|-------------|---------------|---------------------|---------------------------|
| 封面 | `cover` | `head` | 项目、报告、方案、封面、编制 | — |
| 封底 | `back_cover` | `tail` | —（位置主导） | — |
| 目录页 | `table_of_contents` | `head` | 目录 | `第.{1,4}章`、`\.{3,}` |
| 签名页 | `signature` | `tail` | 签字、签章、盖章、经办人 | — |
| 承诺函 | `commitment_letter` | `tail` | 承诺、保证、声明、郑重 | — |
| 版权声明 | `copyright` | `head` | 版权、著作权、保留 | `©` |
| 附录封面 | `appendix_cover` | `any` | 附录、Appendix | — |

### 新增类型示例

```python
class MyNewPageDetector(PageDetector):
    page_type = "my_new_type"
    search_zone = "tail"
    keywords = ["关键词A", "关键词B"]
    keyword_patterns = []   # 无正则关键词时可省略（基类默认空列表）
    # 默认打分和格式提取逻辑自动继承，无需 override

DETECTORS.append(MyNewPageDetector())
```

---

## style_profile.json Schema 变更

### 新增字段结构

```json
{
  "roles": [...],
  "table_roles": [...],
  "special_pages": [
    {
      "type": "cover",
      "confidence": 0.92,
      "detected_by": ["position", "keyword"],
      "skip_style_transfer": true,
      "para_count": 7,
      "paragraphs": [
        {
          "role": "Cover Title",
          "example": "XX项目可行性研究报告",
          "fingerprint": {
            "size": "22.0pt",
            "font": "黑体",
            "bold": true,
            "align": "center",
            "color": null
          }
        },
        {
          "role": "Cover Subtitle",
          "example": "2024年3月",
          "fingerprint": {
            "size": "14.0pt",
            "font": "宋体",
            "bold": false,
            "align": "center",
            "color": null
          }
        }
      ]
    }
  ]
}
```

### 字段说明

| 字段 | 类型 | 说明 |
|------|------|------|
| `type` | string | 页面类型标识 |
| `confidence` | float | 识别置信度 0~1 |
| `detected_by` | list[string] | `"position"` / `"keyword"` 的组合；当无关键词时（如 `back_cover`），仅含 `["position"]` |
| `skip_style_transfer` | bool | 默认 `true`；改为 `false` 可启用精确迁移（未来功能） |
| `para_count` | int | 该页面包含的段落数，便于人工核查 |
| `para_range` | list[int, int] | 段落范围，**闭区间** `[first_para_index, last_para_index]`（两端均包含） |
| `paragraphs` | list | 页面内各段落格式摘要 |

### paragraphs 内 role 命名规范与推断规则

`extract_format()` 默认实现按如下规则推断 role 名（以字体大小为主要依据）：

| 字体大小 | 推断 role（封面类） | 推断 role（签名类） | 推断 role（目录类） |
|---------|-------------------|-------------------|-------------------|
| ≥ 18pt | `Cover Title` | `Signature Title` | `TOC Heading` |
| 12–17pt | `Cover Subtitle` | `Signature Name` | `TOC Heading` |
| < 12pt | `Cover Date` / `Cover Org` | `Signature Date` | `TOC Entry` |
| 无法判断 | `Cover Body` | `Signature Body` | `TOC Body` |

其余 4 种页面类型（`back_cover`、`commitment_letter`、`copyright`、`appendix_cover`）统一按封面类规则命名，将 role 名前缀替换为页面类型标题：

| 页面类型 | 大字 ≥ 18pt | 中字 12–17pt | 小字 < 12pt | 其他 |
|---------|------------|------------|-----------|------|
| `back_cover` | `BackCover Title` | `BackCover Subtitle` | `BackCover Body` | `BackCover Body` |
| `commitment_letter` | `Commitment Title` | `Commitment Body` | `Commitment Footer` | `Commitment Body` |
| `copyright` | `Copyright Title` | `Copyright Body` | `Copyright Footer` | `Copyright Body` |
| `appendix_cover` | `Appendix Title` | `Appendix Subtitle` | `Appendix Body` | `Appendix Body` |

> 封底等页面类型的小字段落（< 12pt）统一归入 Body，不做 Date/Org 细分，因为这些页面通常不需要日期/机构字段的精确迁移语义。

同一页面内按字体大小从大到小依次分配，相同大小的段落归入同一 role（取文档中第一个出现的为代表示例）。

这些 role 名**不进入 `validate_style_profile.py` 白名单校验**，仅供格式参考。

---

## CLI 接口

### `page_classifier.py`（独立脚本）

```bash
# 对任意文档做页面分类，输出到标准输出
python page_classifier.py draft.docx

# 输出到文件
python page_classifier.py draft.docx --output page_map.json

# 只检测指定类型
python page_classifier.py draft.docx --types cover signature

# 调整置信度阈值（默认 0.5）
python page_classifier.py draft.docx --min-score 0.3
```

输出示例：
```json
{
  "cover":     { "para_range": [0, 6],   "confidence": 0.92 },
  "signature": { "para_range": [44, 51], "confidence": 0.85 }
}
```

### `template_analyzer.py`（新增参数）

```bash
# 默认行为：自动运行页面识别，写入 special_pages
python template_analyzer.py template.docx --output style_profile.json

# 关闭特殊页面识别
python template_analyzer.py template.docx --no-special-pages

# 调整置信度阈值
python template_analyzer.py template.docx --special-page-min-score 0.4
```

### `style_transfer.py`（新增参数）

```bash
# 默认行为：读取 profile 中 special_pages，对草稿重新识别并跳过
python style_transfer.py --profile style_profile.json draft.docx output.docx

# 不做特殊页面跳过（恢复旧行为）
python style_transfer.py --profile style_profile.json draft.docx output.docx --no-special-pages
```

---

## 向后兼容保证

| 场景 | 行为 |
|------|------|
| 旧 profile（无 `special_pages`）+ 新 `style_transfer.py` | 自动降级，等同旧行为 |
| `--skip-head N` 与新逻辑并存 | 两者取**段落索引集合的并集**后跳过（即 `skip_set = set(range(N)) | special_page_indices`） |
| 草稿识别置信度低于阈值 | 不跳过该页面，输出 WARN 日志 |
| `--no-special-pages` 显式关闭 | 完全等同旧行为 |

---

## 不在本次范围内

- 精确迁移（`skip_style_transfer: false` 的实际应用逻辑）：本次只提取格式，不实现跨页面样式映射
- 页面类型的用户自定义关键词配置（CLI 参数方式）：留待后续需求
- 对 PDF 或其他格式的支持

---

## 涉及文件

| 文件 | 操作 | 说明 |
|------|------|------|
| `scripts/page_classifier.py` | **NEW** | `PageDetector` 基类 + 7种检测器 + `PageClassifier` + CLI |
| `scripts/template_analyzer.py` | MODIFY | 集成 `PageClassifier`，输出 `special_pages` 到 profile |
| `scripts/style_transfer.py` | MODIFY | 执行前对草稿运行 `PageClassifier`，精确跳过特殊页面段落 |
| `scripts/validate_style_profile.py` | MODIFY | 新增对 `special_pages` 字段的基础结构校验（type/confidence/skip_style_transfer/paragraphs 的类型检查） |
| `tests/test_page_classifier.py` | **NEW** | 各 Detector 打分逻辑 + 分页切割 + 端到端识别测试 |
| `tests/fixtures/test_special_pages_template.docx` | **NEW** | 含封面+正文+签名页的测试固件文档，由测试 setUp 动态生成（python-docx 代码创建，不需要手工制作） |

---

## 验证标准

1. **单元测试**：`CoverPageDetector.score(paras, page_index=0, total_pages=10)` 含"报告"关键词时返回 ≥ 0.7
2. **单元测试**：`CoverPageDetector.score(paras, page_index=9, total_pages=10)` 含"报告"关键词时返回 < 0.5（位置惩罚）
3. **单元测试**：`BackCoverDetector.score(paras=[], page_index=9, total_pages=10)` 返回 ≥ 0.5（仅位置，无关键词时权重重分配为 1.0）
4. **单元测试**：分页切割函数正确在 `<w:br w:type="page"/>` 处切断，切割后页组数 = 分页符数 + 1
5. **单元测试**：文档无分页符、共 30 段时，回退切割产生 2 个虚拟页组（各 15 段），`total_pages=2`，**所有检测器共用此结果**
6. **单元测试**：新增一个最小化子类（只写 `page_type`/`search_zone`/`keywords`），`PageClassifier` 能自动调度且 `keyword_patterns` 默认为空列表
7. **单元测试**：CLI `--min-score 0.3` 传入时，所有检测器的实例 `min_score=0.5` 被忽略，统一使用 0.3
8. **集成测试**：含封面+正文+签名页的测试文档（test_special_pages_template.docx），`page_classifier.py` 正确识别 cover 和 signature 的段落范围
9. **集成测试**：旧版 profile（无 `special_pages`）传入新版 `style_transfer.py`，输出与旧版完全一致
10. **集成测试**：`template_analyzer.py` 对测试模板输出的 profile 包含结构合法的 `special_pages` 字段，`validate_style_profile.py` 通过校验
