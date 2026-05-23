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
    page_type: str         # 唯一类型标识，如 "cover"
    search_zone: str       # "head" / "tail" / "any"
    search_limit: int      # 搜索区域的段落数上限（"any" 时忽略）
    keywords: list[str]    # 内容关键词列表
    min_score: float = 0.5 # 置信度阈值，低于此值不认为匹配

    def score(self, paras: list, zone: str) -> float:
        """
        计算这批段落属于该页面类型的置信度（0~1）。
        默认实现：位置分 × 0.4 + 关键词命中分 × 0.6
        子类可 override 实现特殊逻辑。
        """

    def extract_format(self, paras: list, theme_fonts: dict) -> list[dict]:
        """
        提取该页面内各段落的格式摘要，写入 special_pages.paragraphs。
        默认实现：对每段调用 compute_effective_fingerprint()，
        并根据字体大小推断 role 名（Cover Title / Cover Subtitle 等）。
        """
```

### 默认打分逻辑

```
position_score:
  search_zone == "head" → 页组越靠前分越高（线性衰减）
  search_zone == "tail" → 页组越靠后分越高
  search_zone == "any"  → position_score = 0.5（不加分也不减分）

keyword_score:
  命中关键词数 / 总关键词数（去重，按页组内所有段落文本搜索）

final_score = position_score × 0.4 + keyword_score × 0.6
```

### 注册表（7 种初始类型）

| 类型 | `page_type` | `search_zone` | 核心关键词 |
|------|-------------|---------------|-----------|
| 封面 | `cover` | `head` | 项目、报告、方案、封面、编制 |
| 封底 | `back_cover` | `tail` | —（位置主导，keywords 为空） |
| 目录页 | `table_of_contents` | `head` | 目录、第.*章、……… |
| 签名页 | `signature` | `tail` | 签字、签章、盖章、经办人 |
| 承诺函 | `commitment_letter` | `tail` | 承诺、保证、声明、郑重 |
| 版权声明 | `copyright` | `head` | 版权、著作权、©、保留 |
| 附录封面 | `appendix_cover` | `any` | 附录、Appendix |

### 新增类型示例

```python
class MyNewPageDetector(PageDetector):
    page_type = "my_new_type"
    search_zone = "tail"
    keywords = ["关键词A", "关键词B"]
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
| `detected_by` | list[string] | `"position"` / `"keyword"` 的组合 |
| `skip_style_transfer` | bool | 默认 `true`；改为 `false` 可启用精确迁移（未来功能） |
| `para_count` | int | 该页面包含的段落数，便于人工核查 |
| `paragraphs` | list | 页面内各段落格式摘要 |

### paragraphs 内 role 命名规范

封面类：`Cover Title`、`Cover Subtitle`、`Cover Date`、`Cover Org`  
签名类：`Signature Name`、`Signature Title`、`Signature Date`  
目录类：`TOC Heading`、`TOC Entry`  
（这些 role 名不进入 `validate_style_profile.py` 白名单校验，仅供格式参考）

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
| `--skip-head N` 与新逻辑并存 | 两者取并集跳过 |
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
| `tests/test_page_classifier.py` | **NEW** | 各 Detector 打分逻辑 + 分页切割 + 端到端识别测试 |

---

## 验证标准

1. **单元测试**：`CoverPageDetector.score()` 在文档头部含"报告"关键词时返回 ≥ 0.7
2. **单元测试**：`CoverPageDetector.score()` 在文档尾部含"报告"关键词时返回 < 0.5（位置惩罚）
3. **单元测试**：`BackCoverDetector.score()` 在尾部空白页返回 ≥ 0.5（仅位置，无关键词）
4. **单元测试**：新增一个最小化子类（只写 `page_type`/`search_zone`/`keywords`），`PageClassifier` 能自动调度
5. **集成测试**：含封面+正文+签名页的测试文档，`page_classifier.py` 正确识别 cover 和 signature 的段落范围
6. **集成测试**：旧版 profile（无 `special_pages`）传入新版 `style_transfer.py`，输出与旧版完全一致
7. **集成测试**：`template_analyzer.py` 对测试模板输出的 profile 包含结构合法的 `special_pages` 字段
