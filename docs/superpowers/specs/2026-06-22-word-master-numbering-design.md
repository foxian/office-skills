# word-master 标题编号功能设计

- 日期：2026-06-22
- 范围：为 word-master 新增标题编号能力（add / remove），借鉴 markdown-master 的灵活模板语法
- 不在范围内：Word 原生 list numbering（numPr XML）、heading 升降级、toc 生成、修改 md 侧任何代码

## 背景

markdown-master 有完整的标题编号系统（`structure.py numbering add/remove` + `{N:d}/{N:R}/{N:cn}` 模板占位符 + YAML 配置 + `--start-from` + `--save-config`）。word-master 目前**完全没有**标题编号能力——`md_to_docx.py` 把 `#`/`##` 转成 Word 标题样式但不加编号前缀，`docx_editor.py` 的 DSL 也没有 numbering 操作。本设计为 word-master 补上这层能力。

## 架构定位

word-master 现有脚本职责：
- `md_to_docx.py`：单向转换器（md → docx，从零生成）
- `docx_editor.py`：已有 docx 的 JSON DSL 编辑器（段落级单点操作）
- `docx_reader.py`：只读快照
- `style_transfer.py`：样式套用
- 各 analyzer：只读分析

**word-master 缺一个"结构操作层"**——md 侧有 `structure.py`（heading 升降级、numbering、toc），docx 侧没有对应物。编号属于文档级结构操作，不是段落级编辑，塞进 `docx_editor.py` 的 DSL 会污染其单点操作抽象；只加到 `md_to_docx.py` 又处理不了用户手里已有的 docx。

**结论：新增 `word-master/scripts/docx_structure.py`**，与 md 侧 `structure.py` 对称，未来 heading 升降级、toc 也能挂这个脚本。

## 复用策略：解耦实现，契约对齐

**两个 skill 运行时完全独立**，不通过 sys.path 跨 skill import。编号渲染逻辑（模板解析、罗马数字、中文数字）在 docx 侧独立重写一份。

理由：
1. skill 自治优先——word-master 不应依赖 markdown-master 才能运行
2. 模板渲染逻辑是已收敛代码（占位符语法、罗马数字规则不会变），漂移风险可控
3. **YAML schema 是跨 skill 契约**，不是代码共享——两边读写同一份 YAML，配置文件可互通

代价：md 侧改渲染逻辑（如修罗马数字边界 bug），docx 侧不会自动跟随，需人工同步。spec 中明确记录两份 SKILL.md 的模板语法表必须保持一致。

## 文件结构

### 新增 `word-master/scripts/_numbering_render.py`（私有渲染模块）

与 md 侧 `structure.py` 的渲染函数行为对齐，但独立实现，不 import md 侧任何东西。

| 函数 | 职责 | md 侧对应 |
|------|------|----------|
| `parse_template(template: str) -> list` | 解析 `{N}` / `{N:R}` / `{N:cn}` 为 token 列表 | `parse_template` |
| `format_template(tokens, counters: list[int]) -> str` | 按 6 级计数器渲染前缀 | `format_template` |
| `_to_roman(n, lower=False) -> str` | 十进制→罗马数字（`4=IV`、`9=IX`） | `_to_roman` |
| `_to_letter(n, upper=True) -> str` | 十进制→字母 | `_to_letter` |
| `_to_chinese(n) -> str` | 十进制→中文数字（`一二三...九十`、`十一/十二`） | `_to_chinese` |
| `load_config(path: str) -> dict` | 读 YAML，返回 `{1..6: template, start_from: N}` | `load_config` |
| `save_config(path: str, cfg: dict) -> None` | 写 YAML | `save_config` |

每个函数配 docstring，注明"与 markdown-master/scripts/structure.py 中同名函数行为对齐"，方便将来人工对照同步。纯标准库 + pyyaml。

**为何独立成模块**：渲染层与 docx I/O 分离，`docx_structure.py` 专注"读 docx → 识别标题 → 调渲染 → 写前缀 → 写回"，渲染模块可独立单测（喂 token 和计数器，断言字符串），不需构造 docx。

### 新增 `word-master/scripts/docx_structure.py`（主脚本）

负责 docx I/O + 标题识别 + 前缀写入/剥除 + CLI。

## 标题识别

遍历 `doc.paragraphs`，按 `paragraph.style.name` 判定标题级别：

```python
import re
_HEADING_RE = re.compile(r"^Heading (\d+)$")        # 英文 "Heading 1"
_HEADING_CN_RE = re.compile(r"^标题\s*(\d+)$")      # 中文 "标题 1"

def _detect_heading_level(paragraph):
    name = paragraph.style.name if paragraph.style else ""
    for pat in (_HEADING_RE, _HEADING_CN_RE):
        m = pat.match(name)
        if m:
            level = int(m.group(1))
            if 1 <= level <= 9:
                return level
    return None
```

**与 md 侧根本差异**：
- md 侧靠 `^(#{1,6})\s+` 正则识别，级别 1-6
- docx 侧靠 `style.name` 识别，级别 1-9（Word 原生支持到 Heading 9）
- 模板只支持 h1-h6（与 md 侧对齐）；h7-h9 即使识别出来也不编号（对应模板为空）

**同时识别英文 `Heading N` 和中文 `标题 N`**：因 `style_transfer.py` / `md_to_docx.py` 已处理过中英文样式名，docx 文档两种都常见。

## 前缀写入（numbering add）

对每个识别出的标题段落：
1. 取 `paragraph.text`（python-docx 拼接所有 run 的文本）
2. 先剥除已有编号前缀（见下），得到干净标题文本
3. 用模板渲染出编号前缀（`format_template`）
4. 把"前缀 + 干净标题文本"写回 run

**写回 run 的策略（策略 B）**：所有文本变更只发生在第一个 run，其余 run 一律不动。

```
原段落 "附录A 测试方法" (Heading 1)
├─ run[0]: "附录"     (color=红色)
└─ run[1]: "A 测试方法" (color=默认)

add 编号 "第1章 " 后：
├─ run[0]: "第1章 附录"  (color=红色)  ← 只改 run[0] 的文本（前插前缀）
└─ run[1]: "A 测试方法"  (color=默认)  ← 原样不动
```

**关键约束（贯穿 add 和 remove）**：无论加前缀还是剥前缀，只修改 `run[0].text`，不触碰其余 run。具体：
- add 时：`run[0].text = prefix + (剥除后的 run[0].text)`，其中"剥除"只对 `run[0].text` 做正则替换，不动其他 run
- remove 时：`run[0].text = 剥除后的 run[0].text`，其余 run 不动
- 不做"跨 run 拼接后剥除再重新分配"——那会破坏 run 级格式

**取舍说明（写进 spec）**：
- 编号继承第一个 run 的格式（如上例编号是红色）——合理行为，标题段落整体样式由 Heading 样式决定，run 级覆盖是用户有意为之
- 所有 run 级格式保真——比"塞进第一个 run、清空后续"更优
- 边角情况：若旧编号前缀恰好跨 run 边界（如 run[0]="第1"、run[1]="章 引言"），策略 B 只剥 run[0] 里匹配到的部分，可能剥不干净——此情况罕见且 YAGNI，spec 记录不处理，用户可先 remove 再 add 规避
- 不处理"前缀需跨 run 边界拆分"的边角情况——YAGNI

## 前缀剥除（numbering remove 和 add 的前置步骤）

剥除已有编号前缀的正则，扩展自 md 侧 `files.py` 的 `_sanitize_filename`，并补充 docx 场景常见的中文数字章：

```python
_NUMBER_PREFIX_RE = re.compile(
    r"^("
    r"第[\d]+章\s*"               # 第1章
    r"|[\d]+[、，]\s*"             # 1、 / 1，
    r"|（[\d]+）\s*"               # （1）
    r"|[IVXLCDM]+\s+"              # IV（罗马大写）
    r"|[ivxlcdm]+\s+"              # iv（罗马小写）
    r"|[A-Z]\s+"                   # A
    r"|[a-z]\s+"                   # a
    r"|[\d]+(?:\.[\d]+)*\.?\s+"    # 1 / 1.1 / 1.1.1.
    r"|[零一二三四五六七八九十百千]+章\s*"      # 第一章 / 第三章 / 第十二章（中文数字+章，docx 侧独有）
    r")"
)
```

**剥除策略**：
- `numbering remove`：对所有标题段落剥除前缀，写回
- `numbering add`：先剥除旧前缀（避免重复编号"第1章 第1章 引言"），再渲染新前缀。与 md 侧 `numbering_add_flex` 开头先调 `numbering_remove` 语义一致

**与 md 侧差异**：md 侧 `files.py` 的 `_sanitize_filename` 不覆盖"中文数字章"（`第一章`），docx 侧补上——因中文编号风格（`第{1:cn}章`）在 docx 场景更常见。

## 计数器维护

与 md 侧 `numbering_add_flex` 完全一致：
- `counters = [0] * 6`
- 遇到 level N 标题：`counters[N-1] += 1`，`counters[N..5] = 0`
- `start_from` 仅影响 h1 起始：`counters[0] = start_from - 1`

## CLI 设计

与 md 侧 `structure.py` 镜像：

```bash
# 添加编号
python ${SKILL_DIR}/scripts/docx_structure.py <file.docx> numbering add \
    [--h1 T] [--h2 T] [--h3 T] [--h4 T] [--h5 T] [--h6 T] \
    [--config FILE] [--start-from N] [--save-config FILE] \
    [-o output.docx]

# 移除编号
python ${SKILL_DIR}/scripts/docx_structure.py <file.docx> numbering remove [-o output.docx]
```

**参数语义与 md 侧完全一致**：
- `--h1 .. --h6`：每级模板字符串，空串 `""` 表示该级不编号
- `--config FILE`：从 YAML 加载模板（与 `--hN` 可叠加，CLI 覆盖 config）
- `--start-from N`：h1 起始编号（默认 1）
- `--save-config FILE`：把当前生效的模板存成 YAML，供下次 `--config` 复用
- `-o output.docx`：输出到新文件；不传则覆盖原文件并生成 `.bak` 备份

### .bak 备份策略

与 `docx_editor.py` 现有约定一致（"Original file backed up as `document.docx.bak`"）：
- 覆盖原文件前，`shutil.copy2(file, file + ".bak")`
- `.bak` 只在"覆盖原文件"模式生成；传 `-o` 输出到新文件时不生成（原文件没动）
- `.bak` 已存在则直接覆盖（不做 `.bak2` 链式备份，与 `docx_editor.py` 一致）

### 默认行为

与 md 侧 `structure.py` 一致：**默认覆盖原文件**。docx 是二进制格式，改坏了不像 md 能肉眼看出，备份比 md 侧更有价值（md 侧覆盖不备份，docx 侧覆盖备份——这是与 md 侧的唯一 I/O 差异）。

### 与 md 侧的 CLI 差异

1. 无 `--by` 参数（md 侧 `split` 有，`numbering` 没有；docx 侧 `numbering` 也不需要）
2. `-o` 时 docx 侧额外生成 `.bak`（md 侧 `-o` 不备份，覆盖也不备份）

## YAML 配置格式（跨 skill 契约）

```yaml
h1: "第{1}章 "
h2: "{1}.{2} "
h3: "（{3}）"
h4: ""
h5: ""
h6: ""
start_from: 1
```

两边读写同一份 YAML，配置文件可互通。**这是格式契约，不是代码共享**。

## 模板语法表（与 md 侧 SKILL.md 镜像，须保持一致）

| 占位符 | 含义 | 示例 |
|---------|------|------|
| `{N}` / `{N:d}` | 十进制数字 | 3 |
| `{N:02d}` | 两位补零十进制 | 03 |
| `{N:R}` | 大写罗马数字 | III |
| `{N:r}` | 小写罗马数字 | iii |
| `{N:A}` | 大写字母 | C |
| `{N:a}` | 小写字母 | c |
| `{N:cn}` | 中文数字 | 三 |

N 范围 1-6，对应 H1-H6。空模板字符串表示该级不编号。

## SKILL.md 更新

在 word-master/SKILL.md 的 "Core Scripts" 表格加一行：
```
| `${SKILL_DIR}/scripts/docx_structure.py` | 标题编号管理（添加/移除），支持灵活模板 |
```

文档末尾新增一节 "## Heading Numbering"，包含：CLI 用法、模板语法表（直接照搬 md 侧）、YAML 格式、常见编号示例（技术文档风格、中文章节风格、学术风格、保存/加载配置）。两份 SKILL.md 的模板语法参考必须完全一致。

## 测试策略

### 新增测试文件

- `word-master/tests/test_numbering_render.py` — 渲染层单测
- `word-master/tests/test_docx_structure.py` — docx I/O 层单测

### test_numbering_render.py（渲染层，纯函数）

喂字符串和计数器列表，不需构造 docx。与 md 侧 `test_numbering.py` 渲染部分镜像，**断言两边行为一致**：

1. `test_parse_template_simple` — `"第{1}章 "` 解析出 token 列表
2. `test_format_template_decimal` — `counters=[1,2,0,0,0,0]` + `"{1}.{2} "` → `"1.2 "`
3. `test_format_template_roman_upper` — `"{1:R} "` → `"III "`
4. `test_format_template_roman_lower` — `"{1:r} "` → `"iii "`
5. `test_format_template_letter_upper` — `"{1:A} "` → `"A "`
6. `test_format_template_letter_lower` — `"{1:a} "` → `"a "`
7. `test_format_template_chinese` — `"{1:cn} "` → `"一 "`
8. `test_format_template_zero_padded` — `"{1:02d} "` → `"03 "`
9. `test_format_template_empty_level` — 某级模板为空串，该级不输出
10. `test_load_config` — 读 YAML，返回 `{1..6: template, start_from}`
11. `test_save_config` — 写 YAML，再读回来一致
12. `test_load_config_missing_start_from` — YAML 无 `start_from`，默认 1

### test_docx_structure.py（docx I/O 层）

**动态构造 docx**：用 `python-docx` 在 `tmp_path` 下临时造出符合测试需要的 docx，测完自动清理。不引入 fixture 二进制文件。示例：

```python
def test_add_basic_chinese_chapter(tmp_path):
    doc = docx.Document()
    doc.add_heading("引言", level=1)
    doc.add_heading("方法", level=1)
    docx_path = tmp_path / "test.docx"
    doc.save(str(docx_path))

    from docx_structure import cmd_numbering_add
    cmd_numbering_add(str(docx_path), {1: "第{1}章 "}, start_from=1)

    result = docx.Document(str(docx_path))
    assert result.paragraphs[0].text == "第1章 引言"
    assert result.paragraphs[1].text == "第2章 方法"
```

**标题识别**
1. `test_detect_heading_english` — `Heading 1` / `Heading 3` 识别为 level 1 / 3
2. `test_detect_heading_chinese` — `标题 1` / `标题 2` 识别
3. `test_detect_non_heading` — `Normal` 段落返回 None
4. `test_detect_heading_out_of_range` — `Heading 7` 识别为 level 7，但模板只支持 1-6，不编号

**numbering add**
5. `test_add_basic_chinese_chapter` — `第{1}章 ` / `{1}.{2} `，H1 文本变 `"第1章 原标题"`
6. `test_add_skips_h7_plus` — Heading 7 段落不编号
7. `test_add_start_from` — `--start-from 5`，第一个 H1 是 `"第5章 ..."`
8. `test_add_strips_existing_prefix` — 标题已有 `"第1章 ..."`，重新 add 不会重复
9. `test_add_multi_run_preserves_format` — 多 run 标题（run[0] 红色、run[1] 默认），add 后 run[0] 含前缀且仍红色，run[1] 文本和格式不变
10. `test_add_empty_template_skips_level` — `--h1 ""` 时 H1 不编号，H2 正常编号

**numbering remove**
11. `test_remove_strips_decimal_prefix` — `"1.1 标题"` → `"标题"`
12. `test_remove_strips_chinese_chapter` — `"第一章 引言"` → `"引言"`（docx 侧独有）
13. `test_remove_strips_roman` — `"III 标题"` → `"标题"`
14. `test_remove_no_prefix_unchanged` — 无前缀的标题文本不变

**I/O 与备份**
15. `test_overwrite_creates_bak` — 不传 `-o`，原文件被改、`.bak` 存在且内容是改前的
16. `test_output_flag_no_bak` — 传 `-o`，新文件有编号、原文件未动、无 `.bak`
17. `test_save_config_writes_yaml` — `--save-config cfg.yaml`，文件存在且可被 `load_config` 读回

**CLI 集成（轻量，2 个）**
18. `test_cli_add_via_argv` — `python docx_structure.py doc.docx numbering add --h1 "{1} "` 跑通
19. `test_cli_remove_via_argv` — `numbering remove` 跑通

### 不测的

- 不测 Word 原生 list numbering（numPr XML）——本设计用文本前缀方案
- 不测 `style_transfer.py` 与编号的交互——style_transfer 改样式指纹不改标题文本，两者正交
- 不做 md 侧 vs docx 侧的对照测试——解耦实现决定的不靠测试钉一致，靠 spec 契约 + 人肉同步

### 运行约定

- `cd word-master && python -m pytest tests/test_numbering_render.py tests/test_docx_structure.py -v`
- 与 word-master 现有测试风格一致：pytest + `tmp_path` + 函数级断言
- 不修改 conftest.py；若 word-master 无 conftest.py，测试文件自行 `sys.path.insert` 注入 scripts 目录（与 markdown-master conftest 思路一致，但各自独立）

## 非目标（明确排除）

- 不实现 Word 原生 numPr 自动编号（仅文本前缀方案）
- 不实现 heading 升降级、toc 生成（留给将来扩展 `docx_structure.py`）
- 不修改 md 侧任何代码
- 不引入跨 skill import 或共享模块
- 不做 md/docx 对照测试

## 验收标准

1. 新增 `_numbering_render.py` + `docx_structure.py` + 2 个测试文件后，`python -m pytest tests/` 全绿
2. 不破坏 word-master 现有测试
3. CLI 用法、模板语法、YAML 格式与 md 侧 `structure.py` 对称（两份 SKILL.md 模板语法表完全一致）
4. 渲染层测试覆盖 §test_numbering_render 全部 12 个用例
5. docx I/O 层测试覆盖 §test_docx_structure 全部 19 个用例
6. 多 run 标题写入采用策略 B（保留 run 级格式），测试用例 9 验证
7. `.bak` 备份在覆盖模式下生成、`-o` 模式不生成，测试用例 15/16 验证
