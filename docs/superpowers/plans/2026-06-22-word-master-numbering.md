# word-master 标题编号功能实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 为 word-master 新增 `docx_structure.py`（标题编号 add/remove）和 `_numbering_render.py`（渲染纯函数），与 markdown-master `structure.py` 的模板语法/YAML 格式/CLI 对称，共 31 个测试用例。

**Architecture:** 渲染层独立实现（不跨 skill import），`_numbering_render.py` 镜像 md 侧 `structure.py` 的 `parse_template`/`format_template`/`_to_roman`/`_to_letter`/`_to_chinese`/`load_config`/`save_config`；`docx_structure.py` 负责 docx I/O + 按 `style.name` 识别标题 + 策略 B 写前缀（只改 `run[0]`）+ `.bak` 备份 + argparse CLI。

**Tech Stack:** Python 3、python-docx（项目已用）、pyyaml（md 侧已用，word-master 需安装）、pytest。

**Spec:** `docs/superpowers/specs/2026-06-22-word-master-numbering-design.md`

---

## 文件结构

- **Create:** `word-master/scripts/_numbering_render.py` — 渲染纯函数模块（7 个函数）
- **Create:** `word-master/scripts/docx_structure.py` — 主脚本（标题识别 + 前缀写入/剥除 + I/O + CLI）
- **Create:** `word-master/tests/test_numbering_render.py` — 渲染层 12 个测试
- **Create:** `word-master/tests/test_docx_structure.py` — docx I/O 层 19 个测试
- **Modify:** `word-master/SKILL.md` — Core Scripts 表加一行 + 新增 "Heading Numbering" 节

**被测函数签名（本计划定义）：**
- `_numbering_render.parse_template(template: str) -> list` — token 列表
- `_numbering_render.format_template(tokens, counters: list[int]) -> str`
- `_numbering_render.load_config(path: str) -> dict` — `{1..6: str|None, "start_from": int}`
- `_numbering_render.save_config(path: str, cfg: dict) -> None`
- `docx_structure._detect_heading_level(paragraph) -> int | None`
- `docx_structure._strip_prefix(text: str) -> str`
- `docx_structure.cmd_numbering_add(docx_path, level_templates, start_from=1, output_path=None, save_config_path=None) -> None`
- `docx_structure.cmd_numbering_remove(docx_path, output_path=None) -> None`
- `docx_structure.main()` — argparse 入口

**测试路径注入约定**：word-master 无 conftest.py。新测试文件在文件头自行注入：
```python
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'scripts'))
```
注意：现有 `test_md_to_docx.py` 工作树版本用了 stale 路径 `..','skills','word-master','scripts`（该目录不存在），**新测试不要照搬**，用 `..','scripts`。

---

### Task 1: 创建 _numbering_render.py + test_numbering_render.py（渲染层 12 测试）

**Files:**
- Create: `word-master/scripts/_numbering_render.py`
- Create: `word-master/tests/test_numbering_render.py`

- [ ] **Step 1: 创建 test_numbering_render.py，写入 12 个渲染层测试**

```python
"""_numbering_render.py 的单元测试。

渲染逻辑与 markdown-master/scripts/structure.py 中同名函数行为对齐，
但独立实现、独立测试，不做跨 skill 对照测试。
"""
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'scripts'))

import pytest
from _numbering_render import (
    parse_template, format_template, load_config, save_config,
    _to_roman, _to_letter, _to_chinese,
)


def test_parse_template_simple():
    """简单模板解析为 text + num token 列表。"""
    tokens = parse_template("第{1}章 ")
    assert ("text", "第") in tokens
    assert ("num", 1, "d") in tokens
    assert ("text", "章 ") in tokens


def test_format_template_decimal():
    """十进制占位符 {N} 和 {N:d} 渲染。"""
    tokens = parse_template("{1}.{2} ")
    assert format_template(tokens, [1, 2, 0, 0, 0, 0]) == "1.2 "


def test_format_template_roman_upper():
    """大写罗马 {1:R}。"""
    tokens = parse_template("{1:R} ")
    assert format_template(tokens, [3, 0, 0, 0, 0, 0]) == "III "


def test_format_template_roman_lower():
    """小写罗马 {1:r}。"""
    tokens = parse_template("{1:r} ")
    assert format_template(tokens, [3, 0, 0, 0, 0, 0]) == "iii "


def test_format_template_letter_upper():
    """大写字母 {1:A}（1=A）。"""
    tokens = parse_template("{1:A} ")
    assert format_template(tokens, [1, 0, 0, 0, 0, 0]) == "A "


def test_format_template_letter_lower():
    """小写字母 {1:a}（1=a）。"""
    tokens = parse_template("{1:a} ")
    assert format_template(tokens, [1, 0, 0, 0, 0, 0]) == "a "


def test_format_template_chinese():
    """中文数字 {1:cn}（1=一）。"""
    tokens = parse_template("{1:cn} ")
    assert format_template(tokens, [1, 0, 0, 0, 0, 0]) == "一 "


def test_format_template_zero_padded():
    """两位补零 {1:02d}。"""
    tokens = parse_template("{1:02d} ")
    assert format_template(tokens, [3, 0, 0, 0, 0, 0]) == "03 "


def test_format_template_empty_level():
    """某级模板为空串，parse_template 返回 [("text","")]，format 输出空串。"""
    tokens = parse_template("")
    assert format_template(tokens, [1, 0, 0, 0, 0, 0]) == ""


def test_load_config(tmp_path):
    """读 YAML，返回 {1..6: template, start_from}。"""
    cfg_path = tmp_path / "cfg.yaml"
    cfg_path.write_text(
        'h1: "第{1}章 "\n'
        'h2: "{1}.{2} "\n'
        'h3: ""\n'
        'h4: ""\n'
        'h5: ""\n'
        'h6: ""\n'
        'start_from: 1\n',
        encoding="utf-8",
    )
    cfg = load_config(str(cfg_path))
    assert cfg[1] == "第{1}章 "
    assert cfg[2] == "{1}.{2} "
    assert cfg[3] == ""
    assert cfg["start_from"] == 1


def test_save_config(tmp_path):
    """写 YAML 再读回来一致。"""
    cfg_path = tmp_path / "out.yaml"
    cfg = {1: "第{1}章 ", 2: "{1}.{2} ", 3: "", 4: "", 5: "", 6: "", "start_from": 1}
    save_config(str(cfg_path), cfg)
    loaded = load_config(str(cfg_path))
    assert loaded[1] == "第{1}章 "
    assert loaded[2] == "{1}.{2} "
    assert loaded["start_from"] == 1


def test_load_config_missing_start_from(tmp_path):
    """YAML 无 start_from 字段，默认 1。"""
    cfg_path = tmp_path / "cfg.yaml"
    cfg_path.write_text(
        'h1: "{1} "\n'
        'h2: ""\n'
        'h3: ""\n'
        'h4: ""\n'
        'h5: ""\n'
        'h6: ""\n',
        encoding="utf-8",
    )
    cfg = load_config(str(cfg_path))
    assert cfg["start_from"] == 1
```

- [ ] **Step 2: 运行测试，确认全部失败（模块不存在）**

Run: `cd D:/DevProjects/office-master/word-master && python -m pytest tests/test_numbering_render.py -v`
Expected: 12 errors (ModuleNotFoundError: No module named '_numbering_render')

- [ ] **Step 3: 创建 _numbering_render.py，实现全部 7 个函数**

```python
"""编号渲染纯函数模块。

与 markdown-master/scripts/structure.py 中同名函数行为对齐，独立实现，
不 import md 侧任何东西。纯标准库 + pyyaml。

供 docx_structure.py 调用：解析模板、渲染编号前缀、读写 YAML 配置。
"""
import re
import sys
from pathlib import Path

try:
    import yaml as _yaml
except ImportError:
    _yaml = None

_TEMPLATE_TOKEN_RE = re.compile(r"\{(\d+)(?::([^\}]+))?\}")


def parse_template(template):
    """
    解析模板字符串为 token 列表。

    模板示例: "第{1}章 " -> [("text", "第"), ("num", 1, "d"), ("text", "章 ")]

    token 类型:
        ("text", str)              - 字面文本
        ("num", int 1..6, str fmt) - 数字占位符，fmt 默认为 "d"

    fmt 修饰符:
        d           十进制
        0Nd         N 位补零十进制 (N>=1)
        R / r       大写/小写罗马
        A / a       大写/小写字母
        cn          中文数字

    Raises:
        ValueError: 模板语法错误
    """
    if not template:
        return [("text", "")]

    tokens = []
    pos = 0
    for m in _TEMPLATE_TOKEN_RE.finditer(template):
        if m.start() > pos:
            tokens.append(("text", template[pos:m.start()]))

        level = int(m.group(1))
        fmt_raw = m.group(2) or "d"

        if level < 1 or level > 6:
            raise ValueError("级别必须在 1-6 之间: " + m.group(0))

        fmt = _validate_format(fmt_raw, m.group(0))
        tokens.append(("num", level, fmt))
        pos = m.end()

    if pos < len(template):
        tokens.append(("text", template[pos:]))

    if not tokens:
        return [("text", template)]

    return tokens


def _validate_format(fmt_raw, source):
    """校验并规范化修饰符。"""
    if fmt_raw in ("d", "R", "r", "A", "a", "cn"):
        return fmt_raw
    m = re.match(r"^0(\d+)d$", fmt_raw)
    if m:
        width = int(m.group(1))
        if width < 1:
            raise ValueError("宽度必须为正整数: " + source)
        return fmt_raw
    raise ValueError("未知修饰符: " + source)


_ROMAN_VALS = [(1000, 'M'), (900, 'CM'), (500, 'D'), (400, 'CD'),
                (100, 'C'), (90, 'XC'), (50, 'L'), (40, 'XL'),
                (10, 'X'), (9, 'IX'), (5, 'V'), (4, 'IV'), (1, 'I')]
_ROMAN_LOWER_VALS = [(v, s.lower()) for v, s in _ROMAN_VALS]


def _to_roman(n, lower=False):
    """转换为罗马数字（1..3999 范围，超过则重复 M）。"""
    if n < 1:
        return "0"
    vals = _ROMAN_LOWER_VALS if lower else _ROMAN_VALS
    result = []
    for v, s in vals:
        while n >= v:
            result.append(s)
            n -= v
    if n > 0:
        result.append("m" if lower else "M")
    return "".join(result)


def _to_letter(n, upper=True):
    """转换为字母 (1=A..Z=26, 27=AA..)。"""
    if n < 1:
        return ""
    base = 65 if upper else 97
    result = []
    while n > 0:
        n -= 1
        result.append(chr(base + n % 26))
        n //= 26
    return "".join(reversed(result))


_CN_DIGITS = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"]
_CN_UNITS = ["", "十", "百", "千"]


def _to_chinese(n):
    """转换为中文数字 (0..999 范围)。"""
    if n == 0:
        return "零"
    if n < 0:
        return "负" + _to_chinese(-n)
    if n < 10:
        return _CN_DIGITS[n]
    if n < 20:
        return "十" + (_CN_DIGITS[n - 10] if n > 10 else "")
    parts = []
    s = str(n)
    for i, ch in enumerate(s):
        d = int(ch)
        unit_pos = len(s) - i - 1
        if d == 0:
            if parts and parts[-1] != "零":
                parts.append("零")
        else:
            parts.append(_CN_DIGITS[d] + _CN_UNITS[unit_pos])
    result = "".join(parts)
    if result.endswith("零"):
        result = result[:-1]
    return result


def format_template(tokens, counters):
    """
    根据 token 列表和当前计数器值渲染模板。
    counters: list[int], 长度 6，counters[i] = 第 i+1 级的当前值。
    """
    parts = []
    for tok in tokens:
        if tok[0] == "text":
            parts.append(tok[1])
        elif tok[0] == "num":
            level = tok[1]
            fmt = tok[2]
            n = counters[level - 1]
            if fmt == "d":
                parts.append(str(n))
            elif fmt == "R":
                parts.append(_to_roman(n, lower=False))
            elif fmt == "r":
                parts.append(_to_roman(n, lower=True))
            elif fmt == "A":
                parts.append(_to_letter(n, upper=True))
            elif fmt == "a":
                parts.append(_to_letter(n, upper=False))
            elif fmt == "cn":
                if n > 999:
                    sys.stderr.write("[warning] 中文数字超出 999 范围: " + str(n) + "\n")
                    parts.append("###")
                else:
                    parts.append(_to_chinese(n))
            elif re.match(r"^0\d+d$", fmt):
                parts.append(format(n, fmt))
            else:
                parts.append(str(n))
    return "".join(parts)


def load_config(path):
    """
    从 YAML 文件加载编号配置。
    返回 dict: {1..6: str | None, "start_from": int}
    模板语法在加载时校验，错误抛出 ValueError。
    """
    if _yaml is None:
        raise ImportError("YAML 支持需要 pyyaml，请运行: pip install pyyaml")
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError("配置文件不存在: " + str(path))
    try:
        raw = _yaml.safe_load(p.read_text(encoding="utf-8"))
    except _yaml.YAMLError as e:
        raise ValueError("YAML 解析失败: " + str(e)) from e
    if not isinstance(raw, dict):
        raise ValueError("YAML 顶层必须是 mapping")

    cfg = {i: None for i in range(1, 7)}
    cfg["start_from"] = 1

    for i in range(1, 7):
        val = raw.get("h" + str(i))
        if val is None:
            cfg[i] = None
        elif isinstance(val, str):
            cfg[i] = val
        else:
            raise ValueError("h" + str(i) + " 必须是字符串")

    sf = raw.get("start_from", 1)
    if not isinstance(sf, int) or sf < 1:
        raise ValueError("start_from 必须是 >= 1 的整数")
    cfg["start_from"] = sf

    for i in range(1, 7):
        if cfg[i] is not None and cfg[i] != "":
            parse_template(cfg[i])

    return cfg


def save_config(path, cfg):
    """
    将配置写到 YAML 文件。
    字段顺序固定: h1..h6 -> start_from。None 规范化为 ""。
    """
    if _yaml is None:
        raise ImportError("YAML 支持需要 pyyaml，请运行: pip install pyyaml")
    out = {}
    for i in range(1, 7):
        v = cfg.get(i)
        out["h" + str(i)] = v if v else ""
    out["start_from"] = cfg.get("start_from", 1)

    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(
        _yaml.dump(out, allow_unicode=True, sort_keys=False, default_flow_style=False),
        encoding="utf-8",
    )
```

- [ ] **Step 4: 运行测试，确认 12 个全部通过**

Run: `cd D:/DevProjects/office-master/word-master && python -m pytest tests/test_numbering_render.py -v`
Expected: 12 passed

- [ ] **Step 5: 提交**

```bash
cd D:/DevProjects/office-master
git add word-master/scripts/_numbering_render.py word-master/tests/test_numbering_render.py
git commit -m "feat(word-master): add _numbering_render module with 12 tests"
```

---

### Task 2: docx_structure.py 标题识别 _detect_heading_level（测试 1-4）

**Files:**
- Create: `word-master/scripts/docx_structure.py`
- Create: `word-master/tests/test_docx_structure.py`

- [ ] **Step 1: 创建 test_docx_structure.py，写入标题识别的 4 个测试**

```python
"""docx_structure.py 的单元测试。

动态构造 docx（python-docx 在 tmp_path 下临时生成），不引入 fixture 二进制文件。
cmd_split/cmd_merge 直接 print 到 stdout，测试通过读产出文件断言。
"""
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'scripts'))

import docx
import pytest
from docx.enum.style import WD_STYLE_TYPE
from docx_structure import _detect_heading_level


def _make_doc_with_heading(text, level, style_name=None):
    """构造一个含单个标题段落的 docx Document（不落盘）。

    style_name 默认用英文 "Heading N"；传中文样式名时用 add_style 创建。
    """
    doc = docx.Document()
    if style_name is None:
        p = doc.add_heading(text, level=level)
    else:
        style = doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
        p = doc.add_paragraph(text, style=style)
    return doc, p


def test_detect_heading_english():
    """英文 Heading 1 / Heading 3 识别为 level 1 / 3。"""
    doc, p1 = _make_doc_with_heading("标题一", level=1)
    assert _detect_heading_level(p1) == 1
    doc, p3 = _make_doc_with_heading("标题三", level=3)
    assert _detect_heading_level(p3) == 3


def test_detect_heading_chinese():
    """中文 标题 1 / 标题 2 识别为 level 1 / 2。"""
    _doc, p1 = _make_doc_with_heading("中文标题一", level=1, style_name="标题 1")
    assert _detect_heading_level(p1) == 1
    _doc, p2 = _make_doc_with_heading("中文标题二", level=2, style_name="标题 2")
    assert _detect_heading_level(p2) == 2


def test_detect_non_heading():
    """Normal 段落返回 None。"""
    doc = docx.Document()
    p = doc.add_paragraph("正文段落")
    assert _detect_heading_level(p) is None


def test_detect_heading_out_of_range():
    """Heading 7 识别为 level 7（真实级别），但模板只支持 1-6，调用方不编号。"""
    _doc, p7 = _make_doc_with_heading("第七级", level=7)
    assert _detect_heading_level(p7) == 7
```

- [ ] **Step 2: 运行测试，确认失败（docx_structure 不存在）**

Run: `cd D:/DevProjects/office-master/word-master && python -m pytest tests/test_docx_structure.py -v`
Expected: 4 errors (ModuleNotFoundError: No module named 'docx_structure')

- [ ] **Step 3: 创建 docx_structure.py，实现 _detect_heading_level + 文件头**

```python
"""
DOCX 标题编号管理：添加/移除编号前缀。

与 markdown-master/scripts/structure.py 的 numbering add/remove 对称，
但针对 docx 按 style.name 识别标题（而非 md 的 # 前缀），
编号写为文本前缀（非 Word 原生 numPr）。
"""
import re
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from _numbering_render import (
    parse_template, format_template, load_config, save_config,
)

import docx

_HEADING_RE = re.compile(r"^Heading (\d+)$")
_HEADING_CN_RE = re.compile(r"^标题\s*(\d+)$")


def _detect_heading_level(paragraph):
    """按 paragraph.style.name 判定标题级别（1-9），非标题返回 None。"""
    name = paragraph.style.name if paragraph.style else ""
    for pat in (_HEADING_RE, _HEADING_CN_RE):
        m = pat.match(name)
        if m:
            level = int(m.group(1))
            if 1 <= level <= 9:
                return level
    return None
```

- [ ] **Step 4: 运行测试，确认 4 个通过**

Run: `cd D:/DevProjects/office-master/word-master && python -m pytest tests/test_docx_structure.py -v`
Expected: 4 passed

- [ ] **Step 5: 提交**

```bash
cd D:/DevProjects/office-master
git add word-master/scripts/docx_structure.py word-master/tests/test_docx_structure.py
git commit -m "feat(word-master): add _detect_heading_level with 4 tests"
```

---

### Task 3: docx_structure.py 前缀剥除 + numbering add + I/O（测试 5-10, 15-17）

**Files:**
- Modify: `word-master/scripts/docx_structure.py`
- Modify: `word-master/tests/test_docx_structure.py`

- [ ] **Step 1: 在 test_docx_structure.py 末尾追加 9 个测试（add 6 个 + I/O 3 个）**

```python
from docx.shared import RGBColor
from docx_structure import cmd_numbering_add


def _save_doc(tmp_path, doc, name="test.docx"):
    """保存 docx 到 tmp_path 并返回路径。"""
    p = tmp_path / name
    doc.save(str(p))
    return str(p)


def test_add_basic_chinese_chapter(tmp_path):
    """第{1}章 / {1}.{2} 模板，H1/H2 文本加编号前缀。"""
    doc = docx.Document()
    doc.add_heading("引言", level=1)
    doc.add_heading("背景", level=2)
    doc.add_heading("方法", level=1)
    docx_path = _save_doc(tmp_path, doc)

    cmd_numbering_add(docx_path, {1: "第{1}章 ", 2: "{1}.{2} "}, start_from=1)

    result = docx.Document(docx_path)
    assert result.paragraphs[0].text == "第1章 引言"
    assert result.paragraphs[1].text == "1.1 背景"
    assert result.paragraphs[2].text == "第2章 方法"


def test_add_skips_h7_plus(tmp_path):
    """Heading 7 段落不编号（模板只支持 1-6）。"""
    doc = docx.Document()
    doc.add_heading("H7 标题", level=7)
    doc.add_heading("H1 标题", level=1)
    docx_path = _save_doc(tmp_path, doc)

    cmd_numbering_add(docx_path, {1: "{1} "}, start_from=1)

    result = docx.Document(docx_path)
    assert result.paragraphs[0].text == "H7 标题"
    assert result.paragraphs[1].text == "1 H1 标题"


def test_add_start_from(tmp_path):
    """--start-from 5，第一个 H1 是 第5章。"""
    doc = docx.Document()
    doc.add_heading("引言", level=1)
    docx_path = _save_doc(tmp_path, doc)

    cmd_numbering_add(docx_path, {1: "第{1}章 "}, start_from=5)

    result = docx.Document(docx_path)
    assert result.paragraphs[0].text == "第5章 引言"


def test_add_strips_existing_prefix(tmp_path):
    """标题已有 第1章 前缀，重新 add 不会变成 第1章 第1章。"""
    doc = docx.Document()
    doc.add_heading("第1章 引言", level=1)
    docx_path = _save_doc(tmp_path, doc)

    cmd_numbering_add(docx_path, {1: "第{1}章 "}, start_from=1)

    result = docx.Document(docx_path)
    assert result.paragraphs[0].text == "第1章 引言"


def test_add_multi_run_preserves_format(tmp_path):
    """多 run 标题：编号写进 run[0]，run[1] 文本和格式不变。"""
    doc = docx.Document()
    p = doc.add_heading(level=1)
    run0 = p.add_run("附录")
    run0.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
    p.add_run("A 测试方法")
    docx_path = _save_doc(tmp_path, doc)

    cmd_numbering_add(docx_path, {1: "第{1}章 "}, start_from=1)

    result = docx.Document(docx_path)
    rp = result.paragraphs[0]
    assert rp.runs[0].text == "第1章 附录"
    assert rp.runs[0].font.color.rgb == RGBColor(0xFF, 0x00, 0x00)
    assert rp.runs[1].text == "A 测试方法"


def test_add_empty_template_skips_level(tmp_path):
    """--h1 为空串时 H1 不编号，H2 正常编号。"""
    doc = docx.Document()
    doc.add_heading("引言", level=1)
    doc.add_heading("背景", level=2)
    docx_path = _save_doc(tmp_path, doc)

    cmd_numbering_add(docx_path, {1: "", 2: "{1}.{2} "}, start_from=1)

    result = docx.Document(docx_path)
    assert result.paragraphs[0].text == "引言"
    assert result.paragraphs[1].text == "0.1 背景"


def test_overwrite_creates_bak(tmp_path):
    """不传 output_path，覆盖原文件并生成 .bak（.bak 内容是改前的）。"""
    doc = docx.Document()
    doc.add_heading("引言", level=1)
    docx_path = _save_doc(tmp_path, doc)
    original_bytes = Path(docx_path).read_bytes()

    cmd_numbering_add(docx_path, {1: "第{1}章 "}, start_from=1)

    bak_path = docx_path + ".bak"
    assert Path(bak_path).exists()
    assert Path(bak_path).read_bytes() == original_bytes
    result = docx.Document(docx_path)
    assert result.paragraphs[0].text == "第1章 引言"


def test_output_flag_no_bak(tmp_path):
    """传 output_path，新文件有编号、原文件未动、无 .bak。"""
    doc = docx.Document()
    doc.add_heading("引言", level=1)
    docx_path = _save_doc(tmp_path, doc)
    original_text = docx.Document(docx_path).paragraphs[0].text
    out_path = str(tmp_path / "out.docx")

    cmd_numbering_add(docx_path, {1: "第{1}章 "}, start_from=1, output_path=out_path)

    assert not Path(docx_path + ".bak").exists()
    assert docx.Document(docx_path).paragraphs[0].text == original_text
    assert docx.Document(out_path).paragraphs[0].text == "第1章 引言"


def test_save_config_writes_yaml(tmp_path):
    """--save-config 写出的 YAML 可被 load_config 读回。"""
    from _numbering_render import load_config as _load

    doc = docx.Document()
    doc.add_heading("引言", level=1)
    docx_path = _save_doc(tmp_path, doc)
    cfg_path = str(tmp_path / "cfg.yaml")

    cmd_numbering_add(
        docx_path, {1: "第{1}章 ", 2: "{1}.{2} "}, start_from=1,
        save_config_path=cfg_path,
    )

    loaded = _load(cfg_path)
    assert loaded[1] == "第{1}章 "
    assert loaded[2] == "{1}.{2} "
    assert loaded["start_from"] == 1
```

- [ ] **Step 2: 运行测试，确认 9 个失败（cmd_numbering_add 不存在）**

Run: `cd D:/DevProjects/office-master/word-master && python -m pytest tests/test_docx_structure.py -v`
Expected: 4 passed (Task 2 的), 9 errors (cannot import name 'cmd_numbering_add')

- [ ] **Step 3: 在 docx_structure.py 末尾追加 _strip_prefix / _save_docx / cmd_numbering_add**

```python
_NUMBER_PREFIX_RE = re.compile(
    r"^("
    r"第[\d]+章\s*"
    r"|[\d]+[、，]\s*"
    r"|（[\d]+）\s*"
    r"|[IVXLCDM]+\s+"
    r"|[ivxlcdm]+\s+"
    r"|[A-Z]\s+"
    r"|[a-z]\s+"
    r"|[\d]+(?:\.[\d]+)*\.?\s+"
    r"|[零一二三四五六七八九十百千]+章\s*"
    r")"
)


def _strip_prefix(text):
    """剥除标题文本开头的编号前缀（只剥一次）。"""
    return _NUMBER_PREFIX_RE.sub("", text, count=1)


def _save_docx(doc, docx_path, output_path):
    """保存 docx。output_path 给定则写新文件；否则覆盖原文件并生成 .bak。"""
    if output_path:
        doc.save(output_path)
    else:
        shutil.copy2(docx_path, docx_path + ".bak")
        doc.save(docx_path)


def cmd_numbering_add(docx_path, level_templates, start_from=1,
                      output_path=None, save_config_path=None):
    """
    给 docx 的标题段落添加编号前缀。

    level_templates: dict[1..6] = str | None，空串/None 表示该级不编号。
    start_from: h1 起始编号。
    output_path: 给定则输出到新文件；None 则覆盖原文件并生成 .bak。
    save_config_path: 给定则把当前模板存成 YAML（同时仍执行编号）。
    """
    doc = docx.Document(docx_path)
    counters = [0] * 6
    if level_templates.get(1):
        counters[0] = start_from - 1

    for para in doc.paragraphs:
        level = _detect_heading_level(para)
        if level is None or level > 6:
            continue
        template = level_templates.get(level)
        if not template:
            continue
        counters[level - 1] += 1
        for j in range(level, 6):
            counters[j] = 0

        tokens = parse_template(template)
        prefix = format_template(tokens, counters)

        if not para.runs:
            para.add_run(prefix)
        else:
            stripped = _strip_prefix(para.runs[0].text)
            para.runs[0].text = prefix + stripped

    _save_docx(doc, docx_path, output_path)

    if save_config_path:
        cfg = dict(level_templates)
        cfg["start_from"] = start_from
        save_config(save_config_path, cfg)
```

- [ ] **Step 4: 运行全部 13 个测试，确认通过**

Run: `cd D:/DevProjects/office-master/word-master && python -m pytest tests/test_docx_structure.py -v`
Expected: 13 passed

- [ ] **Step 5: 提交**

```bash
cd D:/DevProjects/office-master
git add word-master/scripts/docx_structure.py word-master/tests/test_docx_structure.py
git commit -m "feat(word-master): add cmd_numbering_add with strip + I/O + 9 tests"
```

---

### Task 4: docx_structure.py numbering remove（测试 11-14）

**Files:**
- Modify: `word-master/scripts/docx_structure.py`
- Modify: `word-master/tests/test_docx_structure.py`

- [ ] **Step 1: 在 test_docx_structure.py 末尾追加 4 个 remove 测试**

```python
from docx_structure import cmd_numbering_remove


def test_remove_strips_decimal_prefix(tmp_path):
    """1.1 标题 → 标题。"""
    doc = docx.Document()
    doc.add_heading("1.1 标题", level=2)
    docx_path = _save_doc(tmp_path, doc)

    cmd_numbering_remove(docx_path)

    assert docx.Document(docx_path).paragraphs[0].text == "标题"


def test_remove_strips_chinese_chapter(tmp_path):
    """第一章 引言 → 引言（docx 侧独有，md 侧不覆盖）。"""
    doc = docx.Document()
    doc.add_heading("第一章 引言", level=1)
    docx_path = _save_doc(tmp_path, doc)

    cmd_numbering_remove(docx_path)

    assert docx.Document(docx_path).paragraphs[0].text == "引言"


def test_remove_strips_roman(tmp_path):
    """III 标题 → 标题。"""
    doc = docx.Document()
    doc.add_heading("III 标题", level=1)
    docx_path = _save_doc(tmp_path, doc)

    cmd_numbering_remove(docx_path)

    assert docx.Document(docx_path).paragraphs[0].text == "标题"


def test_remove_no_prefix_unchanged(tmp_path):
    """无前缀的标题文本不变。"""
    doc = docx.Document()
    doc.add_heading("引言", level=1)
    docx_path = _save_doc(tmp_path, doc)

    cmd_numbering_remove(docx_path)

    assert docx.Document(docx_path).paragraphs[0].text == "引言"
```

- [ ] **Step 2: 运行测试，确认 4 个失败（cmd_numbering_remove 不存在）**

Run: `cd D:/DevProjects/office-master/word-master && python -m pytest tests/test_docx_structure.py -k remove -v`
Expected: 4 errors (cannot import name 'cmd_numbering_remove')

- [ ] **Step 3: 在 docx_structure.py 末尾追加 cmd_numbering_remove**

```python
def cmd_numbering_remove(docx_path, output_path=None):
    """
    移除 docx 所有标题段落的编号前缀。

    output_path: 给定则输出到新文件；None 则覆盖原文件并生成 .bak。
    策略 B：只改 run[0].text，其余 run 不动。
    """
    doc = docx.Document(docx_path)
    for para in doc.paragraphs:
        if _detect_heading_level(para) is None:
            continue
        if not para.runs:
            continue
        para.runs[0].text = _strip_prefix(para.runs[0].text)
    _save_docx(doc, docx_path, output_path)
```

- [ ] **Step 4: 运行全部 17 个测试，确认通过**

Run: `cd D:/DevProjects/office-master/word-master && python -m pytest tests/test_docx_structure.py -v`
Expected: 17 passed

- [ ] **Step 5: 提交**

```bash
cd D:/DevProjects/office-master
git add word-master/scripts/docx_structure.py word-master/tests/test_docx_structure.py
git commit -m "feat(word-master): add cmd_numbering_remove with 4 tests"
```

---

### Task 5: docx_structure.py CLI main()（测试 18-19）

**Files:**
- Modify: `word-master/scripts/docx_structure.py`
- Modify: `word-master/tests/test_docx_structure.py`

- [ ] **Step 1: 在 test_docx_structure.py 末尾追加 2 个 CLI 集成测试**

```python
import subprocess


def test_cli_add_via_argv(tmp_path):
    """python docx_structure.py doc.docx numbering add --h1 "{1} " 跑通。"""
    doc = docx.Document()
    doc.add_heading("引言", level=1)
    docx_path = _save_doc(tmp_path, doc)
    script = os.path.join(os.path.dirname(__file__), '..', 'scripts', 'docx_structure.py')

    result = subprocess.run(
        [sys.executable, script, docx_path, "numbering", "add", "--h1", "{1} "],
        capture_output=True, text=True,
    )
    assert result.returncode == 0, "stderr: " + result.stderr
    assert docx.Document(docx_path).paragraphs[0].text == "1 引言"


def test_cli_remove_via_argv(tmp_path):
    """numbering remove 跑通。"""
    doc = docx.Document()
    doc.add_heading("1 引言", level=1)
    docx_path = _save_doc(tmp_path, doc)
    script = os.path.join(os.path.dirname(__file__), '..', 'scripts', 'docx_structure.py')

    result = subprocess.run(
        [sys.executable, script, docx_path, "numbering", "remove"],
        capture_output=True, text=True,
    )
    assert result.returncode == 0, "stderr: " + result.stderr
    assert docx.Document(docx_path).paragraphs[0].text == "引言"
```

- [ ] **Step 2: 运行测试，确认 2 个失败（main 不存在，subprocess 报错）**

Run: `cd D:/DevProjects/office-master/word-master && python -m pytest tests/test_docx_structure.py -k cli -v`
Expected: 2 failed (subprocess 返回非 0)

- [ ] **Step 3: 在 docx_structure.py 末尾追加 _resolve_templates + main + __main__ 块**

```python
def _resolve_templates(args):
    """
    合并 config 文件与 CLI 参数，得到最终 level_templates dict。
    优先级: config[1..6] < --h1..--h6
    """
    level_templates = {i: None for i in range(1, 7)}
    start_from = 1

    if args.config:
        cfg = load_config(args.config)
        for i in range(1, 7):
            level_templates[i] = cfg[i]
        start_from = cfg.get("start_from", 1)

    for i in range(1, 7):
        cli_val = getattr(args, "h" + str(i))
        if cli_val is not None:
            level_templates[i] = cli_val

    if args.start_from is not None:
        start_from = args.start_from

    return level_templates, start_from


def main():
    import argparse
    parser = argparse.ArgumentParser(description="DOCX 标题编号管理")
    parser.add_argument("input", help="输入 DOCX 文件路径")
    parser.add_argument("action", choices=["numbering"], help="操作类型")
    parser.add_argument("subaction", choices=["add", "remove"], help="子操作")
    parser.add_argument("-o", "--output", help="输出文件路径")
    parser.add_argument("--config", help="编号配置文件 (YAML)")
    parser.add_argument("--save-config", dest="save_config",
                        help="保存编号配置到 YAML 文件")
    parser.add_argument("--start-from", type=int, default=None, dest="start_from",
                        help="h1 起始编号")
    for i in range(1, 7):
        parser.add_argument("--h" + str(i), help="第 " + str(i) + " 级编号模板")

    args = parser.parse_args()

    if args.action == "numbering":
        if args.subaction == "add":
            if args.config and args.save_config:
                parser.error("--config 和 --save-config 不能同时使用")
            level_templates, start_from = _resolve_templates(args)
            cli_h = any(getattr(args, "h" + str(i)) for i in range(1, 7))
            if not (cli_h or args.config):
                parser.error("需要 --h1..--h6 或 --config 参数")
            if not any(v for v in level_templates.values()):
                parser.error("至少需要为一个级别提供模板")

            if args.save_config:
                cfg = dict(level_templates)
                cfg["start_from"] = start_from
                save_config(args.save_config, cfg)
                print("配置已保存到: " + args.save_config)
                return

            cmd_numbering_add(
                args.input, level_templates, start_from,
                output_path=args.output, save_config_path=None,
            )
        elif args.subaction == "remove":
            cmd_numbering_remove(args.input, output_path=args.output)


if __name__ == "__main__":
    main()
```

- [ ] **Step 4: 运行全部 19 个测试，确认通过**

Run: `cd D:/DevProjects/office-master/word-master && python -m pytest tests/test_docx_structure.py -v`
Expected: 19 passed

- [ ] **Step 5: 提交**

```bash
cd D:/DevProjects/office-master
git add word-master/scripts/docx_structure.py word-master/tests/test_docx_structure.py
git commit -m "feat(word-master): add CLI main() with argparse + 2 integration tests"
```

---

### Task 6: 更新 SKILL.md + 全量回归

**Files:**
- Modify: `word-master/SKILL.md`

- [ ] **Step 1: 在 SKILL.md 的 Core Scripts 表格末尾加一行**

用 Edit 工具，把：
```
| `${SKILL_DIR}/scripts/table_analyzer.py` | Extract table style fingerprints (library) |
```
改为：
```
| `${SKILL_DIR}/scripts/table_analyzer.py` | Extract table style fingerprints (library) |
| `${SKILL_DIR}/scripts/docx_structure.py` | 标题编号管理（添加/移除），支持灵活模板 |
```

- [ ] **Step 2: 在 SKILL.md 末尾追加 "Heading Numbering" 节**

```markdown

---

## Heading Numbering

标题编号管理：给 docx 的标题段落添加/移除编号前缀（文本前缀方案，非 Word 原生 numPr）。
与 markdown-master 的 `structure.py numbering` 对称，模板语法和 YAML 配置完全一致。

### CLI

```bash
# 添加编号
python ${SKILL_DIR}/scripts/docx_structure.py <file.docx> numbering add \
    [--h1 T] [--h2 T] [--h3 T] [--h4 T] [--h5 T] [--h6 T] \
    [--config FILE] [--start-from N] [--save-config FILE] \
    [-o output.docx]

# 移除编号
python ${SKILL_DIR}/scripts/docx_structure.py <file.docx> numbering remove [-o output.docx]
```

默认覆盖原文件并生成 `.bak` 备份；`-o` 输出到新文件时不生成 `.bak`。

### 编号模板语法

| 占位符 | 含义 | 示例 |
|---------|------|------|
| `{N}` / `{N:d}` | 十进制数字 | 3 |
| `{N:02d}` | 两位补零十进制数字 | 03 |
| `{N:R}` | 大写罗马数字 | III |
| `{N:r}` | 小写罗马数字 | iii |
| `{N:A}` | 大写字母 | C |
| `{N:a}` | 小写字母 | c |
| `{N:cn}` | 中文数字 | 三 |

N 范围是 1-6，对应标题级别 H1-H6。空模板字符串表示该级标题不添加编号。

### YAML 配置文件格式

```yaml
h1: "第{1}章 "
h2: "{1}.{2} "
h3: "（{3}）"
h4: ""
h5: ""
h6: ""
start_from: 1
```

### 常见编号用法示例

```bash
# 技术文档风格（1/1.1/1.1.1）
python ${SKILL_DIR}/scripts/docx_structure.py doc.docx numbering add \
    --h1 "{1} " --h2 "{1}.{2} " --h3 "{1}.{2}.{3} "

# 中文章节风格（第1章/1.1/（1））
python ${SKILL_DIR}/scripts/docx_structure.py doc.docx numbering add \
    --h1 "第{1}章 " --h2 "{1}.{2} " --h3 "（{3}）"

# 学术风格（I/A/1/a）
python ${SKILL_DIR}/scripts/docx_structure.py doc.docx numbering add \
    --h1 "{1:R} " --h2 "{2:A} " --h3 "{3} " --h4 "{4:a} "

# 保存当前配置为模板
python ${SKILL_DIR}/scripts/docx_structure.py doc.docx numbering add \
    --h1 "第{1}章 " --h2 "{1}.{2} " --save-config chinese_chapter.yaml

# 使用保存的配置
python ${SKILL_DIR}/scripts/docx_structure.py doc.docx numbering add \
    --config chinese_chapter.yaml
```
```

- [ ] **Step 3: 运行新增两个测试文件，确认 31 个用例全绿**

Run: `cd D:/DevProjects/office-master/word-master && python -m pytest tests/test_numbering_render.py tests/test_docx_structure.py -v`
Expected: 31 passed

- [ ] **Step 4: 运行 word-master 全部测试，确认不破坏现有测试**

Run: `cd D:/DevProjects/office-master/word-master && python -m pytest tests/ -v 2>&1 | tail -20`
Expected: 新增 31 个 passed，现有测试无新增 failed（忽略预先存在的失败）

- [ ] **Step 5: 提交**

```bash
cd D:/DevProjects/office-master
git add word-master/SKILL.md
git commit -m "docs(word-master): document heading numbering in SKILL.md"
```

---

## Self-Review

**1. Spec coverage：**
- §文件结构 _numbering_render.py（7 函数）→ Task 1 ✓
- §文件结构 docx_structure.py → Task 2-5 ✓
- §标题识别 _detect_heading_level（英文+中文，1-9）→ Task 2 ✓
- §前缀写入 策略 B（只改 run[0]）→ Task 3 cmd_numbering_add ✓
- §前缀剥除 _strip_prefix（含中文数字章）→ Task 3 ✓
- §计数器维护（counters[0]=start_from-1，level N 清零 N..5）→ Task 3 ✓
- §CLI（--h1..--h6/--config/--start-from/--save-config/-o）→ Task 5 ✓
- §.bak 备份（覆盖时生成，-o 时不生成）→ Task 3 _save_docx ✓
- §默认覆盖原文件 → Task 3 _save_docx ✓
- §YAML 配置格式 → Task 1 load_config/save_config ✓
- §模板语法表 → Task 6 SKILL.md ✓
- §SKILL.md 更新（Core Scripts 表 + Heading Numbering 节）→ Task 6 ✓
- §test_numbering_render 12 用例 → Task 1 ✓
- §test_docx_structure 19 用例（检测 4 + add 6 + remove 4 + I/O 3 + CLI 2）→ Task 2-5 ✓
- §验收标准 1（全绿）→ Task 6 Step 3 ✓
- §验收标准 2（不破坏现有测试）→ Task 6 Step 4 ✓
- §验收标准 3（CLI/模板/YAML 与 md 侧对称）→ Task 5 + Task 6 ✓
- §验收标准 4（渲染层 12 用例）→ Task 1 ✓
- §验收标准 5（docx I/O 层 19 用例）→ Task 2-5 ✓
- §验收标准 6（策略 B，测试 9 验证）→ Task 3 test_add_multi_run_preserves_format ✓
- §验收标准 7（.bak 生成/不生成，测试 15/16 验证）→ Task 3 ✓

**2. Placeholder scan：** 无 TODO / TBD / "类似 Task N"。每个代码步骤含完整代码。

**3. Type consistency：**
- `parse_template(template: str) -> list` — Task 1 定义，Task 3 调用一致
- `format_template(tokens, counters)` — Task 1 定义，Task 3 调用一致
- `load_config(path) -> {1..6, "start_from"}` — Task 1 定义，Task 5 `_resolve_templates` 调用一致
- `save_config(path, cfg)` — Task 1 定义，Task 3/5 调用一致
- `cmd_numbering_add(docx_path, level_templates, start_from, output_path, save_config_path)` — Task 3 定义，Task 5 main() 调用一致
- `cmd_numbering_remove(docx_path, output_path)` — Task 4 定义，Task 5 main() 调用一致
- `_save_docx(doc, docx_path, output_path)` — Task 3 定义，Task 4 调用一致

**潜在风险（执行时注意）：**
- Task 2 `test_detect_heading_chinese`：`doc.styles.add_style("标题 1", WD_STYLE_TYPE.PARAGRAPH)` 创建自定义样式，`_detect_heading_level` 读 `paragraph.style.name` 返回 "标题 1"，正则 `^标题\s*(\d+)$` 匹配。需确认 python-docx 的 `add_style` 后 `style.name` 就是传入的名字（是的，python-docx 用第一个参数作 name）。
- Task 3 `test_add_empty_template_skips_level`：`{1: "", 2: "{1}.{2} "}` 时 H1 不编号（`if not template` 过滤空串），但 counters[0] 仍会因 `level_templates.get(1)` 为空串而不初始化为 start_from-1，保持 0。H2 编号时 counters[0]=0，所以输出 "0.1 背景"。这是预期行为（H1 不编号时其计数器不递增，H2 的父级编号为 0）。
- Task 3 `test_add_multi_run_preserves_format`：`add_heading(level=1)` 不传 text 会创建空标题段落，再 `add_run` 添加 run。python-docx 的 `add_heading` 无 text 参数时创建空段落，需确认。若 `add_heading(level=1)` 报错，改用 `doc.add_paragraph("", style=doc.styles['Heading 1'])`。Task 2 的 `_make_doc_with_heading` 已验证 `add_heading(text, level=N)` 可用。
- Task 5 CLI 测试用 subprocess，`sys.executable` 保证用当前 Python。脚本路径用 `os.path.join(os.path.dirname(__file__), '..', 'scripts', 'docx_structure.py')` 与测试文件的 sys.path 注入一致。
- Task 6 Step 4 现有测试可能有预先存在的失败（与本次无关），执行者需对比"本次改动前后的失败数"，不是要求零失败。

无问题。
