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