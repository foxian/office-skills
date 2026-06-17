
"""
结构操作：标题层级升降、编号管理、目录生成。
"""
import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from _md_utils import read_utf8, write_utf8, is_in_code_block, _precompute_code_state


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
    """校验并规范化修饰符。返回内部表示的 fmt 串。"""
    if fmt_raw in ("d", "R", "r", "A", "a", "cn"):
        return fmt_raw
    # 0Nd 形式
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
        return "0" if lower else "0"
    vals = _ROMAN_LOWER_VALS if lower else _ROMAN_VALS
    result = []
    for v, s in vals:
        while n >= v:
            result.append(s)
            n -= v
    if n > 0:
        result.append("M" if not lower else "m")
    return "".join(result)


def _to_letter(n, upper=True):
    """转换为字母 (1=A..Z=26, 27=AA..)"""
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
_CN_BIG_UNITS = ["", "万", "亿"]


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


def heading_shift(content, delta):
    """delta &gt; 0 升级（减少 #），delta &lt; 0 降级（增加 #）。"""
    lines = content.split("\n")
    code_state = _precompute_code_state(lines)
    result = []
    for i, line in enumerate(lines):
        if code_state[i]:
            result.append(line)
            continue
        m = re.match(r"^(#{1,6})(\s+.*)", line)
        if m:
            level = len(m.group(1))
            new_level = level - delta
            if new_level &lt; 1:
                new_level = 1
            if new_level &gt; 6:
                new_level = 6
            result.append("#" * new_level + m.group(2))
        else:
            result.append(line)
    return "\n".join(result)


_NUMBER_PREFIX = re.compile(r"^(?:第[\d]+章\s*|[\d]+[、，]\s*|（[\d]+）\s*|[IVXLCDM]+\s+|[A-Z]\s+|[a-z]\s+|[\d]+(?:\.[\d]+)*\.?\s+)")


def numbering_remove(content):
    lines = content.split("\n")
    code_state = _precompute_code_state(lines)
    result = []
    for i, line in enumerate(lines):
        if code_state[i]:
            result.append(line)
            continue
        m = re.match(r"^(#{1,6})\s+(.*)", line)
        if m:
            hashes = m.group(1)
            text = m.group(2)
            text = _NUMBER_PREFIX.sub("", text).strip()
            result.append(hashes + " " + text)
        else:
            result.append(line)
    return "\n".join(result)


def numbering_add(content, style, start_from):
    lines = content.split("\n")
    content_clean = numbering_remove(content)
    lines = content_clean.split("\n")

    counters = [0] * 6
    counters[0] = start_from - 1
    result = []

    for i, line in enumerate(lines):
        if is_in_code_block(lines, i):
            result.append(line)
            continue
        m = re.match(r"^(#{1,6})\s+(.*)", line)
        if m:
            level = len(m.group(1))
            text = m.group(2).strip()
            counters[level - 1] += 1
            for j in range(level, 6):
                counters[j] = 0

            prefix = ""
            if style == "technical":
                parts = []
                for k in range(level):
                    parts.append(str(counters[k]))
                prefix = ".".join(parts) + " "
            elif style == "chinese_chapter":
                if level == 1:
                    prefix = "第" + str(counters[0]) + "章 "
                elif level == 2:
                    prefix = str(counters[1]) + "、"
                elif level == 3:
                    prefix = "（" + str(counters[2]) + "）"
                else:
                    prefix = str(counters[level-1]) + ". "
            elif style == "chinese_bidding":
                if level == 1:
                    prefix = str(counters[0]) + "、"
                else:
                    parts = []
                    for k in range(level):
                        parts.append(str(counters[k]))
                    prefix = ".".join(parts) + " "
            elif style == "academic":
                if level == 1:
                    roman = ""
                    n = counters[0]
                    vals = [(1000,'M'),(900,'CM'),(500,'D'),(400,'CD'),(100,'C'),(90,'XC'),
                            (50,'L'),(40,'XL'),(10,'X'),(9,'IX'),(5,'V'),(4,'IV'),(1,'I')]
                    for v, s in vals:
                        while n &gt;= v:
                            roman += s
                            n -= v
                    prefix = roman + " "
                elif level == 2:
                    prefix = chr(64 + counters[1]) + " "
                elif level == 3:
                    prefix = str(counters[2]) + " "
                else:
                    prefix = chr(96 + counters[level-1]) + " "
            else:
                prefix = str(counters[level-1]) + " "

            result.append("#" * level + " " + prefix + text)
        else:
            result.append(line)
    return "\n".join(result)


def _heading_to_anchor(text):
    text = _NUMBER_PREFIX.sub("", text).strip()
    anchor = text.lower()
    anchor = re.sub(r"[^\w一-鿿\- ]", "", anchor)
    anchor = anchor.replace(" ", "-")
    return anchor


def toc_generate(content, depth, position):
    lines = content.split("\n")
    code_state = _precompute_code_state(lines)
    headings = []
    for i, line in enumerate(lines):
        if code_state[i]:
            continue
        m = re.match(r"^(#{1,6})\s+(.*)", line)
        if m:
            level = len(m.group(1))
            text = m.group(2).strip()
            if level &lt;= depth:
                headings.append((level, text))

    if not headings:
        return content

    min_level = min(h[0] for h in headings)
    toc_lines = ["&lt;!-- TOC --&gt;"]
    for level, text in headings:
        indent = "  " * (level - min_level)
        clean_text = _NUMBER_PREFIX.sub("", text).strip()
        anchor = _heading_to_anchor(text)
        toc_lines.append(indent + "- [" + clean_text + "](#" + anchor + ")")
    toc_lines.append("&lt;!-- /TOC --&gt;")
    toc_block = "\n".join(toc_lines)

    content = re.sub(r"&lt;!-- TOC --&gt;.*?&lt;!-- /TOC --&gt;", "", content, flags=re.DOTALL).strip()
    lines = content.split("\n")

    if position == "top":
        return toc_block + "\n\n" + content

    insert_at = len(lines)
    for i, line in enumerate(lines):
        if re.match(r"^#\s+", line):
            insert_at = i + 1
            break
    lines.insert(insert_at, "\n" + toc_block + "\n")
    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(description="Markdown 结构操作工具")
    parser.add_argument("input", help="输入 Markdown 文件路径")
    parser.add_argument("action", choices=["heading", "numbering", "toc"], help="操作类型")
    parser.add_argument("subaction", help="子操作")
    parser.add_argument("-o", "--output", help="输出文件路径")
    parser.add_argument("--levels", type=int, default=1, help="heading 操作的级别数")
    parser.add_argument("--style", help="编号样式")
    parser.add_argument("--start-from", type=int, default=1, dest="start_from", help="h1 起始编号")
    parser.add_argument("--depth", type=int, default=3, help="TOC 最深标题级别")
    parser.add_argument("--position", choices=["top", "after-h1"], default="after-h1", help="TOC 插入位置")

    args = parser.parse_args()
    content = read_utf8(args.input)

    if args.action == "heading":
        if args.subaction == "upgrade":
            result = heading_shift(content, delta=args.levels)
        elif args.subaction == "downgrade":
            result = heading_shift(content, delta=-args.levels)
        else:
            parser.error("不支持的子操作")

    elif args.action == "numbering":
        if args.subaction == "remove":
            result = numbering_remove(content)
        elif args.subaction == "add":
            if not args.style:
                parser.error("需要 --style 参数")
            result = numbering_add(content, args.style, args.start_from)
        else:
            parser.error("不支持的子操作")

    elif args.action == "toc":
        if args.subaction == "generate":
            result = toc_generate(content, args.depth, args.position)
        else:
            parser.error("不支持的子操作")

    out_path = args.output or args.input
    write_utf8(out_path, result)
    print("Successfully written to: " + out_path)


if __name__ == "__main__":
    main()

