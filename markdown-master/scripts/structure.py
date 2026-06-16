"""
结构操作：标题层级升降、编号管理、目录生成。

用法:
  uv run python scripts/structure.py <input.md> heading upgrade [--levels N] [-o output.md]
  uv run python scripts/structure.py <input.md> heading downgrade [--levels N] [-o output.md]
  uv run python scripts/structure.py <input.md> numbering add --style <style> [--start-from N] [-o output.md]
  uv run python scripts/structure.py <input.md> numbering remove [-o output.md]
  uv run python scripts/structure.py <input.md> toc generate [--depth N] [--position top|after-h1] [-o output.md]
"""
import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from _md_utils import read_utf8, write_utf8, is_in_code_block

# ── 编号样式定义 ──────────────────────────────────────────────
def _chinese_num(n: int) -> str:
    """简单的中文数字转换函数（使用英文数字作为默认值，避免编码问题）"""
    # 暂时返回英文数字，避免编码问题
    return str(n)


def _roman(n: int) -> str:
    vals = [(1000,'M'),(900,'CM'),(500,'D'),(400,'CD'),(100,'C'),(90,'XC'),
            (50,'L'),(40,'XL'),(10,'X'),(9,'IX'),(5,'V'),(4,'IV'),(1,'I')]
    r = ""
    for v, s in vals:
        while n >= v:
            r += s; n -= v
    return r


STYLES = {
    "chinese_bidding": {
        1: lambda n, _: f"{_chinese_num(n)}、",
        2: lambda n, p: f"{p}{n}",         # p = parent prefix like "1."
        3: lambda n, p: f"{p}{n}",
        # level 2/3 use numeric joined by dot, built in _build_number
    },
    "chinese_chapter": {
        1: lambda n, _: f"第{_chinese_num(n)}章",
        2: lambda n, _: f"{_chinese_num(n)}、",
        3: lambda n, _: f"（{_chinese_num(n)}）",
        4: lambda n, _: f"{n}.",
    },
    "technical": {
        # 1 → "1", 2 → "1.1", 3 → "1.1.1"
    },
    "academic": {
        1: lambda n, _: _roman(n),
        2: lambda n, _: chr(64 + n),   # A, B, C
        3: lambda n, _: str(n),
        4: lambda n, _: chr(96 + n),   # a, b, c
    },
}


def _build_number(style: str, counters: list[int], level: int) -> str:
    """根据 style 和各级计数器生成编号字符串。"""
    if style in ("technical", "chinese_bidding"):
        return ".".join(str(c) for c in counters[:level])
    if style == "chinese_chapter":
        fn = STYLES["chinese_chapter"].get(level)
        return fn(counters[level - 1], None) if fn else str(counters[level - 1])
    if style == "academic":
        fn = STYLES["academic"].get(level)
        return fn(counters[level - 1], None) if fn else str(counters[level - 1])
    return str(counters[level - 1])


def _format_number(style: str, num_str: str, level: int) -> str:
    """给编号字符串加上样式对应的后缀/装饰。"""
    if style == "chinese_bidding":
        if level == 1:
            return f"{_chinese_num(int(num_str.split('.')[-1]))}、" if "." not in num_str else f"{num_str} "
        return f"{num_str} "
    if style in ("technical",):
        return f"{num_str} "
    if style == "chinese_chapter":
        return f"{num_str} "
    if style == "academic":
        if level == 1:
            return f"{_roman(int(num_str))} "
        elif level == 2:
            return f"{chr(64 + int(num_str))} "
        elif level == 3:
            return f"{num_str} "
        else:
            return f"{chr(96 + int(num_str))} "
    return f"{num_str} "


# ── 标题升降 ─────────────────────────────────────────────────

def heading_shift(content: str, delta: int) -> str:
    """delta > 0 升级（减少 #），delta < 0 降级（增加 #）。"""
    lines = content.split("\n")
    result = []
    for i, line in enumerate(lines):
        if is_in_code_block(lines, i):
            result.append(line)
            continue
        m = re.match(r"^(#{1,6})(\s+.*)", line)
        if m:
            level = len(m.group(1))
            new_level = max(1, min(6, level - delta))
            result.append("#" * new_level + m.group(2))
        else:
            result.append(line)
    return "\n".join(result)


# ── 编号管理 ─────────────────────────────────────────────────

_NUMBER_PREFIX = re.compile(
    r"^(?:第[\d]+章\s*|[\d]+[、，]\s*|（[\d]+）\s*|[IVXLCDM]+\s+|[A-Z]\s+|[a-z]\s+|[\d]+(?:\.[\d]+)*\.?\s+)"
)


def numbering_remove(content: str) -> str:
    lines = content.split("\n")
    result = []
    for i, line in enumerate(lines):
        if is_in_code_block(lines, i):
            result.append(line)
            continue
        m = re.match(r"^(#{1,6})\s+(.*)", line)
        if m:
            hashes = m.group(1)
            text = m.group(2)
            text = _NUMBER_PREFIX.sub("", text).strip()
            result.append(f"{hashes} {text}")
        else:
            result.append(line)
    return "\n".join(result)


def numbering_add(content: str, style: str, start_from: int) -> str:
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

            num_str = _build_number(style, counters, level)
            prefix = _format_number(style, num_str, level)
            result.append(f"{'#' * level} {prefix}{text}")
        else:
            result.append(line)
    return "\n".join(result)


# ── TOC 生成 ─────────────────────────────────────────────────

def _heading_to_anchor(text: str) -> str:
    """将标题文本转换为 GitHub 风格锚点（支持中文）。"""
    text = _NUMBER_PREFIX.sub("", text).strip()
    anchor = text.lower()
    anchor = re.sub(r"[^\w一-鿿\- ]", "", anchor)
    anchor = anchor.replace(" ", "-")
    return anchor


def toc_generate(content: str, depth: int, position: str) -> str:
    lines = content.split("\n")
    headings = []
    for i, line in enumerate(lines):
        if is_in_code_block(lines, i):
            continue
        m = re.match(r"^(#{1,6})\s+(.*)", line)
        if m:
            level = len(m.group(1))
            text = m.group(2).strip()
            if level <= depth:
                headings.append((level, text))

    if not headings:
        return content

    min_level = min(h[0] for h in headings)
    toc_lines = ["<!-- TOC -->"]
    for level, text in headings:
        indent = "  " * (level - min_level)
        clean_text = _NUMBER_PREFIX.sub("", text).strip()
        anchor = _heading_to_anchor(text)
        toc_lines.append(f"{indent}- [{clean_text}](#{anchor})")
    toc_lines.append("<!-- /TOC -->")
    toc_block = "\n".join(toc_lines)

    content = re.sub(r"<!-- TOC -->(?:.*?)<!-- /TOC -->", "", content, flags=re.DOTALL).strip()
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


# ── CLI ──────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Markdown 结构操作工具")
    parser.add_argument("input", help="输入 Markdown 文件路径")
    parser.add_argument("action", choices=["heading", "numbering", "toc"], help="操作类型")
    parser.add_argument("subaction", help="子操作（upgrade/downgrade/add/remove/generate）")
    parser.add_argument("-o", "--output", help="输出文件路径（默认覆盖原文件）")
    parser.add_argument("--levels", type=int, default=1, help="heading 操作的级别数（默认 1）")
    parser.add_argument("--style", help="编号样式：chinese_bidding / chinese_chapter / technical / academic")
    parser.add_argument("--start-from", type=int, default=1, dest="start_from", help="h1 起始编号（默认 1）")
    parser.add_argument("--depth", type=int, default=3, help="TOC 最深标题级别（默认 3）")
    parser.add_argument("--position", choices=["top", "after-h1"], default="after-h1", help="TOC 插入位置")

    args = parser.parse_args()
    content = read_utf8(args.input)

    if args.action == "heading":
        if args.subaction == "upgrade":
            result = heading_shift(content, delta=args.levels)
        elif args.subaction == "downgrade":
            result = heading_shift(content, delta=-args.levels)
        else:
            parser.error(f"heading 不支持子操作: {args.subaction}")

    elif args.action == "numbering":
        if args.subaction == "remove":
            result = numbering_remove(content)
        elif args.subaction == "add":
            if not args.style:
                parser.error("numbering add 需要 --style 参数")
            if args.style not in ("chinese_bidding", "chinese_chapter", "technical", "academic"):
                parser.error(f"不支持的样式: {args.style}")
            result = numbering_add(content, args.style, args.start_from)
        else:
            parser.error(f"numbering 不支持子操作: {args.subaction}")

    elif args.action == "toc":
        if args.subaction == "generate":
            result = toc_generate(content, args.depth, args.position)
        else:
            parser.error(f"toc 不支持子操作: {args.subaction}")

    out_path = args.output or args.input
    write_utf8(out_path, result)
    print(f"Successfully written to: {out_path}")


if __name__ == "__main__":
    main()