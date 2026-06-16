
"""
结构操作：标题层级升降、编号管理、目录生成。
"""
import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from _md_utils import read_utf8, write_utf8, is_in_code_block


def heading_shift(content, delta):
    """delta &gt; 0 升级（减少 #），delta &lt; 0 降级（增加 #）。"""
    lines = content.split("\n")
    result = []
    for i, line in enumerate(lines):
        if is_in_code_block(lines, i):
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
    headings = []
    for i, line in enumerate(lines):
        if is_in_code_block(lines, i):
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

