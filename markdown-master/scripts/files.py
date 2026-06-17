
"""
文件级操作：拆分（split）、合并（merge）。
"""
import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from _md_utils import read_utf8, write_utf8, is_in_code_block, _precompute_code_state


def _sanitize_filename(text):
    """将标题文本转为合法文件名，清理非法字符。"""
    # 移除编号前缀
    text = re.sub(r"^(?:第[\d]+章\s*|[\d]+[、，]\s*|（[\d]+）\s*|[IVXLCDM]+\s+|[A-Z]\s+|[\d]+(?:\.[\d]+)*\.?\s+)", "", text).strip()
    # 替换 Windows/Linux 非法字符
    text = re.sub(r'[\\/:*?"&lt;&gt;|]', "_", text)
    text = text.strip(". ")
    return text[:80] or "untitled"


def cmd_split(input_path, by, output_dir):
    level = int(by[1])
    content = read_utf8(input_path)
    lines = content.split("\n")
    code_state = _precompute_code_state(lines)

    out_dir = Path(output_dir) if output_dir else Path(input_path).parent / "split_output"
    out_dir.mkdir(parents=True, exist_ok=True)

    sections = []
    current_title = "_preamble"
    current_lines = []

    for i, line in enumerate(lines):
        if code_state[i]:
            current_lines.append(line)
            continue
        m = re.match(r"^(#{1,6})\s+(.*)", line)
        if m and len(m.group(1)) == level:
            if current_lines or current_title != "_preamble":
                sections.append((current_title, current_lines))
            current_title = m.group(2).strip()
            current_lines = [line]
        else:
            current_lines.append(line)

    if current_lines:
        sections.append((current_title, current_lines))

    count = 0
    for i, (title, sec_lines) in enumerate(sections):
        if title == "_preamble" and not any(l.strip() for l in sec_lines):
            continue
        if title == "_preamble":
            fname = "00_preamble.md"
        else:
            fname = str(i).zfill(2) + "_" + _sanitize_filename(title) + ".md"
        fpath = out_dir / fname
        write_utf8(str(fpath), "\n".join(sec_lines))
        print("  ✓ " + str(fpath))
        count += 1

    print("\n共拆分为 " + str(count) + " 个文件 -&gt; " + str(out_dir))


def cmd_merge(inputs, output, separator):
    md_files = []

    for inp in inputs:
        p = Path(inp)
        if p.is_dir():
            md_files.extend(sorted(p.rglob("*.md")))
        elif p.is_file() and p.suffix == ".md":
            md_files.append(p)
        else:
            print("  跳过: " + inp)

    if not md_files:
        print("错误：没有找到可合并的 .md 文件")
        sys.exit(1)

    parts = []
    for f in md_files:
        print("  + " + str(f))
        parts.append(read_utf8(str(f)))

    merged = separator.join(parts)
    write_utf8(output, merged)
    print("\n✓ 已合并 " + str(len(parts)) + " 个文件 -&gt; " + output)


def main():
    parser = argparse.ArgumentParser(description="Markdown 文件级操作工具")
    subparsers = parser.add_subparsers(dest="action", required=True)

    # split
    sp = subparsers.add_parser("split", help="按标题级别拆分文档")
    sp.add_argument("input", help="输入 Markdown 文件")
    sp.add_argument("--by", choices=["h1", "h2", "h3"], default="h2", help="按哪一级标题拆分")
    sp.add_argument("--output-dir", dest="output_dir", help="输出目录")

    # merge
    mp = subparsers.add_parser("merge", help="合并多个 Markdown 文件")
    mp.add_argument("inputs", nargs="+", help="输入文件或目录")
    mp.add_argument("-o", "--output", required=True, help="输出文件路径")
    mp.add_argument("--separator", default="\n\n---\n\n", help="文件间分隔符")

    args = parser.parse_args()

    if args.action == "split":
        cmd_split(args.input, args.by, args.output_dir)
    elif args.action == "merge":
        cmd_merge(args.inputs, args.output, args.separator)


if __name__ == "__main__":
    main()

