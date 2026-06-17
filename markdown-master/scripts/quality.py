
"""
质检操作：格式规范检查（lint）、中文排版检查（zhlint）、链接检查（linkcheck）。
"""
import argparse
import re
import sys
import urllib.request
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from _md_utils import read_utf8, write_utf8, is_in_code_block, _precompute_code_state


def run_lint(content, fix):
    lines = content.split("\n")
    code_state = _precompute_code_state(lines)
    issues = []
    result_lines = list(lines)

    # 1. 检查是否有 h1
    has_h1 = False
    for line in lines:
        if re.match(r"^#\s+", line):
            has_h1 = True
            break
    if not has_h1:
        issues.append((0, "缺少 H1 标题"))

    # 2. 标题层级跳跃
    prev_level = 0
    for i, line in enumerate(lines):
        if code_state[i]:
            continue
        m = re.match(r"^(#{1,6})\s+", line)
        if m:
            level = len(m.group(1))
            if prev_level &gt; 0 and level &gt; prev_level + 1:
                issues.append((i + 1, "标题层级跳跃：H" + str(prev_level) + " -&gt; H" + str(level)))
            prev_level = level

    # 3. 连续 3 行以上空行
    i = 0
    while i &lt; len(result_lines):
        if result_lines[i].strip() == "":
            j = i
            while j &lt; len(result_lines) and result_lines[j].strip() == "":
                j += 1
            if j - i &gt;= 3:
                issues.append((i + 1, "连续空行过多"))
                if fix:
                    result_lines[i:j] = ["", ""]
            i = j
        else:
            i += 1

    # 4. 代码块未指定语言
    for i, line in enumerate(lines):
        if re.match(r"^```\s*$", line):
            issues.append((i + 1, "代码块未指定语言"))

    return "\n".join(result_lines), issues


_ZH = r"[一-鿿㐀-䶿豈-﫿]"
_EN = r"[A-Za-z0-9]"
_ZH_BEFORE_EN = re.compile(r"(" + _ZH + ")(" + _EN + r")")
_EN_BEFORE_ZH = re.compile(r"(" + _EN + ")(" + _ZH + r")")
_HALF_PUNCT = re.compile(r"(" + _ZH + ")[,.?!;:]")


def _is_inline_code(line, pos):
    count = 0
    for i, ch in enumerate(line):
        if ch == "`":
            count += 1
        if i == pos:
            break
    return count % 2 == 1


def run_zhlint(content, fix):
    lines = content.split("\n")
    code_state = _precompute_code_state(lines)
    issues = []
    result_lines = []

    for i, line in enumerate(lines):
        if code_state[i]:
            result_lines.append(line)
            continue

        new_line = line

        # 中文与英文之间缺空格
        if _ZH_BEFORE_EN.search(new_line) or _EN_BEFORE_ZH.search(new_line):
            issues.append((i + 1, "中文与英文/数字之间缺少空格"))
            if fix:
                new_line = _ZH_BEFORE_EN.sub(r"\1 \2", new_line)
                new_line = _EN_BEFORE_ZH.sub(r"\1 \2", new_line)

        # 中文语境中使用半角标点
        if _HALF_PUNCT.search(new_line):
            issues.append((i + 1, "中文语境中使用了半角标点"))
            if fix:
                punct_map = {",": "，", ".": "。", "?": "？", "!": "！", ";": "；", ":": "："}
                def replace_fn(m):
                    zh_char = m.group(0)[0]
                    half = m.group(0)[1]
                    return zh_char + punct_map.get(half, half)
                new_line = _HALF_PUNCT.sub(replace_fn, new_line)

        result_lines.append(new_line)

    return "\n".join(result_lines), issues


_MD_LINK = re.compile(r"\[.*?\]\((.*?)\)")
_MD_IMAGE = re.compile(r"!\[.*?\]\((.*?)\)")


def run_linkcheck(content, file_path, check_remote):
    base_dir = Path(file_path).parent
    lines = content.split("\n")
    code_state = _precompute_code_state(lines)
    issues = []

    for i, line in enumerate(lines):
        if code_state[i]:
            continue

        # 图片引用
        for m in _MD_IMAGE.finditer(line):
            src = m.group(1).split(" ")[0]
            if src.startswith("http://") or src.startswith("https://"):
                if check_remote:
                    try:
                        req = urllib.request.Request(src, method="HEAD")
                        req.add_header("User-Agent", "markdown-master-linkcheck/1.0")
                        with urllib.request.urlopen(req, timeout=5) as resp:
                            if resp.status &gt;= 400:
                                issues.append((i + 1, "远程链接返回 " + str(resp.status) + ": " + src))
                    except Exception as e:
                        issues.append((i + 1, "远程链接无法访问: " + src))
            else:
                target = (base_dir / src).resolve()
                if not target.exists():
                    issues.append((i + 1, "图片路径不存在: " + src))

        # 普通链接
        for m in _MD_LINK.finditer(line):
            href = m.group(1).split(" ")[0]
            if href.startswith("#"):
                continue
            if href.startswith("http://") or href.startswith("https://"):
                if check_remote:
                    try:
                        req = urllib.request.Request(href, method="HEAD")
                        req.add_header("User-Agent", "markdown-master-linkcheck/1.0")
                        with urllib.request.urlopen(req, timeout=5) as resp:
                            if resp.status &gt;= 400:
                                issues.append((i + 1, "远程链接返回 " + str(resp.status) + ": " + href))
                    except Exception as e:
                        issues.append((i + 1, "远程链接无法访问: " + href))
            else:
                target = (base_dir / href).resolve()
                if not target.exists():
                    issues.append((i + 1, "本地链接目标不存在: " + href))

    return issues


def print_issues(issues, action):
    if not issues:
        print("✓ No issues found (" + action + ")")
        return
    for line_no, desc in issues:
        loc = "L" + str(line_no) if line_no &gt; 0 else "文档级"
        print("  [" + loc + "] " + desc)
    print("\n共发现 " + str(len(issues)) + " 个问题")


def main():
    parser = argparse.ArgumentParser(description="Markdown 质检工具")
    parser.add_argument("input", help="输入 Markdown 文件路径")
    parser.add_argument("action", choices=["lint", "zhlint", "linkcheck"], help="检查类型")
    parser.add_argument("--fix", action="store_true", help="自动修复可修复的问题")
    parser.add_argument("--check-remote", action="store_true", dest="check_remote", help="linkcheck 时同时检查 HTTP/HTTPS 链接")

    args = parser.parse_args()
    content = read_utf8(args.input)

    if args.action == "lint":
        result, issues = run_lint(content, args.fix)
        print_issues(issues, "lint")
        if args.fix and issues:
            write_utf8(args.input, result)
            print("✓ 已修复并写入: " + args.input)

    elif args.action == "zhlint":
        result, issues = run_zhlint(content, args.fix)
        print_issues(issues, "zhlint")
        if args.fix and issues:
            write_utf8(args.input, result)
            print("✓ 已修复并写入: " + args.input)

    elif args.action == "linkcheck":
        issues = run_linkcheck(content, args.input, args.check_remote)
        print_issues(issues, "linkcheck")


if __name__ == "__main__":
    main()

