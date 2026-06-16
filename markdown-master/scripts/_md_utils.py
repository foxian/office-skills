"""
共用工具库，供 structure.py / quality.py / convert.py / files.py import。
不可直接调用。
"""
import re
from pathlib import Path


def read_utf8(path: str) -> str:
    """读取 UTF-8 文件内容，自动处理 BOM。"""
    return Path(path).read_text(encoding="utf-8-sig")


def write_utf8(path: str, content: str) -> None:
    """将内容写入 UTF-8 文件（无 BOM）。"""
    Path(path).write_text(content, encoding="utf-8")


def parse_file(path: str) -> tuple[str, dict]:
    """
    解析 Markdown 文件，分离 frontmatter 和正文。
    返回: (body_content, frontmatter_dict)
    若无 frontmatter，frontmatter_dict 为空 dict。
    """
    content = read_utf8(path)
    if not content.startswith("---"):
        return content, {}

    lines = content.split("\n")
    end = -1
    for i in range(1, len(lines)):
        if lines[i].strip() == "---":
            end = i
            break
    if end == -1:
        return content, {}

    fm_lines = lines[1:end]
    body = "\n".join(lines[end + 1:])
    fm = {}
    for line in fm_lines:
        if ":" in line:
            k, _, v = line.partition(":")
            fm[k.strip()] = v.strip()
    return body, fm


def extract_headings(content: str) -> list[tuple[int, str, int]]:
    """
    提取正文中所有标题。
    返回: list of (level, text, line_number)  —  line_number 从 1 开始
    跳过代码块内的 # 行。
    """
    lines = content.split("\n")
    headings = []
    for i, line in enumerate(lines):
        if is_in_code_block(lines, i):
            continue
        m = re.match(r"^(#{1,6})\s+(.*)", line)
        if m:
            level = len(m.group(1))
            text = m.group(2).strip()
            headings.append((level, text, i + 1))
    return headings


def is_in_code_block(lines: list[str], line_index: int) -> bool:
    """
    判断 lines[line_index] 是否处于代码块（``` 或 ~~~）内部。
    """
    fence_count = 0
    for i in range(line_index):
        if re.match(r"^(`{3,}|~{3,})", lines[i]):
            fence_count += 1
    return fence_count % 2 == 1
