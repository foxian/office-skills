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