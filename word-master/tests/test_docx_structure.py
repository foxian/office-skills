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