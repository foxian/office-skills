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