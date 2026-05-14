# Edit Operations DSL

AI must output JSON arrays containing edit operations.

## replace_text 操作

查找替换文本，保留 run 级别格式。

```json
{"op": "replace_text", "find": "A", "replace": "B", "target": "all"}
```

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| target | string | 否 | 目标范围：默认 "all" 表示全局替换；传 "p0" 等表示仅在该段落内替换 |
| find | string | ✅ | 要查找的文本 |
| replace | string | ✅ | 替换后的文本 |

## rewrite_paragraph 操作

重写指定段落的内容，保留原有样式。

```json
{"op": "rewrite_paragraph", "target": "p0", "content": "新内容"}
```

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| target | string | ✅ | 段落 ID，如 "p0" |
| content | string | ✅ | 新内容 |

## set_font 操作

设置段落或行内特定文字的字体格式。

```json
{
  "op": "set_font",
  "target": "p0",
  "text_match": "需要加粗的词",
  "match_index": 0,
  "name": "Arial",
  "east_asia": "楷体",
  "size": "14pt",
  "bold": true,
  "italic": false
}
```

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| target | string | ✅ | 段落 ID，如 "p0" |
| text_match | string | 否 | 行内特定文字。如果不填，则修改整个段落的格式 |
| match_index | int/string | 否 | 仅当填了 `text_match` 时有效。传数字 N (0开始) 表示修改第 N 次出现的文字；省略或传 "all" 表示修改段落内所有匹配项 |
| name | string | 否 | 西文字体名 |
| east_asia | string | 否 | 中文字体名 |
| size | string | 否 | 字号，如 "14pt" 或 "28"（half-points） |
| bold | boolean | 否 | 是否加粗 |
| italic | boolean | 否 | 是否斜体 |

**示例：**
- 设置整段为 Arial 14pt 加粗：`{"op": "set_font", "target": "p0", "name": "Arial", "size": "14pt", "bold": true}`
- 局部修改第一个出现的词汇：`{"op": "set_font", "target": "p1", "text_match": "附件", "match_index": 0, "bold": true}`

## set_paragraph_format 操作

设置段落格式（对齐、行距、缩进、段前段后距）。

```json
{
  "op": "set_paragraph_format",
  "target": "p0",
  "alignment": "center",
  "line_spacing": 1.5,
  "first_line_indent": "2em",
  "space_before": "12pt",
  "space_after": "12pt"
}
```

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| target | string | ✅ | 段落 ID |
| alignment | string | 否 | 对齐方式：left/center/right/justify |
| line_spacing | number | 否 | 行距倍数（1.0=单倍，1.5=1.5 倍） |
| first_line_indent | string | 否 | 首行缩进，如 "2em" 或 "24pt" |
| space_before | string | 否 | 段前距，如 "12pt" |
| space_after | string | 否 | 段后距，如 "12pt" |

**示例：**
- 居中对齐：`{"op": "set_paragraph_format", "target": "p0", "alignment": "center"}`
- 公文正文格式：`{"op": "set_paragraph_format", "target": "p1", "line_spacing": 1.5, "first_line_indent": "2em"}`