# Edit Operations DSL

AI must output JSON arrays containing edit operations.

## replace_text 操作

查找替换文本，保留 run 级别格式。

```json
{"op": "replace_text", "find": "A", "replace": "B", "scope": "all"}
```

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| find | string | ✅ | 要查找的文本 |
| replace | string | ✅ | 替换后的文本 |
| scope | string | 否 | 作用域，默认 "all" |

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

设置段落字体格式（段落级别，作用于该段落所有 run）。

```json
{
  "op": "set_font",
  "target": "p0",
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
| name | string | 否 | 西文字体名 |
| east_asia | string | 否 | 中文字体名 |
| size | string | 否 | 字号，如 "14pt" 或 "28"（half-points） |
| bold | boolean | 否 | 是否加粗 |
| italic | boolean | 否 | 是否斜体 |

**示例：**
- 设置整段为 Arial 14pt 加粗：`{"op": "set_font", "target": "p0", "name": "Arial", "size": "14pt", "bold": true}`
- 设置中文字体为楷体：`{"op": "set_font", "target": "p1", "east_asia": "楷体"}`

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