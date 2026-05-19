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

## insert_paragraph 操作

在指定段落锚点前或后插入新段落。

```json
{
  "op": "insert_paragraph",
  "after": "p5",
  "content": "新段落文字",
  "style": "Normal"
}
```

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| after/before | string | ✅ (二选一) | 定位锚点。`before: "p0"` 表示在首段前插入。互斥参数 |
| content | string | ✅ | 新段落文字内容 |
| style | string | 否 | 段落样式名。省略时默认继承锚点段落的样式 |

## delete_paragraph 操作

删除目标段落。

```json
{"op": "delete_paragraph", "target": "p5"}
```

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| target | string | ✅ | 目标段落 ID |

## modify_cell 操作

修改表格指定单元格内容。

```json
{"op": "modify_cell", "table": 0, "row": 1, "col": 2, "content": "新内容"}
```

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| table | int | ✅ | 表格索引 (从 0 开始) |
| row | int | ✅ | 行索引 (从 0 开始) |
| col | int | ✅ | 列索引 (从 0 开始) |
| content | string | ✅ | 单元格新文本内容 |

## insert_row 操作

在表格某行之前或之后插入空行。

```json
{"op": "insert_row", "table": 0, "after": 2}
```

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| table | int | ✅ | 表格索引 (从 0 开始) |
| after/before | int | ✅ (二选一) | 行号锚点 |

## insert_table 操作

在指定段落锚点前或后插入新表格。

```json
{"op": "insert_table", "after": "p3", "rows": 3, "cols": 4}
```

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| after/before | string | ✅ (二选一) | 段落锚点 |
| rows | int | ✅ | 初始行数 |
| cols | int | ✅ | 初始列数 |

## set_header_footer 操作

配置指定节的页眉和页脚。

```json
{
  "op": "set_header_footer",
  "header": "页眉文字",
  "footer": "页脚文字",
  "section": 0,
  "even_page": false
}
```

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| header | string | 否 | 页眉文字内容 |
| footer | string | 否 | 页脚文字内容 |
| section | int | 否 | 节索引，默认 `0` |
| even_page | boolean | 否 | 是否修改偶数页页眉/页脚，默认 `false` |

## insert_page_break 操作

在指定段落锚点前或后插入分页符。

```json
{"op": "insert_page_break", "after": "p10"}
```

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| after/before | string | ✅ (二选一) | 段落锚点 |

## apply_style 操作

强制应用预定义样式。

```json
{"op": "apply_style", "target": "p2", "style": "Heading 1"}
```

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| target | string | ✅ | 目标段落 ID |
| style | string | ✅ | 样式名 |
| clear_run_formats | boolean | 否 | 赋样式后是否清除段落内所有 run 的手动格式覆盖（bold/size/font 等设为 None），默认 false |

## set_page_setup 操作

设置页面边距与纸张方向。

```json
{
  "op": "set_page_setup",
  "margin_top": "2.54cm",
  "margin_bottom": "2.54cm",
  "margin_left": "3.17cm",
  "margin_right": "3.17cm",
  "orientation": "portrait",
  "section": 0
}
```

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| margin_top | string | 否 | 顶部边距，支持 `"Xcm"` 或 `"Xpt"` |
| margin_bottom | string | 否 | 底部边距 |
| margin_left | string | 否 | 左侧边距 |
| margin_right | string | 否 | 右侧边距 |
| orientation | string | 否 | 纸张方向：`portrait`（纵向）或 `landscape`（横向） |
| section | int | 否 | 节索引，默认 `0` |

## update_style_definition 操作

修改文档内置样式的字体与段落格式定义，影响所有应用了该样式的段落。

```json
{
  "op": "update_style_definition",
  "style": "Heading 1",
  "fingerprint": {
    "size": "16.0pt",
    "font": "黑体",
    "bold": true,
    "align": "center",
    "space_before": "12.0pt",
    "space_after": "6.0pt",
    "line_spacing": 1.5,
    "color": "rgb:000000"
  }
}
```

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| style | string | ✅ | 内置样式名，如 "Heading 1"、"Normal" |
| fingerprint | object | ✅ | 格式指纹对象，与 style_profile.json 中的 fingerprint 格式相同；各字段均为可选 |

**fingerprint 字段：**

| 字段 | 类型 | 说明 |
|------|------|------|
| size | string | 字号，如 `"16.0pt"` |
| font | string | 中文/西文字体名（同时写入 w:eastAsia），如 `"黑体"` |
| bold | boolean | 加粗 |
| italic | boolean | 斜体 |
| align | string | 对齐：left/center/right/justify |
| space_before | string | 段前距，如 `"12.0pt"` |
| space_after | string | 段后距 |
| line_spacing | number | 行距倍数（1.5=1.5倍） |
| color | string | 颜色，如 `"rgb:FF0000"` |