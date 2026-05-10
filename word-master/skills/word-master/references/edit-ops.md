# Edit Operations DSL

AI must output JSON arrays containing edit operations.

- `replace_text`: `{"op": "replace_text", "find": "A", "replace": "B", "scope": "all"}`
- `rewrite_paragraph`: `{"op": "rewrite_paragraph", "target": "p0", "content": "New text"}`
- `set_font`: `{"op": "set_font", "target": {"paragraph": "p0"}, "name": "Arial", "size": "14pt"}`
- `set_paragraph_format`: `{"op": "set_paragraph_format", "target": "p0", "alignment": "center"}`
