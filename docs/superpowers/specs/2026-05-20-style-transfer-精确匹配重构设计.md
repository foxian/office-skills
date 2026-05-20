# style_transfer.py 精确匹配重构设计

## 背景

当前 `style_transfer.py` 使用正则表达式匹配文本来识别标题层级，导致正文中的引用文本被误识别为标题。

**问题案例**：
- 目标文档中 `"第五章 技术要求 > 5.2.3 关键技术参数表："` 原样式是 `Body Text`
- 因文本以 `"第五章 "` 开头，被正则 `^(第[一二三四五六七八九十\d]+章)\s` 匹配
- 被错误升级为 `Heading 1`

## 目标

移除文本模式匹配，改为精确名称匹配 + 降级策略，确保样式转换的确定性。

## 设计方案

### 1. 移除的代码

- `_HEADING1_PATTERN`、`_HEADING2_PATTERN`、`_HEADING3_PATTERN` 正则定义
- `_infer_role_from_text` 函数
- Level 2 的文本模式 fallback 逻辑

### 2. 新增降级映射

```python
_FALLBACK_MAP = {
    "Body Text": "Normal",
}
```

### 3. 新的匹配逻辑

```python
def generate_apply_ops(draft_path, profile, skip_head=0, skip_tail=0):
    # Phase 1: update_style_definition ops (保持不变)
    ops = []
    for entry in profile.get("roles", []):
        role = entry.get("role")
        fp = entry.get("fingerprint")
        if role and fp:
            ops.append({
                "op": "update_style_definition",
                "style": role,
                "fingerprint": fp,
            })

    # Phase 2: 精确匹配逻辑
    doc = docx.Document(draft_path)
    paras = doc.paragraphs
    n = len(paras)
    available_roles = {entry["role"] for entry in profile.get("roles", [])}

    for i, para in enumerate(paras):
        if i < skip_head or i >= n - skip_tail:
            continue

        style_name = para.style.name
        role = None

        # 精确匹配：样式名在 profile 中定义
        if style_name in available_roles:
            role = style_name
        # 降级匹配：样式名有降级映射
        elif style_name in _FALLBACK_MAP:
            fallback = _FALLBACK_MAP[style_name]
            if fallback in available_roles:
                role = fallback
        # 其他样式：保持原样，不生成 apply_style 操作

        if role:
            ops.append({
                "op": "apply_style",
                "target": f"p{i}",
                "style": role,
                "clear_run_formats": True,
            })

    return ops
```

### 4. 匹配规则表

| 目标文档样式 | profile 有定义 | profile 无定义 |
|-------------|---------------|---------------|
| `Heading 1/2/3...` | 套用 fingerprint | 保持原样式 |
| `List Bullet/Number` | 套用 fingerprint | 保持原样式 |
| `Body Text` | 降级到 Normal | 降级到 Normal |
| `Normal` | 套用 fingerprint | 保持原样式 |
| 其他样式 | 保持原样式 | 保持原样式 |

### 5. fingerprint 应用

保持不变，`apply_style` 操作会应用 fingerprint 中的所有属性：
- `size`, `bold`, `italic`, `align`, `font`, `color`
- `line_spacing`, `space_before`, `space_after`

## 修改文件

- `d:\melon\Documents\aiwork\office-master\.trae\skills\word-master\scripts\style_transfer.py`

## 测试验证

1. 使用目标文档 `安全加固投标书.docx` 重新执行样式转换
2. 验证 `"第五章 技术要求 > 5.2.3 关键技术参数表："` 段落保持 `Body Text`
3. 验证 `"一、技术方案"` 等真正标题保持 `Heading 1`
