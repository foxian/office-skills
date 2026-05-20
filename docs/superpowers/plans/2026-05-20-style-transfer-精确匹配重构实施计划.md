# style_transfer.py 精确匹配重构实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 重构 style_transfer.py，移除正则文本匹配逻辑，改为精确名称匹配 + 降级策略

**Architecture:** 修改 `generate_apply_ops` 函数的匹配逻辑，移除 `_infer_role_from_text` 和相关的正则匹配，改为简单的名称匹配和降级映射

**Tech Stack:** Python, python-docx

---

## 文件结构

- **Modify**: `d:\melon\Documents\aiwork\office-master\.trae\skills\word-master\scripts\style_transfer.py`

---

## 实施步骤

### Task 1: 移除正则匹配相关代码

**Files:**
- Modify: `d:\melon\Documents\aiwork\office-master\.trae\skills\word-master\scripts\style_transfer.py:1-39`

- [ ] **Step 1: 移除正则模式定义 (lines 12-21)**

删除：
```python
_HEADING1_PATTERN = re.compile(
    r'^(第[一二三四五六七八九十\d]+篇|第[一二三四五六七八九十\d]+章)\s'
)
_HEADING2_PATTERN = re.compile(
    r'^(\d+\.\d*\s|[一二三四五六七八九十]+、\s)'
)
_HEADING3_PATTERN = re.compile(
    r'^(\d+\.\d+\.\d*\s|（[一二三四五六七八九十\d]+）\s|[一二三四五六七八九十]+、\d)'
)
```

- [ ] **Step 2: 移除 `_infer_role_from_text` 函数 (lines 115-136)**

删除整个函数：
```python
def _infer_role_from_text(text, profile):
    ...
```

- [ ] **Step 3: 添加降级映射常量**

在移除正则后，添加：
```python
_FALLBACK_MAP = {
    "Body Text": "Normal",
}
```

- [ ] **Step 4: 提交变更**

```bash
git add .trae/skills/word-master/scripts/style_transfer.py
git commit -m "refactor: remove regex patterns for heading inference"
```

---

### Task 2: 重构 generate_apply_ops 函数

**Files:**
- Modify: `d:\melon\Documents\aiwork\office-master\.trae\skills\word-master\scripts\style_transfer.py:153-215`

- [ ] **Step 1: 读取当前 generate_apply_ops 函数**

确认需要修改的部分。

- [ ] **Step 2: 重构 Level 2 逻辑 (约 lines 200-211)**

将原来的复杂逻辑：
```python
# Level 2: Body-style paragraphs — try text pattern, then default to Normal
if role is None:
    if style_name in ("Normal", "Body Text", "Default Paragraph Style"):
        role = _infer_role_from_text(text, profile)
        if role is None and "Normal" in available_roles:
            role = "Normal"
```

替换为：
```python
# 精确匹配：样式名在 profile 中定义
if style_name in available_roles:
    role = style_name
# 降级匹配：Body Text -> Normal
elif style_name in _FALLBACK_MAP:
    fallback = _FALLBACK_MAP[style_name]
    if fallback in available_roles:
        role = fallback
# 其他样式：保持原样，不生成 apply_style 操作
```

- [ ] **Step 3: 简化变量赋值逻辑**

删除 `text` 变量（不再需要），简化角色赋值。

- [ ] **Step 4: 提交变更**

```bash
git add .trae/skills/word-master/scripts/style_transfer.py
git commit -m "refactor: use precise style name matching with fallback"
```

---

### Task 3: 测试验证

**Files:**
- Test: `d:\melon\Documents\aiwork\office-master\word-master\examples\安全加固投标书.docx`

- [ ] **Step 1: 重新执行样式转换**

```bash
python .trae/skills/word-master/scripts/style_transfer.py \
  --profile word-master/examples/style_profile.json \
  word-master/examples/安全加固投标书.docx \
  word-master/examples/安全加固投标书_v2.docx
```

- [ ] **Step 2: 验证 Body Text 保持不变**

```bash
python .trae/skills/word-master/scripts/docx_reader.py \
  word-master/examples/安全加固投标书_v2.docx \
  --overview | Select-String "第五章 技术要求"
```

预期：`p281` 显示 `{Body Text}`，不是 `{Heading 1}`

- [ ] **Step 3: 验证真正标题保持 Heading 1**

```bash
python .trae/skills/word-master/scripts/docx_reader.py \
  word-master/examples/安全加固投标书_v2.docx \
  --overview | Select-String "一、技术方案"
```

预期：`p5` 显示 `{Heading 1}`

- [ ] **Step 4: 提交测试结果**

```bash
git add word-master/examples/
git commit -m "test: verify style transfer v2 with precise matching"
```

---

### Task 4: 更新 SKILL.md 文档

**Files:**
- Modify: `d:\melon\Documents\aiwork\office-master\word-master\skills\word-master\SKILL.md`

- [ ] **Step 1: 添加精确匹配行为说明**

在 Workflow 3 部分添加说明：
> **匹配规则**：样式转换采用精确名称匹配。目标文档中 `Heading X` 样式会套用模板中对应 `Heading X` 的 fingerprint；`Body Text` 会降级套用 `Normal` 的 fingerprint；其他未匹配样式保持原样。

- [ ] **Step 2: 提交变更**

```bash
git add word-master/skills/word-master/SKILL.md
git commit -m "docs: add precise matching behavior to SKILL.md"
```

---

## 完成后

所有任务完成后，整理提交：
```bash
git log --oneline -5
```

预期输出：
```
test: verify style transfer v2 with precise matching
docs: add precise matching behavior to SKILL.md
refactor: use precise style name matching with fallback
refactor: remove regex patterns for heading inference
```
