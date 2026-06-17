# markdown-master 灵活编号系统设计文档

> 日期：2026-06-17
> 范围：`markdown-master/scripts/structure.py` 编号功能 + `_md_utils.py` 性能优化
> 状态：待用户审核

## 背景与目标

`markdown-master` 当前在 [structure.py](file:///d:/DevProjects/office-master/markdown-master/scripts/structure.py) 中通过单一 `--style` 参数提供 4 种硬编码编号样式（`technical` / `chinese_chapter` / `chinese_bidding` / `academic`）。这存在三个问题：

1. **灵活性差**：用户不能混搭不同级别（例如 h1 用罗马数字、h2 用中文括号）
2. **可扩展性差**：新增样式必须改代码
3. **可复用性差**：写好的样式无法保存为模板给不同文档复用

本次设计目标是：**让每一级标题的编号格式可以独立配置，支持 CLI 参数和 YAML 配置文件两种入口**。

附带优化：解决 `_md_utils.is_in_code_block` 的 O(n²) 性能瓶颈（9 个调用点全部改用预计算 + O(1) 查表）。

---

## 范围与非范围

**在范围内：**
- `structure.py` 的 `numbering add` 子命令重写
- `numbering remove` 的前缀清理正则扩展
- `_md_utils.py` 新增 `_precompute_code_state` 和 `is_fence_line`
- 13 个调用点的性能优化（`structure.py` × 4 + `quality.py` × 3 + `files.py` × 1 + `_md_utils.py` × 1 + 自身 × 1 — 实际为 9 个内联调用点 + 自身递归）
- 新增 YAML 读写（依赖 `pyyaml`）
- `SKILL.md` 同步更新

**不在范围内（明确推迟）：**
- AST 解析方案（设计阶段已讨论并被否决）
- `heading_shift` / `toc_generate` / `extract_headings` 的额外功能
- `quality.py` 自身的功能扩展
- GUI / Web UI

---

## 文件结构

```
markdown-master/
├── SKILL.md                              ← 更新 numbering 章节
├── examples/
│   ├── sample_cn.md                     ← 不动
│   ├── sample_en.md                     ← 不动
│   └── numbering_chinese_chapter.yaml  ← 新增示例配置
└── scripts/
    ├── _md_utils.py                      ← + is_fence_line / + _precompute_code_state
    ├── structure.py                      ← 重写 numbering add / 扩展正则 / 13 处性能优化
    ├── quality.py                        ← 3 处性能优化
    ├── files.py                          ← 1 处性能优化
    └── convert.py                        ← 不动
```

---

## 设计决策摘要

| 维度 | 决策 | 备选 |
|------|------|------|
| 配置入口 | CLI 参数（每级一个）+ YAML 文件 | 仅文件 / 仅 CLI |
| 占位符语法 | `{N}` + `{N:fmt}` 修饰符后缀 | 函数名 / Python f-string |
| 缺失级别 | 不输出编号 | 自动补全 / 报错 |
| 旧 `--style` | 彻底删除 | 保留兼容 / deprecation 警告 |
| 性能优化 | 全 9 处改用预计算 | 仅新功能用 / 不优化 |
| 围栏识别 | 抽 `is_fence_line` 辅助 | 保持内联 |
| AST 解析 | 不采用 | 采用 |
| 配置文件格式 | YAML（依赖 pyyaml） | JSON / INI |
| 起始编号作用 | 仅 `--h1` 模板时影响 h1 | 影响所有级 |
| `--save-config` 行为 | 只写盘不渲染 | 写盘同时渲染 |

---

## 详细设计

### 1. 模板语言

#### 1.1 占位符语法

```
{N}         引用第 N 级计数器的当前值（默认十进制）
{N:fmt}     用指定格式渲染第 N 级计数器
```

`N` 范围 1–6，对应 h1–h6。

#### 1.2 修饰符集合（最终 8 种）

| 写法 | 含义 | 计数器=3 时输出 |
|------|------|------------------|
| `{1}` / `{1:d}` | 十进制 | `3` |
| `{1:02d}` | 两位补零 | `03` |
| `{1:03d}` | 三位补零 | `003` |
| `{1:R}` | 大写罗马 | `III` |
| `{1:r}` | 小写罗马 | `iii` |
| `{1:A}` | 大写字母 | `C` |
| `{1:a}` | 小写字母 | `c` |
| `{1:cn}` | 中文数字 | `三` |

#### 1.3 实现

- `d` 和 `0Nd` 走 Python 内置 `format(n, fmt)`
- `R` / `r` 复用 `numbering_add` 现成罗马转换表
- `A` / `a` 用 `chr(64+n)` / `chr(96+n)`（n ≤ 26）
- `cn` 用 `["零","一","二",…,"十","百",…]` 表（0–999 范围）

#### 1.4 模板渲染失败时

- 加载时（CLI 解析 / YAML 加载）就校验模板，错误直接 `sys.exit(1)`
- 渲染时（计数器 > 999 + cn）输出 `###` + 警告到 stderr，不致命

---

### 2. 配置文件格式

#### 2.1 文件结构

```yaml
# numbering.yaml — 编号样式描述（跨文件复用）
h1: "第{1}章 "
h2: "{1}.{2} "
h3: "（{3}）"
h4: ""
h5: ""
h6: ""
start_from: 1
```

#### 2.2 字段

| 字段 | 类型 | 必填 | 默认 | 说明 |
|------|------|------|------|------|
| `h1`..`h6` | str | 否* | `None` | 6 级模板；空串 = 该级不输出 |
| `start_from` | int | 否 | `1` | h1 起始编号 |

*「否」= 至少一个非空 h1..h6，否则加载报错。

#### 2.3 加载规则

1. **额外字段忽略**（YAML 演化兼容）
2. **类型严格**：错的类型 → `ValueError`
3. **模板语法校验**：加载时跑 `parse_template()`，错就拒
4. **空值规范化**：`h3:` (None) ↔ `h3: ""` 等价

#### 2.4 写盘规整化

- 6 级全部写出（未指定写 `""`）
- `start_from` 写默认值
- 顶层加注释行说明用途
- 字段顺序固定：`h1..h6` → `start_from`
- 字符串双引号

---

### 3. CLI 表面

#### 3.1 三种用法

**纯命令行：**
```bash
python scripts/structure.py <file> numbering add \
  --h1 '第{1}章 ' --h2 '{1}.{2} ' --h3 '（{3}）' \
  --start-from 1 [-o output.md]
```

**配置文件：**
```bash
python scripts/structure.py <file> numbering add \
  --config numbering.yaml [-o output.md]
```

**混合（config 为默认，CLI 覆盖）：**
```bash
python scripts/structure.py <file> numbering add \
  --config numbering.yaml --h2 '（{2}）'
```

#### 3.2 存盘

```bash
python scripts/structure.py <file> numbering add \
  --h1 '第{1}章 ' --h2 '{1}.{2} ' \
  --save-config my-style.yaml
```

**关键**：`--save-config` **不执行** numbering 渲染，只落盘 YAML（避免误覆盖原文件）。

#### 3.3 参数表

| 参数 | 类型 | 默认 | 说明 |
|------|------|------|------|
| `--h1`..`--h6` | str | `None` | 6 级模板 |
| `--config` | path | `None` | 加载 YAML |
| `--save-config` | path | `None` | 存盘 YAML |
| `--start-from` | int | 1 | h1 起始编号 |

#### 3.4 优先级

```
最终模板 = config[1..6] → 用 CLI --h1..--h6 覆盖对应级
```

#### 3.5 已删除

- `--style` 整个参数（彻底废弃，不留 deprecation）

---

### 4. 性能优化

#### 4.1 问题

`_md_utils.is_in_code_block` 每次调用 O(line_index) 扫描。在 9 个调用点的循环里被反复调用，整体 O(n²)。

#### 4.2 方案

新增 `_md_utils._precompute_code_state(lines)`，一次 O(n) 扫描返回 `list[bool]`。调用点改成"循环前预计算 + O(1) 查表"。

```python
def _precompute_code_state(lines):
    state = [False] * len(lines)
    in_block = False
    for i, line in enumerate(lines):
        state[i] = in_block
        if is_fence_line(line):
            in_block = not in_block
    return state
```

#### 4.3 调用点改动清单

| 文件 | 行 | 函数 | 改动 |
|------|-----|------|------|
| `structure.py` | L19 | `heading_shift` | 循环前 `code_state = _precompute_code_state(lines)`，循环里 `if code_state[i]:` |
| `structure.py` | L43 | `numbering_remove` | 同上 |
| `structure.py` | L67 | `numbering_add_flex`（新） | 同上 |
| `structure.py` | L139 | `toc_generate` | 同上 |
| `quality.py` | L32 | lint | 同上 |
| `quality.py` | L87 | zhlint | 同上 |
| `quality.py` | L126 | linkcheck | 同上 |
| `files.py` | L37 | split/merge | 同上 |
| `_md_utils.py` | L58 | `extract_headings` | 同上 |

#### 4.4 向后兼容

- `is_in_code_block` 函数保留不删（外部脚本可能直接 import）
- `_precompute_code_state` 是内部函数（下划线开头）

#### 4.5 围栏识别抽离

```python
_FENCE_RE = re.compile(r"^(`{3,}|~{3,})")

def is_fence_line(line):
    """判断单行是否是代码块围栏（``` 或 ~~~ 开头）。"""
    return bool(_FENCE_RE.match(line))
```

- `is_in_code_block` 内部改用 `is_fence_line`
- `_precompute_code_state` 内部改用 `is_fence_line`
- 正则编译一次，匹配速度微提升

---

### 5. 主流程数据流

```
CLI 解析
  ↓
合并 config + CLI → level_templates: dict[1..6] = str | None
  ↓
若 6 个全 None → 报错退出
  ↓
读取文件 → lines: list[str]
  ↓
预计算 code_state: list[bool]    (O(n))
  ↓
初始化 counters = [start_from-1, 0, 0, 0, 0, 0]
  ↓
遍历 lines
  ├─ 跳过 code_state[i] = True 的行
  ├─ 匹配 (^#{1,6})\s+(.*)
  ├─ level = len(hashes)
  ├─ 清理 _NUMBER_PREFIX 前缀
  ├─ counters[level-1] += 1, counters[level..5] = 0
  ├─ 若 templates[level-1] is None → 原文输出
  └─ 否则 format_template(templates[level-1], counters, level) + text
  ↓
result: str
  ↓
若 --save-config → 写 YAML（不渲染）
否则 → 写 result 到 output 或原文件
```

---

### 6. 错误处理

| 类别 | 场景 | 行为 |
|------|------|------|
| 模板 | `{7}` / `{abc}` / `{1:Q}` / `{1:0d}` | 加载时拒 |
| 配置 | `--config` 文件不存在 / YAML 解析失败 | 报错退出 |
| 配置 | 6 级全空 | 报错 `至少需要为一个级别提供模板` |
| 配置 | `--config` 和 `--save-config` 同时给 | 报错 |
| 渲染 | cn > 999 | 警告 + 输出 `###`（3 个井号作为占位符，提醒用户已超界） |
| 渲染 | 字母 > 26 | 走 `AA, AB, ...` |
| 渲染 | 罗马 > 3999 | 走 `MMMM...` |
| 模板 | h1 缺，h2 有 | h2 模板的 `{1}` 渲染为 `0` |

**非阻塞原则**：加载期任何错误直接 `sys.exit(1)`、不修改原文件；渲染期 warning 不致命、结果仍写入。

---

### 7. `numbering_remove` 扩展

扩展 `_NUMBER_PREFIX` 正则，覆盖更广泛的前缀形式：

```python
_NUMBER_PREFIX = re.compile(
    r"^(?:第[\d零一二三四五六七八九十百千万]+章\s*"
    r"|[\d零一二三四五六七八九十百千万]+[、，.．]\s*"
    r"|[\d零一二三四五六七八九十百千万]+\s+"
    r"|\([\d零一二三四五六七八九十百千万]+\)\s*"
    r"|（[\d零一二三四五六七八九十百千万]+）\s*"
    r"|【[\d零一二三四五六七八九十百千万]+】\s*"
    r"|[IVXLCDMivxlcdm]+\.\s+"
    r"|[A-Za-z]\.\s+"
    r"|\s+"
    r")"
)
```

仍走"行首剥前缀"策略，不完美但 95% 覆盖。漏网的用户可用 `--hN ''` 重置或手工修。

---

## 模块边界与依赖

### 文件改动清单

| 文件 | 改动类型 | 内容 |
|------|----------|------|
| `_md_utils.py` | 新增 + 改 | 加 `is_fence_line`、`_precompute_code_state`；改 `is_in_code_block`、`extract_headings` 内部 |
| `structure.py` | 重写 + 改 | 重写 `numbering_add` 为 `numbering_add_flex`；删 `numbering_add` 旧版；扩展 `_NUMBER_PREFIX`；`main()` 改 argparse；4 处性能优化 |
| `quality.py` | 改 | 3 处性能优化 |
| `files.py` | 改 | 1 处性能优化 |
| `SKILL.md` | 改 | 更新 numbering 章节（4 处示例改成 `--hN` 形式） |
| `examples/numbering_chinese_chapter.yaml` | 新增 | 示例配置文件 |

### 新增依赖

- `pyyaml`（可选，未安装时给清晰错误提示）

---

## 测试策略

### 单元测试（pytest，纯函数）

| 函数 | 用例 |
|------|------|
| `parse_template` | 字面、占位符、修饰符、错误输入（7种、abc、未知修饰符、宽度0） |
| `format_template` | 每个修饰符各一例、嵌套引用、上级未触发（=0） |
| `_precompute_code_state` | 无围栏、单对围栏、多对围栏、未闭合 |
| `is_fence_line` | 三反引号、四反引号、波浪、内嵌 |
| `numbering_add_flex` | 6 级全覆盖、部分覆盖、含代码块、YAML 加载、CLI 覆盖 |

### 集成测试

| 场景 | 期望 |
|------|------|
| `--h1 '第{1}章 '` + 中文文档 | h1 输出 `第1章`，h2/h3 不变 |
| `--config yml` + `--h2 '（{2}）'` | h2 来自 CLI |
| `--save-config` | 写盘，不修改原 md |
| 10 万行文档 numbering add | < 1 秒 |

### 兼容性

- 保留所有旧 CLI 参数（除 `--style`）
- `numbering remove` 接口不变，行为扩展
- `heading_shift` / `toc_generate` 接口完全不变

---

## 风险与回退

| 风险 | 缓解 |
|------|------|
| `_NUMBER_PREFIX` 扩展可能误伤（如把普通句子当编号） | 现有逻辑就是 best-effort，先扩展、用例覆盖 |
| YAML 解析错误信息不友好 | 自定义错误包装，把 PyYAML 原始错误转成"行 N: ..."格式 |
| 9 处性能改动可能引入 bug | 每处改完跑原有测试（lint / zhlint / linkcheck） |
| `pyyaml` 安装失败 | 模块顶部 try/except，import 失败时给清晰错误 |

---

## 实施步骤（高层）

1. `_md_utils.py` 加 `is_fence_line` + `_precompute_code_state`，改 `is_in_code_block` 用新辅助
2. 9 个调用点全部改为"预计算 + 查表"
3. 写 `parse_template` / `format_template` 单元测试
4. 实现 `numbering_add_flex` 替换 `numbering_add`
5. 扩展 `_NUMBER_PREFIX`
6. `main()` 改 argparse（加 `--h1..--h6`、`--config`、`--save-config`；删 `--style`）
7. 实现 YAML 加载/保存
8. 集成测试 + 端到端测试
9. 更新 `SKILL.md`
10. 提交

---

## 待用户审核事项

- [ ] 性能优化范围：9 处全改（vs 仅新功能用）
- [ ] 旧 `--style` 彻底删除（vs deprecation 警告）
- [ ] 起始编号仅作用于 h1
- [ ] `--save-config` 只写不渲染
- [ ] 加载期错误直接退出（不修改原文件）
- [ ] YAML 字段顺序固定为 `h1..h6` → `start_from`

确认后进入实施计划（writing-plans）。
