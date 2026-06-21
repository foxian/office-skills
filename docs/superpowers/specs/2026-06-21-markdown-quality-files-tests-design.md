# markdown-master 格式检测与合并/拆分自动化测试设计

- 日期：2026-06-21
- 范围：为 `markdown-master/scripts/quality.py`（lint / zhlint / linkcheck）和 `markdown-master/scripts/files.py`（split / merge）补自动化单元测试
- 不在范围内：修改源码、远程 HTTP linkcheck 测试、CLI smoke 测试、numbering/convert/md_utils 测试（已被现有测试覆盖）

## 背景

`markdown-master/tests/` 已有 `test_numbering.py`、`test_convert_outputs.py`、`test_pdf_output.py`、`test_yaml_config.py`、`test_template.py`、`test_md_utils.py`，但 `quality.py` 和 `files.py` 两个脚本没有对应单元测试。本次补齐这部分测试覆盖。

## 实现路径

纯函数级测试，与现有 `test_numbering.py` / `test_md_utils.py` 风格一致：
- 直接 `from quality import run_lint, run_zhlint, run_linkcheck`、`from files import cmd_split, cmd_merge`
- 文档内容用 Python 字符串字面量内联，不引入额外 fixture 文件
- 文件 I/O 用 pytest 内置 `tmp_path` fixture
- linkcheck 远程 HTTP 分支不写测试（核心断言逻辑已由本地路径分支覆盖；HTTP 行为不是本工具核心，且易让 CI 不稳定）
- 不写 CLI / subprocess smoke 测试（`main()` 仅为 argparse 薄壳，业务逻辑都在被测函数里）

## 新增文件

- `markdown-master/tests/test_quality.py`
- `markdown-master/tests/test_files.py`

## test_quality.py 用例

### run_lint（检测）

1. `test_lint_clean_doc` — 正常文档无问题，返回 `[]`
2. `test_lint_missing_h1` — 无 `# 标题`，返回 `[(0, "缺少 H1 标题")]`
3. `test_lint_heading_skip` — `# A` 直接跳到 `### C`，报标题层级跳跃，行号正确
4. `test_lint_heading_skip_ignored_in_code_block` — 代码块内的 `###` 不触发跳跃告警（验证 `_precompute_code_state` 集成）
5. `test_lint_excessive_blank_lines` — 4 行连续空行，报"连续空行过多"，行号指向起始空行
6. `test_lint_code_block_no_language` — 单独 ```` ``` ```` 报"代码块未指定语言"；```` ```python ```` 不报

### run_lint（--fix 路径）

7. `test_lint_fix_collapses_blank_lines` — 4 行空行被压缩为 2 行；返回的 issues 仍包含该条（说明"发现并修复"而非"静默修复"）

### run_zhlint（检测 + fix）

8. `test_zhlint_missing_space_zh_en` — `"中文abc"` 报缺空格；fix 后变 `"中文 abc"`
9. `test_zhlint_missing_space_en_zh` — `"abc中文"` 同上
10. `test_zhlint_half_punct` — `"中文,英文"` 报半角标点；fix 后变 `"中文，英文"`
11. `test_zhlint_inline_code_protected` — ``"中文`code`abc"`` 行内代码段不被插入空格。**已知缺陷**：`run_zhlint` 主流程未调用 `_is_inline_code`，此用例会失败，用 `@pytest.mark.xfail(reason="run_zhlint 未调用 _is_inline_code，行内代码段会被错误插入空格")` 标记，不修源码

### run_linkcheck（本地路径）

12. `test_linkcheck_local_image_missing` — 引用不存在的 png，报"图片路径不存在"，行号正确
13. `test_linkcheck_local_image_exists` — `tmp_path` 下真实存在的 png，无报错
14. `test_linkcheck_local_link_missing` — `[text](missing.md)`，报"本地链接目标不存在"
15. `test_linkcheck_anchor_skipped` — `[text](#anchor)` 不报错
16. `test_linkcheck_link_in_code_block_skipped` — 代码块内的 `![](x.png)` 不报错

远程分支：不写测试，文件顶部加注释说明原因。

## test_files.py 用例

### cmd_split

1. `test_split_by_h2_basic` — `# A` + `## B` + 正文 + `## C` + 正文，拆出 `_preamble`、`01_B.md`、`02_C.md`，内容含对应标题行和正文
2. `test_split_by_h2_no_preamble` — 文档以 `##` 开头，无 preamble，第一个文件名是 `00_xxx.md`
3. `test_split_preamble_only_blank_skipped` — 只有正文无标题，preamble 全是空行 → 不输出任何文件（验证 `if title == "_preamble" and not any(...)` 分支）
4. `test_split_heading_in_code_block_ignored` — 代码块内的 `## xxx` 不触发拆分
5. `test_split_filename_sanitization` — 标题含 `\\/:*?"<>|` 等非法字符，被替换为 `_`，长度截断到 80 字符
6. `test_split_custom_output_dir` — `--output-dir` 指向 `tmp_path` 子目录，目录自动创建
7. `test_split_by_h1` — 验证 `--by h1` 按 h1 拆（`##` 不拆）

### cmd_merge

8. `test_merge_two_files` — 合并两个文件，输出含两份内容，中间用默认分隔符 `\n\n---\n\n`
9. `test_merge_custom_separator` — `--separator` 生效
10. `test_merge_directory_input` — 输入是目录，按文件名字母序合并所有 `.md`（用 `a.md` / `b.md` / `c.md` 验证顺序）
11. `test_merge_skips_non_md` — 目录里混入 `.txt` / `.png`，被跳过，只合并 `.md`
12. `test_merge_empty_inputs_exits_nonzero` — 没有任何 `.md` 文件时 `sys.exit(1)`（`pytest.raises(SystemExit)` 捕获）

### 副作用处理

- `cmd_split` / `cmd_merge` 直接 `print` 到 stdout 且 `cmd_merge` 会 `sys.exit`，不返回值
- split：调用后读 `tmp_path` 下的产出文件做断言，不验证 stdout
- merge：调用后读 `-o` 输出文件做断言
- merge 空输入：`pytest.raises(SystemExit)` 捕获退出码

## 运行约定与验收标准

### 运行方式

- `cd markdown-master && python -m pytest tests/test_quality.py tests/test_files.py -v`
- 不引入新 pytest 插件或依赖
- 不修改 conftest.py（现有 sys.path 配置已足够）

### xfail 约定

- `test_zhlint_inline_code_protected` 用 `@pytest.mark.xfail(reason="run_zhlint 未调用 _is_inline_code，行内代码段会被错误插入空格")` 标记
- spec 中明确记录这是已知缺陷，不在本次实现里修复

### 验收标准

1. 新增两文件后，`python -m pytest tests/` 全绿（含 xfail，xfail 不算失败）
2. 不破坏现有 6 个测试文件
3. 不修改 `scripts/` 下任何源码（本次只写测试）
4. 测试覆盖 test_quality.py 用例（16 个）和 test_files.py 用例（12 个）列出的全部编号
5. 每个测试函数有中文 docstring 说明被测行为

### 非目标（明确排除）

- 不修 `_is_inline_code` 的 bug（留待后续单独处理）
- 不写 linkcheck 远程 HTTP 测试
- 不写 CLI / subprocess smoke 测试
- 不补 numbering / convert / md_utils 的测试（已被现有测试覆盖）
