# 灵活编号系统实施计划

**Goal**: 用每级独立模板系统替换 markdown-master 硬编码的 4 种编号样式，并修复所有 9 处 `is_in_code_block` 调用的 O(n²) 性能问题。

## 验收清单（最终确认）

实施完成后请确认以下项目：

- [x] `python -m pytest tests/ -v` 全部通过
- [x] `python scripts/structure.py doc.md numbering add --h1 "第{1}章 " --h2 "{1}.{2} "` 正常工作
- [x] `python scripts/structure.py doc.md numbering add --config examples/numbering_chinese_chapter.yaml` 正常工作
- [x] `python scripts/structure.py doc.md numbering add --h1 "A" --save-config out.yaml` 只写盘不修改原文档
- [x] 代码块内的 `#` 不会被编号
- [x] `--style` 完全消失（`python scripts/structure.py --help` 中不再出现）
- [x] `SKILL.md` 反映新 CLI 用法
- [x] Git 提交历史清晰，每步可独立回滚
