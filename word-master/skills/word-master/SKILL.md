---
name: word-master
description: "AI-driven Word document creation and editing system. Use when user wants to edit, review, or format a .docx file, or create a new one."
---

# Word Master Skill

## 1. Create Mode
- User discusses outline. AI writes `draft.md`.
- AI generates DOCX using: `python scripts/md_to_docx.py draft.md --template default`

## 2. Edit Mode
- Read: `python scripts/docx_reader.py <file.docx> --overview`
- AI generates operations JSON.
- Edit: `python scripts/docx_editor.py <file.docx> ops.json`
- Diff: `python scripts/docx_diff.py <file.docx> <modified.docx>`
