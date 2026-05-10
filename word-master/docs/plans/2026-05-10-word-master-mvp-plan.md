# Word Master MVP Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build the Word Master MVP featuring a dual-layer architecture for creating and editing Word documents via AI, supporting a robust "Read-Modify-Verify" cycle.

**Architecture:** 
- **Create Mode:** `md_to_docx.py` converts standard Markdown into DOCX using predefined format rules.
- **Edit Mode:** `docx_reader.py` extracts a rich Markdown summary with paragraph IDs. AI generates JSON operations. `docx_editor.py` executes operations (replace, rewrite, formatting) directly on the DOCX using `python-docx`. `docx_diff.py` generates a summary of changes.

**Tech Stack:** Python 3.10+, `python-docx` (Word manipulation), `pytest` (testing).

---

### Task 1: Project Setup & Skill Definition

**Files:**
- Create: `word-master/skills/word-master/requirements.txt`
- Create: `word-master/skills/word-master/SKILL.md`
- Create: `word-master/skills/word-master/references/edit-ops.md`

- [ ] **Step 1: Write requirements.txt**

```text
python-docx>=1.1.0
pytest>=7.4.0
```

- [ ] **Step 2: Write SKILL.md**

```markdown
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
```

- [ ] **Step 3: Write edit-ops.md**

```markdown
# Edit Operations DSL

AI must output JSON arrays containing edit operations.

- `replace_text`: `{"op": "replace_text", "find": "A", "replace": "B", "scope": "all"}`
- `rewrite_paragraph`: `{"op": "rewrite_paragraph", "target": "p0", "content": "New text"}`
- `set_font`: `{"op": "set_font", "target": {"paragraph": "p0"}, "name": "Arial", "size": "14pt"}`
- `set_paragraph_format`: `{"op": "set_paragraph_format", "target": "p0", "alignment": "center"}`
```

- [ ] **Step 4: Commit**

```bash
cd d:/melon/Documents/aiwork/office-master
mkdir -p word-master/skills/word-master/references
git add word-master/skills/word-master
git commit -m "feat(word-master): initialize project structure and skill definitions"
```

---

### Task 2: Implement docx_reader.py

**Files:**
- Create: `word-master/skills/word-master/scripts/docx_reader.py`
- Create: `word-master/tests/test_docx_reader.py`

- [ ] **Step 1: Write the failing test**

```python
# word-master/tests/test_docx_reader.py
import docx
from docx_reader import extract_rich_markdown

def test_extract_rich_markdown(tmp_path):
    doc = docx.Document()
    doc.add_heading('Chapter 1', level=1)
    doc.add_paragraph('Hello world.')
    doc_path = tmp_path / "test.docx"
    doc.save(doc_path)
    
    md_output = extract_rich_markdown(str(doc_path))
    assert "[p0]" in md_output
    assert "Chapter 1" in md_output
    assert "[p1]" in md_output
    assert "Hello world." in md_output
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest word-master/tests/test_docx_reader.py`
Expected: FAIL (ModuleNotFoundError for `docx_reader`)

- [ ] **Step 3: Write minimal implementation**

```python
# word-master/skills/word-master/scripts/docx_reader.py
import docx
import sys

def extract_rich_markdown(filepath, overview=False):
    doc = docx.Document(filepath)
    lines = []
    lines.append(f"<!-- DOCUMENT: {filepath} -->\n")
    
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        style_name = para.style.name
        
        if overview and not style_name.startswith('Heading'):
            continue
            
        lines.append(f"### [p{i}] {text} {{{style_name}}}")
        
    return "\n\n".join(lines)

if __name__ == "__main__":
    filepath = sys.argv[1]
    overview = "--overview" in sys.argv
    print(extract_rich_markdown(filepath, overview))
```

- [ ] **Step 4: Run test to verify it passes**

Run: `PYTHONPATH=word-master/skills/word-master/scripts pytest word-master/tests/test_docx_reader.py`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add word-master/tests/test_docx_reader.py word-master/skills/word-master/scripts/docx_reader.py
git commit -m "feat(word-master): implement docx_reader to extract rich markdown"
```

---

### Task 3: Implement docx_editor.py

**Files:**
- Create: `word-master/skills/word-master/scripts/docx_editor.py`
- Create: `word-master/tests/test_docx_editor.py`

- [ ] **Step 1: Write the failing test**

```python
# word-master/tests/test_docx_editor.py
import docx
import json
from docx_editor import apply_operations

def test_apply_operations(tmp_path):
    doc = docx.Document()
    doc.add_paragraph('Old text')
    in_path = tmp_path / "in.docx"
    out_path = tmp_path / "out.docx"
    doc.save(in_path)
    
    ops = [{"op": "rewrite_paragraph", "target": "p0", "content": "New text"}]
    apply_operations(str(in_path), ops, str(out_path))
    
    doc_out = docx.Document(out_path)
    assert doc_out.paragraphs[0].text == "New text"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest word-master/tests/test_docx_editor.py`
Expected: FAIL

- [ ] **Step 3: Write minimal implementation**

```python
# word-master/skills/word-master/scripts/docx_editor.py
import docx
import json
import sys

def apply_operations(filepath, ops, outpath=None):
    doc = docx.Document(filepath)
    if outpath is None:
        outpath = filepath
        
    for op in ops:
        if op['op'] == 'rewrite_paragraph':
            idx = int(op['target'].replace('p', ''))
            if idx < len(doc.paragraphs):
                # Simple replacement (clears inline styles for MVP)
                doc.paragraphs[idx].text = op['content']
        elif op['op'] == 'replace_text':
            find_txt = op['find']
            repl_txt = op['replace']
            for p in doc.paragraphs:
                if find_txt in p.text:
                    p.text = p.text.replace(find_txt, repl_txt)
                    
    doc.save(outpath)

if __name__ == "__main__":
    filepath = sys.argv[1]
    with open(sys.argv[2], 'r', encoding='utf-8') as f:
        ops = json.load(f)
    outpath = sys.argv[3] if len(sys.argv) > 3 else filepath
    apply_operations(filepath, ops, outpath)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `PYTHONPATH=word-master/skills/word-master/scripts pytest word-master/tests/test_docx_editor.py`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add word-master/tests/test_docx_editor.py word-master/skills/word-master/scripts/docx_editor.py
git commit -m "feat(word-master): implement docx_editor with basic paragraph rewriting and text replacement"
```

---

### Task 4: Implement docx_diff.py

**Files:**
- Create: `word-master/skills/word-master/scripts/docx_diff.py`
- Create: `word-master/tests/test_docx_diff.py`

- [ ] **Step 1: Write the failing test**

```python
# word-master/tests/test_docx_diff.py
import docx
from docx_diff import generate_diff

def test_generate_diff(tmp_path):
    d1 = docx.Document()
    d1.add_paragraph('Old text')
    p1 = tmp_path / "1.docx"
    d1.save(p1)
    
    d2 = docx.Document()
    d2.add_paragraph('New text')
    p2 = tmp_path / "2.docx"
    d2.save(p2)
    
    diff = generate_diff(str(p1), str(p2))
    assert "- Old text" in diff
    assert "+ New text" in diff
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest word-master/tests/test_docx_diff.py`
Expected: FAIL

- [ ] **Step 3: Write minimal implementation**

```python
# word-master/skills/word-master/scripts/docx_diff.py
import docx
import sys
import difflib

def generate_diff(path1, path2):
    d1 = docx.Document(path1)
    d2 = docx.Document(path2)
    
    texts1 = [p.text for p in d1.paragraphs if p.text.strip()]
    texts2 = [p.text for p in d2.paragraphs if p.text.strip()]
    
    diff = difflib.unified_diff(texts1, texts2, lineterm='', n=0)
    return '\n'.join(diff)

if __name__ == "__main__":
    print(generate_diff(sys.argv[1], sys.argv[2]))
```

- [ ] **Step 4: Run test to verify it passes**

Run: `PYTHONPATH=word-master/skills/word-master/scripts pytest word-master/tests/test_docx_diff.py`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add word-master/tests/test_docx_diff.py word-master/skills/word-master/scripts/docx_diff.py
git commit -m "feat(word-master): implement docx_diff for simple text comparisons"
```

---

### Task 5: Implement md_to_docx.py (Create Mode)

**Files:**
- Create: `word-master/skills/word-master/scripts/md_to_docx.py`
- Create: `word-master/tests/test_md_to_docx.py`

- [ ] **Step 1: Write the failing test**

```python
# word-master/tests/test_md_to_docx.py
import docx
from md_to_docx import convert_markdown

def test_convert_markdown(tmp_path):
    md_content = "# Title\n\nSome paragraph text."
    md_path = tmp_path / "draft.md"
    md_path.write_text(md_content)
    
    out_path = tmp_path / "out.docx"
    convert_markdown(str(md_path), str(out_path))
    
    doc = docx.Document(out_path)
    assert doc.paragraphs[0].text == "Title"
    assert doc.paragraphs[0].style.name == "Heading 1"
    assert doc.paragraphs[1].text == "Some paragraph text."
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest word-master/tests/test_md_to_docx.py`
Expected: FAIL

- [ ] **Step 3: Write minimal implementation**

```python
# word-master/skills/word-master/scripts/md_to_docx.py
import docx
import sys

def convert_markdown(md_path, out_path, template=None):
    doc = docx.Document()
    
    with open(md_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        
    for line in lines:
        line = line.strip()
        if not line:
            continue
        if line.startswith('# '):
            doc.add_heading(line[2:], level=1)
        elif line.startswith('## '):
            doc.add_heading(line[3:], level=2)
        else:
            doc.add_paragraph(line)
            
    doc.save(out_path)

if __name__ == "__main__":
    convert_markdown(sys.argv[1], sys.argv[2])
```

- [ ] **Step 4: Run test to verify it passes**

Run: `PYTHONPATH=word-master/skills/word-master/scripts pytest word-master/tests/test_md_to_docx.py`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add word-master/tests/test_md_to_docx.py word-master/skills/word-master/scripts/md_to_docx.py
git commit -m "feat(word-master): implement md_to_docx for basic markdown to docx conversion"
```
