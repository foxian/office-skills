import docx
import json
import shutil
import sys
import os
from docx.enum.text import WD_ALIGN_PARAGRAPH

def _replace_text_in_runs(paragraph, find_txt, repl_txt):
    """Run-level text replacement to preserve inline formatting (bold, italic, font, etc.)."""
    full_text = paragraph.text
    if find_txt not in full_text:
        return False

    replaced = False
    for run in paragraph.runs:
        if find_txt in run.text:
            run.text = run.text.replace(find_txt, repl_txt)
            replaced = True

    if not replaced:
        paragraph.text = full_text.replace(find_txt, repl_txt)
        replaced = True

    return replaced

def _apply_set_font(paragraph, params):
    """Apply font settings to all runs in a paragraph."""
    for run in paragraph.runs:
        if "name" in params:
            run.font.name = params["name"]
        if "east_asia" in params:
            run.font.east_asia = params["east_asia"]
        if "size" in params:
            size_str = params["size"]
            if size_str.endswith("pt"):
                from docx.shared import Pt
                run.font.size = Pt(float(size_str.rstrip("pt")))
            elif size_str.isdigit():
                from docx.shared import Pt
                run.font.size = Pt(int(size_str) / 2)
        if "bold" in params:
            run.bold = params["bold"]
        if "italic" in params:
            run.italic = params["italic"]

def _apply_set_paragraph_format(paragraph, params):
    """Apply paragraph format settings."""
    pf = paragraph.paragraph_format
    if "alignment" in params:
        align_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY
        }
        pf.alignment = align_map.get(params["alignment"], WD_ALIGN_PARAGRAPH.LEFT)
    if "line_spacing" in params:
        pf.line_spacing = params["line_spacing"]
    if "first_line_indent" in params:
        indent_str = params["first_line_indent"]
        if indent_str.endswith("em"):
            from docx.shared import Emu
            pf.first_line_indent = Emu(int(float(indent_str.rstrip("em")) * 914400 / 2))
        elif indent_str.endswith("pt"):
            from docx.shared import Pt
            pf.first_line_indent = Pt(float(indent_str.rstrip("pt")))
    if "space_before" in params:
        space_str = params["space_before"]
        if space_str.endswith("pt"):
            from docx.shared import Pt
            pf.space_before = Pt(float(space_str.rstrip("pt")))
    if "space_after" in params:
        space_str = params["space_after"]
        if space_str.endswith("pt"):
            from docx.shared import Pt
            pf.space_after = Pt(float(space_str.rstrip("pt")))

def _backup_file(filepath):
    """Create a .bak backup of the original file before modifying."""
    bak_path = filepath + '.bak'
    shutil.copy2(filepath, bak_path)
    return bak_path

def apply_operations(filepath, ops, outpath=None):
    doc = docx.Document(filepath)
    if outpath is None:
        outpath = filepath
    
    # Backup original before any modification
    if os.path.exists(filepath):
        _backup_file(filepath)
        
    for op in ops:
        if op['op'] == 'rewrite_paragraph':
            idx = int(op['target'].replace('p', ''))
            if idx < len(doc.paragraphs):
                para = doc.paragraphs[idx]
                saved_style = para.style
                para.text = op['content']
                para.style = saved_style
        elif op['op'] == 'replace_text':
            find_txt = op['find']
            repl_txt = op['replace']
            for p in doc.paragraphs:
                _replace_text_in_runs(p, find_txt, repl_txt)
        elif op['op'] == 'set_font':
            idx = int(op['target'].replace('p', ''))
            if idx < len(doc.paragraphs):
                _apply_set_font(doc.paragraphs[idx], op)
        elif op['op'] == 'set_paragraph_format':
            idx = int(op['target'].replace('p', ''))
            if idx < len(doc.paragraphs):
                _apply_set_paragraph_format(doc.paragraphs[idx], op)

    doc.save(outpath)

if __name__ == "__main__":
    filepath = sys.argv[1]
    with open(sys.argv[2], 'r', encoding='utf-8') as f:
        ops = json.load(f)
    outpath = sys.argv[3] if len(sys.argv) > 3 else filepath
    apply_operations(filepath, ops, outpath)
