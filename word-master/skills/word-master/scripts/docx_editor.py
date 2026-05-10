import docx
import json
import shutil
import sys
import os

def _replace_text_in_runs(paragraph, find_txt, repl_txt):
    """Run-level text replacement to preserve inline formatting (bold, italic, font, etc.)."""
    # Build the full paragraph text from runs
    full_text = paragraph.text
    if find_txt not in full_text:
        return False
    
    # Strategy: for each run, try to replace within that single run first.
    # This handles the common case where the search text falls within one run.
    replaced = False
    for run in paragraph.runs:
        if find_txt in run.text:
            run.text = run.text.replace(find_txt, repl_txt)
            replaced = True
    
    # If the search text spans multiple runs, fall back to paragraph-level replacement.
    # This is a known limitation — cross-run replacement loses inline formatting.
    if not replaced:
        paragraph.text = full_text.replace(find_txt, repl_txt)
        replaced = True
    
    return replaced

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
                # Preserve paragraph-level style, replace content only
                saved_style = para.style
                para.text = op['content']
                para.style = saved_style
        elif op['op'] == 'replace_text':
            find_txt = op['find']
            repl_txt = op['replace']
            for p in doc.paragraphs:
                _replace_text_in_runs(p, find_txt, repl_txt)
                    
    doc.save(outpath)

if __name__ == "__main__":
    filepath = sys.argv[1]
    with open(sys.argv[2], 'r', encoding='utf-8') as f:
        ops = json.load(f)
    outpath = sys.argv[3] if len(sys.argv) > 3 else filepath
    apply_operations(filepath, ops, outpath)
