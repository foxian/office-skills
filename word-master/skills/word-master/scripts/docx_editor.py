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

def _copy_run_format(src_run, dest_run):
    """Copy formatting from src_run to dest_run"""
    dest_run.bold = src_run.bold
    dest_run.italic = src_run.italic
    dest_run.underline = src_run.underline
    dest_run.style = src_run.style
    if src_run.font:
        if src_run.font.name:
            dest_run.font.name = src_run.font.name
        if src_run.font.size:
            dest_run.font.size = src_run.font.size
        if src_run.font.color and src_run.font.color.rgb:
            dest_run.font.color.rgb = src_run.font.color.rgb
        if src_run.font.highlight_color:
            dest_run.font.highlight_color = src_run.font.highlight_color
        dest_run.font.strike = src_run.font.strike
        dest_run.font.subscript = src_run.font.subscript
        dest_run.font.superscript = src_run.font.superscript

def _apply_font_attributes(run, params):
    if "name" in params:
        run.font.name = params["name"]
    if "east_asia" in params:
        try:
            run.font.east_asia = params["east_asia"]
        except AttributeError:
            pass
    if "size" in params:
        size_str = str(params["size"])
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

def _apply_set_font(paragraph, params):
    """Apply font settings to paragraph runs or specific matched text."""
    text_match = params.get("text_match")
    
    if not text_match:
        for run in paragraph.runs:
            _apply_font_attributes(run, params)
        return

    full_text = paragraph.text
    if text_match not in full_text:
        return
        
    match_index = params.get("match_index", "all")
    match_ranges = []
    start = 0
    while True:
        idx = full_text.find(text_match, start)
        if idx == -1:
            break
        match_ranges.append((idx, idx + len(text_match)))
        start = idx + len(text_match)
        
    if str(match_index) != "all":
        try:
            mi = int(match_index)
            if mi < 0 or mi >= len(match_ranges):
                return
            match_ranges = [match_ranges[mi]]
        except ValueError:
            pass

    if not match_ranges:
        return

    original_runs_info = []
    current_idx = 0
    for r in paragraph.runs:
        r_len = len(r.text)
        original_runs_info.append({
            'text': r.text,
            'start': current_idx,
            'end': current_idx + r_len,
            'run': r
        })
        current_idx += r_len

    paragraph.clear()

    for r_info in original_runs_info:
        r_start = r_info['start']
        r_end = r_info['end']
        r_text = r_info['text']
        
        intersections = []
        for m_start, m_end in match_ranges:
            if m_start < r_end and m_end > r_start:
                intersections.append((max(r_start, m_start), min(r_end, m_end)))
                
        if not intersections:
            new_run = paragraph.add_run(r_text)
            _copy_run_format(r_info['run'], new_run)
            continue
            
        curr_offset = 0
        for i_start, i_end in intersections:
            rel_start = i_start - r_start
            rel_end = i_end - r_start
            
            if rel_start > curr_offset:
                new_run = paragraph.add_run(r_text[curr_offset:rel_start])
                _copy_run_format(r_info['run'], new_run)
                
            new_run = paragraph.add_run(r_text[rel_start:rel_end])
            _copy_run_format(r_info['run'], new_run)
            _apply_font_attributes(new_run, params)
            
            curr_offset = rel_end
            
        if curr_offset < len(r_text):
            new_run = paragraph.add_run(r_text[curr_offset:])
            _copy_run_format(r_info['run'], new_run)

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

def _apply_insert_paragraph(doc, params):
    """Insert a new paragraph before or after the anchor paragraph."""
    content = params.get('content')
    if content is None:
        print("[WARNING] op=insert_paragraph: missing content, skipping.")
        return

    before_id = params.get('before')
    after_id = params.get('after')

    if bool(before_id) == bool(after_id):
        print("[WARNING] op=insert_paragraph: must provide exactly one of 'before' or 'after', skipping.")
        return

    target_id = before_id or after_id
    try:
        idx = int(target_id.replace('p', ''))
        if idx < 0 or idx >= len(doc.paragraphs):
            print(f"[WARNING] op=insert_paragraph: paragraph index {idx} out of bounds, skipping.")
            return
    except (ValueError, AttributeError):
        print("[WARNING] op=insert_paragraph: invalid target format, skipping.")
        return

    anchor_p = doc.paragraphs[idx]
    style_name = params.get('style')

    if before_id:
        new_p = anchor_p.insert_paragraph_before(content)
        if style_name:
            new_p.style = style_name
        elif idx == 0:
            new_p.style = 'Normal'
        else:
            new_p.style = anchor_p.style.name
    else:
        new_p = doc.add_paragraph(content)
        anchor_p._p.addnext(new_p._p)
        new_p.style = style_name if style_name else anchor_p.style.name


def _apply_delete_paragraph(doc, params):
    """Delete the paragraph at the given index."""
    target = params.get('target')
    if not target:
        print("[WARNING] op=delete_paragraph: missing target, skipping.")
        return
    try:
        idx = int(target.replace('p', ''))
        if idx < 0 or idx >= len(doc.paragraphs):
            print(f"[WARNING] op=delete_paragraph: paragraph index {idx} out of bounds, skipping.")
            return
        p = doc.paragraphs[idx]
        p._element.getparent().remove(p._element)
    except (ValueError, AttributeError):
        print("[WARNING] op=delete_paragraph: invalid target format, skipping.")


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
            target = op.get('target', 'all')
            find_txt = op['find']
            repl_txt = op['replace']
            if target == 'all':
                for p in doc.paragraphs:
                    _replace_text_in_runs(p, find_txt, repl_txt)
            else:
                idx = int(target.replace('p', ''))
                if idx < len(doc.paragraphs):
                    _replace_text_in_runs(doc.paragraphs[idx], find_txt, repl_txt)
        elif op['op'] == 'set_font':
            idx = int(op['target'].replace('p', ''))
            if idx < len(doc.paragraphs):
                _apply_set_font(doc.paragraphs[idx], op)
        elif op['op'] == 'set_paragraph_format':
            idx = int(op['target'].replace('p', ''))
            if idx < len(doc.paragraphs):
                _apply_set_paragraph_format(doc.paragraphs[idx], op)
        elif op['op'] == 'insert_paragraph':
            _apply_insert_paragraph(doc, op)
        elif op['op'] == 'delete_paragraph':
            _apply_delete_paragraph(doc, op)

    doc.save(outpath)

if __name__ == "__main__":
    filepath = sys.argv[1]
    with open(sys.argv[2], 'r', encoding='utf-8') as f:
        ops = json.load(f)
    outpath = sys.argv[3] if len(sys.argv) > 3 else filepath
    apply_operations(filepath, ops, outpath)
