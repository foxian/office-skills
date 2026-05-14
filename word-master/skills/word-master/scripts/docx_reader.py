import docx
import sys
from docx.enum.text import WD_ALIGN_PARAGRAPH

def _extract_style_props(style):
    props = {}
    if not hasattr(style, 'font'):
        return props
    
    # Font properties
    if style.font.name: props['font'] = style.font.name
    if style.font.size: props['size'] = f"{style.font.size.pt}pt"
    if style.font.bold is not None: props['bold'] = style.font.bold
    if style.font.italic is not None: props['italic'] = style.font.italic
    
    # Paragraph properties
    if hasattr(style, 'paragraph_format'):
        pf = style.paragraph_format
        if pf.alignment is not None:
            align_map = {
                WD_ALIGN_PARAGRAPH.LEFT: 'LEFT',
                WD_ALIGN_PARAGRAPH.CENTER: 'CENTER',
                WD_ALIGN_PARAGRAPH.RIGHT: 'RIGHT',
                WD_ALIGN_PARAGRAPH.JUSTIFY: 'JUSTIFY'
            }
            props['align'] = align_map.get(pf.alignment, str(pf.alignment))
        if pf.line_spacing is not None: props['line_spacing'] = pf.line_spacing
        if pf.first_line_indent is not None:
            props['first_line_indent'] = f"{pf.first_line_indent.pt}pt"
        if pf.space_before is not None:
            props['space_before'] = f"{pf.space_before.pt}pt"
        if pf.space_after is not None:
            props['space_after'] = f"{pf.space_after.pt}pt"
            
    return props

def _extract_run_overrides(paragraph):
    overrides = []
    for run in paragraph.runs:
        text = run.text.strip()
        if not text:
            continue
        
        run_props = {}
        if run.font.name is not None: run_props['font'] = run.font.name
        if run.font.size is not None: run_props['size'] = f"{run.font.size.pt}pt"
        if run.font.bold is not None: run_props['bold'] = run.font.bold
        if run.font.italic is not None: run_props['italic'] = run.font.italic
        
        if run_props:
            prop_str = ", ".join(f"{k}={v}" for k, v in run_props.items())
            overrides.append(f"{prop_str} on \"{text}\"")
    return overrides

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
