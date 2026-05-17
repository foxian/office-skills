import docx
import json
import sys
from docx.enum.text import WD_ALIGN_PARAGRAPH


def _get_effective_value(run_val, style_val):
    """Return run value if explicitly set, else fall back to style value."""
    if run_val is not None:
        return run_val
    return style_val


def compute_effective_fingerprint(paragraph):
    """
    Compute effective format fingerprint for a paragraph, merging style attributes and run-level overrides.
    For multi-run paragraphs, uses the first run with an explicit value.
    Returns dict with: size, bold, italic, align.
    """
    style = paragraph.style
    style_font = style.font if hasattr(style, 'font') else None
    style_pf = style.paragraph_format if hasattr(style, 'paragraph_format') else None

    # Collect run-level overrides (first run with explicit value wins)
    run_size = None
    run_bold = None
    run_italic = None
    for run in paragraph.runs:
        if run.font.size is not None and run_size is None:
            run_size = run.font.size
        if run.bold is not None and run_bold is None:
            run_bold = run.bold
        if run.italic is not None and run_italic is None:
            run_italic = run.italic

    # Font size
    style_size = style_font.size if style_font else None
    eff_size = run_size or style_size
    size_str = f"{eff_size.pt}pt" if eff_size else None

    # Bold
    style_bold = style_font.bold if style_font else None
    eff_bold = _get_effective_value(run_bold, style_bold)

    # Italic
    style_italic = style_font.italic if style_font else None
    eff_italic = _get_effective_value(run_italic, style_italic)

    # Alignment
    align_map = {
        WD_ALIGN_PARAGRAPH.LEFT: "left",
        WD_ALIGN_PARAGRAPH.CENTER: "center",
        WD_ALIGN_PARAGRAPH.RIGHT: "right",
        WD_ALIGN_PARAGRAPH.JUSTIFY: "justify",
    }
    pf_align = paragraph.paragraph_format.alignment if hasattr(paragraph, 'paragraph_format') and paragraph.paragraph_format else None
    style_align = style_pf.alignment if style_pf else None
    eff_align_raw = _get_effective_value(pf_align, style_align)
    eff_align = align_map.get(eff_align_raw) if eff_align_raw else None

    return {
        "size": size_str,
        "bold": eff_bold,
        "italic": eff_italic,
        "align": eff_align,
    }
