import docx
import json
import re
import sys
from docx.enum.text import WD_ALIGN_PARAGRAPH

CHAPTER_PATTERN = re.compile(
    r'^(第[一二三四五六七八九十\d]+章|[一二三四五六七八九十]+、|（[一二三四五六七八九十\d]+）|\d+\.\s)'
)


def _fingerprint_key(fp):
    """Convert fingerprint dict to a hashable key for grouping."""
    return (fp.get("size"), fp.get("bold"), fp.get("italic"), fp.get("align"))


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

    # Font family
    run_font = None
    for run in paragraph.runs:
        if run.font.name is not None:
            run_font = run.font.name
            break
    style_font_name = style_font.name if style_font and hasattr(style_font, 'name') else None
    eff_font = run_font if run_font is not None else style_font_name

    return {
        "size": size_str,
        "bold": eff_bold,
        "italic": eff_italic,
        "align": eff_align,
        "font": eff_font,
    }


def extract_fingerprints(filepath, min_cluster_size=4):
    """
    Analyze a DOCX document, extract body format fingerprint clusters.
    Filter out clusters with < min_cluster_size members (cover page, etc).
    Each cluster retains one representative sample text.
    Returns list of {id, fingerprint, example}.
    """
    doc = docx.Document(filepath)

    clusters = {}  # key -> list of (paragraph_text, fingerprint)
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        fp = compute_effective_fingerprint(para)
        key = _fingerprint_key(fp)
        if key not in clusters:
            clusters[key] = []
        clusters[key].append((text, fp))

    # Strategy A: filter small clusters; Strategy C: chapter pattern clusters exempt from filtering
    result = []
    for i, (key, members) in enumerate(clusters.items()):
        has_chapter_pattern = any(CHAPTER_PATTERN.match(text) for text, _ in members)
        if len(members) < min_cluster_size and not has_chapter_pattern:
            continue
        representative_text, representative_fp = members[0]
        result.append({
            "id": i,
            "fingerprint": representative_fp,
            "example": representative_text,
        })
    return result


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Extract style fingerprints from a DOCX template.")
    parser.add_argument("filepath", help="Source template DOCX path")
    parser.add_argument("--min-cluster-size", type=int, default=4)
    parser.add_argument("--output", default="fingerprints.json")
    args = parser.parse_args()
    fingerprints = extract_fingerprints(args.filepath, args.min_cluster_size)
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(fingerprints, f, ensure_ascii=False, indent=2)
    print(f"[INFO] Extracted {len(fingerprints)} fingerprint groups → {args.output}")

