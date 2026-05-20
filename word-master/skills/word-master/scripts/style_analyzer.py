import docx
import json
import re
import sys
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

CHAPTER_PATTERN = re.compile(
    r'^(第[一二三四五六七八九十\d]+章|[一二三四五六七八九十]+、|（[一二三四五六七八九十\d]+）|\d+\.\s)'
)


def _fingerprint_key(fp):
    """Convert fingerprint dict to a hashable key for grouping."""
    return (fp.get("size"), fp.get("bold"), fp.get("italic"), fp.get("align"), fp.get("font"), fp.get("color"))


def _get_effective_value(run_val, style_val):
    """Return run value if explicitly set, else fall back to style value."""
    if run_val is not None:
        return run_val
    return style_val


def compute_effective_fingerprint(paragraph):
    """
    Compute effective format fingerprint for a paragraph, merging style attributes and run-level overrides.
    For multi-run paragraphs, uses the first run with an explicit value.
    Returns dict with: size, bold, italic, align, font, color, space_before, space_after, line_spacing.
    """
    style = paragraph.style
    style_font = style.font if hasattr(style, 'font') else None
    style_pf = style.paragraph_format if hasattr(style, 'paragraph_format') else None

    # Collect run-level overrides (first run with explicit value wins)
    run_size = None
    run_bold = None
    run_italic = None
    run_color_rgb = None
    run_color_theme = None
    for run in paragraph.runs:
        if run.font.size is not None and run_size is None:
            run_size = run.font.size
        if run.bold is not None and run_bold is None:
            run_bold = run.bold
        if run.italic is not None and run_italic is None:
            run_italic = run.italic
        # Color: extract rgb and theme_color
        if run.font.color:
            if run.font.color.rgb is not None and run_color_rgb is None:
                run_color_rgb = str(run.font.color.rgb)
            if run.font.color.theme_color is not None and run_color_theme is None:
                run_color_theme = str(run.font.color.theme_color)

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

    # Font color: prefer run-level, fall back to style
    style_color_rgb = None
    style_color_theme = None
    if style_font and hasattr(style_font, 'color') and style_font.color:
        if style_font.color.rgb is not None:
            style_color_rgb = str(style_font.color.rgb)
        if style_font.color.theme_color is not None:
            style_color_theme = str(style_font.color.theme_color)
    eff_color_rgb = _get_effective_value(run_color_rgb, style_color_rgb)
    eff_color_theme = _get_effective_value(run_color_theme, style_color_theme)
    # Combine into a single color string: "theme:NAME" or "rgb:XXXXXX" or None
    # theme_color has higher priority — it represents semantic intent (e.g., ACCENT_1)
    if eff_color_theme:
        eff_color = f"theme:{eff_color_theme}"
    elif eff_color_rgb:
        eff_color = f"rgb:{eff_color_rgb}"
    else:
        eff_color = None

    # Spacing fields (paragraph-level)

    def _get_pt_str(val):
        """Convert Pt value to 'Xpt' string, defaults to '0.0pt' if not set."""
        if val is None:
            return "0.0pt"
        try:
            return f"{val.pt}pt"
        except AttributeError:
            return "0.0pt"

    sb_val = paragraph.paragraph_format.space_before
    if sb_val is None and style_pf:
        sb_val = style_pf.space_before
    eff_space_before = _get_pt_str(sb_val)

    sa_val = paragraph.paragraph_format.space_after
    if sa_val is None and style_pf:
        sa_val = style_pf.space_after
    eff_space_after = _get_pt_str(sa_val)

    # Line spacing: only support multiplier (float), not absolute Pt
    ls_val = paragraph.paragraph_format.line_spacing
    if ls_val is None and style_pf:
        ls_val = style_pf.line_spacing
    # python-docx: if line_spacing is a Pt object, it's absolute mode — not supported
    if ls_val is not None:
        try:
            ls_val.pt  # this only works if it's a Pt (absolute), which we don't support
            ls_val = None
        except AttributeError:
            pass  # it's a float (multiplier mode) — keep it
    eff_line_spacing = ls_val

    return {
        "size": size_str,
        "bold": eff_bold,
        "italic": eff_italic,
        "align": eff_align,
        "font": eff_font,
        "color": eff_color,
        "space_before": eff_space_before,
        "space_after": eff_space_after,
        "line_spacing": eff_line_spacing,
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


def _detect_list_type(paragraph, doc):
    """
    Detect if a paragraph is a Word native list.
    Returns "List Bullet", "List Number", or None.
    """
    pPr = paragraph._element.get_or_add_pPr()
    numPr = pPr.find(qn('w:numPr'))

    # 1. 尝试使用 XML 物理节点识别
    if numPr is not None:
        try:
            numId_elem = numPr.find(qn('w:numId'))
            ilvl_elem = numPr.find(qn('w:ilvl'))
            if numId_elem is not None:
                numId = numId_elem.get(qn('w:val'))
                ilvl = ilvl_elem.get(qn('w:val')) if ilvl_elem is not None else "0"

                # 追溯 numbering.xml 定义
                numbering_part = doc.part.numbering_part
                if numbering_part is not None:
                    numbering = numbering_part.numbering_definitions._numbering

                    # 遍历查找 w:num
                    num_node = None
                    for num in numbering.findall(qn('w:num')):
                        if num.get(qn('w:numId')) == numId:
                            num_node = num
                            break

                    if num_node is not None:
                        abstractNumId_elem = num_node.find(qn('w:abstractNumId'))
                        if abstractNumId_elem is not None:
                            abstractNumId = abstractNumId_elem.get(qn('w:val'))

                            # 遍历查找 w:abstractNum
                            abstractNum_node = None
                            for abNum in numbering.findall(qn('w:abstractNum')):
                                if abNum.get(qn('w:abstractNumId')) == abstractNumId:
                                    abstractNum_node = abNum
                                    break

                            if abstractNum_node is not None:
                                # 遍历查找 w:lvl
                                lvl_node = None
                                for lvl in abstractNum_node.findall(qn('w:lvl')):
                                    if lvl.get(qn('w:ilvl')) == ilvl:
                                        lvl_node = lvl
                                        break

                                if lvl_node is not None:
                                    numFmt_elem = lvl_node.find(qn('w:numFmt'))
                                    if numFmt_elem is not None:
                                        numFmt = numFmt_elem.get(qn('w:val'))
                                        if numFmt == 'bullet':
                                            return "List Bullet"
                                        else:
                                            return "List Number"
        except Exception:
            pass

    # 2. 降级：样式名关键字判定
    style_name = paragraph.style.name.lower()
    if 'bullet' in style_name:
        return "List Bullet"
    if 'list' in style_name:
        if re.match(r'^([0-9a-zA-Z]+|[一二三四五六七八九十]+)[\.、\s]', paragraph.text.strip()):
            return "List Number"
        return "List Bullet"

    # 3. 降级：文本前缀正则匹配
    stripped = paragraph.text.strip()
    if stripped.startswith(('•', '·', '-', '*', 'o', '▪', '■', '◆', '▲')):
        return "List Bullet"
    if re.match(r'^([0-9a-zA-Z]+|[一二三四五六七八九十]+)[\.、\s]', stripped):
        return "List Number"

    return None


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
