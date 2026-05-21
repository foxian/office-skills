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
    Returns dict with: size, bold, italic, align, font, color, space_before, space_after, first_line_indent, line_spacing.
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

    # First line indent
    fli_val = paragraph.paragraph_format.first_line_indent
    if fli_val is None and style_pf:
        fli_val = style_pf.first_line_indent
    eff_first_line_indent = _get_pt_str(fli_val)

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
        "first_line_indent": eff_first_line_indent,
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


def _get_outline_level(paragraph) -> int | None:
    """
    Read paragraph outline level (0=Heading 1, ..., 8=Heading 9).
    Strategy:
      1. Check paragraph-level XML directly.
      2. Walk style inheritance chain (para.style -> base_style -> ...),
         up to 10 hops to guard against circular references.
    Returns int 0-8 or None (not a heading).
    """
    def _read_outline_lvl_from_pPr(pPr_elem):
        if pPr_elem is None:
            return None
        ol = pPr_elem.find(qn('w:outlineLvl'))
        if ol is None:
            return None
        val = ol.get(qn('w:val'))
        if val is None:
            return None
        try:
            level = int(val)
            return level if 0 <= level <= 8 else None
        except ValueError:
            return None

    # 1. Paragraph-level direct attribute
    para_pPr = paragraph._element.find(qn('w:pPr'))
    level = _read_outline_lvl_from_pPr(para_pPr)
    if level is not None:
        return level

    # 2. Style inheritance chain
    style = paragraph.style
    for _ in range(10):
        if style is None:
            break
        style_pPr = style.element.find(qn('w:pPr'))
        level = _read_outline_lvl_from_pPr(style_pPr)
        if level is not None:
            return level
        style = style.base_style

    return None


def _average_fingerprints(fps: list[dict]) -> dict:
    """
    Aggregate a list of fingerprint dicts into one representative fingerprint.
    Aggregation rules per field: see Spec §3.
    """
    if not fps:
        return {}

    def _pt_to_float(s):
        if s is None:
            return 0.0
        try:
            return float(s.rstrip("pt"))
        except (ValueError, AttributeError):
            return 0.0

    def _vote(values, tiebreak_order=None, skip_none=False):
        from collections import Counter
        filtered = [v for v in values if v is not None] if skip_none else values
        if not filtered:
            return None
        counts = Counter(filtered)
        max_count = max(counts.values())
        winners = [k for k, v in counts.items() if v == max_count]
        if len(winners) == 1:
            return winners[0]
        if tiebreak_order is not None:
            for priority in tiebreak_order:
                if priority in winners:
                    return priority
        # Fallback: first occurrence in fps
        for v in values:
            if v in winners:
                return v
        return winners[0]

    bold_tiebreak = [False, None, True]

    size_vals = [_pt_to_float(fp.get("size")) for fp in fps if fp.get("size") is not None]
    ls_vals = [fp.get("line_spacing") for fp in fps if fp.get("line_spacing") is not None]

    def _avg_pt(field):
        vals = [_pt_to_float(fp.get(field)) for fp in fps]
        return f"{sum(vals) / len(vals):.1f}pt"

    return {
        "size": f"{sum(size_vals) / len(size_vals):.1f}pt" if size_vals else None,
        "bold": _vote([fp.get("bold") for fp in fps], tiebreak_order=bold_tiebreak),
        "italic": _vote([fp.get("italic") for fp in fps], tiebreak_order=bold_tiebreak),
        "align": _vote([fp.get("align") for fp in fps]),
        "font": _vote([fp.get("font") for fp in fps], skip_none=True),
        "color": _vote([fp.get("color") for fp in fps]),
        "space_before": _avg_pt("space_before"),
        "space_after": _avg_pt("space_after"),
        "first_line_indent": _avg_pt("first_line_indent"),
        "line_spacing": sum(ls_vals) / len(ls_vals) if ls_vals else None,
    }


def _detect_list_type(paragraph, doc) -> str | None:
    """
    Detect if a paragraph is a Word native list.
    Returns "List Bullet", "List Number", or None.
    Uses <w:numPr> XML node as primary signal; falls back to
    style name keywords and text prefix characters.
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
        except (AttributeError, KeyError, TypeError):
            pass

    # 2. 降级：样式名关键字判定
    style_name = paragraph.style.name.lower()
    if 'list bullet' in style_name or 'list paragraph' in style_name:
        return "List Bullet"
    if 'list number' in style_name:
        return "List Number"
    if 'list' in style_name:
        if re.match(r'^([0-9a-zA-Z]+|[一二三四五六七八九十]+)[\.、\s]', paragraph.text.strip()):
            return "List Number"
        return "List Bullet"

    # 3. 降级：文本前缀匹配（仅匹配 bullet 符号；不对数字前缀判定以防止拦截标题级联）
    stripped = paragraph.text.strip()
    if stripped.startswith(('•', '·', '-', '*', 'o', '▪', '■', '◆', '▲')):
        return "List Bullet"

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
