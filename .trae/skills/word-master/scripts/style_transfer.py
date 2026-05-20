import docx
import json
import sys
import os

# Reuse fingerprint computation from style_analyzer
sys.path.insert(0, os.path.dirname(__file__))
from style_analyzer import _detect_list_type
from validate_style_profile import validate_profile

_FALLBACK_MAP = {
    "Body Text": "Normal",
}


def _score(fp, role_fp):
    """
    Calculate similarity score between two fingerprints (0.0 ~ 1.0).
    Weights: size=0.25, bold=0.125, italic=0.125, align=0.125,
             font=0.125, color=0.125, space_before=0.0625, space_after=0.0625.
    line_spacing is recorded in fingerprint but not scored.
    Size uses linear decay: max(0, 1 - diff/10).

    When a field in fp is None (not set, relies on style default), that
    dimension is skipped and the score is renormalized so the max is 1.0.
    This prevents "fake style" documents (no run-level overrides) from
    being penalized for missing font/size information.
    """
    score = 0.0
    weight_sum = 0.0  # tracks total weight of scored dimensions

    # size (weight=0.25): linear decay
    s1 = fp.get("size")
    s2 = role_fp.get("size")
    if s1 is not None:
        weight_sum += 0.25
        if s1 and s2:
            try:
                diff = abs(float(s1.rstrip("pt")) - float(s2.rstrip("pt")))
                score += max(0.0, 1.0 - diff / 10.0) * 0.25
            except ValueError:
                pass
        elif s1 == s2:  # both None or both empty
            score += 0.25

    # bold (weight=0.125)
    if fp.get("bold") is not None:
        weight_sum += 0.125
        score += 0.125 if fp.get("bold") == role_fp.get("bold") else 0.0

    # italic (weight=0.125)
    if fp.get("italic") is not None:
        weight_sum += 0.125
        score += 0.125 if fp.get("italic") == role_fp.get("italic") else 0.0

    # align (weight=0.125)
    if fp.get("align") is not None:
        weight_sum += 0.125
        score += 0.125 if fp.get("align") == role_fp.get("align") else 0.0

    # font (weight=0.125): exact match
    f1 = fp.get("font")
    f2 = role_fp.get("font")
    if f1 is not None:
        weight_sum += 0.125
        if f2 is None:
            pass  # fp has font, role doesn't — no score
        elif f1 == f2:
            score += 0.125

    # color (weight=0.125): exact match on combined "rgb:XXX" or "theme:XXX" string
    c1 = fp.get("color")
    c2 = role_fp.get("color")
    if c1 is not None:
        weight_sum += 0.125
        if c2 is not None and c1 == c2:
            score += 0.125

    # space_before (weight=0.0625): numeric comparison within 0.5pt tolerance
    sb1 = fp.get("space_before")
    if sb1 is not None:
        weight_sum += 0.0625
        sb2 = role_fp.get("space_before") or "0.0pt"
        try:
            if abs(float(sb1.rstrip("pt")) - float(sb2.rstrip("pt"))) < 0.5:
                score += 0.0625
        except (ValueError, AttributeError):
            pass

    # space_after (weight=0.0625): numeric comparison within 0.5pt tolerance
    sa1 = fp.get("space_after")
    if sa1 is not None:
        weight_sum += 0.0625
        sa2 = role_fp.get("space_after") or "0.0pt"
        try:
            if abs(float(sa1.rstrip("pt")) - float(sa2.rstrip("pt"))) < 0.5:
                score += 0.0625
        except (ValueError, AttributeError):
            pass

    # Renormalize: if some dimensions were skipped, scale score to 0-1 range
    if weight_sum == 0:
        return 0.0
    # Penalize when too few dimensions contributed: if less than half the max
    # weight (1.0) was available, scale down proportionally. This prevents
    # paragraphs with all-None fingerprints from scoring 1.0 on just 1-2 dims.
    confidence = min(1.0, weight_sum / 0.5)  # need at least 0.5 weight to be fully trusted
    return (score / weight_sum) * confidence


def match_fingerprint_to_role(fp, profile, threshold=0.6):
    """
    Match a paragraph fingerprint to the nearest role in style_profile.
    Returns best-matching Word style name, or None if best score < threshold.
    """
    best_role = None
    best_score = 0.0
    for entry in profile.get("roles", []):
        s = _score(fp, entry["fingerprint"])
        if s > best_score:
            best_score = s
            best_role = entry["role"]
    return best_role if best_score >= threshold else None


def generate_apply_ops(draft_path, profile, skip_head=0, skip_tail=0):
    """
    Walk draft.docx paragraphs, match roles via deterministic cascade pipeline.
    Pipeline:
      1. Level 0: Direct Heading Name Match
      2. Level 1: XML Numbering (numPr) & Text Bullet prefix Match (Strict Handoff)
      3. Level 2: Body-style (Normal/Body Text) — text pattern first, then default Normal
      4. Level 3: Absolute Fallback to 'Normal'
    """
    # --- Phase 1: generate update_style_definition ops ---
    ops = []
    for entry in profile.get("roles", []):
        role = entry.get("role")
        fp = entry.get("fingerprint")
        if role and fp:
            ops.append({
                "op": "update_style_definition",
                "style": role,
                "fingerprint": fp,
            })

    # --- Phase 2: generate apply_style ops per paragraph ---
    doc = docx.Document(draft_path)
    paras = doc.paragraphs
    n = len(paras)
    available_roles = {entry["role"] for entry in profile.get("roles", [])}

    for i, para in enumerate(paras):
        if i < skip_head or i >= n - skip_tail:
            continue
        style_name = para.style.name
        role = None

        # Level 0: Direct Heading Match
        if style_name.startswith("Heading ") and style_name in available_roles:
            role = style_name

        # Level 1: XML List / Bullet Prefix Match with Strict Handoff
        if role is None:
            list_type = _detect_list_type(para, doc)
            if list_type is not None:
                role = list_type if list_type in available_roles else "Normal"

        # Level 2: 精确匹配（仅当 role 未被设置时）
        if role is None:
            if style_name in available_roles:
                role = style_name
            elif style_name in _FALLBACK_MAP:
                fallback = _FALLBACK_MAP[style_name]
                if fallback in available_roles:
                    role = fallback

        # Level 3: Absolute Fallback
        if role is None:
            role = "Normal" if "Normal" in available_roles else None

        if role:
            ops.append({
                "op": "apply_style",
                "target": f"p{i}",
                "style": role,
                "clear_run_formats": True,
            })
    return ops


def run(template_path, draft_path, output_path,
        profile_path=None, review=False, skip_head=0, skip_tail=0):
    """
    Main entry: two-phase style transfer workflow.
    profile_path: if provided, skip analysis and LLM inference, reuse existing profile.
    review: if True, pause after profile load for user confirmation.
    """
    # --- File existence validation ---
    if not os.path.exists(draft_path):
        raise FileNotFoundError(f"[ERROR] Draft file not found: {draft_path}")
    if profile_path and not os.path.exists(profile_path):
        raise FileNotFoundError(f"[ERROR] Profile file not found: {profile_path}")
    if not profile_path and template_path and not os.path.exists(template_path):
        raise FileNotFoundError(f"[ERROR] Template file not found: {template_path}")

    # --- Phase 1: Get style_profile ---

    # --- Validate profile format ---

    if profile_path:
        print(f"[INFO] Loading existing profile: {profile_path}")
        with open(profile_path, "r", encoding="utf-8") as f:
            profile = json.load(f)

        if not validate_profile(profile):
            raise ValueError(f"[ERROR] style_profile.json 格式验证失败，请检查上述错误")
    else:
        from style_analyzer import extract_fingerprints
        print(f"[INFO] Analyzing template: {template_path}")
        fingerprints = extract_fingerprints(template_path)

        # --- LLM inference placeholder ---
        # Current version saves fingerprints to JSON for external LLM to generate style_profile.
        # TODO: integrate API call to auto-generate style_profile.
        fingerprints_path = "fingerprints.json"
        with open(fingerprints_path, "w", encoding="utf-8") as f:
            json.dump(fingerprints, f, ensure_ascii=False, indent=2)
        print(f"[INFO] Fingerprints saved to {fingerprints_path}")
        print("[INFO] Please review fingerprints.json, ask LLM to generate style_profile.json,")
        print("       then re-run with: --profile style_profile.json")
        return

    # --- Review mode: pause for user confirmation ---
    if review:
        print("[REVIEW] Profile loaded. Edit style_profile.json if needed, then press Enter to continue...")
        input()

    # --- Phase 2: Generate DSL and delegate to docx_editor ---
    print(f"[INFO] Generating apply_style operations for: {draft_path}")
    ops = generate_apply_ops(draft_path, profile,
                              skip_head=skip_head, skip_tail=skip_tail)
    print(f"[INFO] Generated {len(ops)} operations")

    # Delegate to docx_editor (reuses backup and safe execution)
    editor_path = os.path.join(os.path.dirname(__file__), "docx_editor.py")
    from docx_editor import apply_operations
    apply_operations(draft_path, ops, output_path)
    print(f"[INFO] Done → {output_path}")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Word document style transfer tool.")
    parser.add_argument("template", nargs="?", help="Source template DOCX (skip if using --profile)")
    parser.add_argument("draft", help="Target draft DOCX to apply styles to")
    parser.add_argument("output", help="Output DOCX path")
    parser.add_argument("--profile", help="Reuse existing style_profile.json (skip analysis)")
    parser.add_argument("--review", action="store_true", help="Pause after profile load for manual review")
    parser.add_argument("--skip-head", type=int, default=0, metavar="N", help="Skip first N paragraphs")
    parser.add_argument("--skip-tail", type=int, default=0, metavar="N", help="Skip last N paragraphs")
    args = parser.parse_args()
    run(
        template_path=args.template,
        draft_path=args.draft,
        output_path=args.output,
        profile_path=args.profile,
        review=args.review,
        skip_head=args.skip_head,
        skip_tail=args.skip_tail,
    )
