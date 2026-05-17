import docx
import json
import sys
import os

# Reuse fingerprint computation from style_analyzer
sys.path.insert(0, os.path.dirname(__file__))
from style_analyzer import compute_effective_fingerprint


def _score(fp, role_fp):
    """
    Calculate similarity score between two fingerprints (0.0 ~ 1.0).
    Equal weights: size=1/4, bold=1/4, italic=1/4, align=1/4.
    Size uses linear decay: max(0, 1 - diff/10) * (1/4).
    """
    total = 4.0
    w = 1.0 / total

    # size similarity: linear decay, diff=0 -> 1.0, diff=10 -> 0.0
    s1 = fp.get("size")
    s2 = role_fp.get("size")
    if s1 and s2:
        try:
            diff = abs(float(s1.rstrip("pt")) - float(s2.rstrip("pt")))
            size_score = max(0.0, 1.0 - diff / 10.0) * w
        except ValueError:
            size_score = 0.0
    elif s1 == s2:  # both None
        size_score = w
    else:
        size_score = 0.0

    bold_score = w if fp.get("bold") == role_fp.get("bold") else 0.0
    italic_score = w if fp.get("italic") == role_fp.get("italic") else 0.0
    align_score = w if fp.get("align") == role_fp.get("align") else 0.0

    return size_score + bold_score + italic_score + align_score


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


def generate_apply_ops(draft_path, profile, threshold=0.6, skip_head=0, skip_tail=0):
    """
    Walk draft.docx paragraphs, match fingerprints, generate apply_style DSL list.
    skip_head/skip_tail: skip first/last N paragraphs (human fallback).
    """
    doc = docx.Document(draft_path)
    paras = doc.paragraphs
    n = len(paras)
    ops = []

    for i, para in enumerate(paras):
        if i < skip_head or i >= n - skip_tail:
            continue
        if not para.text.strip():
            continue
        fp = compute_effective_fingerprint(para)
        role = match_fingerprint_to_role(fp, profile, threshold)
        if role:
            ops.append({"op": "apply_style", "target": f"p{i}", "style": role})
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
    if profile_path:
        print(f"[INFO] Loading existing profile: {profile_path}")
        with open(profile_path, "r", encoding="utf-8") as f:
            profile = json.load(f)
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
