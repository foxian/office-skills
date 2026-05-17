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
