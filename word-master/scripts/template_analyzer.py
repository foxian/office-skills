import os
import sys
import json
import argparse

# Ensure the directory of this script is added to sys.path to allow correct imports of peer modules
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

import style_analyzer
import table_analyzer

def main():
    parser = argparse.ArgumentParser(
        description="Extract template profile (paragraphs and tables) from a DOCX template."
    )
    parser.add_argument(
        "filepath",
        help="Source template DOCX path"
    )
    parser.add_argument(
        "--output",
        default="style_profile.json",
        help="Path to output style profile JSON (default: style_profile.json)"
    )
    parser.add_argument(
        "--para-min-cluster-size",
        type=int,
        default=4,
        help="Minimum cluster size for paragraph grouping (default: 4)"
    )
    group = parser.add_mutually_exclusive_group()
    group.add_argument(
        "--para-heading-aware",
        action="store_true",
        dest="para_heading_aware",
        default=True,
        help="Enable heading-aware mode for paragraph analysis (default)."
    )
    group.add_argument(
        "--para-no-heading-aware",
        action="store_false",
        dest="para_heading_aware",
        help="Disable heading-aware mode (use normal format clustering instead)."
    )

    args = parser.parse_args()

    # 1. 提取段落样式指纹
    raw_paras = style_analyzer.extract_fingerprints(
        args.filepath,
        min_cluster_size=args.para_min_cluster_size,
        heading_aware=args.para_heading_aware
    )

    # 2. 映射归一化段落键：将 heading_role 键重命名为 role，其余键原样保留
    norm_paras = []
    for entry in raw_paras:
        new_entry = {}
        for k, v in entry.items():
            if k == "heading_role":
                new_entry["role"] = v
            else:
                new_entry[k] = v
        norm_paras.append(new_entry)

    # 3. 提取表格样式指纹
    table_roles = table_analyzer.extract_table_fingerprints(args.filepath)

    # 4. 结合段落与表格，输出 JSON 配置文件
    output_data = {
        "roles": norm_paras,
        "table_roles": table_roles
    }

    # 确保输出目录存在
    output_dir = os.path.dirname(os.path.abspath(args.output))
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(output_data, f, ensure_ascii=False, indent=2)

    print(f"[INFO] Successfully analyzed template -> {args.output}")
    print(f"       Paragraph roles: {len(norm_paras)}")
    print(f"       Table roles: {len(table_roles)}")

if __name__ == "__main__":
    main()
