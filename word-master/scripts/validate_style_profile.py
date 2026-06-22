import json
import sys

# 声明合法角色白名单
VALID_ROLES = {
    "Heading 1", "Heading 2", "Heading 3", "Heading 4",
    "Heading 5", "Heading 6", "Heading 7", "Heading 8", "Heading 9",
    "Normal", "List Bullet", "List Number"
}

def validate_profile(profile_or_path):
    """
    Validate style_profile.json format.
    Args:
        profile_or_path: Either a dict (already parsed) or str (file path)
    Returns:
        bool: True if valid, False otherwise
    """
    if isinstance(profile_or_path, str):
        with open(profile_or_path, "r", encoding="utf-8") as f:
            profile = json.load(f)
    else:
        profile = profile_or_path

    errors = []

    if "roles" not in profile:
        errors.append("缺少 'roles' 根键（必须是数组）")
    elif not isinstance(profile["roles"], list):
        errors.append("'roles' 必须是数组")
    else:
        for i, entry in enumerate(profile["roles"]):
            if not isinstance(entry, dict):
                errors.append(f"roles[{i}] 必须是对象")
                continue

            if "role" not in entry:
                errors.append(f"roles[{i}] 缺少 'role' 字段")
            elif not isinstance(entry["role"], str):
                errors.append(f"roles[{i}].role 必须是字符串")
            elif entry["role"] not in VALID_ROLES:
                errors.append(f"roles[{i}].role '{entry['role']}' 不是受支持的语义角色类型")

            if "fingerprint" not in entry:
                errors.append(f"roles[{i}] 缺少 'fingerprint' 字段")
            elif not isinstance(entry["fingerprint"], dict):
                errors.append(f"roles[{i}].fingerprint 必须是对象")

    # 校验 table_roles 字段
    if "table_roles" in profile:
        if not isinstance(profile["table_roles"], list):
            errors.append("'table_roles' 必须是数组")
        else:
            for j, item in enumerate(profile["table_roles"]):
                if not isinstance(item, dict):
                    errors.append(f"table_roles[{j}] 必须是对象")
                    continue

                # 校验 id
                if "id" not in item:
                    errors.append(f"table_roles[{j}] 缺少 'id' 字段")
                elif not isinstance(item["id"], int):
                    errors.append(f"table_roles[{j}].id 必须是整型")

                # 校验 structure
                if "structure" not in item:
                    errors.append(f"table_roles[{j}] 缺少 'structure' 字段")
                elif not isinstance(item["structure"], dict):
                    errors.append(f"table_roles[{j}].structure 必须是对象")
                else:
                    struct = item["structure"]
                    if "cols" not in struct:
                        errors.append(f"table_roles[{j}].structure 缺少 'cols' 字段")
                    elif not isinstance(struct["cols"], int):
                        errors.append(f"table_roles[{j}].structure.cols 必须是整型")

                    if "has_header_row" not in struct:
                        errors.append(f"table_roles[{j}].structure 缺少 'has_header_row' 字段")
                    elif not isinstance(struct["has_header_row"], bool):
                        errors.append(f"table_roles[{j}].structure.has_header_row 必须是布尔型")

                    if "rows" not in struct:
                        errors.append(f"table_roles[{j}].structure 缺少 'rows' 字段")
                    elif not isinstance(struct["rows"], int):
                        errors.append(f"table_roles[{j}].structure.rows 必须是整型")

                # 校验 border（如果存在）
                if "border" in item:
                    border = item["border"]
                    if border is not None:
                        if not isinstance(border, dict):
                            errors.append(f"table_roles[{j}].border 必须是对象")
                        else:
                            for border_dir, border_val in border.items():
                                if border_dir not in {"top", "bottom", "left", "right", "insideH", "insideV"}:
                                    errors.append(f"table_roles[{j}].border 包含未知的边框方向: {border_dir}")
                                if border_val is not None:
                                    if not isinstance(border_val, dict):
                                        errors.append(f"table_roles[{j}].border.{border_dir} 必须是对象")
                                    else:
                                        if "val" in border_val and not isinstance(border_val["val"], str):
                                            errors.append(f"table_roles[{j}].border.{border_dir}.val 必须是字符串")
                                        if "sz" in border_val and not isinstance(border_val["sz"], int):
                                            errors.append(f"table_roles[{j}].border.{border_dir}.sz 必须是整型")
                                        if "color" in border_val and border_val["color"] is not None and not isinstance(border_val["color"], str):
                                            errors.append(f"table_roles[{j}].border.{border_dir}.color 必须是字符串或 None")

                # 校验 shading（如果存在）
                if "shading" in item:
                    shading = item["shading"]
                    if shading is not None:
                        if not isinstance(shading, dict):
                            errors.append(f"table_roles[{j}].shading 必须是对象")
                        else:
                            for shading_key, shading_val in shading.items():
                                if shading_key not in {"header", "body"}:
                                    errors.append(f"table_roles[{j}].shading 包含未知的阴影键: {shading_key}")
                                if shading_val is not None and not isinstance(shading_val, str):
                                    errors.append(f"table_roles[{j}].shading.{shading_key} 必须是字符串或 None")

                # 校验 text format 字段（header_text, body_text）
                for text_field in ["header_text", "body_text"]:
                    if text_field in item:
                        text_fmt = item[text_field]
                        if text_fmt is not None:
                            if not isinstance(text_fmt, dict):
                                errors.append(f"table_roles[{j}].{text_field} 必须是对象")
                            else:
                                if "font" in text_fmt and text_fmt["font"] is not None and not isinstance(text_fmt["font"], str):
                                    errors.append(f"table_roles[{j}].{text_field}.font 必须是字符串或 None")
                                if "size" in text_fmt and text_fmt["size"] is not None and not isinstance(text_fmt["size"], str):
                                    errors.append(f"table_roles[{j}].{text_field}.size 必须是字符串或 None")
                                if "bold" in text_fmt and text_fmt["bold"] is not None and not isinstance(text_fmt["bold"], bool):
                                    errors.append(f"table_roles[{j}].{text_field}.bold 必须是布尔值或 None")
                                if "align" in text_fmt and text_fmt["align"] is not None and not isinstance(text_fmt["align"], str):
                                    errors.append(f"table_roles[{j}].{text_field}.align 必须是字符串或 None")

    if errors:
        print("[ERROR] style_profile.json 格式验证失败:")
        for e in errors:
            print(f"  - {e}")
        return False

    print("[INFO] style_profile.json 格式验证通过")
    return True

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python validate_style_profile.py <profile.json>")
        sys.exit(1)

    profile_path = sys.argv[1]
    if not validate_profile(profile_path):
        sys.exit(1)
