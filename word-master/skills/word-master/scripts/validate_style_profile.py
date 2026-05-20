import json
import sys

# Valid semantic role names whitelist
VALID_ROLES = {
    "Heading 1", "Heading 2", "Heading 3", "Heading 4",
    "Heading 5", "Heading 6", "Heading 7", "Heading 8", "Heading 9",
    "Normal", "List Bullet"
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