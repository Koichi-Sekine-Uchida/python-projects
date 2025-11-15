import json

# チェック対象のパラメータ（実際にBacklog APIに送信する前に確認）
def check_parameters(payload):
    """
    Backlogの課題作成APIに送るパラメータをチェックする関数
    """
    errors = []

    # 必須パラメータのチェック
    required_keys = ["projectId", "summary", "description", "issueTypeId", "priorityId"]
    for key in required_keys:
        if key not in payload:
            errors.append(f"❌ 必須パラメータ '{key}' が見つかりません")

    # `customFields` のフォーマットチェック
    if "customFields" in payload:
        if not isinstance(payload["customFields"], list):
            errors.append("❌ 'customFields' はリスト形式である必要があります")
        else:
            for field in payload["customFields"]:
                if not isinstance(field, dict):
                    errors.append("❌ 'customFields' の各要素は辞書である必要があります")
                elif "id" not in field or "value" not in field:
                    errors.append(f"❌ 'customFields' に必要なキー ('id', 'value') が不足しています: {field}")

    # `issueTypeId` の値チェック
    valid_issue_types = [450457, 692971]  # 実際のIDに置き換え
    if payload.get("issueTypeId") not in valid_issue_types:
        errors.append(f"❌ 'issueTypeId' が不正です: {payload.get('issueTypeId')} (有効な値: {valid_issue_types})")

    # `priorityId` の値チェック
    valid_priority_ids = [1, 2, 3, 4]  # 1: 高, 2: 中, 3: 低, 4: 最低
    if payload.get("priorityId") not in valid_priority_ids:
        errors.append(f"❌ 'priorityId' が不正です: {payload.get('priorityId')} (有効な値: {valid_priority_ids})")

    # エラーがある場合は表示
    if errors:
        print("\n".join(errors))
        return False
    else:
        print("✅ パラメータは問題ありません")
        return True

# **テスト用のパラメータ（エラーがあるか確認）**
test_payload = {
    "projectId": 51948,
    "summary": "テスト課題",
    "description": "テスト用の説明",
    "issueTypeId": 450457,  # 配信
    "priorityId": 3,  # 中
    "customFields": [
        {"id": 72497, "value": 2}  # インターネットエッジ
    ]
}

# **パラメータチェックを実行**
print("🔍 パラメータチェック開始...")
is_valid = check_parameters(test_payload)

# JSON表示（デバッグ用）
print("\n送信するJSON:")
print(json.dumps(test_payload, indent=4, ensure_ascii=False))

if is_valid:
    print("🚀 APIリクエストを送信できます！")
else:
    print("⚠️ APIリクエストを送信できません。パラメータを修正してください。")
