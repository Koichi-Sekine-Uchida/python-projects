import requests
import json

# Backlog API 設定
BACKLOG_SPACE_ID = "ucdprj"
API_KEY = "d9J1kvSFf3oFVhIJESxjJ0rKfGRkEea7Fr2K2eRPcZwU7zRzb60DOVlDanFoLfdv"
PROJECT_ID = 51948  # プロジェクトID
ISSUE_TYPE_ID_EDGE = 692971  # 「インターネットエッジ」
EDGE_CUSTOM_FIELD_ID = 72497  # エッジ種別のカスタムフィールドID
EDGE_TYPE_INTERNET = 2  # 「インターネットエッジ」

# Backlog API エンドポイント
BACKLOG_API_URL = f"https://{BACKLOG_SPACE_ID}.backlog.com/api/v2/issues"

# 課題作成データ
payload = {
    "projectId": PROJECT_ID,
    "summary": "ベーシック_【東京都あきる野市】小学校_1校_発注明細20250225.xlsx",
    "description": "自動アップロードにより添付されたExcelファイルです",
    "issueTypeId": ISSUE_TYPE_ID_EDGE,
    "priorityId": 3,
    "customFields": [
        {
            "id": int(72497),  # ここで明示的に整数にする
            "value": int(2)     # これも整数
        }
    ]
}

# APIリクエスト送信
params = {"apiKey": API_KEY}
headers = {"Content-Type": "application/json"}
response = requests.post(BACKLOG_API_URL, params=params, headers=headers, data=json.dumps(payload))

# レスポンスを出力
print(f"🚀 送信データ: {json.dumps(payload, indent=4, ensure_ascii=False)}")
print(f"レスポンスコード: {response.status_code}")
print(f"レスポンス内容: {response.text}")
