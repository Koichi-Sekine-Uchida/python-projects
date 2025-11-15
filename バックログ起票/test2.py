import requests
import json

# Backlog API 設定
BACKLOG_SPACE_ID = "ucdprj"
API_KEY = "d9J1kvSFf3oFVhIJESxjJ0rKfGRkEea7Fr2K2eRPcZwU7zRzb60DOVlDanFoLfdv"
PROJECT_ID = 51948

# 課題タイプ (692971 または 692972)
ISSUE_TYPE_ID_EDGE = 692972  # 692971でも試す

# カスタムフィールド (エッジ種別)
EDGE_CUSTOM_FIELD_ID = 72497
EDGE_TYPE_INTERNET = 2  # インターネットエッジ

# APIエンドポイント
BACKLOG_API_URL = f"https://{BACKLOG_SPACE_ID}.backlog.com/api/v2/issues"

# **パラメータ (クエリパラメータ方式に変更)**
params = {
    "apiKey": API_KEY,
    f"customField_{EDGE_CUSTOM_FIELD_ID}": EDGE_TYPE_INTERNET  # クエリパラメータとして送る
}

# **リクエストボディ**
payload = {
    "projectId": PROJECT_ID,
    "summary": "ベーシック_【東京都あきる野市】小学校_1校_発注明細20250225.xlsx",
    "description": "自動アップロードにより添付されたExcelファイルです",
    "issueTypeId": ISSUE_TYPE_ID_EDGE,
    "priorityId": 3
}

# **デバッグ用に送信データを表示**
print("\n📤 送信データ:")
print(json.dumps(payload, indent=4, ensure_ascii=False))
print("\n📡 送信するパラメータ:")
print(json.dumps(params, indent=4, ensure_ascii=False))

# **APIリクエスト送信**
response = requests.post(BACKLOG_API_URL, params=params, json=payload)

# **レスポンスを表示**
print("\n📡 レスポンス:")
print(f"レスポンスコード: {response.status_code}")
print(f"レスポンス内容: {response.text}")

