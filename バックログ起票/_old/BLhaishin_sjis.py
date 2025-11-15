
# -*- coding: shift_jis -*-
import requests
import json

# Backlog APIの設定
BACKLOG_SPACE_ID = "ucdprj"  # スペースID
API_KEY = "d9J1kvSFf3oFVhIJESxjJ0rKfGRkEea7Fr2K2eRPcZwU7zRzb60DOVlDanFoLfdv"  # APIキー
PROJECT_ID = 51948  # プロジェクトID
ISSUE_TYPE_ID = 450457  # 課題タイプID
PRIORITY_ID = 3  # 優先度ID
FILE_PATH = r"C:\tools\selenium\配信BackLog\テスト.xlsx"

# APIエンドポイント
BACKLOG_API_URL = f"https://{BACKLOG_SPACE_ID}.backlog.com/api/v2/issues"

# 課題作成データ
DATA = {
    "projectId": PROJECT_ID,
    "summary": "Python API 課題作成テスト",
    "description": "APIを使って課題を作成するテスト",
    "issueTypeId": ISSUE_TYPE_ID,
    "priorityId": PRIORITY_ID,
    "apiKey": API_KEY
}

# ヘッダー設定
HEADERS = {"Content-Type": "application/json"}

# 課題を作成
response = requests.post(BACKLOG_API_URL, headers=HEADERS, json=DATA)

# 課題作成が成功したらissue_idを取得
if response.status_code in [200, 201]:
    issue_data = response.json()
    issue_id = issue_data["id"]
    issue_key = issue_data["issueKey"]
    print(f"課題作成成功: {issue_key} (ID: {issue_id})")
else:
    print("課題作成に失敗しました。")
    print(f"レスポンス: {response.text}")
    exit()

# ファイル添付
UPLOAD_API_URL = f"https://{BACKLOG_SPACE_ID}.backlog.com/api/v2/issues/{issue_id}/attachments"
files = {"file": open(FILE_PATH, "rb")}
params = {"apiKey": API_KEY}

upload_response = requests.post(UPLOAD_API_URL, params=params, files=files)

if upload_response.status_code in [200, 201]:
    print("ファイルのアップロードに成功しました。")
else:
    print("ファイルのアップロードに失敗しました。")
    print(f"レスポンス: {upload_response.text}")
