import requests
import os

# Backlog API の設定
BACKLOG_SPACE_ID = "ucdprj"
API_KEY = "d9J1kvSFf3oFVhIJESxjJ0rKfGRkEea7Fr2K2eRPcZwU7zRzb60DOVlDanFoLfdv"
PROJECT_ID = 51948  # プロジェクトID
ISSUE_TYPE_ID_DEFAULT = 450457  # 通常の課題タイプ
ISSUE_TYPE_ID_EDGE = 692971  # ベーシックファイル用の課題タイプ
PRIORITY_ID = 3  # 優先度 (中)

# カスタムフィールド (エッジ種別)
EDGE_CUSTOM_FIELD_ID = 72497
EDGE_TYPE_INTERNET = 2  # インターネットエッジ

# **対象フォルダ**
TARGET_FOLDER = r"C:\tools\バックログ起票\EXCEL"

# **課題作成 API URL**
BACKLOG_API_URL = f"https://{BACKLOG_SPACE_ID}.backlog.com/api/v2/issues"

# **コメント追加 API URL**
COMMENT_API_URL_TEMPLATE = f"https://{BACKLOG_SPACE_ID}.backlog.com/api/v2/issues/{{issue_id}}/comments"

# **ファイルアップロード API URL**
UPLOAD_TEMP_FILE_API_URL = f"https://{BACKLOG_SPACE_ID}.backlog.com/api/v2/space/attachment"

# **課題を作成**
def create_issue(file_name):
    """Backlog に課題を作成する"""
    
    # ファイル名の先頭が「ベーシック_」ならEDGEの課題タイプを使用
    if file_name.startswith("ベーシック_"):
        issue_type_id = ISSUE_TYPE_ID_EDGE
        params = {"apiKey": API_KEY, f"customField_{EDGE_CUSTOM_FIELD_ID}": EDGE_TYPE_INTERNET}
    else:
        issue_type_id = ISSUE_TYPE_ID_DEFAULT
        params = {"apiKey": API_KEY}
    
    payload = {
        "projectId": PROJECT_ID,
        "summary": file_name,  # ファイル名を課題タイトルに
        "description": "自動アップロードにより添付されたExcelファイルです",
        "issueTypeId": issue_type_id,
        "priorityId": PRIORITY_ID
    }
    
    response = requests.post(BACKLOG_API_URL, params=params, json=payload)
    print(f"課題作成レスポンス: {response.status_code}")
    print(f"レスポンス内容: {response.text}")

    if response.status_code in [200, 201]:
        issue_data = response.json()
        return issue_data.get("id")  # 課題IDを返す
    else:
        print("❌ 課題作成に失敗しました。")
        return None

# **ファイルを一時アップロードし、ファイルIDを取得**
def upload_temp_file(file_path):
    """Backlogの一時ファイル領域にアップロードし、ファイルIDを取得"""
    params = {"apiKey": API_KEY}
    
    # **ファイルの存在を確認**
    if not os.path.exists(file_path):
        print(f"❌ エラー: ファイルが見つかりません → {file_path}")
        return None
    
    with open(file_path, "rb") as file:
        files = {"file": file}
        response = requests.post(UPLOAD_TEMP_FILE_API_URL, params=params, files=files)
    
    print(f"一時ファイルアップロードレスポンス: {response.status_code}")
    print(f"レスポンス内容: {response.text}")
    
    if response.status_code in [200, 201]:
        upload_data = response.json()
        return upload_data["id"]  # 一時ファイルIDを取得
    else:
        print(f"❌ ファイル {file_path} のアップロードに失敗しました。")
        print(f"エラー詳細: {response.json()}")
        return None

# **コメントにファイルを添付**
def upload_file_as_comment(issue_id, file_id):
    """Backlogの課題にコメントとしてファイルを添付"""
    upload_url = COMMENT_API_URL_TEMPLATE.format(issue_id=issue_id)
    params = {"apiKey": API_KEY}
    
    data = {
        "content": "ファイルを添付しました。",
        "attachmentId[]": file_id  # 一時アップロードしたファイルのIDを指定
    }
    
    response = requests.post(upload_url, params=params, data=data)
    
    print(f"コメント追加レスポンス: {response.status_code}")
    print(f"レスポンス内容: {response.text}")
    
    if response.status_code in [200, 201]:
        print(f"✅ 課題 {issue_id} にファイルを添付しました。")
    else:
        print(f"❌ 課題 {issue_id} へのファイル添付に失敗しました。")
        print(f"エラー詳細: {response.json()}")

# **メイン処理**
if __name__ == "__main__":
    # **対象フォルダの全 `.xlsx` ファイルを取得**
    files = [f for f in os.listdir(TARGET_FOLDER) if f.endswith(".xlsx")]
    
    if not files:
        print("⚠️ 添付対象のファイルが見つかりません。")
    else:
        for file_name in files:
            file_path = os.path.join(TARGET_FOLDER, file_name)
            
            # **課題を作成**
            issue_id = create_issue(file_name)
            
            # **課題が作成できた場合のみファイルをアップロード**
            if issue_id:
                file_id = upload_temp_file(file_path)
                
                if file_id:
                    upload_file_as_comment(issue_id, file_id)
