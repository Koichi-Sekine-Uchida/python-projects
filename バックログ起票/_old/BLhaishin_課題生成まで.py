import os
import requests

# ======== 設定部分 ========
api_key = 'd9J1kvSFf3oFVhIJESxjJ0rKfGRkEea7Fr2K2eRPcZwU7zRzb60DOVlDanFoLfdv'
# プロジェクトの数値ID（Backlogの管理画面で確認）
project_id = 51948

# 課題作成のエンドポイント（必要に応じてドメインなど修正）
url = f"https://ucdprj.backlog.com/api/v2/issues?apiKey={api_key}"

# 監視対象のフォルダパス（例: 生文字列を使用）
folder_path = r'C:\tools\selenium\配信BackLog\配信用フォルダ'

def main():
    # フォルダが存在するかチェック
    if not os.path.exists(folder_path):
        print("指定されたフォルダパスが存在しません:", folder_path)
        return

    # フォルダ内の各ファイルを処理
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        if os.path.isfile(file_path):
            data = {
                'projectId': project_id,
                'summary': file_name,
                'description': '自動アップロードにより添付されたファイルです',
                'issueTypeId': 450457,  # 例: 「配信」の課題種別ID
                'priorityId': 3       # 例: 「中」など、適切な優先度ID
            }
            with open(file_path, 'rb') as f:
                files = {'file': f}
                response = requests.post(url, data=data, files=files)

            if response.status_code == 201:
                print(f"{file_name} の課題が正常に作成されました。")
            else:
                print(f"{file_name} の作成中にエラーが発生しました: {response.text}")

if __name__ == "__main__":
    main()
