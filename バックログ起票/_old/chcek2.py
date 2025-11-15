import requests

# Backlog API の設定
API_KEY = "d9J1kvSFf3oFVhIJESxjJ0rKfGRkEea7Fr2K2eRPcZwU7zRzb60DOVlDanFoLfdv"
PROJECT_ID = 51948  # プロジェクトID

# カスタムフィールド取得APIエンドポイント
URL_CUSTOM_FIELDS = f"https://ucdprj.backlog.com/api/v2/projects/{PROJECT_ID}/customFields?apiKey={API_KEY}"

# APIリクエスト送信
response = requests.get(URL_CUSTOM_FIELDS)

# APIレスポンス処理
if response.status_code == 200:
    custom_fields = response.json()

    # "エッジ種別" に該当するカスタムフィールドを探す
    for field in custom_fields:
        if "エッジ種別" in field["name"]:  # フィールド名に「エッジ種別」が含まれるものを抽出
            print("カスタムフィールド名:", field["name"])
            print("カスタムフィールドID:", field["id"])
            print("フィールドタイプ:", field["typeId"])
            
            # エッジ種別の選択肢がある場合、表示
            if "items" in field:
                print("エッジ種別の選択肢:")
                for item in field["items"]:
                    print(f" - {item['name']} (ID: {item['id']})")
            print("--------")
else:
    print(f"❌ APIリクエストに失敗しました。ステータスコード: {response.status_code}")
    print(f"レスポンス内容: {response.text}")
