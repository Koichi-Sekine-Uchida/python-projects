from azure.monitor.query import LogsQueryClient
from azure.identity import AzureCliCredential
from datetime import timedelta
import pandas as pd

# Azure 認証情報
credential = AzureCliCredential(tenant_id="9748a44c-e147-44bc-9854-6d875f421853")
client = LogsQueryClient(credential)

# **ワークスペース ID を固定**
workspace_id = "57684c92-42c0-467e-99c4-03eeaea0ecad"

# 📌 クエリ修正（AzureDiagnostics を使用）
query = """
AzureDiagnostics
| where Category == "FrontDoorAccessLog"
| where TimeGenerated > ago(4h)
| where toint(httpStatusCode_d) == 500
| project TimeGenerated, httpStatusCode_d, requestUri_s
"""

# クエリ実行
response = client.query_workspace(
    workspace_id=workspace_id,
    query=query,
    timespan=None  # `ago()` を使っているため `timespan` は不要
)

# 📌 結果を処理
if response.tables and len(response.tables) > 0 and len(response.tables[0].rows) > 0:
    table = response.tables[0]  # 最初のテーブルのみを使用
    print("Table columns:", table.columns)  # デバッグ出力

    # **エラー回避: columns をそのまま使う**
    columns = table.columns  # `column_names` は不要なので削除

    # データフレームに変換
    df = pd.DataFrame(table.rows, columns=columns)

    # **カラムが存在するかチェックして選択**
    selected_columns = [col for col in ["TimeGenerated", "httpStatusCode_d", "requestUri_s"] if col in df.columns]
    df = df[selected_columns]

    # **テーブル形式で表示**
    print("\n🚀 取得した 500 エラーログ:")
    print(df.to_string(index=False))  # 表形式で出力

else:
    print("\n🚀 指定した時間内に 500 エラーは検出されませんでした。")
