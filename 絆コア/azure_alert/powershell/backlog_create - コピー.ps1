# PowerShellのエンコーディングをUTF-8に設定
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = $OutputEncoding

# Backlog APIの基本設定
$apiKey = "KDqNG9Ozvmwq8J7O3EiDaygtE9oDGSzhXzB6v4oyQhodgiB3xCSAsC9MZ89mOZuk"  
$baseUrl = "https://ucdprj.backlog.com/api/v2"
$projectId = "596029"  
$issueTypeIdParent = "3144619"  
$issueTypeIdChild = "3144662"   
$statusIdProcessing = "3"  # 「処理中」のステータスID（要確認）
$assigneeIdSekine = "12345"  # 「関根」のユーザーID（要確認）

# 今日の日付を取得
$today = Get-Date -Format "yyyy年MM月dd日"

# **1. 親課題の作成**
$parentIssueTitle = "$today 絆Core日次監視"
$parentIssueData = @{
    "projectId" = $projectId
    "summary" = $parentIssueTitle
    "issueTypeId" = $issueTypeIdParent
    "priorityId" = 3
    "statusId" = $statusIdProcessing
    "assigneeId" = $assigneeIdSekine
} | ConvertTo-Json -Depth 10

Write-Output "親課題を作成中..."

try {
    $parentIssueResponse = Invoke-RestMethod -Uri "$baseUrl/issues?apiKey=$apiKey" -Method Post -Body $parentIssueData -ContentType "application/json; charset=utf-8"
    $parentIssueKey = $parentIssueResponse.id
    Write-Output "✅ 親課題作成成功: $parentIssueKey"
} catch {
    Write-Output "❌ 親課題作成エラー: $_"
    exit 1
}

# **2. 子課題の作成**
$childIssueTitle = "09時00分 日次監視"
$childIssueData = @{
    "projectId" = $projectId
    "summary" = $childIssueTitle
    "parentIssueId" = $parentIssueKey
    "issueTypeId" = $issueTypeIdChild
    "priorityId" = 3
    "statusId" = $statusIdProcessing
    "assigneeId" = $assigneeIdSekine
} | ConvertTo-Json -Depth 10

Write-Output "子課題を作成中..."

try {
    $childIssueResponse = Invoke-RestMethod -Uri "$baseUrl/issues?apiKey=$apiKey" -Method Post -Body $childIssueData -ContentType "application/json; charset=utf-8"
    $childIssueKey = $childIssueResponse.id
    Write-Output "✅ 子課題作成成功: $childIssueKey"
} catch {
    Write-Output "❌ 子課題作成エラー: $_"
    exit 1
}

# **3. Azureの500エラーログを取得**
Write-Output "Azureの500エラーログを取得中..."

$workspaceId = "57684c92-42c0-467e-99c4-03eeaea0ecad"  # Azure Log AnalyticsのワークスペースID

$query = @"
AzureDiagnostics
| where Category == "FrontDoorAccessLog"
| where toint(httpStatusCode_d) == 500
| where TimeGenerated > ago(4h)
| project TimeGenerated, originUrl_s
"@

try {
    $result = Invoke-AzOperationalInsightsQuery -WorkspaceId $workspaceId -Query $query
    $logEntries = $result.Results | ForEach-Object { "$($_.TimeGenerated) $($_.originUrl_s)" }
    $logText = ($logEntries -join "`n") -replace '"', '\"'  # 改行処理
    Write-Output "✅ Azureログ取得成功"
} catch {
    $logText = "❌ Azureログ取得失敗: $_"
    Write-Output $logText
}

# **4. 子課題にAzureのログをコメント追加**
$commentData = @{
    "content" = "以下のURLで500エラーが発生しました:`n$logText"
} | ConvertTo-Json -Depth 10

Write-Output "コメントを追加中..."

try {
    Invoke-RestMethod -Uri "$baseUrl/issues/$childIssueKey/comments?apiKey=$apiKey" -Method Post -Body $commentData -ContentType "application/json; charset=utf-8"
    Write-Output "✅ コメント追加成功"
} catch {
    Write-Output "❌ コメント追加エラー: $_"
    exit 1
}

Write-Output "🎉 Backlog課題が作成されました！"
