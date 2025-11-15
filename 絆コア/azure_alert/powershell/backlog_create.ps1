# **エンコーディングをUTF-8に設定**
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = $OutputEncoding

# **Backlog APIの基本設定**
$apiKey = "KDqNG9Ozvmwq8J7O3EiDaygtE9oDGSzhXzB6v4oyQhodgiB3xCSAsC9MZ89mOZuk"
$baseUrl = "https://ucdprj.backlog.com/api/v2"
$projectId = "596029"

# **課題タイプID**
$issueTypeIdParent = 3144619  # 親課題
$issueTypeIdChild = 3144662   # 子課題

# **ステータスID（処理中）**
$statusIdProcessing = 2       

# **担当者（関根）**
$assigneeIdSekine = 10005898  

# **今日の日付**
$today = Get-Date -Format "yyyy年MM月dd日"

# **📌 親課題の作成**
$parentIssueData = @{
    projectId   = $projectId
    summary     = "$today 絆コア日次監視"
    issueTypeId = $issueTypeIdParent
    priorityId  = 3
    statusId    = $statusIdProcessing
    assigneeId  = $assigneeIdSekine
} | ConvertTo-Json -Depth 10 -Compress

Write-Output "Backlog Main Create"

try {
    $parentIssueResponse = Invoke-RestMethod -Uri "$baseUrl/issues?apiKey=$apiKey" -Method Post -Body $parentIssueData -ContentType "application/json"
    $parentIssueKey = $parentIssueResponse.id
    Write-Output "✅ 親課題作成成功: $parentIssueKey"
} catch {
    Write-Output "❌ 親課題作成エラー: $_"
    exit 1
}

# **📌 子課題の作成**
$childIssueData = @{
    projectId     = $projectId
    summary       = "09時00分 日次監視"
    parentIssueId = $parentIssueKey
    issueTypeId   = $issueTypeIdChild
    priorityId    = 3
    statusId      = $statusIdProcessing
    assigneeId    = $assigneeIdSekine
} | ConvertTo-Json -Depth 10 -Compress

Write-Output "Backlog Sub Create"

try {
    $childIssueResponse = Invoke-RestMethod -Uri "$baseUrl/issues?apiKey=$apiKey" -Method Post -Body $childIssueData -ContentType "application/json"
    $childIssueKey = $childIssueResponse.id
    Write-Output "✅ 子課題作成成功: $childIssueKey"
} catch {
    Write-Output "❌ 子課題作成エラー: $_"
    exit 1
}

# **📌 ステータスを「処理中」に変更 & 担当者を関根に設定**
$updateData = @{
    statusId = $statusIdProcessing
    assigneeId = $assigneeIdSekine
} | ConvertTo-Json -Depth 10 -Compress

Invoke-RestMethod -Uri "$baseUrl/issues/$parentIssueKey?apiKey=$apiKey" -Method Patch -Body $updateData -ContentType "application/json"
Invoke-RestMethod -Uri "$baseUrl/issues/$childIssueKey?apiKey=$apiKey" -Method Patch -Body $updateData -ContentType "application/json"

Write-Output "✅ ステータス変更完了（処理中）& 担当者設定"

# **📌 Azureログを取得してBacklogにコメント追加**
$azureLogs = Get-Content "C:\tools\logs\azure_log.txt" -Raw  # Azureログファイルを読み込む

$commentData = @{
    content = "Azureログ: `n$azureLogs"
} | ConvertTo-Json -Depth 10 -Compress

Invoke-RestMethod -Uri "$baseUrl/issues/$childIssueKey/comments?apiKey=$apiKey" -Method Post -Body $commentData -ContentType "application/json"

Write-Output "✅ AzureログをBacklogに追加完了"

Write-Output "Backlog 課題が正常に作成されました！"
