# ==============================================================================
# deploy_sync.ps1 : Life & Work OS ローカル反映・GAS固定デプロイ・Git同期スクリプト
# ==============================================================================
param(
    [string]$DriveRoot = "",
    [string]$LocalRepo = "C:\company-log",
    [string]$LocalMirror = "C:\lifework-os-web",
    [string]$BackupRoot = "C:\company-log_backup",
    [string]$CommitMessage = "",
    [switch]$NoPause
)

$ErrorActionPreference = "Stop"

# 1. Google Drive ソースフォルダの検出
if ([string]::IsNullOrWhiteSpace($DriveRoot)) {
    $candidates = @(
        "G:\マイドライブ\gemini",
        "$env:USERPROFILE\Google ドライブ\マイドライブ\gemini",
        "$env:USERPROFILE\Google Drive\My Drive\gemini"
    )
    foreach ($c in $candidates) {
        if (Test-Path (Join-Path $c "01_現行ソースコード")) { $DriveRoot = $c; break }
    }
}

$sourceDir = if ($DriveRoot) { Join-Path $DriveRoot "01_現行ソースコード" } else { "" }
if ([string]::IsNullOrWhiteSpace($sourceDir) -or -not (Test-Path $sourceDir)) {
    throw "Google Drive ソースフォルダ (01_現行ソースコード) が見つかりません。-DriveRoot を指定してください。"
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# 2. バックアップ取得（ローカル退避 & Driveアーカイブ）
Write-Host "=== [1/4] バックアップ退避中... ===" -ForegroundColor Cyan
$localBackupDir = Join-Path $BackupRoot "snap_$timestamp"
New-Item -ItemType Directory -Path $localBackupDir -Force | Out-Null

# C:\ 直下の旧バックアップを整理
Get-ChildItem -Path "C:\" -Filter "*.bak*" -File -ErrorAction SilentlyContinue | ForEach-Object {
    Move-Item -Path $_.FullName -Destination (Join-Path $localBackupDir $_.Name) -Force
}

# ローカル現行ソースコードをバックアップ
if (Test-Path $LocalRepo) {
    Copy-Item "$LocalRepo\*.gs", "$LocalRepo\*.html", "$LocalRepo\*.json" $localBackupDir -ErrorAction SilentlyContinue
}

# Google Drive 側のアーカイブ作成
$driveArchiveDir = Join-Path $DriveRoot "03_更新履歴_アーカイブ\backup_$timestamp"
if (Test-Path (Join-Path $DriveRoot "03_更新履歴_アーカイブ")) {
    New-Item -ItemType Directory -Path $driveArchiveDir -Force | Out-Null
    Copy-Item "$LocalRepo\*.gs", "$LocalRepo\*.html", "$LocalRepo\*.json" $driveArchiveDir -ErrorAction SilentlyContinue
    Write-Host "Drive側アーカイブ保存: $driveArchiveDir" -ForegroundColor Green
}
Write-Host "ローカル退避完了: $localBackupDir" -ForegroundColor Green

# 3. Google Drive から最新ソースコードをローカルへフラット反映
Write-Host "=== [2/4] Google Drive から最新コードを $LocalRepo へ反映中... ===" -ForegroundColor Cyan
$updatedFiles = Get-ChildItem -Path $sourceDir -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in ".gs", ".html", ".json" }
if (-not $updatedFiles -or $updatedFiles.Count -eq 0) {
    Write-Warning "01_現行ソースコードに対象ファイルが見つかりませんでした。"
} else {
    foreach ($f in $updatedFiles) {
        Copy-Item -Path $f.FullName -Destination (Join-Path $LocalRepo $f.Name) -Force
        if (Test-Path $LocalMirror) {
            Copy-Item -Path $f.FullName -Destination (Join-Path $LocalMirror $f.Name) -Force
        }
        Write-Host "反映: $($f.Name)" -ForegroundColor Gray
    }
    Write-Host "ソースコード反映完了 ($($updatedFiles.Count) ファイル)" -ForegroundColor Green
}

# 4. clasp による Google Apps Script 本番デプロイ（固定デプロイID維持）
Write-Host "=== [3/4] clasp による Google Apps Script 本番デプロイ中... ===" -ForegroundColor Cyan
if (Get-Command clasp -ErrorAction SilentlyContinue) {
    Push-Location $LocalRepo
    try {
        if (-not (Test-Path ".clasp.json")) {
            throw ".clasp.json が見つかりません: $LocalRepo\.clasp.json"
        }
        if (-not (Test-Path ".deploy_id")) {
            throw ".deploy_id が見つかりません: $LocalRepo\.deploy_id"
        }

        $deployId = (Get-Content ".deploy_id" -Raw).Trim()
        Write-Host "固定デプロイID: $deployId" -ForegroundColor Yellow

        Write-Host "clasp push --force 実行中..."
        clasp push --force
        if ($LASTEXITCODE -ne 0) { throw "clasp push に失敗しました。" }

        Write-Host "clasp deploy (固定ID上書き更新) 実行中..."
        clasp deploy -i $deployId -d "Update $timestamp"
        if ($LASTEXITCODE -ne 0) { throw "clasp deploy に失敗しました。" }

        Write-Host "GAS 本番デプロイ完了（固定デプロイID維持）" -ForegroundColor Green
    } finally {
        Pop-Location
    }
} else {
    Write-Warning "clasp コマンドが見つかりませんでした。"
}

# 5. Git ステージング・コミット・プッシュ
Write-Host "=== [4/4] Git ステージング・コミット・プッシュ中... ===" -ForegroundColor Cyan
$gitRepos = @($LocalRepo)
if ((Test-Path $LocalMirror) -and (Test-Path (Join-Path $LocalMirror ".git"))) {
    $gitRepos += $LocalMirror
}

foreach ($repo in $gitRepos) {
    if ((Get-Command git -ErrorAction SilentlyContinue) -and (Test-Path (Join-Path $repo ".git"))) {
        Push-Location $repo
        try {
            Write-Host "Git 同期実行中: $repo" -ForegroundColor Yellow
            git add .
            $status = git status --porcelain
            if (-not [string]::IsNullOrWhiteSpace($status)) {
                $msg = if ([string]::IsNullOrWhiteSpace($CommitMessage)) { "Update deploy ($timestamp)" } else { $CommitMessage }
                git commit -m $msg
                git push -u origin main
                Write-Host "Git push 完了: $repo" -ForegroundColor Green
            } else {
                Write-Host "Git 変更差分はありませんでした: $repo" -ForegroundColor Gray
            }
        } catch {
            Write-Warning "Git 同期中に警告/エラーが発生しました ($repo): $_"
        } finally {
            Pop-Location
        }
    }
}

Write-Host "`n=== すべての反映・デプロイ・同期が正常に完了しました ($timestamp) ===" -ForegroundColor Green

if (-not $NoPause -and [Environment]::UserInteractive) {
    Write-Host ""
    # non-interactive 環境でブロックされないよう配慮
}
