Set-Location "C:\company-log"
$deployId = (Get-Content ".\.deploy_id" -Raw).Trim()
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm"

Write-Host "=== 1. Pushing code to GAS ===" -ForegroundColor Cyan
clasp push --force
if ($LASTEXITCODE -ne 0) { Write-Host "clasp push failed." -ForegroundColor Red; exit 1 }

Write-Host "`n=== 2. Updating GAS deployment ===" -ForegroundColor Cyan
clasp deploy -i $deployId -d "Update $timestamp"
if ($LASTEXITCODE -ne 0) { Write-Host "clasp deploy failed." -ForegroundColor Red; exit 1 }

if (Test-Path ".\.git") {
  Write-Host "`n=== 3. Pushing to GitHub ===" -ForegroundColor Cyan
  git add -A
  $changes = git status --porcelain
  if ($changes) {
    git commit -m "Update app ($timestamp)"
    git push -u origin main
    if ($LASTEXITCODE -ne 0) { Write-Host "git push failed." -ForegroundColor Red; exit 1 }
  } else {
    Write-Host "No changes to commit."
  }
}

Write-Host "`n=== Deploy and Git push completed ===" -ForegroundColor Green