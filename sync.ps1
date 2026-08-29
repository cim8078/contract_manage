# 一键同步：提交本地修改并推送到 GitHub
# 用法：
#   .\sync.ps1                  # 自动生成提交说明（当前时间）
#   .\sync.ps1 "修复登录问题"    # 使用自定义提交说明

param(
    [string]$Message = ""
)

Set-Location $PSScriptRoot

if ([string]::IsNullOrWhiteSpace($Message)) {
    $Message = "同步更新 $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
}

git add -A
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

git commit -m $Message
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

git push
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Write-Host "`n已同步到 GitHub: https://github.com/cim8078/contract_manage" -ForegroundColor Green
