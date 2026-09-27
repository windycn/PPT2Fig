$ErrorActionPreference = "Stop"
$root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$helper = Join-Path $root "dist\helper\ppt2fig.exe"
$destination = Join-Path $env:LOCALAPPDATA "PPT2Fig"
New-Item -ItemType Directory -Force -Path $destination | Out-Null
if (Test-Path $helper) {
    Copy-Item $helper (Join-Path $destination "ppt2fig.exe") -Force
} else {
    py -3 -m venv (Join-Path $destination "runtime")
    $python = Join-Path $destination "runtime\Scripts\python.exe"
    & $python -m pip install "$root[gui]"
    $source = Join-Path $destination "runtime\Scripts\ppt2fig.exe"
    Copy-Item $source (Join-Path $destination "ppt2fig.exe") -Force
}
& (Join-Path $destination "ppt2fig.exe") doctor
if ($LASTEXITCODE -ne 0) {
    throw "Poppler 不完整。请安装 Poppler，或设置 PPT2FIG_POPPLER_DIR 指向其 bin 目录。"
}
Write-Host "辅助程序已安装。然后在 PowerPoint 的加载项管理中添加 addin/PPT2Fig-v1.ppam。"
