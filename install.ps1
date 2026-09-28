$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

if (-not (Test-Path ".venv\\Scripts\\python.exe")) {
    uv venv .venv
}

uv pip install `
    --python ".venv\\Scripts\\python.exe" `
    --index-url "https://pypi.tuna.tsinghua.edu.cn/simple" `
    -r ".\\requirements.txt"

Write-Host "依赖安装完成。启动命令："
Write-Host '& ".\\.venv\\Scripts\\python.exe" -m uvicorn app.main:app --reload --port 8000'
