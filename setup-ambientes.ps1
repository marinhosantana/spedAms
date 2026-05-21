$ErrorActionPreference = "Stop"

Set-Location -Path $PSScriptRoot

if (Get-Command py -ErrorAction SilentlyContinue) {
    $pythonLauncher = "py"
} else {
    $pythonLauncher = "python"
}

if (-not (Test-Path ".\.venv-dev\Scripts\python.exe")) {
    & $pythonLauncher -m venv .venv-dev
}

if (-not (Test-Path ".\.venv-prod\Scripts\python.exe")) {
    & $pythonLauncher -m venv .venv-prod
}

& ".\.venv-dev\Scripts\python.exe" -m pip install --upgrade pip
& ".\.venv-dev\Scripts\python.exe" -m pip install -r requirements.txt
& ".\.venv-dev\Scripts\python.exe" -m pip install -r requirements-build.txt

& ".\.venv-prod\Scripts\python.exe" -m pip install --upgrade pip
& ".\.venv-prod\Scripts\python.exe" -m pip install -r requirements.txt
& ".\.venv-prod\Scripts\python.exe" -m pip install -r requirements-build.txt

Write-Host "Ambientes preparados com sucesso."
Write-Host "Use .\run-dev.ps1 para desenvolvimento."
Write-Host "Use .\run-prod.ps1 para producao."
