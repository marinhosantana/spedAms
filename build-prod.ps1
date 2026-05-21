$ErrorActionPreference = "Stop"

Set-Location -Path $PSScriptRoot

$pythonExe = Join-Path $PSScriptRoot ".venv-prod\Scripts\python.exe"
if (-not (Test-Path $pythonExe)) {
    $pythonExe = "python"
}

& $pythonExe -m pip install -r requirements.txt
& $pythonExe -m pip install -r requirements-build.txt
& $pythonExe -m PyInstaller `
    --noconfirm `
    --clean `
    --onedir `
    --windowed `
    --name "RevisorSPED_PROD" `
    --paths ".\Sped" `
    ".\Sped\build_entry_prod.py"

Copy-Item -Path ".\Sped\mysql_schema.sql" -Destination ".\dist\RevisorSPED_PROD\mysql_schema.sql" -Force

Write-Host "EXE de producao gerado em: dist\RevisorSPED_PROD\RevisorSPED_PROD.exe"
