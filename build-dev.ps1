$ErrorActionPreference = "Stop"

Set-Location -Path $PSScriptRoot

$pythonExe = Join-Path $PSScriptRoot ".venv-dev\Scripts\python.exe"
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
    --name "RevisorSPED_DEV" `
    --paths ".\Sped" `
    ".\Sped\build_entry_dev.py"

Copy-Item -Path ".\Sped\mysql_schema.sql" -Destination ".\dist\RevisorSPED_DEV\mysql_schema.sql" -Force

Write-Host "EXE de desenvolvimento gerado em: dist\RevisorSPED_DEV\RevisorSPED_DEV.exe"
