$ErrorActionPreference = "Stop"

$projectRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
Set-Location -Path $projectRoot

$pythonExe = Join-Path $projectRoot ".venv-prod\Scripts\python.exe"
if (-not (Test-Path $pythonExe)) {
    $pythonExe = "python"
}

& $pythonExe -m pip install -r .\requirements\base.txt
& $pythonExe -m pip install -r .\requirements\build.txt
& $pythonExe -m PyInstaller `
    --noconfirm `
    --clean `
    --onedir `
    --windowed `
    --name "RevisorSPED_PROD" `
    --paths ".\Sped" `
    --add-data ".\Sped\assets;assets" `
    ".\Sped\build_entries\prod.py"

Copy-Item -Path ".\Sped\mysql_schema.sql" -Destination ".\dist\RevisorSPED_PROD\mysql_schema.sql" -Force

Write-Host "EXE de producao gerado em: dist\RevisorSPED_PROD\RevisorSPED_PROD.exe"
