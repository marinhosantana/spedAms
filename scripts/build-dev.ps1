$ErrorActionPreference = "Stop"

$projectRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
Set-Location -Path $projectRoot

$pythonExe = Join-Path $projectRoot ".venv-dev\Scripts\python.exe"
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
    --name "RevisorSPED_DEV" `
    --paths ".\Sped" `
    --add-data ".\Sped\assets;assets" `
    ".\Sped\build_entries\dev.py"

Copy-Item -Path ".\Sped\mysql_schema.sql" -Destination ".\dist\RevisorSPED_DEV\mysql_schema.sql" -Force

Write-Host "EXE de desenvolvimento gerado em: dist\RevisorSPED_DEV\RevisorSPED_DEV.exe"
