$ErrorActionPreference = "Stop"

$env:SPED_ENV = "prod"
Set-Location -Path $PSScriptRoot

$pythonExe = Join-Path $PSScriptRoot ".venv-prod\Scripts\python.exe"
if (-not (Test-Path $pythonExe)) {
    $pythonExe = "python"
}

& $pythonExe .\Sped\main.py
