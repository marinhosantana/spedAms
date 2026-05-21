$ErrorActionPreference = "Stop"

$env:SPED_ENV = "dev"
Set-Location -Path $PSScriptRoot

$pythonExe = Join-Path $PSScriptRoot ".venv-dev\Scripts\python.exe"
if (-not (Test-Path $pythonExe)) {
    $pythonExe = "python"
}

& $pythonExe .\Sped\main.py
