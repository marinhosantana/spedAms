$ErrorActionPreference = "Stop"

$projectRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$env:SPED_ENV = "dev"
Set-Location -Path $projectRoot

$pythonExe = Join-Path $projectRoot ".venv-dev\Scripts\python.exe"
if (-not (Test-Path $pythonExe)) {
    $pythonExe = "python"
}

& $pythonExe .\Sped\main.py
