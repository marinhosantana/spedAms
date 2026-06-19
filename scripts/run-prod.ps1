$ErrorActionPreference = "Stop"

$projectRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$env:SPED_ENV = "prod"
Set-Location -Path $projectRoot

$pythonExe = Join-Path $projectRoot ".venv-prod\Scripts\python.exe"
if (-not (Test-Path $pythonExe)) {
    $pythonExe = "python"
}

& $pythonExe .\Sped\NovoRevisorQt.py
