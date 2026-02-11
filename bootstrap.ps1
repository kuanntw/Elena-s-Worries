Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvPath = Join-Path $ProjectRoot ".venv"
$PythonCmd = "py"

if (-not (Get-Command $PythonCmd -ErrorAction SilentlyContinue)) {
    throw "Python launcher 'py' not found."
}

if (-not (Test-Path $VenvPath)) {
    & $PythonCmd -m venv $VenvPath
}

$VenvPython = Join-Path $VenvPath "Scripts\python.exe"
if (-not (Test-Path $VenvPython)) {
    throw "Virtual environment python not found at $VenvPython"
}

& $VenvPython -m pip install --upgrade pip
& $VenvPython -m pip install -r (Join-Path $ProjectRoot "requirements.txt")

Write-Host "Bootstrap complete: $VenvPath"
