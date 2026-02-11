Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvPython = Join-Path $ProjectRoot ".venv\Scripts\python.exe"

if (-not (Test-Path $VenvPython)) {
    & (Join-Path $ProjectRoot "bootstrap.ps1")
}

$DistDir = Join-Path $ProjectRoot "dist"
$BuildDir = Join-Path $ProjectRoot "build"
$ReleaseRoot = Join-Path $ProjectRoot "release"
$ReleaseDir = Join-Path $ReleaseRoot "ResumeMailer"

if (Test-Path $DistDir) { Remove-Item -Recurse -Force $DistDir }
if (Test-Path $BuildDir) { Remove-Item -Recurse -Force $BuildDir }
if (Test-Path $ReleaseDir) { Remove-Item -Recurse -Force $ReleaseDir }
New-Item -ItemType Directory -Path $ReleaseRoot -Force | Out-Null

$SplashImage = Join-Path $ProjectRoot "photo\elena.png"
if (-not (Test-Path $SplashImage)) {
    throw "Splash image not found: $SplashImage"
}

& $VenvPython -m PyInstaller `
    --noconfirm `
    --clean `
    --onedir `
    --windowed `
    --name ResumeMailer `
    --add-data "$SplashImage;photo" `
    (Join-Path $ProjectRoot "app\main.py")

if (-not (Test-Path (Join-Path $DistDir "ResumeMailer"))) {
    throw "Build failed: dist\\ResumeMailer not found."
}

Copy-Item -Recurse -Force (Join-Path $DistDir "ResumeMailer") $ReleaseDir

Write-Host "Build complete."
Write-Host "EXE: $(Join-Path $ReleaseDir 'ResumeMailer.exe')"
