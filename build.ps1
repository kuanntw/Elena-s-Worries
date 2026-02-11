Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvPython = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
$VersionFile = Join-Path $ProjectRoot "VERSION"

if (-not (Test-Path $VersionFile)) {
    Set-Content -Path $VersionFile -Value "v0.9.0" -Encoding utf8
}

$CurrentVersion = (Get-Content -Path $VersionFile -Raw).Trim()
if ($CurrentVersion -notmatch '^v(\d+)\.(\d+)\.(\d+)$') {
    throw "Invalid VERSION format: $CurrentVersion. Expected v<major>.<minor>.<patch>"
}

$Major = [int]$Matches[1]
$Minor = [int]$Matches[2]
$Patch = [int]$Matches[3]
$BuildVersion = $CurrentVersion
$NextVersion = "v$Major.$Minor.$($Patch + 1)"

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
$PrimaryExe = Join-Path $ReleaseDir "ResumeMailer.exe"
$DirectDownloadExe = Join-Path $ReleaseRoot "ResumeMailer.exe"
Copy-Item -Force $PrimaryExe $DirectDownloadExe
Set-Content -Path (Join-Path $ReleaseRoot "VERSION.txt") -Value $BuildVersion -Encoding utf8
Set-Content -Path $VersionFile -Value $NextVersion -Encoding utf8

Write-Host "Build complete."
Write-Host "EXE: $(Join-Path $ReleaseDir 'ResumeMailer.exe')"
Write-Host "Direct EXE: $DirectDownloadExe"
Write-Host "Built version: $BuildVersion"
Write-Host "Next version: $NextVersion"
