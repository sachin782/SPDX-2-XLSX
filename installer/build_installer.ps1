#Requires -Version 5.1
<#
.SYNOPSIS
    Builds SPDX-Excel-Converter.exe and the Windows setup installer.

.DESCRIPTION
    1. Verifies Python 3.11+
    2. Installs Python dependencies and PyInstaller
    3. Downloads Microsoft Visual C++ Redistributable (prerequisite)
    4. Builds the standalone application executable
    5. Compiles the Inno Setup installer (publisher: Sachin Rawat)

.OUTPUTS
    dist\SPDX-Excel-Converter.exe
    dist\installer\SPDX-Excel-Converter-Setup.exe
#>

$ErrorActionPreference = "Stop"

$InstallerDir = $PSScriptRoot
$ProjectRoot = Split-Path -Parent $InstallerDir
$DistDir = Join-Path $ProjectRoot "dist"
$RedistDir = Join-Path $InstallerDir "redist"
$VcRedistUrl = "https://aka.ms/vs/17/release/vc_redist.x64.exe"
$VcRedistPath = Join-Path $RedistDir "vc_redist.x64.exe"
$AppExe = Join-Path $DistDir "SPDX-Excel-Converter.exe"
$SetupExe = Join-Path $DistDir "installer\SPDX-Excel-Converter-Setup.exe"
$SpecFile = Join-Path $InstallerDir "spdx_converter.spec"
$IssFile = Join-Path $InstallerDir "setup.iss"

function Write-Step([string]$Message) {
    Write-Host ""
    Write-Host "==> $Message" -ForegroundColor Cyan
}

function Find-Python {
    $candidates = @("python", "py")
    foreach ($cmd in $candidates) {
        try {
            $versionOutput = & $cmd -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')" 2>$null
            if ($LASTEXITCODE -eq 0 -and $versionOutput) {
                $parts = $versionOutput.Trim().Split(".")
                $major = [int]$parts[0]
                $minor = [int]$parts[1]
                if ($major -gt 3 -or ($major -eq 3 -and $minor -ge 11)) {
                    return $cmd
                }
            }
        } catch {
            continue
        }
    }
    throw "Python 3.11 or newer is required. Install from https://www.python.org/downloads/"
}

function Find-InnoSetupCompiler {
    $paths = @(
        "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
        "$env:ProgramFiles\Inno Setup 6\ISCC.exe",
        "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
    )
    foreach ($path in $paths) {
        if (Test-Path $path) {
            return $path
        }
    }
    return $null
}

Write-Host "SPDX JSON to Excel Converter - Installer Build" -ForegroundColor Green
Write-Host "Publisher: Sachin Rawat"

Write-Step "Checking Python"
$Python = Find-Python
Write-Host "Using: $Python"

Write-Step "Installing Python dependencies"
& $Python -m pip install --upgrade pip | Out-Null
& $Python -m pip install -r (Join-Path $ProjectRoot "requirements.txt") pyinstaller
if ($LASTEXITCODE -ne 0) {
    throw "Failed to install Python dependencies."
}

Write-Step "Downloading prerequisite: Microsoft Visual C++ Redistributable"
New-Item -ItemType Directory -Force -Path $RedistDir | Out-Null
if (-not (Test-Path $VcRedistPath)) {
    Write-Host "Downloading $VcRedistUrl ..."
    Invoke-WebRequest -Uri $VcRedistUrl -OutFile $VcRedistPath -UseBasicParsing
} else {
    Write-Host "Already present: $VcRedistPath"
}

Write-Step "Building standalone executable with PyInstaller"
New-Item -ItemType Directory -Force -Path $DistDir | Out-Null
Push-Location $ProjectRoot
try {
    & $Python -m PyInstaller --noconfirm --clean $SpecFile
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller build failed."
    }
} finally {
    Pop-Location
}

if (-not (Test-Path $AppExe)) {
    throw "Expected executable not found: $AppExe"
}
Write-Host "Built: $AppExe"

Write-Step "Compiling Windows installer with Inno Setup"
$Iscc = Find-InnoSetupCompiler
if (-not $Iscc) {
    Write-Host ""
    Write-Warning "Inno Setup 6 was not found. Install from https://jrsoftware.org/isdl.php"
    Write-Warning "The standalone executable is ready at: $AppExe"
    Write-Warning "After installing Inno Setup, re-run this script or compile manually:"
    Write-Warning "  `"$env:ProgramFiles(x86)\Inno Setup 6\ISCC.exe`" `"$IssFile`""
    exit 0
}

New-Item -ItemType Directory -Force -Path (Split-Path $SetupExe) | Out-Null
& $Iscc $IssFile
if ($LASTEXITCODE -ne 0) {
    throw "Inno Setup compilation failed."
}

Write-Host ""
Write-Host "Build complete!" -ForegroundColor Green
Write-Host "  Application: $AppExe"
Write-Host "  Installer:   $SetupExe"
Write-Host ""
Write-Host "End users can run SPDX-Excel-Converter-Setup.exe to install prerequisites and the application."
