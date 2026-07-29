# Build the Windows "full" installer payload:
#   1) PyInstaller one-file with pikepdf + ocrmypdf
#   2) Stage portable Tesseract + Ghostscript next to the exe
#   3) Compile Inno Setup → dist/coversheets-<ver>-windows-x64-setup.exe
#
# Usage (from repo root, with a venv that has the project installed):
#   powershell -ExecutionPolicy Bypass -File scripts/build_windows_full.ps1
#   powershell -File scripts/build_windows_full.ps1 -SkipInstallDeps
#
# CI installs choco packages + Inno Setup before calling this script.

[CmdletBinding()]
param(
    [switch]$SkipInstallDeps,
    [switch]$SkipInno,
    [string]$Version = ""
)

$ErrorActionPreference = "Stop"
$RepoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
Set-Location $RepoRoot

function Write-Step([string]$Message) {
    Write-Host ""
    Write-Host "==> $Message" -ForegroundColor Cyan
}

function Get-PackageVersion {
    if ($Version) { return $Version.TrimStart("v") }
    $init = Join-Path $RepoRoot "coversheets\__init__.py"
    $line = Select-String -Path $init -Pattern '__version__\s*=\s*["'']([^"'']+)["'']' | Select-Object -First 1
    if (-not $line) { throw "Could not read __version__ from coversheets/__init__.py" }
    return $line.Matches[0].Groups[1].Value
}

function Ensure-Command([string]$Name) {
    if (-not (Get-Command $Name -ErrorAction SilentlyContinue)) {
        throw "Required command not found on PATH: $Name"
    }
}

$AppVersion = Get-PackageVersion
$StageDir = Join-Path $RepoRoot "dist\windows-full"
$ToolsCache = Join-Path $RepoRoot "build\windows-tools"
$InnoScript = Join-Path $RepoRoot "installer\windows\coversheets.iss"

Write-Host "Windows full installer build"
Write-Host "  version : $AppVersion"
Write-Host "  stage   : $StageDir"

# --- Python deps + PyInstaller ----------------------------------------------
Write-Step "Install Python package extras (full + dev)"
if (-not $SkipInstallDeps) {
    Ensure-Command python
    python -m pip install --upgrade pip
    python -m pip install -e ".[full,dev]"
}

Write-Step "PyInstaller one-file (with collected extras)"
Ensure-Command pyinstaller
if (Test-Path (Join-Path $RepoRoot "dist\coversheets.exe")) {
    Remove-Item (Join-Path $RepoRoot "dist\coversheets.exe") -Force
}
pyinstaller coversheets.spec --noconfirm --clean
$BuiltExe = Join-Path $RepoRoot "dist\coversheets.exe"
if (-not (Test-Path $BuiltExe)) {
    throw "PyInstaller did not produce dist\coversheets.exe"
}

# --- Stage payload ----------------------------------------------------------
Write-Step "Stage payload directory"
if (Test-Path $StageDir) {
    Remove-Item $StageDir -Recurse -Force
}
New-Item -ItemType Directory -Path $StageDir | Out-Null
Copy-Item $BuiltExe (Join-Path $StageDir "coversheets.exe")

$ReadmeInstalled = @"
Automatic Exhibit Cover Sheets v$AppVersion (Windows full install)

Getting started
  1. Open "Automatic Exhibit Cover Sheets" from the Start Menu
  2. Click Open Folder… or Add PDFs…
  3. Edit Cover Label cells if needed
  4. Click Generate Cover Sheets

This install includes:
  • Linearize (pikepdf) — web-optimized PDFs
  • OCR (English) — ocrmypdf + Tesseract + Ghostscript, bundled next to the app

You do not need to install Python, Tesseract, or Ghostscript yourself.

Outputs are named +OriginalName.pdf next to each source (or in your chosen output folder).
Original PDFs are never modified.

Project: https://github.com/uchi-ansonfung/Auto-Cover-Sheets-Script
"@
Set-Content -Path (Join-Path $StageDir "README-INSTALLED.txt") -Value $ReadmeInstalled -Encoding UTF8

# --- Tesseract + Ghostscript ------------------------------------------------
function Install-ChocoIfNeeded {
    if (Get-Command choco -ErrorAction SilentlyContinue) { return }
    Write-Step "Install Chocolatey (package manager)"
    Set-ExecutionPolicy Bypass -Scope Process -Force
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
    Invoke-Expression ((New-Object System.Net.WebClient).DownloadString("https://community.chocolatey.org/install.ps1"))
    $env:Path = "C:\ProgramData\chocolatey\bin;" + $env:Path
    Ensure-Command choco
}

function Copy-TesseractTree {
    param([string[]]$SearchRoots)

    $exe = $null
    foreach ($root in $SearchRoots) {
        if (-not $root -or -not (Test-Path $root)) { continue }
        $found = Get-ChildItem -Path $root -Filter "tesseract.exe" -Recurse -ErrorAction SilentlyContinue |
            Select-Object -First 1
        if ($found) {
            $exe = $found
            break
        }
    }
    if (-not $exe) { return $false }

    $srcDir = $exe.Directory.FullName
    $dest = Join-Path $StageDir "tesseract"
    Write-Host "  Copying Tesseract from $srcDir"
    New-Item -ItemType Directory -Path $dest -Force | Out-Null
    Copy-Item -Path (Join-Path $srcDir "*") -Destination $dest -Recurse -Force

    # Ensure eng traineddata exists (some installs put tessdata elsewhere).
    $eng = Join-Path $dest "tessdata\eng.traineddata"
    if (-not (Test-Path $eng)) {
        $engSrc = Get-ChildItem -Path $SearchRoots -Filter "eng.traineddata" -Recurse -ErrorAction SilentlyContinue |
            Select-Object -First 1
        if ($engSrc) {
            $tessdata = Join-Path $dest "tessdata"
            New-Item -ItemType Directory -Path $tessdata -Force | Out-Null
            Copy-Item $engSrc.FullName $eng -Force
        }
    }
    if (-not (Test-Path $eng)) {
        Write-Warning "eng.traineddata not found next to tesseract; downloading official eng data"
        $tessdata = Join-Path $dest "tessdata"
        New-Item -ItemType Directory -Path $tessdata -Force | Out-Null
        $url = "https://github.com/tesseract-ocr/tessdata_fast/raw/main/eng.traineddata"
        Invoke-WebRequest -Uri $url -OutFile $eng -UseBasicParsing
    }
    return (Test-Path (Join-Path $dest "tesseract.exe"))
}

function Copy-GhostscriptTree {
    param([string[]]$SearchRoots)

    $exe = $null
    foreach ($root in $SearchRoots) {
        if (-not $root -or -not (Test-Path $root)) { continue }
        foreach ($name in @("gswin64c.exe", "gswin32c.exe")) {
            $found = Get-ChildItem -Path $root -Filter $name -Recurse -ErrorAction SilentlyContinue |
                Select-Object -First 1
            if ($found) {
                $exe = $found
                break
            }
        }
        if ($exe) { break }
    }
    if (-not $exe) { return $false }

    # Prefer the Ghostscript version root (parent of bin).
    $binDir = $exe.Directory.FullName
    $gsRoot = if ($binDir.ToLower().EndsWith("\bin")) {
        (Resolve-Path (Join-Path $binDir "..")).Path
    } else {
        $binDir
    }

    $dest = Join-Path $StageDir "ghostscript"
    Write-Host "  Copying Ghostscript from $gsRoot"
    if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
    New-Item -ItemType Directory -Path $dest -Force | Out-Null
    Copy-Item -Path (Join-Path $gsRoot "*") -Destination $dest -Recurse -Force
    return $true
}

New-Item -ItemType Directory -Path $ToolsCache -Force | Out-Null

$tessSearch = @(
    "C:\Program Files\Tesseract-OCR"
    "C:\Program Files (x86)\Tesseract-OCR"
    $ToolsCache
)
$gsSearch = @(
    "C:\Program Files\gs"
    "C:\Program Files (x86)\gs"
    $ToolsCache
)

Write-Step "Locate or install Tesseract"
if (-not (Copy-TesseractTree -SearchRoots $tessSearch)) {
    if ($SkipInstallDeps) {
        throw "Tesseract not found and -SkipInstallDeps was set"
    }
    Install-ChocoIfNeeded
    choco install tesseract -y --no-progress
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
                [System.Environment]::GetEnvironmentVariable("Path", "User")
    if (-not (Copy-TesseractTree -SearchRoots $tessSearch)) {
        throw "Failed to stage Tesseract after choco install"
    }
}

Write-Step "Locate or install Ghostscript"
if (-not (Copy-GhostscriptTree -SearchRoots $gsSearch)) {
    if ($SkipInstallDeps) {
        throw "Ghostscript not found and -SkipInstallDeps was set"
    }
    Install-ChocoIfNeeded
    choco install ghostscript -y --no-progress
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
                [System.Environment]::GetEnvironmentVariable("Path", "User")
    if (-not (Copy-GhostscriptTree -SearchRoots $gsSearch)) {
        throw "Failed to stage Ghostscript after choco install"
    }
}

# Sanity checks
$mustExist = @(
    (Join-Path $StageDir "coversheets.exe")
    (Join-Path $StageDir "tesseract\tesseract.exe")
    (Join-Path $StageDir "tesseract\tessdata\eng.traineddata")
)
foreach ($p in $mustExist) {
    if (-not (Test-Path $p)) { throw "Missing staged file: $p" }
}
$gsOk = Get-ChildItem (Join-Path $StageDir "ghostscript") -Filter "gswin*.exe" -Recurse -ErrorAction SilentlyContinue |
    Select-Object -First 1
if (-not $gsOk) { throw "Ghostscript CLI not found under dist\windows-full\ghostscript" }

# Portable Ghostscript needs lib/ (and usually gsdll*.dll) so jbig2dec can
# decode scanned PDFs that use /JBIG2Decode. A bin-only copy is not enough.
$gsLib = Join-Path $StageDir "ghostscript\lib"
if (-not (Test-Path $gsLib)) {
    # Some layouts nest version dirs; accept any lib folder under ghostscript.
    $gsLib = Get-ChildItem (Join-Path $StageDir "ghostscript") -Filter "lib" -Directory -Recurse -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if (-not $gsLib) {
        throw "Ghostscript lib/ folder missing under dist\windows-full\ghostscript (needed for JBIG2/jbig2dec)"
    }
}
$gsDll = Get-ChildItem (Join-Path $StageDir "ghostscript") -Filter "gsdll*.dll" -Recurse -ErrorAction SilentlyContinue |
    Select-Object -First 1
if (-not $gsDll) {
    Write-Warning "gsdll*.dll not found under staged Ghostscript; JBIG2 decode may fail on some systems"
} else {
    Write-Host "  Ghostscript DLL: $($gsDll.FullName)"
}

Write-Host "Staged payload:"
Get-ChildItem $StageDir | Format-Table Name, Length -AutoSize

# --- Inno Setup -------------------------------------------------------------
if ($SkipInno) {
    Write-Step "Skipping Inno Setup (-SkipInno)"
    Write-Host "Payload ready at $StageDir"
    exit 0
}

Write-Step "Compile Inno Setup installer"
$iscc = @(
    "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe"
    "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
    "${env:LOCALAPPDATA}\Programs\Inno Setup 6\ISCC.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $iscc) {
    if ($SkipInstallDeps) {
        throw "ISCC.exe not found and -SkipInstallDeps was set"
    }
    Install-ChocoIfNeeded
    choco install innosetup -y --no-progress
    $iscc = @(
        "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe"
        "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
}
if (-not $iscc) { throw "Inno Setup compiler (ISCC.exe) not found" }

$payloadAbs = (Resolve-Path $StageDir).Path
& $iscc `
    "/DMyAppVersion=$AppVersion" `
    "/DPayloadDir=$payloadAbs" `
    $InnoScript

$SetupName = "coversheets-$AppVersion-windows-x64-setup.exe"
$SetupPath = Join-Path $RepoRoot "dist\$SetupName"
if (-not (Test-Path $SetupPath)) {
    throw "Inno Setup did not produce $SetupPath"
}

Write-Step "Done"
Write-Host "Installer: $SetupPath"
Write-Host "Size: $([math]::Round((Get-Item $SetupPath).Length / 1MB, 1)) MB"
