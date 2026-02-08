#!/usr/bin/env pwsh
# Build Script for SherpaOnnx SAPI5 Installer

$ErrorActionPreference = "Stop"

Write-Host "=== SherpaOnnx SAPI5 Installer Build Script ===" -ForegroundColor Cyan
Write-Host ""

# Check prerequisites
Write-Host "Checking prerequisites..." -ForegroundColor Yellow

# Check for .NET 6 SDK
if (-not (Get-Command "dotnet" -ErrorAction SilentlyContinue)) {
    Write-Host "ERROR: .NET 6 SDK not found. Please install from https://dotnet.microsoft.com/download" -ForegroundColor Red
    exit 1
}
Write-Host "  ✓ .NET SDK found" -ForegroundColor Green

# Check for WiX Toolset
$wixTargets = "${env:ProgramFiles(x86)}\WiX Toolset v3.11\bin\WixProjects.targets"
if (-not (Test-Path $wixTargets)) {
    Write-Host "WARNING: WiX Toolset not found at standard location" -ForegroundColor Yellow
    Write-Host "  Please install WiX Toolset v3.11 or later from: https://wixtoolset.org/" -ForegroundColor Yellow
    Write-Host ""
    $continue = Read-Host "Continue anyway? (y/n)"
    if ($continue -ne "y") { exit 1 }
} else {
    Write-Host "  ✓ WiX Toolset found" -ForegroundColor Green
}

# Check for Visual Studio (for building native DLL)
if (-not (Test-Path "${env:ProgramFiles}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe") -and
    -not (Test-Path "${env:ProgramFiles}\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe") -and
    -not (Test-Path "${env:ProgramFiles}\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe")) {
    Write-Host "WARNING: Visual Studio 2022 not found at standard location" -ForegroundColor Yellow
    Write-Host "  You will need a pre-built NativeTTSWrapper.dll" -ForegroundColor Yellow
} else {
    Write-Host "  ✓ Visual Studio 2022 found" -ForegroundColor Green
}

Write-Host ""

# Build steps
Write-Host "Building components..." -ForegroundColor Cyan
Write-Host ""

# Step 1: Build Native DLL
Write-Host "[1/4] Building NativeTTSWrapper.dll..." -ForegroundColor Yellow
$vsPath = "${env:ProgramFiles}\Microsoft Visual Studio\2022"
$msBuildPaths = @(
    "$vsPath\Community\MSBuild\Current\Bin\MSBuild.exe",
    "$vsPath\Professional\MSBuild\Current\Bin\MSBuild.exe",
    "$vsPath\Enterprise\MSBuild\Current\Bin\MSBuild.exe"
)

$msBuild = $msBuildPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

if ($msBuild) {
    & $msBuild "NativeTTSWrapper\NativeTTSWrapper.sln" /t:Build /p:Configuration=Release /p:Platform=x64 /v:minimal
    if ($LASTEXITCODE -ne 0) {
        Write-Host "ERROR: Failed to build NativeTTSWrapper.dll" -ForegroundColor Red
        exit 1
    }
    Write-Host "  ✓ Native DLL built successfully" -ForegroundColor Green
} else {
    Write-Host "  ⚠ Skipping (no MSBuild found, using pre-built DLL)" -ForegroundColor Yellow
}

# Step 2: Build Installer Console App
Write-Host "[2/4] Building SherpaOnnxSAPIInstaller.exe..." -ForegroundColor Yellow
dotnet build "Installer\Installer.csproj" -c Release -o "Installer\bin\Release\net6.0"
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Failed to build installer console app" -ForegroundColor Red
    exit 1
}
Write-Host "  ✓ Installer console app built successfully" -ForegroundColor Green

# Step 3: Build WiX Installer
Write-Host "[3/4] Building MSI installer..." -ForegroundColor Yellow
& $env:WIX\bin\candle.exe "Installer\Product.wxs" -out "Installer\obj\Release\Product.wixlib" -ext "WixUIExtension" -ext "WixUtilExtension"
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Failed to compile WiX source" -ForegroundColor Red
    exit 1
}

& $env:WIX\bin\light.exe "Installer\obj\Release\Product.wixlib" -out "dist\SherpaOnnxSAPI.msi" -ext "WixUIExtension" -ext "WixUtilExtension" -sval
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Failed to link WiX installer" -ForegroundColor Red
    exit 1
}
Write-Host "  ✓ MSI installer built successfully" -ForegroundColor Green

# Step 4: Copy voice database
Write-Host "[4/4] Copying voice database..." -ForegroundColor Yellow
if (Test-Path "dist\merged_models.json") {
    Write-Host "  ✓ Voice database already exists" -ForegroundColor Green
} else {
    Write-Host "  ⚠ Voice database not found. Downloading..." -ForegroundColor Yellow
    Invoke-WebRequest -Uri "https://github.com/willwade/tts-wrapper/raw/main/tts_wrapper/engines/sherpaonnx/merged_models.json" -OutFile "dist\merged_models.json"
    Write-Host "  ✓ Voice database downloaded" -ForegroundColor Green
}

Write-Host ""
Write-Host "=== Build Complete ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "Output files:" -ForegroundColor Green
Write-Host "  - dist\SherpaOnnxSAPI.msi          (Installer)" -ForegroundColor White
Write-Host "  - Installer\bin\Release\net6.0\   (Console app)" -ForegroundColor White
Write-Host "  - NativeTTSWrapper\x64\Release\   (Native DLL)" -ForegroundColor White
Write-Host ""
Write-Host "To test the installer:" -ForegroundColor Yellow
Write-Host "  msiexec /i dist\SherpaOnnxSAPI.msi /l*v install.log" -ForegroundColor White
Write-Host ""
