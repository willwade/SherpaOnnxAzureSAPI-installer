#!/usr/bin/env pwsh
<#
.SYNOPSIS
    SherpaOnnx SAPI5 TTS Engine - Build Script
.DESCRIPTION
    Complete build script that replicates the GitHub Actions CI/CD pipeline.
    Builds Native DLL, ConfigApp, Installer, and MSI.
.EXAMPLE
    .\build.ps1
#>

$ErrorActionPreference = "Stop"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "SherpaOnnx SAPI5 Build Script" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptDir

$SuccessCount = 0
$ErrorCount = 0

function Test-Step {
    param(
        [string]$Name,
        [scriptblock]$Script
    )

    Write-Host "`n[$($Name)]" -ForegroundColor Yellow
    try {
        & $Script
        Write-Host "  ✓ SUCCESS" -ForegroundColor Green
        $script:SuccessCount++
        return $true
    }
    catch {
        Write-Host "  ✗ FAILED: $_" -ForegroundColor Red
        $script:ErrorCount++
        return $false
    }
}

# Step 1: Download SherpaOnnx dependencies
Test-Step "Step 1: Download SherpaOnnx Dependencies" {
    Write-Host "  Running: scripts\Download-SherpaOnnx.ps1" -ForegroundColor Gray
    & pwsh -File "scripts\Download-SherpaOnnx.ps1"
    if ($LASTEXITCODE -ne 0) { throw "Download failed with exit code $LASTEXITCODE" }
}

# Step 2: Build NativeTTSWrapper.dll
Test-Step "Step 2: Build NativeTTSWrapper.dll" {
    Write-Host "  Locating MSBuild..." -ForegroundColor Gray

    # Try to find MSBuild in standard locations
    $msbuild = $null
    $vsWhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"

    if (Test-Path $vsWhere) {
        # Find latest VS installation with MSBuild
        $vsPath = & $vsWhere -latest -products * -requires Microsoft.Component.MSBuild -property installationPath
        if ($vsPath) {
            $msbuildPath = Join-Path $vsPath "MSBuild\Current\Bin\MSBuild.exe"
            if (Test-Path $msbuildPath) {
                $msbuild = $msbuildPath
                Write-Host "  Found: $msbuildPath" -ForegroundColor Cyan
            }
        }
    }

    # Fallback to PATH
    if (-not $msbuild) {
        $msbuild = Get-Command msbuild -ErrorAction SilentlyContinue
        if ($msbuild) {
            Write-Host "  Found in PATH: $($msbuild.Source)" -ForegroundColor Cyan
        }
    }

    if (-not $msbuild) {
        throw "MSBuild not found. Please install Visual Studio 2022 with C++ Desktop Development workload."
    }

    Write-Host "  Building: NativeTTSWrapper\NativeTTSWrapper.sln" -ForegroundColor Gray

    & $msbuild "NativeTTSWrapper\NativeTTSWrapper.sln" /t:Build /p:Configuration=Release /p:Platform=x64 /v:minimal /nologo
    if ($LASTEXITCODE -ne 0) { throw "Build failed with exit code $LASTEXITCODE" }

    # Verify DLL was created
    $dllPath = "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll"
    if (-not (Test-Path $dllPath)) {
        throw "DLL not found at $dllPath"
    }

    $dllSize = (Get-Item $dllPath).Length / 1MB
    Write-Host "  ✓ DLL Size: $([math]::Round($dllSize, 2)) MB" -ForegroundColor Green
}

# Step 3: Build ConfigApp
Test-Step "Step 3: Build ConfigApp" {
    Write-Host "  Running: dotnet build ConfigApp\ConfigApp.csproj" -ForegroundColor Gray

    # Check if dotnet is available
    $dotnet = Get-Command dotnet -ErrorAction SilentlyContinue
    if (-not $dotnet) {
        throw "dotnet CLI not found. Please install .NET 8 SDK from https://dotnet.microsoft.com/download"
    }

    & dotnet build "ConfigApp\ConfigApp.csproj" -c Release --no-dependencies /nologo
    if ($LASTEXITCODE -ne 0) { throw "Build failed with exit code $LASTEXITCODE" }

    # Verify executable was created
    $exePath = "ConfigApp\bin\Release\net8.0-windows\SherpaOnnxConfig.exe"
    if (-not (Test-Path $exePath)) {
        throw "Executable not found at $exePath"
    }
    Write-Host "  Created: SherpaOnnxConfig.exe" -ForegroundColor Cyan
}

# Step 4: Build Console Installer
Test-Step "Step 4: Build Console Installer" {
    Write-Host "  Running: dotnet build Installer\Installer.csproj" -ForegroundColor Gray

    & dotnet build "Installer\Installer.csproj" -c Release --no-dependencies /nologo
    if ($LASTEXITCODE -ne 0) { throw "Build failed with exit code $LASTEXITCODE" }

    # Verify executable was created
    $exePath = "Installer\bin\Release\net8.0\SherpaOnnxSAPIInstaller.exe"
    if (-not (Test-Path $exePath)) {
        throw "Executable not found at $exePath"
    }
    Write-Host "  Created: SherpaOnnxSAPIInstaller.exe" -ForegroundColor Cyan
}

# Step 5: Build MSI
Test-Step "Step 5: Build MSI Installer" {
    Write-Host "  Checking for WiX Toolset..." -ForegroundColor Gray

    $wixCandle = Get-Command candle.exe -ErrorAction SilentlyContinue
    $wixLight = Get-Command light.exe -ErrorAction SilentlyContinue

    if (-not $wixCandle -or -not $wixLight) {
        Write-Host "  WiX Toolset not found. Downloading..." -ForegroundColor Yellow

        $wixDir = Join-Path $ScriptDir "wix-tools"
        $wixZip = Join-Path $ScriptDir "wix314-binaries.zip"

        # Download WiX if not already present
        if (-not (Test-Path $wixDir)) {
            Write-Host "  Downloading WiX Toolset v3.14..." -ForegroundColor Gray
            $url = "https://github.com/wixtoolset/wix3/releases/download/wix3141rtm/wix314-binaries.zip"
            Invoke-WebRequest -Uri $url -OutFile $wixZip -UseBasicParsing

            Write-Host "  Extracting..." -ForegroundColor Gray
            Expand-Archive -Path $wixZip -DestinationPath $wixDir -Force
            Remove-Item $wixZip
        }

        $wixCandle = Join-Path $wixDir "candle.exe"
        $wixLight = Join-Path $wixDir "light.exe"

        if (-not (Test-Path $wixCandle)) {
            throw "WiX candle.exe not found at $wixCandle"
        }
    }

    Write-Host "  Running WiX candle.exe..." -ForegroundColor Gray
    & $wixCandle "Installer\Product.wxs" -out "Installer\Product.wixobj" -ext WixUIExtension.dll -ext WixUtilExtension.dll
    if ($LASTEXITCODE -ne 0) { throw "WiX candle failed with exit code $LASTEXITCODE" }

    Write-Host "  Running WiX light.exe..." -ForegroundColor Gray
    & $wixLight "Installer\Product.wixobj" -out "SherpaOnnxSAPI.msi" -ext WixUIExtension.dll -ext WixUtilExtension.dll -sval
    if ($LASTEXITCODE -ne 0) { throw "WiX light failed with exit code $LASTEXITCODE" }

    # Verify MSI was created
    $msiPath = "SherpaOnnxSAPI.msi"
    if (-not (Test-Path $msiPath)) {
        throw "MSI not found at $msiPath"
    }

    $msiSize = (Get-Item $msiPath).Length / 1MB
    Write-Host "  MSI Size: $([math]::Round($msiSize, 2)) MB" -ForegroundColor Cyan

    # Warn if MSI is suspiciously small
    if ($msiSize -lt 10) {
        Write-Host "  ⚠ WARNING: MSI is smaller than expected (< 10 MB)" -ForegroundColor Yellow
        Write-Host "  This may indicate NativeTTSWrapper.dll is not included." -ForegroundColor Yellow
    }
}

# Summary
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Build Summary" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Steps Passed: $SuccessCount" -ForegroundColor Green
if ($ErrorCount -gt 0) {
    Write-Host "  Steps Failed: $ErrorCount" -ForegroundColor Red
    Write-Host ""
    Write-Host "Build FAILED with $ErrorCount error(s)" -ForegroundColor Red
    exit 1
}
else {
    Write-Host "  Steps Failed: 0" -ForegroundColor Green
    Write-Host ""
    Write-Host "✓ Build completed successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Output files:" -ForegroundColor Cyan
    Write-Host "  - NativeTTSWrapper\NativeTTSWrapper.dll" -ForegroundColor White
    Write-Host "  - ConfigApp\bin\Release\net8.0-windows\SherpaOnnxConfig.exe" -ForegroundColor White
    Write-Host "  - Installer\bin\Release\net8.0\SherpaOnnxSAPIInstaller.exe" -ForegroundColor White
    Write-Host "  - SherpaOnnxSAPI.msi" -ForegroundColor White
    Write-Host ""
}
