# Complete Build Script for SherpaOnnx SAPI Installer
# This script builds everything from scratch into a single executable

param(
    [string]$Configuration = "Release",
    [switch]$SkipNative = $false,
    [switch]$SkipTests = $false,
    [switch]$Clean = $false
)

Write-Host "BUILDING COMPLETE SHERPAONNX SAPI INSTALLER" -ForegroundColor Cyan
Write-Host "===========================================" -ForegroundColor Cyan

# Check if running as administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "WARNING: Not running as Administrator. Some build steps may fail." -ForegroundColor Yellow
    Write-Host "Consider running: sudo .\BuildCompleteInstaller.ps1" -ForegroundColor Yellow
    Write-Host ""
}

$ErrorActionPreference = "Stop"
$startTime = Get-Date

# Build configuration
$buildConfig = $Configuration
$platform = "x64"
$outputDir = ".\dist"
$tempDir = ".\temp"

Write-Host "Build Configuration: $buildConfig" -ForegroundColor Green
Write-Host "Platform: $platform" -ForegroundColor Green
Write-Host "Output Directory: $outputDir" -ForegroundColor Green
Write-Host ""

# Step 1: Clean previous builds
if ($Clean) {
    Write-Host "1. Cleaning previous builds..." -ForegroundColor Cyan
    
    if (Test-Path $outputDir) {
        Remove-Item $outputDir -Recurse -Force
        Write-Host "   Cleaned output directory" -ForegroundColor Gray
    }
    
    if (Test-Path $tempDir) {
        Remove-Item $tempDir -Recurse -Force
        Write-Host "   Cleaned temp directory" -ForegroundColor Gray
    }
    
    # Clean project build outputs
    $binDirs = Get-ChildItem -Path . -Recurse -Directory -Name "bin" -ErrorAction SilentlyContinue
    foreach ($binDir in $binDirs) {
        $binPath = Join-Path (Get-Location) $binDir
        if (Test-Path $binPath) {
            Remove-Item $binPath -Recurse -Force
            Write-Host "   Cleaned $binPath" -ForegroundColor Gray
        }
    }

    $objDirs = Get-ChildItem -Path . -Recurse -Directory -Name "obj" -ErrorAction SilentlyContinue
    foreach ($objDir in $objDirs) {
        $objPath = Join-Path (Get-Location) $objDir
        if (Test-Path $objPath) {
            Remove-Item $objPath -Recurse -Force
            Write-Host "   Cleaned $objPath" -ForegroundColor Gray
        }
    }
    
    Write-Host "   Clean completed" -ForegroundColor Green
    Write-Host ""
}

# Create output directories
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

# Step 2: Find and verify build tools
Write-Host "2. Verifying build tools..." -ForegroundColor Cyan

# Check for .NET 6.0
$dotnetPath = "C:\Program Files\dotnet\dotnet.exe"
if (Test-Path $dotnetPath) {
    try {
        $dotnetVersion = & $dotnetPath --version
        Write-Host "   .NET SDK: $dotnetVersion" -ForegroundColor Green
    } catch {
        Write-Host "   ERROR: Failed to run .NET SDK." -ForegroundColor Red
        exit 1
    }
} else {
    Write-Host "   ERROR: .NET 6.0 SDK not found. Please install .NET 6.0 SDK." -ForegroundColor Red
    exit 1
}

# Check for MSBuild (for native projects)
if (-not $SkipNative) {
    $vswherePath = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vswherePath) {
        $vsInstallPath = & $vswherePath -latest -products * -requires Microsoft.Component.MSBuild -property installationPath
        if ($vsInstallPath) {
            $msbuild = Join-Path $vsInstallPath "MSBuild\Current\Bin\MSBuild.exe"
            if (Test-Path $msbuild) {
                Write-Host "   MSBuild: $msbuild" -ForegroundColor Green
            } else {
                Write-Host "   WARNING: MSBuild not found. Native wrapper build will be skipped." -ForegroundColor Yellow
                $SkipNative = $true
            }
        } else {
            Write-Host "   WARNING: Visual Studio not found. Native wrapper build will be skipped." -ForegroundColor Yellow
            $SkipNative = $true
        }
    } else {
        Write-Host "   WARNING: Visual Studio Installer not found. Native wrapper build will be skipped." -ForegroundColor Yellow
        $SkipNative = $true
    }
}

Write-Host ""

# Step 3: Restore NuGet packages
Write-Host "3. Restoring NuGet packages..." -ForegroundColor Cyan
try {
    & $dotnetPath restore TTSInstaller.sln
    Write-Host "   NuGet packages restored successfully" -ForegroundColor Green
} catch {
    Write-Host "   ERROR: Failed to restore NuGet packages" -ForegroundColor Red
    exit 1
}
Write-Host ""

# Step 4: Build managed projects
Write-Host "4. Building managed projects..." -ForegroundColor Cyan

# Build OpenSpeechTTS (managed COM objects)
Write-Host "   Building OpenSpeechTTS..." -ForegroundColor White
try {
    & $dotnetPath build "OpenSpeechTTS\OpenSpeechTTS.csproj" --configuration $buildConfig --no-restore
    Write-Host "   OpenSpeechTTS built successfully" -ForegroundColor Green
} catch {
    Write-Host "   ERROR: Failed to build OpenSpeechTTS" -ForegroundColor Red
    exit 1
}

# Build SherpaWorker
Write-Host "   Building SherpaWorker..." -ForegroundColor White
try {
    & $dotnetPath build "SherpaWorker\SherpaWorker.csproj" --configuration $buildConfig --no-restore
    Write-Host "   SherpaWorker built successfully" -ForegroundColor Green
} catch {
    Write-Host "   ERROR: Failed to build SherpaWorker" -ForegroundColor Red
    exit 1
}

# Build main installer
Write-Host "   Building TTSInstaller..." -ForegroundColor White
try {
    & $dotnetPath build "TTSInstaller.csproj" --configuration $buildConfig --no-restore
    Write-Host "   TTSInstaller built successfully" -ForegroundColor Green
} catch {
    Write-Host "   ERROR: Failed to build TTSInstaller" -ForegroundColor Red
    exit 1
}

Write-Host ""

# Step 5: Build native COM wrapper
if (-not $SkipNative) {
    Write-Host "5. Building native COM wrapper..." -ForegroundColor Cyan
    try {
        & $msbuild "NativeTTSWrapper\NativeTTSWrapper.vcxproj" /p:Configuration=$buildConfig /p:Platform=$platform /verbosity:minimal
        
        $nativeDllPath = "NativeTTSWrapper\$platform\$buildConfig\NativeTTSWrapper.dll"
        if (Test-Path $nativeDllPath) {
            $dllSize = (Get-Item $nativeDllPath).Length
            Write-Host "   Native COM wrapper built successfully ($([math]::Round($dllSize/1KB, 1)) KB)" -ForegroundColor Green
        } else {
            Write-Host "   ERROR: Native DLL not found at $nativeDllPath" -ForegroundColor Red
            exit 1
        }
    } catch {
        Write-Host "   ERROR: Failed to build native COM wrapper" -ForegroundColor Red
        Write-Host "   $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
} else {
    Write-Host "5. Skipping native COM wrapper build" -ForegroundColor Yellow
}
Write-Host ""

# Step 6: Publish installer as single executable
Write-Host "6. Publishing installer as single executable..." -ForegroundColor Cyan
try {
    & $dotnetPath publish "TTSInstaller.csproj" `
        --configuration $buildConfig `
        --runtime win-x64 `
        --self-contained true `
        --output $tempDir `
        /p:PublishSingleFile=true `
        /p:IncludeNativeLibrariesForSelfExtract=true `
        /p:DebugType=embedded `
        --no-restore

    Write-Host "   Installer published successfully" -ForegroundColor Green
} catch {
    Write-Host "   ERROR: Failed to publish installer" -ForegroundColor Red
    exit 1
}
Write-Host ""

# Step 7: Copy all components to output directory
Write-Host "7. Copying components to output directory..." -ForegroundColor Cyan

# Copy main installer executable
$installerExe = "$tempDir\TTSInstaller.exe"
if (Test-Path $installerExe) {
    Copy-Item $installerExe "$outputDir\SherpaOnnxSAPIInstaller.exe" -Force
    $exeSize = (Get-Item "$outputDir\SherpaOnnxSAPIInstaller.exe").Length
    Write-Host "   Copied installer executable ($([math]::Round($exeSize/1MB, 1)) MB)" -ForegroundColor Green
} else {
    Write-Host "   ERROR: Installer executable not found" -ForegroundColor Red
    exit 1
}

# Copy managed COM DLL
$managedDll = "OpenSpeechTTS\bin\$buildConfig\net472\OpenSpeechTTS.dll"
if (Test-Path $managedDll) {
    Copy-Item $managedDll "$outputDir\" -Force
    Write-Host "   Copied managed COM DLL" -ForegroundColor Green
} else {
    Write-Host "   WARNING: Managed COM DLL not found at $managedDll" -ForegroundColor Yellow
}

# Copy native COM wrapper
if (-not $SkipNative) {
    $nativeDll = "NativeTTSWrapper\$platform\$buildConfig\NativeTTSWrapper.dll"
    if (Test-Path $nativeDll) {
        Copy-Item $nativeDll "$outputDir\" -Force
        Write-Host "   Copied native COM wrapper" -ForegroundColor Green
    } else {
        Write-Host "   WARNING: Native COM wrapper not found at $nativeDll" -ForegroundColor Yellow
    }
}

# Copy SherpaWorker and dependencies
$sherpaWorkerExe = "SherpaWorker\bin\$buildConfig\net6.0\win-x64\SherpaWorker.exe"
if (Test-Path $sherpaWorkerExe) {
    Copy-Item $sherpaWorkerExe "$outputDir\" -Force
    Write-Host "   Copied SherpaWorker.exe" -ForegroundColor Green

    # Copy SherpaWorker dependencies
    $sherpaWorkerDir = "SherpaWorker\bin\$buildConfig\net6.0\win-x64"
    $dependencies = @(
        "sherpa-onnx.dll",
        "SherpaNative.dll",
        "onnxruntime.dll",
        "onnxruntime_providers_shared.dll",
        "SherpaWorker.deps.json",
        "SherpaWorker.runtimeconfig.json"
    )

    foreach ($dep in $dependencies) {
        $depPath = "$sherpaWorkerDir\$dep"
        if (Test-Path $depPath) {
            Copy-Item $depPath "$outputDir\" -Force
            Write-Host "   Copied $dep" -ForegroundColor Gray
        } else {
            Write-Host "   WARNING: Dependency $dep not found" -ForegroundColor Yellow
        }
    }
} else {
    Write-Host "   WARNING: SherpaWorker.exe not found at $sherpaWorkerExe" -ForegroundColor Yellow
}

Write-Host ""

# Step 8: Create deployment package
Write-Host "8. Creating deployment package..." -ForegroundColor Cyan

# Create README for the package
$readmeContent = @"
# SherpaOnnx SAPI Installer

This package contains the complete SherpaOnnx SAPI installer with Azure TTS support.

## Contents:
- SherpaOnnxSAPIInstaller.exe - Main installer (single executable)
- OpenSpeechTTS.dll - Managed COM objects for Azure TTS
- NativeTTSWrapper.dll - Native COM wrapper for SherpaOnnx (100% SAPI compatibility)
- SherpaWorker.exe - ProcessBridge worker for SherpaOnnx
- Dependencies - Required DLLs for SherpaOnnx functionality

## Usage:

### Install SherpaOnnx voice:
```
sudo .\SherpaOnnxSAPIInstaller.exe install amy
```

### Install Azure TTS voice:
```
sudo .\SherpaOnnxSAPIInstaller.exe install-azure en-US-JennyNeural --key YOUR_KEY --region eastus
```

### Interactive mode:
```
sudo .\SherpaOnnxSAPIInstaller.exe
```

### Uninstall all voices:
```
sudo .\SherpaOnnxSAPIInstaller.exe uninstall all
```

## Requirements:
- Windows 10/11
- Administrator privileges
- .NET Framework 4.7.2+ (for Azure TTS)
- .NET 6.0 Runtime (included in installer)

Built on: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
"@

$readmeContent | Out-File "$outputDir\README.md" -Encoding UTF8
Write-Host "   Created README.md" -ForegroundColor Green

# Create version info
$versionInfo = @{
    BuildDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Configuration = $buildConfig
    Platform = $platform
    Components = @{
        Installer = Test-Path "$outputDir\SherpaOnnxSAPIInstaller.exe"
        ManagedCOM = Test-Path "$outputDir\OpenSpeechTTS.dll"
        NativeCOM = Test-Path "$outputDir\NativeTTSWrapper.dll"
        SherpaWorker = Test-Path "$outputDir\SherpaWorker.exe"
    }
} | ConvertTo-Json -Depth 3

$versionInfo | Out-File "$outputDir\build-info.json" -Encoding UTF8
Write-Host "   Created build-info.json" -ForegroundColor Green

Write-Host ""

# Step 9: Summary
$endTime = Get-Date
$buildDuration = $endTime - $startTime

Write-Host "BUILD COMPLETED SUCCESSFULLY!" -ForegroundColor Cyan
Write-Host "=============================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Build Duration: $($buildDuration.TotalMinutes.ToString("F1")) minutes" -ForegroundColor Green
Write-Host "Output Directory: $outputDir" -ForegroundColor Green
Write-Host ""

Write-Host "BUILT COMPONENTS:" -ForegroundColor Yellow
Get-ChildItem $outputDir | ForEach-Object {
    $size = if ($_.PSIsContainer) { "DIR" } else { "$([math]::Round($_.Length/1KB, 1)) KB" }
    Write-Host "  - $($_.Name) ($size)" -ForegroundColor White
}

Write-Host ""
Write-Host "READY FOR DEPLOYMENT!" -ForegroundColor Cyan
Write-Host "Copy the contents of '$outputDir' to deploy the complete installer." -ForegroundColor Yellow

# Cleanup temp directory
Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
