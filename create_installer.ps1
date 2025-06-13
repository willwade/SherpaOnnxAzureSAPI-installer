# Create Complete Installer for C++ SAPI Bridge to AACSpeakHelper
# This script creates a comprehensive installer package

param(
    [string]$OutputDir = "installer-package",
    [string]$Version = "1.0.0",
    [switch]$SkipBuild,
    [switch]$CreateMSI
)

Write-Host "🎯 Creating Complete Installer Package" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "Output Directory: $OutputDir" -ForegroundColor White
Write-Host "Version: $Version" -ForegroundColor White
Write-Host ""

$ErrorActionPreference = "Stop"

# Function to check if a command exists
function Test-Command {
    param([string]$Command)
    try {
        Get-Command $Command -ErrorAction Stop | Out-Null
        return $true
    } catch {
        return $false
    }
}

# Step 1: Prerequisites Check
Write-Host "📋 Step 1: Prerequisites Check" -ForegroundColor Yellow
Write-Host "==============================" -ForegroundColor Yellow

$prerequisites = @{
    "uv" = "Python package manager"
    "dotnet" = ".NET SDK"
    "msbuild" = "MSBuild (Visual Studio)"
}

$allPrereqsMet = $true
foreach ($prereq in $prerequisites.Keys) {
    if (Test-Command $prereq) {
        Write-Host "✅ $prereq`: Available" -ForegroundColor Green
    } else {
        Write-Host "❌ $prereq`: Not found - $($prerequisites[$prereq])" -ForegroundColor Red
        $allPrereqsMet = $false
    }
}

if (-not $allPrereqsMet) {
    Write-Host "❌ Prerequisites not met. Please install missing components." -ForegroundColor Red
    exit 1
}

Write-Host ""

# Step 2: Create Output Directory
Write-Host "📋 Step 2: Creating Output Directory" -ForegroundColor Yellow
Write-Host "====================================" -ForegroundColor Yellow

if (Test-Path $OutputDir) {
    Write-Host "Removing existing output directory..." -ForegroundColor Cyan
    Remove-Item $OutputDir -Recurse -Force
}

New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
New-Item -ItemType Directory -Path "$OutputDir\bin" -Force | Out-Null
New-Item -ItemType Directory -Path "$OutputDir\voice_configs" -Force | Out-Null
New-Item -ItemType Directory -Path "$OutputDir\docs" -Force | Out-Null
New-Item -ItemType Directory -Path "$OutputDir\tests" -Force | Out-Null

Write-Host "✅ Output directory structure created" -ForegroundColor Green
Write-Host ""

# Step 3: Build Components (if not skipped)
if (-not $SkipBuild) {
    Write-Host "📋 Step 3: Building Components" -ForegroundColor Yellow
    Write-Host "==============================" -ForegroundColor Yellow

    # Build C++ COM Wrapper
    Write-Host "Building C++ COM wrapper..." -ForegroundColor Cyan
    try {
        msbuild "NativeTTSWrapper\NativeTTSWrapper.vcxproj" /p:Configuration=Release /p:Platform=x64
        
        if (Test-Path "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll") {
            Copy-Item "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll" "$OutputDir\bin\"
            Copy-Item "NativeTTSWrapper\x64\Release\NativeTTSWrapper.pdb" "$OutputDir\bin\" -ErrorAction SilentlyContinue
            Write-Host "✅ C++ COM wrapper built and copied" -ForegroundColor Green
        } else {
            throw "C++ COM wrapper DLL not found"
        }
    } catch {
        Write-Host "❌ C++ build failed: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }

    # Build .NET Installer
    Write-Host "Building .NET installer..." -ForegroundColor Cyan
    try {
        Set-Location "Installer"
        dotnet publish -c Release -o "..\$OutputDir\bin\installer" --self-contained true -r win-x64
        Set-Location ".."
        
        if (Test-Path "$OutputDir\bin\installer\SherpaOnnxSAPIInstaller.exe") {
            Write-Host "✅ .NET installer built and copied" -ForegroundColor Green
        } else {
            throw ".NET installer executable not found"
        }
    } catch {
        Write-Host "❌ .NET build failed: $($_.Exception.Message)" -ForegroundColor Red
        Set-Location ".."
        exit 1
    }

    # Build Python CLI with PyInstaller
    Write-Host "Building Python CLI..." -ForegroundColor Cyan
    try {
        uv venv
        uv sync --extra build
        uv run pyinstaller --onefile --name "SapiVoiceManager" --distpath "$OutputDir\bin" SapiVoiceManager.py
        
        if (Test-Path "$OutputDir\bin\SapiVoiceManager.exe") {
            Write-Host "✅ Python CLI built and copied" -ForegroundColor Green
        } else {
            throw "Python CLI executable not found"
        }
    } catch {
        Write-Host "❌ Python CLI build failed: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }

    Write-Host ""
}

# Step 4: Copy Voice Configurations
Write-Host "📋 Step 4: Copying Voice Configurations" -ForegroundColor Yellow
Write-Host "=======================================" -ForegroundColor Yellow

if (Test-Path "voice_configs") {
    Copy-Item "voice_configs\*" "$OutputDir\voice_configs\" -Recurse
    $voiceCount = (Get-ChildItem "$OutputDir\voice_configs" -Filter "*.json").Count
    Write-Host "✅ Copied $voiceCount voice configurations" -ForegroundColor Green
} else {
    Write-Host "⚠️ Voice configurations directory not found" -ForegroundColor Yellow
}

Write-Host ""

# Step 5: Copy Documentation and Tests
Write-Host "📋 Step 5: Copying Documentation and Tests" -ForegroundColor Yellow
Write-Host "===========================================" -ForegroundColor Yellow

# Copy documentation
$docFiles = @("README.md", "INTEGRATION_TESTING.md", "TODO.md")
foreach ($docFile in $docFiles) {
    if (Test-Path $docFile) {
        Copy-Item $docFile "$OutputDir\docs\"
        Write-Host "✅ Copied $docFile" -ForegroundColor Green
    }
}

# Copy test scripts
$testFiles = @("test_complete_workflow.ps1", "test_windows_integration.ps1")
foreach ($testFile in $testFiles) {
    if (Test-Path $testFile) {
        Copy-Item $testFile "$OutputDir\tests\"
        Write-Host "✅ Copied $testFile" -ForegroundColor Green
    }
}

Write-Host ""

# Step 6: Create Installation Scripts
Write-Host "📋 Step 6: Creating Installation Scripts" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow

# Create install script
$installScript = @'
@echo off
echo C++ SAPI Bridge to AACSpeakHelper - Installation
echo ===============================================
echo.

echo Checking administrator privileges...
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo ERROR: This installer requires administrator privileges.
    echo Please run as administrator.
    pause
    exit /b 1
)

echo ✅ Administrator privileges confirmed
echo.

echo Registering C++ COM wrapper...
regsvr32 /s "%~dp0bin\NativeTTSWrapper.dll"
if %errorLevel% equ 0 (
    echo ✅ C++ COM wrapper registered successfully
) else (
    echo ❌ Failed to register C++ COM wrapper
    pause
    exit /b 1
)

echo.
echo Installation completed successfully!
echo.
echo Next steps:
echo 1. Set up AACSpeakHelper service
echo 2. Install voices using: bin\SapiVoiceManager.exe --install English-SherpaOnnx-Jenny
echo 3. Test integration using: tests\test_windows_integration.ps1
echo.
echo See docs\INTEGRATION_TESTING.md for complete instructions.
pause
'@

$installScript | Out-File "$OutputDir\install.bat" -Encoding ASCII

# Create uninstall script
$uninstallScript = @'
@echo off
echo C++ SAPI Bridge to AACSpeakHelper - Uninstallation
echo =================================================
echo.

echo Checking administrator privileges...
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo ERROR: This installer requires administrator privileges.
    echo Please run as administrator.
    pause
    exit /b 1
)

echo ✅ Administrator privileges confirmed
echo.

echo Removing installed voices...
bin\SapiVoiceManager.exe --remove-all

echo Unregistering C++ COM wrapper...
regsvr32 /s /u "%~dp0bin\NativeTTSWrapper.dll"
if %errorLevel% equ 0 (
    echo ✅ C++ COM wrapper unregistered successfully
) else (
    echo ⚠️ Failed to unregister C++ COM wrapper (may not have been registered)
)

echo.
echo Uninstallation completed!
pause
'@

$uninstallScript | Out-File "$OutputDir\uninstall.bat" -Encoding ASCII

Write-Host "✅ Installation scripts created" -ForegroundColor Green
Write-Host ""

# Step 7: Create Package Information
Write-Host "📋 Step 7: Creating Package Information" -ForegroundColor Yellow
Write-Host "=======================================" -ForegroundColor Yellow

# Create package README
$packageReadme = @"
# C++ SAPI Bridge to AACSpeakHelper - Installation Package

## 🎯 Overview

This package contains the complete C++ SAPI Bridge to AACSpeakHelper implementation, ready for installation and testing.

## 📦 Package Contents

### Core Components:
- **bin/NativeTTSWrapper.dll** - C++ SAPI COM wrapper with AACSpeakHelper integration
- **bin/installer/SherpaOnnxSAPIInstaller.exe** - .NET installer for voice registration
- **bin/SapiVoiceManager.exe** - Python CLI tool for voice management

### Voice Configurations:
- **voice_configs/** - Pre-configured voice definitions for multiple TTS engines

### Documentation:
- **docs/INTEGRATION_TESTING.md** - Complete testing and integration guide
- **docs/README.md** - Main project documentation
- **docs/TODO.md** - Project status and roadmap

### Testing Framework:
- **tests/test_windows_integration.ps1** - Complete integration testing
- **tests/test_complete_workflow.ps1** - Build and test workflow

## 🚀 Quick Installation

### 1. Install Components (as Administrator):
```
install.bat
```

### 2. Set up AACSpeakHelper Service:
```
git clone https://github.com/AceCentre/AACSpeakHelper
cd AACSpeakHelper
uv venv && uv sync --all-extras
uv run python AACSpeakHelperServer.py
```

### 3. Install a Voice:
```
bin\SapiVoiceManager.exe --install English-SherpaOnnx-Jenny
```

### 4. Test Integration:
```
tests\test_windows_integration.ps1
```

## 🧪 Testing

Run the complete integration test:
```
tests\test_windows_integration.ps1
```

## 📋 Requirements

- Windows 10/11
- Administrator privileges for COM registration
- AACSpeakHelper service running
- .NET 6.0 Runtime (included in installer)

## 🔧 Uninstallation

To remove all components:
```
uninstall.bat
```

## 🔗 Repository

https://github.com/willwade/SherpaOnnxAzureSAPI-installer

## 📊 Package Information

- **Version**: $Version
- **Build Date**: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
- **Architecture**: Windows x64
- **Components**: C++ COM wrapper, .NET installer, Python CLI, Voice configs, Tests

---

**Ready for production testing!** 🎉
"@

$packageReadme | Out-File "$OutputDir\README.md" -Encoding UTF8

# Create package info JSON
$packageInfo = @{
    "name" = "C++ SAPI Bridge to AACSpeakHelper"
    "version" = $Version
    "build_date" = (Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC")
    "platform" = "Windows x64"
    "components" = @{
        "cpp_com_wrapper" = "bin/NativeTTSWrapper.dll"
        "dotnet_installer" = "bin/installer/SherpaOnnxSAPIInstaller.exe"
        "python_cli" = "bin/SapiVoiceManager.exe"
        "voice_configs" = "voice_configs/*.json"
        "installation_scripts" = @("install.bat", "uninstall.bat")
        "test_framework" = "tests/*.ps1"
        "documentation" = "docs/*.md"
    }
    "installation" = @{
        "requires_admin" = $true
        "install_script" = "install.bat"
        "uninstall_script" = "uninstall.bat"
        "test_script" = "tests/test_windows_integration.ps1"
    }
    "integration" = @{
        "architecture" = "C++ SAPI Bridge to AACSpeakHelper"
        "pipe_service" = "AACSpeakHelper"
        "supported_engines" = @("SherpaOnnx", "Google TTS", "Azure TTS", "ElevenLabs")
        "ready_for_testing" = $true
    }
} | ConvertTo-Json -Depth 4

$packageInfo | Out-File "$OutputDir\package-info.json" -Encoding UTF8

Write-Host "✅ Package information created" -ForegroundColor Green
Write-Host ""

# Step 8: Create ZIP Archive
Write-Host "📋 Step 8: Creating ZIP Archive" -ForegroundColor Yellow
Write-Host "===============================" -ForegroundColor Yellow

$zipName = "C++ SAPI Bridge to AACSpeakHelper v$Version.zip"
try {
    Compress-Archive -Path "$OutputDir\*" -DestinationPath $zipName -Force
    $zipSize = [math]::Round((Get-Item $zipName).Length / 1MB, 2)
    Write-Host "✅ ZIP archive created: $zipName ($zipSize MB)" -ForegroundColor Green
} catch {
    Write-Host "❌ Failed to create ZIP archive: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""

# Step 9: Summary
Write-Host "📋 Step 9: Installation Package Summary" -ForegroundColor Yellow
Write-Host "=======================================" -ForegroundColor Yellow

Write-Host "🎯 Package Created Successfully!" -ForegroundColor Green
Write-Host "  📁 Directory: $OutputDir" -ForegroundColor White
Write-Host "  📦 Archive: $zipName" -ForegroundColor White
Write-Host "  🏷️ Version: $Version" -ForegroundColor White
Write-Host ""

Write-Host "📋 Package Contents:" -ForegroundColor Cyan
if (Test-Path "$OutputDir\bin\NativeTTSWrapper.dll") {
    Write-Host "  ✅ C++ COM Wrapper" -ForegroundColor Green
} else {
    Write-Host "  ❌ C++ COM Wrapper" -ForegroundColor Red
}

if (Test-Path "$OutputDir\bin\installer\SherpaOnnxSAPIInstaller.exe") {
    Write-Host "  ✅ .NET Installer" -ForegroundColor Green
} else {
    Write-Host "  ❌ .NET Installer" -ForegroundColor Red
}

if (Test-Path "$OutputDir\bin\SapiVoiceManager.exe") {
    Write-Host "  ✅ Python CLI" -ForegroundColor Green
} else {
    Write-Host "  ❌ Python CLI" -ForegroundColor Red
}

$voiceConfigCount = (Get-ChildItem "$OutputDir\voice_configs" -Filter "*.json" -ErrorAction SilentlyContinue).Count
Write-Host "  ✅ Voice Configurations: $voiceConfigCount" -ForegroundColor Green

Write-Host ""
Write-Host "🚀 Next Steps:" -ForegroundColor Cyan
Write-Host "  1. Extract/distribute the package" -ForegroundColor White
Write-Host "  2. Run install.bat as Administrator" -ForegroundColor White
Write-Host "  3. Set up AACSpeakHelper service" -ForegroundColor White
Write-Host "  4. Run tests\test_windows_integration.ps1" -ForegroundColor White

Write-Host ""
Write-Host "🎉 Installation package creation completed!" -ForegroundColor Green
