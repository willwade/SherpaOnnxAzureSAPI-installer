# SherpaOnnx SAPI5 TTS Engine - Build Instructions

Complete guide to build the SherpaOnnx SAPI5 TTS Engine.

## Quick Start

```powershell
# 1. Clone the repository
git clone https://github.com/willwade/SherpaOnnxAzureSAPI-installer.git
cd SherpaOnnxAzureSAPI-installer

# 2. Download dependencies (runs automatically in CI, run locally for dev)
pwsh -File scripts\Download-SherpaOnnx.ps1

# 3. Build the native DLL
msbuild NativeTTSWrapper\NativeTTSWrapper.sln /p:Configuration=Release /p:Platform=x64

# 4. Register the DLL (requires Administrator)
regsvr32 "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll"
```

## Prerequisites

### Required Software

1. **Visual Studio 2019/2022** (Community Edition works)
   - C++ Desktop Development workload
   - Windows SDK (10.0 or later)
   - Platform Toolset: v143 or later

2. **Git for Windows**
   - Required for `bash` tar extraction (included with Git)
   - Download: https://git-scm.com/download/win

3. **.NET 8 SDK** (for ConfigApp and Installer)
   - Download: https://dotnet.microsoft.com/download/dotnet/8.0

### Optional Software

- **7-Zip** - Faster extraction (script will auto-detect)
  - Download: https://www.7-zip.org/

## Automated Dependency Download

The `scripts\Download-SherpaOnnx.ps1` script handles everything:

| Feature | Description |
|---------|-------------|
| **Download** | Fetches SherpaOnnx v1.12.10 from GitHub release |
| **Extract** | Auto-detects 7-Zip, bash+tar, or PowerShell tar |
| **Verify** | Checks all required libraries are present |
| **Cross-platform** | Works in GitHub Actions CI and local development |

**Run the script:**
```powershell
pwsh -File scripts\Download-SherpaOnnx.ps1
```

**What it downloads:**
- SherpaOnnx v1.12.10 Windows x64 static library (~170 MB)
- Includes all required dependencies: cppinyin_core.lib, onnxruntime.lib, etc.

**If the script fails:**
```
ERROR: Extraction failed. No sherpa-onnx directory found.

Please install one of the following:
  1. 7-Zip: https://www.7-zip.org/
  2. Git for Windows (includes bash): https://git-scm.com/download/win
```

## Manual Build Steps

### 1. Download SherpaOnnx Library

**Option A: Use the script (recommended)**
```powershell
pwsh -File scripts\Download-SherpaOnnx.ps1
```

**Option B: Manual download**
```powershell
# Download from our GitHub release
Invoke-WebRequest -Uri "https://github.com/willwade/SherpaOnnxAzureSAPI-installer/releases/download/v1.0.0-deps/sherpa-onnx-win-x64-static.tar.bz2" -OutFile "NativeTTSWrapper\libs-win\archive.tar.bz2"

# Extract (requires 7-Zip or Git Bash)
cd NativeTTSWrapper\libs-win
tar -xf archive.tar.bz2
```

Expected structure:
```
NativeTTSWrapper/libs-win/sherpa-onnx-v1.12.10-win-x64-static/
├── include/
│   └── sherpa-onnx/
│       └── c-api/
│           └── c-api.h
└── lib/
    ├── cppinyin_core.lib
    ├── sherpa-onnx-c-api.lib
    ├── sherpa-onnx-core.lib
    ├── onnxruntime.lib
    └── ... (other dependencies)
```

### 2. Build Native DLL

**Using Visual Studio:**
1. Open `NativeTTSWrapper\NativeTTSWrapper.sln`
2. Select **Release | x64** configuration
3. Build → Build Solution (F7)

**Using MSBuild command line:**
```bash
msbuild NativeTTSWrapper\NativeTTSWrapper.sln /t:Build /p:Configuration=Release /p:Platform=x64
```

Output:
- `NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll` (~15 MB)

### 3. Register DLL (Administrator required)

```powershell
# Navigate to build output
cd NativeTTSWrapper\x64\Release

# Register the DLL
regsvr32 NativeTTSWrapper.dll

# Or with elevation
Start-Process regsvr32 -ArgumentList 'NativeTTSWrapper.dll' -Verb RunAs -Wait
```

### 4. Verify Installation

```powershell
# Check if voice is registered
Add-Type -AssemblyName System.Speech
$synthesizer = New-Object System.Speech.Synthesis.SpeechSynthesizer
$synthesizer.GetInstalledVoices() | Where-Object { $_.VoiceInfo.Name -like "*Sherpa*" }
```

Expected output:
```
Name     : SherpaOnnx TTS Voice
Culture  : en-US
Gender   : Female
Age      : Adult
```

### 5. Test the Voice

```powershell
# Test with PowerShell
$synthesizer = New-Object System.Speech.Synthesis.SpeechSynthesizer
$synthesizer.SelectVoice("SherpaOnnx TTS Voice")
$synthesizer.Speak("Hello, this is a test of the SherpaOnnx text to speech engine.")
```

Or use the test script:
```powershell
powershell -ExecutionPolicy Bypass -File test_sapi5_extended.ps1
```

## Building the Installer

### Build ConfigApp (GUI)
```powershell
dotnet build ConfigApp\ConfigApp.csproj -c Release
```

### Build Console Installer
```powershell
dotnet build Installer\Installer.csproj -c Release
```

### Build MSI Installer (requires WiX)
```powershell
# Download WiX Toolset v3.14
Invoke-WebRequest -Uri "https://github.com/wixtoolset/wix3/releases/download/wix3141rtm/wix314-binaries.zip" -OutFile wix314-binaries.zip
Expand-Archive -Path wix314-binaries.zip -DestinationPath wix-tools -Force

# Build MSI
.\wix-tools\candle.exe Installer\Product.wxs -out Installer/Product.wixobj
.\wix-tools\light.exe Installer/Product.wixobj -out SherpaOnnxSAPI.msi
```

## Troubleshooting

### Build Errors

**"Cannot open sherpa-onnx/c-api/c-api.h"**
- Run: `pwsh -File scripts\Download-SherpaOnnx.ps1`
- Verify: `NativeTTSWrapper\libs-win\sherpa-onnx-v1.12.10-win-x64-static\include\sherpa-onnx\c-api\c-api.h` exists

**"error MSB8020: The build tools for v145 cannot be found"**
- The project now uses v143 (Visual Studio 2022)
- Update Visual Studio or use the v143 toolset

**"fatal error LNK1181: cannot open input file 'cppinyin_core.lib'"**
- Run: `pwsh -File scripts\Download-SherpaOnnx.ps1`
- Verify the script completed successfully

### Registration Errors

**"Class not registered" (0x80040154)**
- Run regsvr32 with Administrator privileges
- Use full path to DLL

**"DllRegisterServer entry point not found"**
- Ensure Release build (not Debug)
- Verify DLL exports: `dumpbin /EXPORTS NativeTTSWrapper.dll`

### Runtime Errors

**"Model file not found"**
- Download and configure a voice model using the ConfigApp
- Or manually update `engines_config.json`

**"Failed to load configuration"**
- Ensure `engines_config.json` is in the same directory as the DLL

## Project Structure

```
SherpaOnnxAzureSAPI-installer/
├── scripts/
│   └── Download-SherpaOnnx.ps1    # Dependency downloader
├── NativeTTSWrapper/               # Native C++ DLL project
│   ├── *.cpp, *.h                  # Source files
│   ├── libs-win/                   # SherpaOnnx libraries (downloaded by script)
│   ├── deps/                       # Header-only dependencies
│   │   ├── nlohmann/json.hpp       # JSON library
│   │   └── spdlog/                 # Logging library
│   └── x64/Release/                # Built DLL
├── ConfigApp/                      # WinForms GUI for voice installation
├── Installer/                      # Console installer
├── Installer/Product.wxs           # WiX installer configuration
└── test_sapi5_extended.ps1        # Test script
```

## CI/CD

The GitHub Actions workflow (`.github/workflows/build-and-release.yml`) automatically:

1. Downloads SherpaOnnx dependencies using `scripts\Download-SherpaOnnx.ps1`
2. Builds NativeTTSWrapper.dll
3. Builds ConfigApp and Console Installer
4. Builds MSI installer
5. Runs SAPI5 integration tests
6. Creates GitHub releases with artifacts

## See Also

- [SETUP.md](SETUP.md) - Voice model configuration
- [ARCHITECTURE.md](ARCHITECTURE.md) - System design
- [TROUBLESHOOTING.md](TROUBLESHOOTING.md) - Common issues
