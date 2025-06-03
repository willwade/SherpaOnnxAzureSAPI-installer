# Complete Build Instructions for SherpaOnnx SAPI Installer

This document provides step-by-step instructions to build the complete SherpaOnnx SAPI installer from scratch.

## Prerequisites

### 1. Install .NET 6.0 SDK
Download and install from: https://dotnet.microsoft.com/download/dotnet/6.0
- Choose "SDK x64" for Windows
- This is required for building the main installer and SherpaWorker

### 2. Install .NET Framework 4.7.2 Developer Pack
Download from: https://dotnet.microsoft.com/download/dotnet-framework/net472
- This is required for building the managed COM objects (OpenSpeechTTS)

### 3. Install Visual Studio Build Tools 2022
Download from: https://visualstudio.microsoft.com/downloads/#build-tools-for-visual-studio-2022
- Install with "C++ build tools" workload
- Include "Windows 10/11 SDK" and "ATL for v143 build tools"
- This is required for building the native COM wrapper

### 4. Verify Prerequisites
Run these commands to verify installation:
```powershell
dotnet --version          # Should show 6.0.x or later
dotnet --list-sdks        # Should include 6.0.x
where msbuild            # Should find MSBuild.exe
```

## Build Process

### Option 1: Automated Build Script (Recommended)
```powershell
# Run the complete build script
sudo .\BuildCompleteInstaller.ps1

# Or with options:
sudo .\BuildCompleteInstaller.ps1 -Clean -Configuration Release
```

### Option 2: Manual Build Steps

#### Step 1: Restore NuGet Packages
```powershell
dotnet restore TTSInstaller.sln
```

#### Step 2: Build Managed Projects
```powershell
# Build OpenSpeechTTS (managed COM objects for Azure TTS)
dotnet build "OpenSpeechTTS\OpenSpeechTTS.csproj" --configuration Release

# Build SherpaWorker (ProcessBridge worker)
dotnet build "SherpaWorker\SherpaWorker.csproj" --configuration Release

# Build main installer
dotnet build "TTSInstaller.csproj" --configuration Release
```

#### Step 3: Build Native COM Wrapper
```powershell
# Find MSBuild path
$msbuild = "${env:ProgramFiles}\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"

# Build native wrapper
& $msbuild "NativeTTSWrapper\NativeTTSWrapper.vcxproj" /p:Configuration=Release /p:Platform=x64
```

#### Step 4: Publish Single Executable
```powershell
dotnet publish "TTSInstaller.csproj" `
    --configuration Release `
    --runtime win-x64 `
    --self-contained true `
    --output ".\dist" `
    /p:PublishSingleFile=true
```

#### Step 5: Copy Components
```powershell
# Create output directory
New-Item -ItemType Directory -Path ".\dist" -Force

# Copy main installer
Copy-Item ".\dist\TTSInstaller.exe" ".\dist\SherpaOnnxSAPIInstaller.exe"

# Copy managed COM DLL
Copy-Item "OpenSpeechTTS\bin\Release\net472\OpenSpeechTTS.dll" ".\dist\"

# Copy native COM wrapper
Copy-Item "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll" ".\dist\"

# Copy SherpaWorker and dependencies
Copy-Item "SherpaWorker\bin\Release\net6.0\SherpaWorker.exe" ".\dist\"
Copy-Item "SherpaWorker\bin\Release\net6.0\*.dll" ".\dist\"
Copy-Item "SherpaWorker\bin\Release\net6.0\*.json" ".\dist\"
```

## Build Output

After successful build, you'll have in the `.\dist` directory:

### Core Components
- **SherpaOnnxSAPIInstaller.exe** - Main installer (single executable, ~50MB)
- **OpenSpeechTTS.dll** - Managed COM objects for Azure TTS
- **NativeTTSWrapper.dll** - Native COM wrapper for SherpaOnnx (108.5 KB)

### ProcessBridge Components
- **SherpaWorker.exe** - ProcessBridge worker (.NET 6.0, ~58MB)
- **sherpa-onnx.dll** - SherpaOnnx native library
- **SherpaNative.dll** - Native wrapper for SherpaOnnx
- **onnxruntime.dll** - ONNX Runtime
- **onnxruntime_providers_shared.dll** - ONNX Runtime providers
- **SherpaWorker.deps.json** - Dependency manifest
- **SherpaWorker.runtimeconfig.json** - Runtime configuration

## Usage

### Install SherpaOnnx Voice
```powershell
sudo .\SherpaOnnxSAPIInstaller.exe install amy
```

### Install Azure TTS Voice
```powershell
sudo .\SherpaOnnxSAPIInstaller.exe install-azure en-US-JennyNeural --key YOUR_KEY --region eastus
```

### Interactive Mode
```powershell
sudo .\SherpaOnnxSAPIInstaller.exe
```

### Test Installation
```powershell
# Test with PowerShell
$voice = New-Object -ComObject SAPI.SpVoice
$voice.Speak("Hello from SherpaOnnx!")
```

## Troubleshooting

### Common Issues

1. **"dotnet not found"**
   - Install .NET 6.0 SDK from Microsoft
   - Restart PowerShell after installation

2. **"MSBuild not found"**
   - Install Visual Studio Build Tools 2022
   - Include C++ build tools workload

3. **"regasm not found"**
   - Install .NET Framework 4.7.2 Developer Pack
   - Verify installation in Windows Features

4. **Build fails with "missing dependencies"**
   - Run `dotnet restore TTSInstaller.sln`
   - Check internet connection for NuGet packages

5. **Native wrapper build fails**
   - Verify Visual Studio Build Tools installation
   - Check that ATL libraries are installed
   - Try building in Visual Studio IDE first

### GitHub Actions Build

The project includes a GitHub Actions workflow that automatically builds the installer on every push. However, it may need updates to match the current project structure.

To fix the GitHub Actions build:
1. Update the workflow file to include all projects
2. Add proper MSBuild steps for the native wrapper
3. Ensure all dependencies are properly copied

## Architecture Summary

The complete installer provides:

1. **SherpaOnnx Engine**: Native COM wrapper → ProcessBridge → SherpaWorker → SherpaOnnx
2. **Azure TTS Engine**: Managed COM objects → Azure TTS API
3. **100% SAPI Compatibility**: Both engines work with standard SAPI calls
4. **Complete Voice Management**: Install, uninstall, verify voices
5. **Multiple Interfaces**: Command line and interactive modes

This creates a production-ready, multi-engine TTS system with full Windows SAPI integration.
