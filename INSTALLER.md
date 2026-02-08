# SherpaOnnx SAPI5 Installer Guide

Complete guide for building and using the SherpaOnnx SAPI5 TTS installer.

## Prerequisites

### Required Software

1. **Visual Studio 2022** (Community Edition works)
   - C++ Desktop Development workload
   - Windows SDK (10.0 or later)
   - MSVC v143 or later

2. **.NET 6 SDK**
   - Download from: https://dotnet.microsoft.com/download/dotnet/6.0

3. **WiX Toolset v3.11**
   - Download from: https://wixtoolset.org/releases/
   - Required for MSI creation

4. **SherpaOnnx Windows Libraries**
   - Downloaded via `download_sherpa.ps1` or manually from HuggingFace

## Building the Installer

### Quick Build

```powershell
# Run the automated build script
powershell -ExecutionPolicy Bypass -File build-installer.ps1
```

This will:
1. Build NativeTTSWrapper.dll
2. Build SherpaOnnxSAPIInstaller.exe
3. Create SherpaOnnxSAPI.msi
4. Copy voice database

### Manual Build Steps

#### 1. Build Native DLL

```bash
cd NativeTTSWrapper
msbuild NativeTTSWrapper.sln /t:Build /p:Configuration=Release /p:Platform=x64
```

#### 2. Build Console Installer

```bash
dotnet build Installer\Installer.csproj -c Release
```

#### 3. Build MSI

```bash
cd Installer
candle Product.wxs -ext WixUIExtension -ext WixUtilExtension
light -out ..\dist\SherpaOnnxSAPI.msi -ext WixUIExtension -ext WixUtilExtension
```

## Using the Installer

### Installation

**Option A: MSI Installer (Recommended)**
```powershell
msiexec /i dist\SherpaOnnxSAPI.msi
```

**Option B: Command Line with Logging**
```powershell
msiexec /i dist\SherpaOnnxSAPI.msi /l*v install.log
```

**Option C: Silent Installation**
```powershell
msiexec /i dist\SherpaOnnxSAPI.msi /quiet /norestart
```

### What Gets Installed

```
C:\Program Files\OpenAssistive\OpenSpeech\
├── NativeTTSWrapper.dll          # SAPI5 COM engine
├── sherpa-onnx-core.dll          # SherpaOnnx runtime
├── sherpa-onnx-c-api.dll         # SherpaOnnx C API
├── onnxruntime.dll               # ONNX Runtime
├── engines_config.json           # Engine configuration
└── SherpaOnnxSAPIInstaller.exe   # Console installer
```

### Registry Entries Created

```
HKLM\SOFTWARE\Classes\CLSID\{A1B2C3D4-E5F6-4A5B-8C9D-1E2F3A4B5C6D}
└── InprocServer32
    ├── (default) = "C:\Program Files\OpenAssistive\OpenSpeech\NativeTTSWrapper.dll"
    └── ThreadingModel = "Both"

HKLM\SOFTWARE\Microsoft\Speech\Voices\Tokens\TestSherpaVoice\
├── (default) = "Test Sherpa Voice"
├── CLSID = "{A1B2C3D4-E5F6-4A5B-8C9D-1E2F3A4B5C6D}"
└── Attributes\
    ├── Language = "409"
    ├── Gender = "Female"
    ├── Age = "Adult"
    └── Vendor = "OpenAssistive"
```

## Installing Additional Voices

After installation, use the console installer to add more voices:

```powershell
cd "C:\Program Files\OpenAssistive\OpenSpeech"
.\SherpaOnnxSAPIInstaller.exe install piper-en-amy-low
```

Or interactively:
```powershell
.\SherpaOnnxSAPIInstaller.exe
```

## Uninstallation

### Via Programs and Features
1. Open Settings → Apps → Installed Apps
2. Find "SherpaOnnx SAPI5 TTS Engine"
3. Click Uninstall

### Via Command Line
```powershell
msiexec /x {ProductCode}
```

### What Gets Removed
- All installed files
- Registry entries
- SAPI5 voice registrations
- **NOT** downloaded voice models (preserved in `C:\Program Files\OpenAssistive\OpenSpeech\models\`)

## Testing the Installation

### Quick Test

```powershell
powershell -ExecutionPolicy Bypass -File test_sapi5_extended.ps1
```

### Manual Test

```powershell
Add-Type -AssemblyName System.Speech
$voice = New-Object System.Speech.Synthesis.SpeechSynthesizer
$voice.GetInstalledVoices() | Where-Object { $_.VoiceInfo.Description -like "*Sherpa*" }
$voice.SelectVoice("TestSherpaVoice")
$voice.Speak("Hello world!")
```

## Troubleshooting

### Build Issues

**"WiX Toolset not found"**
- Install WiX Toolset v3.11 from https://wixtoolset.org/
- Ensure WIX environment variable is set

**"MSBuild not found"**
- Install Visual Studio 2022 with C++ Desktop Development workload
- Or use Developer Command Prompt for VS

**"dotnet command not found"**
- Install .NET 6 SDK from https://dotnet.microsoft.com/download

### Installation Issues

**"Class not registered" (0x80040154)**
- Run installer as Administrator
- Check Windows Event Viewer for detailed error

**"Voice not appearing in SAPI5"**
- Verify registry entries were created
- Restart any applications using SAPI5

**"Model file not found"**
- Download voice models using SherpaOnnxSAPIInstaller.exe
- Or manually place models in: `C:\Program Files\OpenAssistive\OpenSpeech\models\`

### Runtime Issues

**"DllRegisterServer entry point not found"**
- Ensure Release build was used (not Debug)
- Rebuild NativeTTSWrapper with /MT runtime

See [TROUBLESHOOTING.md](TROUBLESHOOTING.md) for more details.

## Project Structure

```
Installer/
├── Installer.csproj              # Console app project
├── SherpaOnnxSAPI.wixproj        # WiX installer project
├── Product.wxs                   # WiX product definition
├── Program.cs                    # Main installer code
├── ModelInstaller.cs             # Model download/install
├── Sapi5RegistrarExtended.cs    # SAPI5 registration
├── AzureConfigManager.cs         # Azure configuration
├── AzureTtsService.cs            # Azure TTS service
├── AzureVoiceInstaller.cs        # Azure voice installer
├── EngineConfigManager.cs        # Config file management
└── Shared/                       # Shared types
    ├── TtsModel.cs
    ├── AzureTtsModel.cs
    ├── AzureConfig.cs
    ├── LanguageInfo.cs
    └── LanguageCodeConverter.cs
```

## Advanced Configuration

### Custom Install Directory

```powershell
msiexec /i dist\SherpaOnnxSAPI.msi INSTALLFOLDER="D:\Custom\Path"
```

### Silent Install with Custom Path

```powershell
msiexec /i dist\SherpaOnnxSAPI.msi /quiet INSTALLFOLDER="D:\Custom\Path"
```

### Creating Custom Voice Configurations

Edit `engines_config.json` after installation:

```json
{
  "engines": {
    "my-voice": {
      "type": "sherpaonnx",
      "config": {
        "modelPath": "C:/path/to/model.onnx",
        "tokensPath": "C:/path/to/tokens.txt",
        "lengthScale": 1.15,
        "noiseScale": 0.667,
        "noiseScaleW": 0.8
      }
    }
  },
  "voices": {
    "my-voice": "my-voice"
  }
}
```

See [SETUP.md](SETUP.md) for more configuration options.

## Support

- Documentation: [README.md](README.md), [BUILD.md](BUILD.md), [SETUP.md](SETUP.md)
- Issues: Check [TROUBLESHOOTING.md](TROUBLESHOOTING.md)
- Source: https://github.com/OpenAssistive/SherpaOnnxSAPI-installer
