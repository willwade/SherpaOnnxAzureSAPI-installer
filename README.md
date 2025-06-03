# SherpaOnnx SAPI Installer

A complete, production-ready installer for multi-engine Text-to-Speech (TTS) voices with full Windows SAPI compatibility. Supports both **SherpaOnnx** (offline, high-quality) and **Azure TTS** (cloud-based) engines with 100% SAPI integration.

## 🎯 Key Features

- **🎵 100% SAPI Compatibility**: Works with any Windows application that uses SAPI
- **🚀 Dual Engine Support**: SherpaOnnx (offline) + Azure TTS (cloud)
- **⚡ Native Performance**: C++ COM wrapper for maximum compatibility
- **📦 Complete Installer**: Single executable with all dependencies
- **🎛️ Multiple Interfaces**: Command line, interactive mode, and programmatic API
- **🔧 Voice Management**: Install, uninstall, verify, and test voices
- **🌐 Dynamic Models**: Automatic download from online voice repositories

## 🏗️ Architecture

### SherpaOnnx Engine (Offline TTS)
```
SAPI Application
       ↓
Native COM Wrapper (C++)     ← 100% SAPI Compatible
       ↓
ProcessBridge (JSON IPC)
       ↓
SherpaWorker (.NET 6.0)
       ↓
SherpaOnnx (Native C++)
       ↓
High-Quality Audio Output
```

### Azure TTS Engine (Cloud TTS)
```
SAPI Application
       ↓
Managed COM Objects (.NET)   ← Full SAPI Integration
       ↓
Azure TTS API
       ↓
Cloud-Generated Audio
```

## 🚀 Quick Start

### Prerequisites
- Windows 10/11
- Administrator privileges
- .NET 6.0 Runtime (included in installer)
- .NET Framework 4.7.2+ (for Azure TTS)

### Installation

#### Option 1: Download Release (Recommended)
1. Download the latest release from [GitHub Releases](../../releases)
2. Extract the package
3. Run as Administrator:
   ```powershell
   sudo .\SherpaOnnxSAPIInstaller.exe
   ```

#### Option 2: Build from Source
```powershell
# Install prerequisites (see Build Instructions below)
git clone https://github.com/willwade/SherpaOnnxAzureSAPI-installer.git
cd SherpaOnnxAzureSAPI-installer

# Build complete installer
sudo .\BuildCompleteInstaller.ps1

# Or build native wrapper only
sudo .\BuildNativeOnly.ps1
```

## 📖 Usage

### Interactive Mode
```powershell
sudo .\SherpaOnnxSAPIInstaller.exe
```
Provides a user-friendly menu for:
- Installing SherpaOnnx voices
- Installing Azure TTS voices  
- Managing voice configurations
- Uninstalling voices

### Command Line Interface

#### Install SherpaOnnx Voice
```powershell
sudo .\SherpaOnnxSAPIInstaller.exe install amy
sudo .\SherpaOnnxSAPIInstaller.exe install jenny
```

#### Install Azure TTS Voice
```powershell
# Configure Azure credentials
sudo .\SherpaOnnxSAPIInstaller.exe save-azure-config --key YOUR_KEY --region eastus

# Install Azure voice
sudo .\SherpaOnnxSAPIInstaller.exe install-azure en-US-JennyNeural

# With style and role
sudo .\SherpaOnnxSAPIInstaller.exe install-azure en-US-AriaNeural --style cheerful --role YoungAdultFemale
```

#### List Available Voices
```powershell
# List Azure voices
sudo .\SherpaOnnxSAPIInstaller.exe list-azure-voices

# Verify installation
sudo .\SherpaOnnxSAPIInstaller.exe verify amy
```

#### Uninstall Voices
```powershell
# Uninstall specific voice
sudo .\SherpaOnnxSAPIInstaller.exe uninstall amy

# Uninstall all voices
sudo .\SherpaOnnxSAPIInstaller.exe uninstall all
```

### Testing Installation
```powershell
# Test with PowerShell
$voice = New-Object -ComObject SAPI.SpVoice
$voice.Speak("Hello from SherpaOnnx!")

# Test with specific voice
$voices = $voice.GetVoices()
$amyVoice = $voices | Where-Object { $_.GetDescription() -like "*amy*" }
$voice.Voice = $amyVoice
$voice.Speak("This is Amy speaking!")
```

## 🔧 Build Instructions

### Prerequisites for Building
1. **Visual Studio Build Tools 2022**
   - Download: https://visualstudio.microsoft.com/downloads/#build-tools-for-visual-studio-2022
   - Install with "C++ build tools" workload
   - Include "Windows 10/11 SDK" and "ATL for v143 build tools"

2. **.NET 6.0 SDK**
   - Download: https://dotnet.microsoft.com/download/dotnet/6.0
   - Choose "SDK x64" for Windows

3. **.NET Framework 4.7.2 Developer Pack**
   - Download: https://dotnet.microsoft.com/download/dotnet-framework/net472

### Build Process

#### Complete Build (All Components)
```powershell
# Clean build with all components
sudo .\BuildCompleteInstaller.ps1 -Clean

# Build specific configuration
sudo .\BuildCompleteInstaller.ps1 -Configuration Release
```

#### Native Wrapper Only
```powershell
# Build just the native COM wrapper (when .NET 6.0 not available)
sudo .\BuildNativeOnly.ps1
```

#### Manual Build Steps
```powershell
# Restore packages
dotnet restore TTSInstaller.sln

# Build managed components
dotnet build "OpenSpeechTTS\OpenSpeechTTS.csproj" --configuration Release
dotnet build "SherpaWorker\SherpaWorker.csproj" --configuration Release
dotnet build "TTSInstaller.csproj" --configuration Release

# Build native wrapper
$msbuild = "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
& $msbuild "NativeTTSWrapper\NativeTTSWrapper.vcxproj" /p:Configuration=Release /p:Platform=x64

# Publish single executable
dotnet publish "TTSInstaller.csproj" --configuration Release --runtime win-x64 --self-contained true /p:PublishSingleFile=true
```

## 📁 Project Structure

```
SherpaOnnxAzureSAPI-installer/
├── Installer/                    # Main installer logic
│   ├── Program.cs                # Entry point and CLI interface
│   ├── ModelInstaller.cs         # SherpaOnnx model management
│   ├── AzureVoiceInstaller.cs    # Azure TTS integration
│   └── Sapi5RegistrarExtended.cs # Voice registration system
├── NativeTTSWrapper/             # Native C++ COM wrapper
│   ├── NativeTTSWrapper.cpp      # Main implementation
│   ├── NativeTTSWrapper.h        # Interface definitions
│   └── NativeTTSWrapper.vcxproj  # Visual Studio project
├── OpenSpeechTTS/                # Managed COM objects
│   ├── Sapi5VoiceImpl.cs         # SherpaOnnx SAPI implementation
│   └── AzureSapi5VoiceImpl.cs    # Azure TTS SAPI implementation
├── SherpaWorker/                 # ProcessBridge worker
│   └── Program.cs                # .NET 6.0 worker process
├── BuildCompleteInstaller.ps1    # Complete build script
├── BuildNativeOnly.ps1           # Native-only build script
└── TTSInstaller.sln              # Visual Studio solution
```

## 🎵 How It Works

### SherpaOnnx Voice Pipeline
1. **SAPI Application** calls standard `voice.Speak()` method
2. **Native COM Wrapper** receives the call with 100% SAPI compatibility
3. **ProcessBridge** communicates via JSON IPC for isolation
4. **SherpaWorker** processes the request using .NET 6.0
5. **SherpaOnnx** generates high-quality audio using neural models
6. **Audio Output** is returned through the SAPI pipeline

### Azure TTS Voice Pipeline
1. **SAPI Application** calls standard `voice.Speak()` method
2. **Managed COM Objects** handle the SAPI interface
3. **Azure TTS API** processes the request in the cloud
4. **Audio Output** is streamed back through SAPI

### Key Advantages
- **100% SAPI Compatibility**: Works with any Windows application
- **No Application Changes**: Existing software works immediately
- **High Performance**: Native C++ wrapper for optimal speed
- **Robust Architecture**: Process isolation prevents crashes
- **Dual Engine Support**: Best of both offline and cloud TTS

## 🔧 Technical Details

### Components
- **Native COM Wrapper**: 108.5 KB C++ DLL with full SAPI interfaces
- **ProcessBridge System**: .NET 6.0 SherpaWorker (58.7 MB)
- **Managed COM Objects**: .NET Framework 4.7.2 for Azure integration
- **Voice Models**: Downloaded automatically from online repositories
- **Registry Integration**: Complete SAPI voice registration

### Supported Voice Formats
- **SherpaOnnx**: ONNX neural models (Piper, MMS, VITS, Coqui)
- **Azure TTS**: All Azure Neural voices with styles and roles
- **Languages**: 100+ languages supported across both engines

### System Requirements
- **OS**: Windows 10/11 (x64)
- **Memory**: 4GB RAM minimum, 8GB recommended
- **Storage**: 500MB for installer, additional space for voice models
- **Network**: Internet connection for model downloads and Azure TTS

## 🐛 Troubleshooting

### Common Issues

#### Voice Not Appearing
```powershell
# Check voice registration
Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Speech\Voices\Tokens\*" | Select-Object PSChildName

# Re-register COM objects
sudo regsvr32 "C:\Program Files\OpenAssistive\OpenSpeech\NativeTTSWrapper.dll"
```

#### Build Failures
```powershell
# Check prerequisites
dotnet --version          # Should show 6.0.x+
where msbuild            # Should find MSBuild.exe

# Clean and rebuild
sudo .\BuildCompleteInstaller.ps1 -Clean
```

#### Audio Issues
- Check Windows audio settings
- Verify voice model files in `C:\Program Files\OpenSpeech\models\`
- Review logs in `C:\OpenSpeech\*.log`

### Debug Logs
- **Native Wrapper**: `C:\OpenSpeech\native_tts_debug.log`
- **ProcessBridge**: `C:\OpenSpeech\sherpa_debug.log`
- **Azure TTS**: `C:\OpenSpeech\azure_debug.log`

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📄 License

This project is licensed under the Apache 2.0 License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- [SherpaOnnx](https://github.com/k2-fsa/sherpa-onnx) - High-quality neural TTS
- [Microsoft Azure TTS](https://azure.microsoft.com/services/cognitive-services/text-to-speech/) - Cloud TTS service
- [Piper](https://github.com/rhasspy/piper) - Neural voice models
- Windows SAPI - Speech API framework

---

**🎉 Ready to give your applications a voice? Install SherpaOnnx SAPI today!**
