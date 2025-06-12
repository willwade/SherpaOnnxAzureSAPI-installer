# C++ SAPI Bridge to AACSpeakHelper

A **C++ SAPI COM wrapper** that bridges Windows SAPI applications to the **AACSpeakHelper pipe service**. This allows any SAPI application to use multiple TTS engines (Azure TTS, SherpaOnnx, Google TTS, etc.) through a unified interface.

## 🎯 Key Features

- **🎵 100% SAPI Compatibility**: Works with any Windows application that uses SAPI
- **🚀 Multi-Engine Support**: Azure TTS, SherpaOnnx, Google TTS, ElevenLabs, and more
- **⚡ Native C++ Performance**: C++ COM wrapper for maximum compatibility
- **🔧 Unified Interface**: All TTS engines accessible through AACSpeakHelper pipe service
- **📦 Easy Installation**: CLI installer with voice management
- **🎛️ Configuration-Based**: JSON voice configurations for easy management
- **🌐 Extensible**: Easy to add new TTS engines via AACSpeakHelper

## 🏗️ Architecture

```
SAPI Application (Notepad, Screen Readers, etc.)
       ↓
C++ COM Wrapper (NativeTTSWrapper.dll)
       ↓
Named Pipe Communication (\\.\pipe\AACSpeakHelper)
       ↓
AACSpeakHelper Python Service
       ↓
Multiple TTS Engines (Azure, SherpaOnnx, Google, etc.)
```

### Key Components

- **C++ COM Wrapper**: Native SAPI interface with full compatibility
- **AACSpeakHelper Service**: Python-based TTS engine manager with pipe interface
- **Voice Configurations**: JSON files defining TTS engine settings
- **CLI Installer**: Python tool for voice management (matches AACSpeakHelper pattern)

## 🚀 Quick Start

### Prerequisites
- Windows 10/11
- Administrator privileges
- Python 3.11+ (for AACSpeakHelper service)
- Visual Studio Build Tools (for C++ compilation)
- .NET 6.0 SDK (for CLI installer)

### Installation

#### Step 1: Set up AACSpeakHelper Service
```bash
# Clone and set up AACSpeakHelper
git clone https://github.com/AceCentre/AACSpeakHelper
cd AACSpeakHelper
uv venv
uv sync --all-extras

# Start the service
uv run python AACSpeakHelperServer.py
```

#### Step 2: Build and Install SAPI Bridge
```powershell
# Clone this repository
git clone https://github.com/willwade/SherpaOnnxAzureSAPI-installer.git
cd SherpaOnnxAzureSAPI-installer

# Build the C++ COM wrapper
# (Build instructions below)

# Install voices using CLI tool
uv run python SapiVoiceManager.py
```

## 📖 Usage

### CLI Voice Manager (Interactive Mode)
```bash
# Start the interactive voice manager
uv run python SapiVoiceManager.py
```

This provides a user-friendly menu for:
- Installing voices from AACSpeakHelper configurations
- Managing voice registrations in Windows SAPI
- Testing voice synthesis
- Uninstalling voices

### Command Line Interface

#### Install Voice
```bash
# Install a voice by configuration name
uv run python SapiVoiceManager.py --install British-English-Azure-Libby

# Install multiple voices
uv run python SapiVoiceManager.py --install American-English-Azure-Jenny
uv run python SapiVoiceManager.py --install British-English-SherpaOnnx-Amy
```

#### List Available Voices
```bash
# List all available voice configurations
uv run python SapiVoiceManager.py --list

# List installed SAPI voices
uv run python SapiVoiceManager.py --list-installed
```

#### Remove Voice
```bash
# Remove a specific voice
uv run python SapiVoiceManager.py --remove British-English-Azure-Libby

# Remove all installed voices
uv run python SapiVoiceManager.py --remove-all
```

### Testing Installation
```powershell
# Test with PowerShell SAPI
$voice = New-Object -ComObject SAPI.SpVoice
$voice.Speak("Hello from AACSpeakHelper!")

# Test with specific voice
$voices = $voice.GetVoices()
$azureVoice = $voices | Where-Object { $_.GetDescription() -like "*Azure*Libby*" }
$voice.Voice = $azureVoice
$voice.Speak("This is Libby from Azure TTS!")
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

3. **Python 3.11+ with uv**
   - Download Python: https://www.python.org/downloads/
   - Install uv: `python -m pip install uv`

### Build Process

#### Build C++ COM Wrapper
```powershell
# Build the native SAPI COM wrapper
$msbuild = "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
& $msbuild "NativeTTSWrapper\NativeTTSWrapper.vcxproj" /p:Configuration=Release /p:Platform=x64

# Register the COM wrapper
sudo regsvr32 "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll"
```

#### Build CLI Installer
```powershell
# Build the .NET CLI installer
cd Installer
dotnet build --configuration Release
```

#### Set up Python Environment
```bash
# Set up Python dependencies
uv venv
uv sync
```

## 📁 Project Structure

```
SherpaOnnxAzureSAPI-installer/
├── NativeTTSWrapper/             # C++ SAPI COM wrapper
│   ├── NativeTTSWrapper.cpp      # Main SAPI implementation
│   ├── NativeTTSWrapper.h        # Interface definitions
│   └── NativeTTSWrapper.vcxproj  # Visual Studio project
├── Installer/                    # .NET CLI installer
│   ├── Program.cs                # Entry point and CLI interface
│   ├── ConfigBasedVoiceManager.cs # Voice configuration management
│   ├── PipeServiceBridge.cs      # AACSpeakHelper communication
│   └── PipeServiceComWrapper.cs  # SAPI COM integration
├── voice_configs/                # Voice configuration files
│   ├── British-English-Azure-Libby.json
│   ├── American-English-Azure-Jenny.json
│   └── British-English-SherpaOnnx-Amy.json
├── SapiVoiceManager.py           # Python CLI tool (AACSpeakHelper pattern)
├── SherpaOnnxAzureSAPI-installer.sln # Visual Studio solution
└── README.md                     # This file
```

## 🎵 How It Works

### Voice Synthesis Pipeline
1. **SAPI Application** calls standard `voice.Speak()` method
2. **C++ COM Wrapper** receives the call with 100% SAPI compatibility
3. **Named Pipe Communication** sends JSON request to AACSpeakHelper
4. **AACSpeakHelper Service** processes the request using configured TTS engine
5. **TTS Engine** (Azure, SherpaOnnx, Google, etc.) generates audio
6. **Audio Output** is returned through the pipe back to SAPI

### Voice Configuration System
1. **Voice Configs** define TTS engine settings in JSON format
2. **CLI Installer** registers voices with Windows SAPI registry
3. **COM Wrapper** loads voice config and communicates with AACSpeakHelper
4. **AACSpeakHelper** handles the actual TTS engine communication

### Key Advantages
- **100% SAPI Compatibility**: Works with any Windows application
- **No Application Changes**: Existing software works immediately
- **Multi-Engine Support**: Easy to add new TTS engines via AACSpeakHelper
- **Configuration-Based**: JSON configs make voice management simple
- **Unified Interface**: All TTS engines accessible through one pipe service

## 🔧 Technical Details

### Components
- **C++ COM Wrapper**: Native SAPI interface with full compatibility
- **AACSpeakHelper Service**: Python-based TTS engine manager
- **Named Pipe Communication**: `\\.\pipe\AACSpeakHelper` for IPC
- **Voice Configurations**: JSON files defining TTS engine settings
- **CLI Installer**: Python tool for voice management

### Supported TTS Engines (via AACSpeakHelper)
- **Azure TTS**: All Azure Neural voices with styles and roles
- **SherpaOnnx**: ONNX neural models (Piper, MMS, VITS, Coqui)
- **Google TTS**: Google Cloud Text-to-Speech
- **ElevenLabs**: High-quality AI voices
- **OpenAI TTS**: OpenAI's text-to-speech models
- **And more**: Extensible via AACSpeakHelper

### System Requirements
- **OS**: Windows 10/11 (x64)
- **Memory**: 4GB RAM minimum, 8GB recommended
- **Python**: 3.11+ for AACSpeakHelper service
- **Network**: Internet connection for cloud TTS services

## 🐛 Troubleshooting

### Common Issues

#### Voice Not Appearing in SAPI
```powershell
# Check voice registration
Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Speech\Voices\Tokens\*" | Select-Object PSChildName

# Re-register COM wrapper
sudo regsvr32 "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll"
```

#### AACSpeakHelper Service Not Running
```bash
# Start AACSpeakHelper service
cd AACSpeakHelper
uv run python AACSpeakHelperServer.py

# Check if service is listening on pipe
# (Service should show "Server listening on pipe..." message)
```

#### Voice Synthesis Fails
- Ensure AACSpeakHelper service is running
- Check voice configuration file exists in `voice_configs/`
- Verify TTS engine credentials are configured in AACSpeakHelper
- Check Windows Event Viewer for COM errors

#### Build Issues
```powershell
# Check prerequisites
dotnet --version          # Should show 6.0.x+
where msbuild            # Should find MSBuild.exe
python --version         # Should show 3.11+
uv --version             # Should show uv version
```

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📄 License

This project is licensed under the Apache 2.0 License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- [AACSpeakHelper](https://github.com/AceCentre/AACSpeakHelper) - Multi-engine TTS service
- [Ace Centre](https://acecentre.org.uk/) - Funding and supporting AACSpeakHelper
- [SherpaOnnx](https://github.com/k2-fsa/sherpa-onnx) - High-quality neural TTS
- [Microsoft Azure TTS](https://azure.microsoft.com/services/cognitive-services/text-to-speech/) - Cloud TTS service
- Windows SAPI - Speech API framework

---

**🎉 Ready to bridge SAPI applications to multiple TTS engines? Install the C++ SAPI Bridge today!**
