# SherpaOnnxAzureSAPI-installer

A **simplified Python + C++ SAPI bridge** that connects Windows SAPI applications to the **AACSpeakHelper pipe service**. This allows any SAPI application to use multiple TTS engines (Azure TTS, SherpaOnnx, Google TTS, etc.) through a unified interface.

## 🎯 **Simplified Architecture**

**Clean, maintainable codebase with two main components:**
- ✅ **C++ COM Wrapper** - Pure SAPI interface, communicates only with AACSpeakHelper pipe
- ✅ **Python Voice Installer** - Unified installer replacing .NET complexity
- ✅ **AACSpeakHelper Service** - Multi-engine TTS backend (separate repository)
- ✅ **No fallback chains** - Simple, reliable pipe communication only
- ✅ **Configuration-based voices** - JSON voice definitions

## 🎯 Key Features

- **🎵 100% SAPI Compatibility**: Works with any Windows application that uses SAPI
- **🚀 Multi-Engine Support**: Azure TTS, SherpaOnnx, Google TTS, ElevenLabs, and more
- **⚡ Native C++ Performance**: C++ COM wrapper for maximum compatibility and speed
- **🔧 Unified Interface**: All TTS engines accessible through AACSpeakHelper pipe service
- **📦 Easy Installation**: Automated installer with one-click setup
- **🎛️ Configuration-Based**: JSON voice configurations for easy management
- **🌐 Extensible**: Easy to add new TTS engines via AACSpeakHelper
- **🧪 Comprehensive Testing**: Complete integration testing framework
- **🏗️ CI/CD Ready**: Automated builds with GitHub Actions

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

## 🔑 Component CLSIDs

**IMPORTANT**: All voices must be registered with the correct C++ COM wrapper CLSID:

| Component | CLSID | Purpose |
|-----------|-------|---------|
| **C++ NativeTTSWrapper** | `{E1C4A8F2-9B3D-4A5E-8F7C-2D1B3E4F5A6B}` | ✅ **CORRECT** - Use this for all voices |
| ~~Old .NET Wrapper~~ | ~~`{4A8B9C2D-1E3F-4567-8901-234567890ABC}`~~ | ❌ **WRONG** - Do not use |

**All SAPI voices must use the C++ wrapper CLSID**: `{E1C4A8F2-9B3D-4A5E-8F7C-2D1B3E4F5A6B}`

### Key Components

- **C++ COM Wrapper**: Native SAPI interface with full compatibility
- **AACSpeakHelper Service**: Python-based TTS engine manager with pipe interface
- **Voice Configurations**: JSON files defining TTS engine settings
- **CLI Installer**: Python tool for voice management (matches AACSpeakHelper pattern)

## 🚀 Quick Start

### Prerequisites
- Windows 10/11
- Administrator privileges (for COM registration)
- Python 3.11+ with required packages
- AACSpeakHelper service running

### Installation Steps

#### Step 1: Set up AACSpeakHelper Service
```bash
# Clone and set up AACSpeakHelper
git clone https://github.com/AceCentre/AACSpeakHelper
cd AACSpeakHelper
uv venv && uv sync --all-extras

# Start the service
uv run python AACSpeakHelperServer.py
```

#### Step 2: Install SAPI Bridge
```powershell
# Clone this repository
git clone https://github.com/willwade/SherpaOnnxAzureSAPI-installer.git
cd SherpaOnnxAzureSAPI-installer

# Install a voice using the Python installer (requires admin PowerShell)
sudo python sapi_voice_installer.py install English-SherpaOnnx-Jenny
```

### Quick Test
```powershell
# Test the voice installation
.\test-voice.ps1 Jenny

# Or test SAPI synthesis directly
$voice = New-Object -ComObject SAPI.SpVoice
$voice.Speak("Hello from SherpaOnnx via AACSpeakHelper!")
```

## 📖 Usage

### Voice Management

#### Available Voice Configurations
- **`English-SherpaOnnx-Jenny`** - High-quality neural TTS (no credentials needed)
- **`English-Google-Basic`** - Google TTS (no credentials needed)
- **`British-English-Azure-Libby`** - Azure TTS British voice (requires API key)
- **`American-English-Azure-Jenny`** - Azure TTS American voice (requires API key)

#### Voice Installation Commands

**Install Voice**:
```powershell
# Install SherpaOnnx voice (no credentials needed)
sudo python sapi_voice_installer.py install English-SherpaOnnx-Jenny

# Install Google TTS voice (no credentials needed)
sudo python sapi_voice_installer.py install English-Google-Basic

# Install Azure TTS voice (requires API key in AACSpeakHelper)
sudo python sapi_voice_installer.py install British-English-Azure-Libby
```

**List Voices**:
```powershell
# List all installed SAPI voices
python sapi_voice_installer.py list
```

**Test Voices**:
```powershell
# Test specific voice
.\test-voice.ps1 Jenny

# List all available SAPI voices
.\test-voice.ps1
```

**Remove Voices**:
```powershell
# Remove a voice
sudo python sapi_voice_installer.py uninstall English-SherpaOnnx-Jenny
```

### Testing Installation

#### Automated Integration Testing
```powershell
# Run the complete integration test
.\test_windows_integration.ps1

# Test with Google TTS as well
.\test_windows_integration.ps1 -TestGoogle
```

#### Manual SAPI Testing
```powershell
# Test with PowerShell SAPI
$voice = New-Object -ComObject SAPI.SpVoice
$voice.Speak("Hello from AACSpeakHelper!")

# Test with specific voice
$voices = $voice.GetVoices()
$sherpaVoice = $voices | Where-Object { $_.GetDescription() -like "*SherpaOnnx*" }
$voice.Voice = $sherpaVoice
$voice.Speak("This is Jenny from SherpaOnnx neural TTS!")
```

#### Real Application Testing
- **Notepad**: Select text → Right-click → "Speak selected text"
- **Windows Narrator**: Use installed voices in narrator settings
- **Screen Readers**: NVDA, JAWS, and other screen readers can use the voices

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
   - Install uv: `pip install uv`

### Automated Build Process

#### Complete Build and Test
```powershell
# Run the complete build and test workflow
.\test_complete_workflow.ps1

# Or run Windows integration test (includes build)
.\test_windows_integration.ps1
```

#### Create Installer Package
```powershell
# Create complete installer package
.\create_installer.ps1

# Create with custom version
.\create_installer.ps1 -Version "1.1.0"

# Skip build if already built
.\create_installer.ps1 -SkipBuild
```

### Manual Build Process

#### Build C++ COM Wrapper
```powershell
# Build the native SAPI COM wrapper
msbuild "NativeTTSWrapper\NativeTTSWrapper.vcxproj" /p:Configuration=Release /p:Platform=x64

# Register the COM wrapper (as Administrator)
regsvr32 "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll"
```

#### Build .NET Installer
```powershell
# Build the .NET installer
cd Installer
dotnet build -c Release
dotnet publish -c Release -o "..\build\installer" --self-contained true -r win-x64
cd ..
```

#### Build Python CLI with PyInstaller
```bash
# Set up Python environment
uv venv
uv sync --extra build

# Build standalone executable
uv run pyinstaller --onefile --name "SapiVoiceManager" SapiVoiceManager.py
```

## 📁 Project Structure

```
SherpaOnnxAzureSAPI-installer/
├── NativeTTSWrapper/             # C++ SAPI COM wrapper
│   ├── NativeTTSWrapper.cpp      # Main SAPI implementation with AACSpeakHelper pipe
│   ├── NativeTTSWrapper.h        # Interface definitions
│   ├── NativeTTSWrapper.idl      # COM interface definition
│   └── NativeTTSWrapper.vcxproj  # Visual Studio project
├── Installer/                    # .NET installer components
│   ├── Program.cs                # Entry point and CLI interface
│   ├── ConfigBasedVoiceManager.cs # Voice configuration management
│   └── Installer.csproj          # .NET project file
├── AACSpeakHelper/               # Python TTS service (submodule)
│   └── AACSpeakHelperServer.py   # Main service entry point
├── voice_configs/                # Voice configuration files (AACSpeakHelper format)
│   ├── English-SherpaOnnx-Jenny.json      # SherpaOnnx neural voice
│   ├── English-Google-Basic.json          # Google TTS voice
│   ├── British-English-Azure-Libby.json   # Azure TTS British voice
│   └── American-English-Azure-Jenny.json  # Azure TTS American voice
├── archive/                      # Archived old code
│   └── old-dotnet-projects/      # Legacy .NET implementations
├── install-voice.bat             # Voice installation script
├── register-com-wrapper.bat      # COM wrapper registration
├── test-voice.ps1               # Voice testing script
├── INTEGRATION_TESTING.md        # Complete testing guide
├── TODO.md                       # Project status and roadmap
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
- **SherpaOnnx**: High-quality neural models (offline, no credentials needed)
- **Google TTS**: Basic text-to-speech (online, no credentials needed)
- **Azure TTS**: All Azure Neural voices with styles and roles (requires API key)
- **ElevenLabs**: High-quality AI voices (requires API key)
- **OpenAI TTS**: OpenAI's text-to-speech models (requires API key)
- **And more**: Extensible via AACSpeakHelper plugin system

### System Requirements
- **OS**: Windows 10/11 (x64)
- **Memory**: 4GB RAM minimum, 8GB recommended for neural TTS
- **Python**: 3.11+ for AACSpeakHelper service
- **Network**: Internet connection for cloud TTS services (SherpaOnnx works offline)
- **Privileges**: Administrator access for COM registration

## 🐛 Troubleshooting

### Common Issues

#### Voice Not Appearing in SAPI
```powershell
# Check voice registration
Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\SPEECH\Voices\Tokens\*" | Select-Object PSChildName

# Re-register COM wrapper (as Administrator)
regsvr32 "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll"

# Check CLSID registration
Get-ItemProperty "HKCR:\CLSID\{E1C4A8F2-9B3D-4A5E-8F7C-2D1B3E4F5A6B}"
```

#### AACSpeakHelper Service Not Running
```bash
# Start AACSpeakHelper service
cd AACSpeakHelper
uv run python AACSpeakHelperServer.py

# Check if service is listening on pipe
# (Service should show "Waiting for client connection..." message)

# Test pipe connectivity
# The service creates: \\.\pipe\AACSpeakHelper
```

#### Voice Synthesis Fails
- Ensure AACSpeakHelper service is running and listening on pipe
- Check voice configuration file exists in `voice_configs/`
- Verify TTS engine credentials are configured in AACSpeakHelper (for cloud services)
- Check Windows Event Viewer for COM errors
- Run integration test: `.\test_windows_integration.ps1`

#### Build Issues
```powershell
# Check prerequisites
dotnet --version          # Should show 6.0.x+
where msbuild            # Should find MSBuild.exe
python --version         # Should show 3.11+
uv --version             # Should show uv version

# Run automated build test
.\test_complete_workflow.ps1
```

#### Integration Testing
```powershell
# Run complete integration test
.\test_windows_integration.ps1

# Test specific voice
.\test_windows_integration.ps1 -TestVoice "English-SherpaOnnx-Jenny"

# Test with Google TTS as well
.\test_windows_integration.ps1 -TestGoogle
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
