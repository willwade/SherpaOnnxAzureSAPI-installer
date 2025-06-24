# OpenSpeechSAPI

A **universal SAPI bridge** that connects Windows SAPI applications to multiple TTS engines through the **AACSpeakHelper pipe service**. This allows any SAPI application to use multiple TTS engines (Azure TTS, SherpaOnnx, Google TTS, ElevenLabs, OpenAI, and more) through a unified interface.


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

## 🔑 (Notes!) Component CLSIDs

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
# Clone this repository
git clone https://github.com/willwade/SherpaOnnxAzureSAPI-installer.git
cd SherpaOnnxAzureSAPI-installer

# Set up configuration (IMPORTANT: Never commit real API keys!)
cp settings.cfg.example settings.cfg
# Edit settings.cfg with your real Azure TTS keys and other credentials

uv venv && uv sync --all-extras

# Start the service (AACSpeakHelper is integrated in root directory)
uv run AACSpeakHelperServer.py
```

#### Step 2: Install SAPI Bridge
```powershell
# In a new admin PowerShell window, install a voice
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

#### Complete Build
```powershell
# Run the complete build workflow
.\build-all.bat
```

### Manual Build Process

#### Build C++ COM Wrapper
```powershell
# Build the native SAPI COM wrapper
.\build_com_wrapper.bat

# Register the COM wrapper (as Administrator) - USE THE BATCH SCRIPT!
sudo .\register-com-wrapper.bat

# Alternative manual build (if needed)
# msbuild "NativeTTSWrapper\NativeTTSWrapper.vcxproj" /p:Configuration=Release /p:Platform=x64
```

**⚠️ Important**: Always use `sudo .\register-com-wrapper.bat` for COM registration. Direct `regsvr32` calls often fail due to missing dependencies.

#### Build Python CLI with PyInstaller
```bash
# Set up Python environment
uv venv
uv sync --extra build

# Build standalone executable
uv run pyinstaller --onefile --name "sapi_voice_installer" sapi_voice_installer.py
```

## 📁 Project Structure

```
OpenSpeechSAPI/
├── NativeTTSWrapper/             # C++ SAPI COM wrapper
│   ├── NativeTTSWrapper.cpp      # Main SAPI implementation
│   ├── NativeTTSWrapper.h        # Interface definitions
│   ├── NativeTTSWrapper.idl      # COM interface definition
│   ├── NativeTTSWrapper.vcxproj  # Visual Studio project
│   └── x64/Release/              # Build output directory
├── voice_configs/                # Voice configuration files
│   ├── English-SherpaOnnx-Jenny.json      # SherpaOnnx neural voice
│   ├── English-Google-Basic.json          # Google TTS voice
│   ├── British-English-Azure-Libby.json   # Azure TTS British voice
│   └── American-English-Azure-Jenny.json  # Azure TTS American voice
├── assets/                       # Application assets
│   ├── configure.ico             # Configuration icon
│   └── translate.ico             # Application icon
├── AACSpeakHelperServer.py       # Python TTS service (integrated in root)
├── sapi_voice_installer.py       # Python voice installer CLI
├── test_pipe.py                  # Testing script for pipe service
├── build-all.bat                 # Complete build script
├── build_com_wrapper.bat         # C++ COM wrapper build script
├── register-com-wrapper.bat      # COM wrapper registration
├── uninstall-sapi-voices.bat     # Voice uninstaller script
├── install-voice.bat             # Voice installation script
├── test-voice.ps1                # Voice testing script
├── test-sapi-libby.ps1          # Specific Libby voice test
├── installer.nsi                 # NSIS installer script
├── pyproject.toml                # Python project configuration
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

#### COM Registration Issues
**⚠️ IMPORTANT**: Use the batch script for reliable COM registration:

```powershell
# ✅ RECOMMENDED: Use the registration script (most reliable)
sudo .\register-com-wrapper.bat

# ❌ AVOID: Direct regsvr32 often fails
# regsvr32 "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll"
```

**Why the batch script works better:**
- Handles dependency copying automatically
- Uses proper error handling
- Includes required DLL dependencies
- More reliable than direct regsvr32 calls

**If COM registration fails:**
```powershell
# Clean registry first
sudo reg delete "HKEY_CLASSES_ROOT\CLSID\{4A8B9C2D-1E3F-4567-8901-234567890ABC}" /f

# Then use the batch script
sudo .\register-com-wrapper.bat
```

#### AACSpeakHelper Service Not Running
```bash
# Start AACSpeakHelper service (from project root)
uv run AACSpeakHelperServer.py

# Check if service is listening on pipe
# (Service should show "Waiting for client connection..." message)

# Test pipe connectivity
# The service creates: \\.\pipe\AACSpeakHelper
```

#### Voice Synthesis Fails
- Ensure AACSpeakHelper service is running and listening on pipe
- Check voice configuration file exists in `voice_configs/`
- Verify TTS engine credentials are configured in `settings.cfg` (for cloud services)
- Check Windows Event Viewer for COM errors
- Test with: `.\test-voice.ps1 Jenny`

#### Build Issues
```powershell
# Check prerequisites
where msbuild            # Should find MSBuild.exe
python --version         # Should show 3.11+
uv --version             # Should show uv version

# Run complete build
.\build-all.bat
```

#### Manual Testing
```powershell
# Test specific voice
.\test-voice.ps1 Jenny

# Test Libby voice specifically
.\test-sapi-libby.ps1

# List all SAPI voices
.\test-voice.ps1
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

**🎉 Ready to bridge SAPI applications to multiple TTS engines? Install OpenSpeechSAPI today!**
