# SherpaOnnx SAPI5 TTS Engine - Build Instructions

Complete guide to build the NativeTTSWrapper DLL with SherpaOnnx integration for SAPI5.

## Prerequisites

### Required Software

1. **Visual Studio 2019/2022** (Community Edition works)
   - C++ Desktop Development workload
   - Windows SDK (10.0 or later)
   - MSVC v143 or later (v145 used in this project)

2. **Git**
   - For cloning repositories

3. **7-Zip** (optional but recommended)
   - For extracting downloaded archives

### Project Dependencies

All dependencies are included in the repository:
- SherpaOnnx v1.12.10 Windows binaries (pre-compiled)
- Azure Speech SDK C++ libraries
- spdlog (logging library)
- nlohmann/json (JSON parsing, included as header-only)

## Quick Start (Pre-built)

If you have the pre-built binaries:

1. Register the DLL:
   ```powershell
   regsvr32 "C:\path\to\NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll"
   ```

2. Test the voice:
   ```powershell
   powershell -ExecutionPolicy Bypass -File test_sapi5_extended.ps1
   ```

## Building from Source

### Step 1: Clone Repository

```bash
git clone https://github.com/yourusername/SherpaOnnxAzureSAPI-installer.git
cd SherpaOnnxAzureSAPI-installer
```

### Step 2: Download SherpaOnnx Model

Download the vits-piper-en_US-amy-low model:

```powershell
# Create model directory
New-Item -ItemType Directory -Force -Path "models\amy\vits-piper-en_US-amy-low"

# Download model (63 MB)
Invoke-WebRequest -Uri "https://huggingface.co/csukuangfj/sherpa-onnx-apk/resolve/main/tts-engine-new/1.10.17/sherpa-onnx-1.10.17-arm64-v8a-en-tts-vits-piper-en_US-amy-medium.tar.bz2" -OutFile "models\amy\sherpa-amy-model.tar.bz2"

# Extract with 7-Zip
& "${env:ProgramFiles}\7-Zip\7z.exe" x "models\amy\sherpa-amy-model.tar.bz2" -O"models\amy" -y
```

**Or manually download from:**
https://huggingface.co/csukuangfj/vits-piper-en_US-amy-low

### Step 3: Download SherpaOnnx Windows Libraries

```powershell
# Download sherpa-onnx-v1.12.10-win-x64-static
Invoke-WebRequest -Uri "https://huggingface.co/csukuangfj/sherpa-onnx-libs/resolve/main/win64/1.12.10/sherpa-onnx-v1.12.10-win-x64-static.tar.bz2" -OutFile "NativeTTSWrapper\libs-win\sherpa-onnx-win-x64-static.tar.bz2"

# Extract
& "${env:ProgramFiles}\7-Zip\7z.exe" x "NativeTTSWrapper\libs-win\sherpa-onnx-win-x64-static.tar.bz2" -O"NativeTTSWrapper\libs-win" -y
```

Expected structure after extraction:
```
NativeTTSWrapper/libs-win/sherpa-onnx-v1.12.10-win-x64-static/
├── include/
│   └── sherpa-onnx/
│       └── c-api/
│           └── c-api.h
└── lib/
    ├── sherpa-onnx-c-api.lib
    ├── sherpa-onnx-core.lib
    ├── onnxruntime.lib
    └── ... (other dependencies)
```

### Step 4: Configure Project

The project is configured to use **MT (static runtime)** to match SherpaOnnx libraries.

**Important:** Do not change these settings:
- Configuration Type: DynamicLibrary
- Runtime Library: MultiThreaded (MT)
- Use of ATL: Static
- Platform Toolset: v143 or later

### Step 5: Build Solution

**Using Visual Studio:**
1. Open `NativeTTSWrapper\NativeTTSWrapper.sln`
2. Select Release | x64 configuration
3. Build → Build Solution (F7)

**Using MSBuild command line:**
```bash
cd NativeTTSWrapper
msbuild NativeTTSWrapper.sln /t:Build /p:Configuration=Release /p:Platform=x64
```

Expected output:
- DLL: `x64\Release\NativeTTSWrapper.dll`
- LIB: `x64\Release\NativeTTSWrapper.lib`
- EXP: `x64\Release\NativeTTSWrapper.exp`

### Step 6: Register DLL

**Administrator privileges required:**
```powershell
# Unregister old version (if exists)
regsvr32 /u "x64\Release\NativeTTSWrapper.dll"

# Register new version
regsvr32 "x64\Release\NativeTTSWrapper.dll"
```

**Or with elevation:**
```powershell
Start-Process regsvr32 -ArgumentList '"x64\Release\NativeTTSWrapper.dll"' -Verb RunAs -Wait
```

### Step 7: Verify Installation

```powershell
# Check voice is registered
Add-Type -AssemblyName System.Speech
$synthesizer = New-Object System.Speech.Synthesis.SpeechSynthesizer
$synthesizer.GetInstalledVoices() | Where-Object { $_.VoiceInfo.Description -like "*Sherpa*" }
```

Expected output:
```
Name     : Test Sherpa Voice
Culture   : en-US
Gender    : Female
Age       : Adult
Description : Test Sherpa Voice
```

## Troubleshooting

### Build Errors

**"Runtime Library mismatch"**
- Solution: Ensure project uses MT (static runtime), not MD
- Location: Project Properties → C/C++ → Code Generation → Runtime Library

**"Cannot open sherpa-onnx-c-api.h"**
- Solution: Run download_sherpa.ps1 to get Windows binaries
- Verify: `NativeTTSWrapper\libs-win\sherpa-onnx-v1.12.10-win-x64-static\include\sherpa-onnx\c-api\c-api.h` exists

### Registration Errors

**"Class not registered" (0x80040154)**
- Solution: Run regsvr32 with Administrator privileges
- Use full path to DLL

**"DllRegisterServer entry point not found"**
- Solution: Ensure Release build (not Debug)
- Check DLL exports: `dumpbin /EXPORTS NativeTTSWrapper.dll`

### Runtime Errors

**"Model file not found"**
- Solution: Update `engines_config.json` with correct model paths
- Check model files exist at specified locations

**"Failed to load configuration"**
- Solution: Ensure `engines_config.json` is in same directory as DLL
- Check JSON syntax is valid

### Voice Sounds Too Fast/Slow ("Minnie Mouse" effect)

The voice speed can be adjusted in `engines_config.json`:

```json
{
  "config": {
    "lengthScale": 1.0,    // Default speed (1.0 = normal)
    "noiseScale": 0.667,    // Pitch variation
    "noiseScaleW": 0.8      // Pitch stability
  }
}
```

- **lengthScale > 1.0**: Slower speech
- **lengthScale < 1.0**: Faster speech
- Recommended range: 0.8 to 1.2

## Advanced: Building SherpaOnnx from Source

If you need to build SherpaOnnx with MSVC (for MD runtime compatibility), see [SHERPAONNX_BUILD.md](SHERPAONNX_BUILD.md).

## Project Structure

```
SherpaOnnxAzureSAPI-installer/
├── NativeTTSWrapper/           # Main C++ project
│   ├── *.cpp, *.h              # Source files
│   ├── azure-speech-sdk/        # Azure Speech SDK
│   ├── libs-win/                # SherpaOnnx Windows binaries
│   ├── x64/Release/             # Built DLL
│   └── engines_config.json     # Engine configuration
├── models/                      # Downloaded TTS models
│   └── amy/                     # vits-piper-en_US-amy-low
├── test_sapi5_extended.ps1     # Test script
└── docs/                       # Documentation (this file)
```

## Configuration

### engines_config.json

Located in same directory as DLL. Contains engine definitions:

```json
{
  "engines": {
    "sherpa-amy": {
      "type": "sherpaonnx",
      "config": {
        "modelPath": "C:/path/to/model.onnx",
        "tokensPath": "C:/path/to/tokens.txt",
        "dataDir": "C:/path/to/espeak-ng-data",
        "noiseScale": 0.667,
        "noiseScaleW": 0.8,
        "lengthScale": 1.0,
        "numThreads": 1
      }
    }
  },
  "voices": {
    "amy": "sherpa-amy"
  }
}
```

### Registry Entries

Voice registration under:
```
HKLM\SOFTWARE\Microsoft\Speech\Voices\Tokens\TestSherpaVoice\
```

CLSID: `{A1B2C3D4-E5F6-4A5B-8C9D-1E2F3A4B5C6D}`

## Next Steps

After building and installing:

1. Test with various text samples
2. Adjust voice parameters for natural speech
3. Download additional voices (see [MODELS.md](MODELS.md))
4. Create installer for distribution

## See Also

- [SETUP.md](SETUP.md) - Model download and configuration
- [ARCHITECTURE.md](ARCHITECTURE.md) - System design
- [MODELS.md](MODELS.md) - Available TTS models
- [TROUBLESHOOTING.md](TROUBLESHOOTING.md) - Common issues
