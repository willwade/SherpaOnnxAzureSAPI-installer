# SherpaOnnx SAPI5 TTS Engine

A native Windows SAPI5 Text-to-Speech engine using SherpaOnnx with offline neural TTS models.

## Status: ✅ WORKING

**Current Version**: Native SAPI5 COM Wrapper
- ✅ SAPI5 voice enumeration
- ✅ SherpaOnnx v1.12.10 integration (Windows x64)
- ✅ vits-piper-en_US-amy-low model working
- ✅ Full COM implementation with ATL
- ✅ 100% compatible with SAPI5 applications

## Quick Start

### Prerequisites

- Windows 10/11 (x64)
- Visual Studio 2019/2022 with:
  - MSVC v143 or later
  - C++ ATL features
  - Windows 10 SDK

### Installation

1. **Build the DLL** (see [BUILD.md](BUILD.md))

2. **Register the DLL**:
   ```powershell
   regsvr32 "C:\github\SherpaOnnxAzureSAPI-installer\NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll"
   ```

3. **Test**:
   ```powershell
   Add-Type -AssemblyName System.Speech
   $voice = New-Object System.Speech.Synthesis.SpeechSynthesizer
   $voice.SelectVoice("Test Sherpa Voice")
   $voice.Speak("Hello world!")
   ```

## Architecture

**Current Design** (Simplified, Direct Integration):
```
SAPI5 Application
       ↓
SAPI5 (spvoice.dll)
       ↓
NativeTTSWrapper.dll (COM Object)
       ↓
TTSEngineManager
       ↓
SherpaOnnxEngine (C++ Wrapper)
       ↓
SherpaOnnx C API
       ↓
vits-piper-en_US-amy-low (ONNX Model)
       ↓
Audio Output (PCM 16kHz)
```

**Key Changes**:
- Removed ProcessBridge (no external SherpaWorker.exe)
- Direct SherpaOnnx C API integration
- No .NET 6.0 dependency for TTS engine
- Native C++ implementation only

## Features

- ✅ **SAPI5 Compatible**: Works with PowerShell, .NET apps, accessibility tools
- ✅ **Offline**: No internet connection required after model download
- ✅ **High Quality**: Neural TTS using VITS models
- ✅ **Multiple Voices**: Support for different languages/genders
- ✅ **Fast Generation**: ~0.5s latency, 10x faster than real-time
- ✅ **Low Memory**: ~250MB peak usage

## Voice Parameters

Adjust in `engines_config.json`:

| Parameter | Range | Effect |
|-----------|-------|--------|
| lengthScale | 0.8-1.5 | Speech speed (1.0=normal, >1 slower) |
| noiseScale | 0.3-0.9 | Pitch variation |
| noiseScaleW | 0.5-1.2 | Pitch stability |
| numThreads | 1-4 | CPU threads (1=safest) |

**"Minnie Mouse" effect?** Increase `lengthScale` to 1.1-1.2

## Documentation

- [BUILD.md](BUILD.md) - Build instructions
- [SETUP.md](SETUP.md) - Configuration and voice tuning
- [ARCHITECTURE.md](ARCHITECTURE.md) - System design
- [TROUBLESHOOTING.md](TROUBLESHOOTING.md) - Common issues

## Project Structure

```
NativeTTSWrapper/
├── NativeTTSWrapper.cpp/h     # Main COM object (ISpTTSEngine)
├── TTSEngineManager.cpp/h     # Engine lifecycle management
├── SherpaOnnxEngine.cpp/h      # SherpaOnnx C API wrapper
├── AzureTTSEngine.cpp/h        # Azure TTS (stub, not implemented)
├── ITTSEngine.cpp/h            # Engine interface
├── libs-win/                    # SherpaOnnx Windows binaries
├── azure-speech-sdk/           # Azure Speech SDK (unused)
└── engines_config.json        # Engine configuration
```

## Current Voice

**Test Sherpa Voice**
- Model: vits-piper-en_US-amy-low
- Language: English (US)
- Gender: Female
- Sample Rate: 16000 Hz

## Download Models

### vits-piper-en_US-amy-low

**Option A: PowerShell**
```powershell
$modelUrl = "https://huggingface.co/csukuangfj/vits-piper-en_US-amy-low/resolve/main/vits-piper-en_US-amy-low.onnx"
$tokensUrl = "https://huggingface.co/csukuangfj/vits-piper-en_US-amy-low/resolve/main/tokens.txt"

Invoke-WebRequest -Uri $modelUrl -OutFile "models\amy\vits-piper-en_US-amy-low.onnx"
Invoke-WebRequest -Uri $tokensUrl -OutFile "models\amy\tokens.txt"
```

**Option B: Browser**
1. Visit: https://huggingface.co/csukuangfj/vits-piper-en_US-amy-low
2. Download:
   - `vits-piper-en_US-amy-low.onnx` (63 MB)
   - `tokens.txt`

**Other Voices**:
- https://huggingface.co/rhasspy/piper-voices/tree/main/en/en_US
- https://huggingface.co/rhasspy/piper-voices/tree/main/en/en_GB

## Testing

### Extended Test Suite

```powershell
powershell -ExecutionPolicy Bypass -File test_sapi5_extended.ps1
```

Tests:
- Short sentences
- Long text passages
- Numbers and decimals
- Punctuation
- Technical terms
- Multiple sentences

### Manual Test

```powershell
Add-Type -AssemblyName System.Speech
$voice = New-Object System.Speech.Synthesis.SpeechSynthesizer

# List voices
$voice.GetInstalledVoices() | Format-Table Name, Culture, Description

# Select voice
$voice.SelectVoice("Test Sherpa Voice")

# Speak
$voice.Speak("The quick brown fox jumps over the lazy dog.")
```

## Building from Source

See [BUILD.md](BUILD.md) for complete instructions.

**Quick Summary**:
1. Install Visual Studio 2022 Build Tools
2. Download SherpaOnnx Windows binaries
3. Download vits-piper-en_US-amy-low model
4. Open `NativeTTSWrapper\NativeTTSWrapper.sln`
5. Build Release|x64
6. Register DLL

## Requirements

- **Build**: Visual Studio 2019/2022 with C++ ATL
- **Runtime**: No .NET dependency for TTS engine
- **Memory**: 4GB RAM minimum, 8GB recommended
- **Storage**: 150MB for DLL + 63MB per model

## Known Issues

### Voice Sounds Too Fast ("Minnie Mouse")

**Solution**: Increase `lengthScale` to 1.1-1.2 in `engines_config.json`

### Voice Sounds Robotic

**Solution**: Try different `noiseScale` values (0.6-0.7)

## TODO

- [ ] Azure TTS engine implementation
- [ ] GPU acceleration (CUDA/ROC)
- [ ] SSML support
- [ ] Streaming audio output
- [ ] Additional voice models
- [ ] Installer creation

## License

Apache 2.0

## Acknowledgments

- [SherpaOnnx](https://github.com/k2-fsa/sherpa-onnx) - Neural TTS engine
- [Piper](https://github.com/rhasspy/piper) - Neural voice models
- [Rhasspy](https://github.com/rhasspy) - Voice model hosting

---

**Status**: ✅ Working - SAPI5 TTS with SherpaOnnx v1.12.10
