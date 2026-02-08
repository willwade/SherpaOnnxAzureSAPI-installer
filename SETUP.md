# Setup Guide - SherpaOnnx SAPI5 TTS Engine

Guide to download models, configure the engine, and adjust voice parameters.

## Quick Setup

### 1. Download TTS Model

The vits-piper-en_US-amy-low model is recommended for English female voice.

**Option A: Automated Download (PowerShell)**
```powershell
cd C:\github\SherpaOnnxAzureSAPI-installer

# Create model directory
New-Item -ItemType Directory -Force -Path "models\amy\vits-piper-en_US-amy-low"

# Download model (63 MB)
$modelUrl = "https://huggingface.co/csukuangfj/vits-piper-en_US-amy-low/resolve/main/vits-piper-en_US-amy-low.onnx"
$tokensUrl = "https://huggingface.co/csukuangfj/vits-piper-en_US-amy-low/resolve/main/tokens.txt"

Invoke-WebRequest -Uri $modelUrl -OutFile "models\amy\vits-piper-en_US-amy-low\en_US-amy-low.onnx"
Invoke-WebRequest -Uri $tokensUrl -OutFile "models\amy\vits-piper-en_US-amy-low\tokens.txt"
```

**Option B: Manual Download**
1. Visit: https://huggingface.co/csukuangfj/vits-piper-en_US-amy-low
2. Download:
   - `vits-piper-en_US-amy-low.onnx` (63 MB)
   - `tokens.txt`
3. Place in: `models\amy\vits-piper-en_US-amy-low\`

### 2. Update Configuration

Edit `NativeTTSWrapper\x64\Release\engines_config.json`:

```json
{
  "engines": {
    "sherpa-amy": {
      "type": "sherpaonnx",
      "config": {
        "modelPath": "C:\\github\\SherpaOnnxAzureSAPI-installer\\models\\amy\\vits-piper-en_US-amy-low\\en_US-amy-low.onnx",
        "tokensPath": "C:\\github\\SherpaOnnxAzureSAPI-installer\\models\\amy\\vits-piper-en_US-amy-low\\tokens.txt",
        "dataDir": "C:\\github\\SherpaOnnxAzureSAPI-installer\\models\\amy\\vits-piper-en_US-amy-low\\espeak-ng-data",
        "noiseScale": 0.667,
        "noiseScaleW": 0.8,
        "lengthScale": 1.0,
        "numThreads": 1,
        "provider": "cpu",
        "debug": true
      }
    }
  },
  "voices": {
    "amy": "sherpa-amy",
    "sherpa-amy": "sherpa-amy"
  }
}
```

### 3. Register Voice

```powershell
regsvr32 "C:\github\SherpaOnnxAzureSAPI-installer\NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll"
```

### 4. Test

```powershell
powershell -ExecutionPolicy Bypass -File test_sapi5_extended.ps1
```

## Voice Parameters

### Speed Control

Adjust `lengthScale` in `engines_config.json`:

| Value | Speed | Description |
|-------|-------|-------------|
| 0.8   | 25% faster | Fast speech, may sound hurried |
| 1.0   | Normal | Default speed |
| 1.2   | 20% slower | Slower, more deliberate |
| 1.5   | 50% slower | Very slow, good for accessibility |

### Pitch Control

Adjust `noiseScale` and `noiseScaleW`:

| Parameter | Range | Effect |
|-----------|-------|--------|
| noiseScale | 0.3-0.9 | Higher = more pitch variation |
| noiseScaleW | 0.5-1.2 | Higher = more stable pitch |

**Recommended for "Minnie Mouse" effect:**
```json
"noiseScale": 0.6,
"noiseScaleW": 0.9,
"lengthScale": 1.1
```

### Thread Count

- `numThreads: 1` - Single-threaded (default, safest)
- `numThreads: 2` - Dual-core (faster, uses more CPU)
- `numThreads: 4` - Quad-core (fastest, may cause conflicts)

## Model Details

### vits-piper-en_US-amy-low

- **Language**: English (US)
- **Gender**: Female
- **Sample Rate**: 16000 Hz
- **Model Size**: 63 MB
- **Quality**: Low (faster, less accurate)
- **License**: CC-BY-4.0 (Piper/VITS)

**Alternatives:**
- `vits-piper-en_US-amy-medium` - Better quality, slower (95 MB)
- `vits-piper-en_US-lessac-low` - Male voice (66 MB)
- `vits-piper-en_US-lessac-medium` - Male, better quality (99 MB)

## Alternative Models

### Other English Voices

Download from: https://huggingface.co/rhasspy/piper-voices/tree/main/en/en_US

| Model | Size | Quality | Gender |
|-------|------|--------|--------|
| amy-low | 63 MB | Low | Female |
| amy-medium | 95 MB | Medium | Female |
| lessac-low | 66 MB | Low | Male |
| lessac-medium | 99 MB | Medium | Male |
| kathleen-low | 65 MB | Low | Female |
| kathleen-medium | 98 MB | Medium | Female |

### British English

Download from: https://huggingface.co/rhasspy/piper-voices/tree/main/en/en_GB

| Model | Size | Quality | Gender |
|-------|------|--------|--------|
| alan-low | 68 MB | Low | Male |
| alan-medium | 102 MB | Medium | Male |
| amy-mid-low | 67 MB | Low | Female |
| amy-mid-medium | 101 MB | Medium | Female |
| sweetbbak-amy | 178 MB | High | Female |

### Other Languages

See full list: https://huggingface.co/rhasspy/piper-voices

## Configuration Examples

### Multiple Voices

```json
{
  "engines": {
    "sherpa-amy": {
      "type": "sherpaonnx",
      "config": {
        "modelPath": "C:/models/amy/en_US-amy-low.onnx",
        "tokensPath": "C:/models/amy/tokens.txt",
        "noiseScale": 0.667,
        "noiseScaleW": 0.8,
        "lengthScale": 1.0,
        "numThreads": 1
      }
    },
    "sherpa-alan": {
      "type": "sherpaonnx",
      "config": {
        "modelPath": "C:/models/alan/en_GB-alan-low.onnx",
        "tokensPath": "C:/models/alan/tokens.txt",
        "noiseScale": 0.667,
        "noiseScaleW": 0.8,
        "lengthScale": 1.0,
        "numThreads": 1
      }
    }
  },
  "voices": {
    "amy-us": "sherpa-amy",
    "alan-uk": "sherpa-alan"
  }
}
```

### Slower, More Natural Speech

```json
{
  "config": {
    "lengthScale": 1.15,
    "noiseScale": 0.7,
    "noiseScaleW": 0.85,
    "sampleRate": 22050
  }
}
```

### Fast Speech (Accessibility)

```json
{
  "config": {
    "lengthScale": 0.8,
    "noiseScale": 0.6,
    "noiseScaleW": 0.9,
    "sampleRate": 22050
  }
}
```

## Troubleshooting

### Model Not Found

**Error**: "Model file not found: C:\path\to\model.onnx"

**Solutions**:
1. Verify model paths in `engines_config.json` use double backslashes `\\`
2. Check model files exist at specified locations
3. Verify file permissions (read access required)

### DLL Registration Fails

**Error**: "Class not registered" (0x80040154)

**Solutions**:
1. Run regsvr32 as Administrator
2. Check Windows Event Viewer for detailed error
3. Verify MSVC redistributable is installed

### Voice Sounds Distorted

**Symptoms**: Cracking, popping, or "robotic" sound

**Solutions**:
1. Try different `lengthScale` values (0.9-1.2)
2. Reduce `numThreads` to 1
3. Check CPU usage (close other applications)
4. Try lower quality model (faster generation)

### High Memory Usage

**Symptoms**: System slowdown when speaking

**Solutions**:
1. Reduce `numThreads` to 1
2. Use smaller model (low vs medium)
3. Close and reopen TTS application between long texts
4. Increase system RAM if possible

## Performance Optimization

### CPU Usage

| Setting | CPU Usage | Quality |
|---------|-----------|--------|
| 1 thread | 25-50% | Best |
| 2 threads | 40-80% | Good |
| 4 threads | 60-95% | Fair |

### Latency

| Model | Quality | Latency | Real-time Factor |
|-------|--------|--------|------------------|
| amy-low | Low | ~0.5s | 0.1x |
| amy-medium | Medium | ~1.0s | 0.05x |
| amy-high | High | ~3.0s | 0.015x |

Real-time Factor < 0.1 means faster than real-time.

## See Also

- [BUILD.md](BUILD.md) - Build instructions
- [ARCHITECTURE.md](ARCHITECTURE.md) - System design
- [MODELS.md](MODELS.md) - Available models
