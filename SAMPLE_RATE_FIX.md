# Sample Rate Fix - "Minnie Mouse" Voice Issue

## Problem

The SherpaOnnx SAPI5 voice was sounding like "Minnie Mouse" - faster and higher-pitched than normal.

### Root Cause

**Sample rate mismatch** between what SAPI5 expects and what SherpaOnnx produces:

| Component | Sample Rate | Result |
|-----------|-------------|--------|
| GetOutputFormat() | 22050 Hz (hardcoded) | SAPI5 expects 22.05kHz |
| SherpaOnnx Output | 16000 Hz (amy model) | Audio is 16kHz |
| **Playback** | **22.05kHz playback of 16kHz audio** | **Fast + high pitch** |

When 16kHz audio is played at 22.05kHz:
- Plays at **1.38x speed** (22050/16000)
- Pitch increases by **~38%**
- Creates the "Minnie Mouse" effect

## Solution

Instead of hardcoding the output format to 22050Hz, **query the engine** for its actual sample rate and return that.

### Changes Made

**1. Added member variable to store actual sample rate** (`NativeTTSWrapper.h:66`):
```cpp
int m_actualSampleRate;  // Actual sample rate from the engine (e.g., 16000 for SherpaOnnx)
```

**2. Initialize to default** (`NativeTTSWrapper.cpp:31`):
```cpp
CNativeTTSWrapper::CNativeTTSWrapper() : m_engineInitialized(false), m_actualSampleRate(16000)
```

**3. Query engine after initialization** (`NativeTTSWrapper.cpp:350-370`):
```cpp
ITTSEngine* engine = manager.GetEngine(m_currentEngineId);
if (engine)
{
    int channels, bitsPerSample;
    HRESULT hr = engine->GetSupportedFormat(m_actualSampleRate, channels, bitsPerSample);
    // ...
}
```

**4. Use actual sample rate in GetOutputFormat()** (`NativeTTSWrapper.cpp:134`):
```cpp
// OLD: pFormat->nSamplesPerSec = 22050;
// NEW:
pFormat->nSamplesPerSec = m_actualSampleRate;  // e.g., 16000
```

## Result

After rebuild:
- ✅ Sample rate matches engine output
- ✅ Audio plays at correct speed
- ✅ No pitch artifacts
- ✅ Works with any model sample rate (16kHz, 22.05kHz, 24kHz, etc.)

## Voice Parameter Recommendations

In `engines_config.json`:

```json
{
  "engines": {
    "sherpa-amy": {
      "config": {
        "lengthScale": 1.15,     // 15% slower (was 1.0)
        "noiseScale": 0.667,     // Pitch variation
        "noiseScaleW": 0.8       // Pitch stability
      }
    }
  }
}
```

| Parameter | Range | Effect |
|-----------|-------|--------|
| lengthScale | 0.8-1.5 | Speech speed (1.15 = 15% slower) |
| noiseScale | 0.3-0.9 | Higher = more pitch variation |
| noiseScaleW | 0.5-1.2 | Higher = more stable pitch |

## Building the Fix

### Prerequisites
- Visual Studio 2019/2022 with C++ ATL
- Windows 10 SDK
- SherpaOnnx Windows binaries in `libs-win/`

### Build Commands
```bash
cd NativeTTSWrapper
msbuild NativeTTSWrapper.sln /t:Build /p:Configuration=Release /p:Platform=x64
```

### After Build
```powershell
# Unregister old version
regsvr32 /u "x64\Release\NativeTTSWrapper.dll"

# Register new version
regsvr32 "x64\Release\NativeTTSWrapper.dll"

# Test
powershell -ExecutionPolicy Bypass -File test_sapi5_extended.ps1
```

## Technical Details

### Why the WAV Header Was Correct

The WAV header written in `ConvertFloatSamplesToBytes()` always used the actual sample rate from the engine:

```cpp
*reinterpret_cast<uint32_t*>(pData + 24) = static_cast<uint32_t>(sampleRate);
```

So the audio data was correctly labeled as 16kHz. The problem was that `GetOutputFormat()` told SAPI5 to expect 22.05kHz!

### The SAPI5 Format Negotiation

```
SAPI5: "What format do you support?"
GetOutputFormat(): "22.05kHz, 16-bit, mono"  // ❌ Wrong!

SAPI5: "Here's your 22.05kHz buffer"
Speak(): *generates 16kHz audio, writes it as 22.05kHz*  // ❌ Mismatch!

SAPI5: *plays 16kHz audio at 22.05kHz*  // ❌ Fast + high pitch!
```

### After Fix

```
SAPI5: "What format do you support?"
GetOutputFormat(): "16kHz, 16-bit, mono"  // ✅ Correct!

SAPI5: "Here's your 16kHz buffer"
Speak(): *generates 16kHz audio*  // ✅ Match!

SAPI5: *plays 16kHz audio at 16kHz*  // ✅ Normal speed and pitch!
```

## Alternative Solutions (Not Used)

### Option A: Resample 16kHz → 22.05kHz
- **Pros**: Could keep 22.05kHz output
- **Cons**: Adds CPU overhead, quality loss from resampling
- **Complexity**: Requires libsamplerate or similar

### Option B: Force 22.05kHz Models
- **Pros**: No code change needed
- **Cons**: Limits model selection, not all models available at 22.05kHz

### Option C: Selected Solution ✅
- **Pros**: Works with any sample rate, no overhead
- **Cons**: Requires rebuild of DLL
- **Implementation**: ~30 lines of code

## Testing

After rebuilding and registering, test with:

```powershell
Add-Type -AssemblyName System.Speech
$voice = New-Object System.Speech.Synthesis.SpeechSynthesizer
$voice.SelectVoice("TestSherpaVoice")
$voice.Speak("The quick brown fox jumps over the lazy dog.")
```

Expected result: **Natural speed and pitch**, no "Minnie Mouse" effect.

## References

- Sample rate in audio: https://en.wikipedia.org/wiki/Sampling_rate
- SAPI5 format negotiation: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ee125170(v=vs.85)
- SherpaOnnx models: https://github.com/k2-fsa/sherpa-onnx
