# Architecture - SherpaOnnx SAPI5 TTS Engine

System design and code flow for the NativeTTSWrapper SAPI5 engine.

## Overview

```
┌─────────────────────────────────────────────────────────────────┐
│                        SAPI5 Application                        │
│  (PowerShell, .NET apps, accessibility tools, etc.)            │
└────────────────────────────┬────────────────────────────────┘
                             │
                             ▼
┌─────────────────────────────────────────────────────────────────┐
│                       SAPI5 Layer (spvoice.dll)                   │
│  - Enumerates voices                                            │
│  - Calls ISpTTSEngine::Speak()                                  │
│  - Manages audio playback                                        │
└────────────────────────────┬────────────────────────────────┘
                             │
                             ▼
┌─────────────────────────────────────────────────────────────────┐
│                  NativeTTSWrapper.dll (COM Object)              │
│  - Implements ISpTTSEngine, ISpObjectWithToken                  │
│  - Registered as Test Sherpa Voice                              │
│  - CLSID: {A1B2C3D4-E5F6-4A5B-8C9D-1E2F3A4B5C6D}                    │
└───────────────┬─────────────────────────────────┬───────────────┘
                │                                 │
                ▼                                 ▼
┌───────────────────────────────┐   ┌─────────────────────────────────┐
│   TTSEngineManager            │   │   SherpaOnnxEngine             │
│   - Manages engine instances    │   │   - SherpaOnnx C API wrapper   │
│   - Parses config files        │   │   - Audio generation          │
│   - Voice-to-engine mapping   │   │   - Model loading              │
└───────────────┬───────────────┘   └───────────────┬───────────────┘
                │                                   │
                ▼                                   ▼
        ┌──────────────────────┐      ┌─────────────────────────────┐
        │  SherpaOnnx C API    │      │  vits-piper-en_US-amy-low   │
        │  (Windows binaries)  │      │  - Model: en_US-amy-low.onnx  │
        └──────────────────────┘      │  - Tokens: tokens.txt        │
                                         │  - Sample rate: 16000 Hz      │
                                         └─────────────────────────────┘
```

## Components

### 1. NativeTTSWrapper (C++ ATL COM Object)

**File**: `NativeTTSWrapper.cpp/h`

**Responsibilities**:
- Implements SAPI5 `ISpTTSEngine` interface
- Implements SAPI5 `ISpObjectWithToken` interface
- COM registration and class factory
- Audio data streaming to SAPI5
- Text extraction from SAPI5 fragments

**Key Methods**:
```cpp
STDMETHODIMP Speak(DWORD dwSpeakFlags, REFGUID rguidFormatId,
                   const WAVEFORMATEX* pWaveFormatEx,
                   const SPVTEXTFRAG* pTextFragList,
                   ISpTTSEngineSite* pOutputSite);

STDMETHODIMP GetOutputFormat(const GUID* pTargetFormatId,
                           const WAVEFORMATEX* pTargetWaveFormatEx,
                           GUID* pOutputFormatId,
                           WAVEFORMATEX** ppCoMemOutputWaveFormatEx);

STDMETHODIMP SetObjectToken(ISpObjectToken* pToken);
```

### 2. TTSEngineManager

**File**: `TTSEngineManager.cpp/h`

**Responsibilities**:
- Singleton pattern for global engine management
- Engine lifecycle (creation, initialization, shutdown)
- Configuration file parsing (JSON)
- Voice-to-engine mapping
- Thread-safe engine access

**Key Methods**:
```cpp
HRESULT InitializeEngine(const std::wstring& engineId, EngineType type,
                         const std::wstring& config);
ITTSEngine* GetEngine(const std::wstring& engineId);
HRESULT LoadConfiguration(const std::wstring& configPath);
HRESULT ParseConfiguration(const std::wstring& jsonConfig);
```

### 3. SherpaOnnxEngine

**File**: `SherpaOnnxEngine.cpp/h`

**Responsibilities**:
- Wraps SherpaOnnx C API
- Model loading and validation
- Audio generation (text → PCM samples)
- Voice parameter handling

**Key Methods**:
```cpp
HRESULT Initialize(const std::wstring& config);
HRESULT Generate(const std::wstring& text,
                 std::vector<float>& samples, int& sampleRate);
bool ValidateModelFiles();
bool IsInitialized();
```

### 4. AzureTTSEngine (Stub)

**File**: `AzureTTSEngine.cpp/h`

**Status**: Placeholder for future Azure Speech SDK integration

## Data Flow

### Speak Request Flow

```
1. Application calls voice.Speak("Hello world")
   │
2. SAPI5 looks up voice token in registry
   │
3. SAPI5 loads NativeTTSWrapper.dll
   │
4. SAPI5 calls SetObjectToken() with voice configuration
   │   └─ Extracts engine ID from token
   │   └─ Stores for later use
   │
5. SAPI5 calls GetOutputFormat()
   │   └─ Returns: 22050 Hz, 16-bit, mono
   │
6. SAPI5 calls Speak() with text fragments
   │
7. NativeTTSWrapper::Speak() processes request
   │   ├─ Extract text from fragments
   │   ├─ Call GenerateAudioViaNativeEngine()
   │   │   ├─ Get engine from TTSEngineManager
   │   │   ├─ If not found, InitializeEngineFromToken()
   │   │   │   ├─ Load engines_config.json
   │   │   │   ├─ ParseConfiguration()
   │   │   │   └─ Initialize each engine
   │   │   └─ engine->Generate(text, samples, sampleRate)
   │   │       └─ SherpaOnnxOfflineTtsGenerate()
   │   ├─ Convert float samples to 16-bit PCM
   │   ├─ Send SPEVENT_START_INPUT_STREAM
   │   ├─ Write audio data via pOutputSite->Write()
   │   └─ Send SPEVENT_END_INPUT_STREAM
   │
8. SAPI5 plays audio to speakers
```

### Engine Initialization Flow

```
1. SetObjectToken() called
   │
2. Extract engine ID from token path
   │   Token ID: "HKEY_LOCAL_MACHINE\...\Tokens\TestSherpaVoice"
   │   └─ Engine ID: "sherpa-amy"
   │
3. Store engine ID in m_currentEngineId
   │
4. (Later) Speak() called
   │
5. GenerateAudioViaNativeEngine() checks for existing engine
   │
6. If not found, InitializeEngineFromToken()
   │   ├─ Get TTSEngineManager singleton
   │   ├─ GetModuleDirectory() for config path
   │   ├─ manager.LoadConfiguration(configPath)
   │   │   ├─ Read JSON file
   │   │   ├─ Parse JSON
   │   │   ├─ For each engine in config:
   │   │   │   ├─ Create engine (SherpaOnnxEngine)
   │   │   │   ├─ engine->Initialize(config)
   │   │   │   │   ├─ Parse configuration
   │   │   │   │   ├─ Validate model files exist
   │   │   │   │   ├─ Create SherpaOnnx config
   │   │   │   │   ├─ SherpaOnnxCreateOfflineTts()
   │   │   │   │   └─ Get sample rate
   │   │   │   └─ Add to m_engines map
   │   │   └─ Parse voices mapping
   │   └─ Return S_OK
   │
7. Engine ready for audio generation
```

## Configuration System

### engines_config.json Structure

```json
{
  "engines": {
    "<engineId>": {
      "type": "sherpaonnx",
      "config": {
        "modelPath": "C:\\path\\to\\model.onnx",
        "tokensPath": "C:\\path\\to\\tokens.txt",
        "dataDir": "C:\\path\\to\\espeak-ng-data",
        "noiseScale": 0.667,
        "noiseScaleW": 0.8,
        "lengthScale": 1.0,
        "numThreads": 1
      }
    }
  },
  "voices": {
    "<voiceName>": "<engineId>"
  }
}
```

### Configuration Locations

The config file is searched in this order:
1. DLL directory: `GetModuleDirectory() + "\engines_config.json"`
2. If not found, fallback to hardcoded paths

## Memory Management

### Engine Lifecycle

- **Creation**: `TTSEngineManager::InitializeEngine()`
- **Storage**: `std::unique_ptr` in `m_engines` map
- **Access**: Thread-safe via mutex lock
- **Destruction**: `TTSEngineManager::~TTSEngineManager()`

### Audio Buffer Management

```
SherpaOnnx generates: std::vector<float> (32-bit float samples)
    ↓ ConvertFloatSamplesToBytes()
SAPI5 expects:    std::vector<BYTE> (16-bit PCM)
    ↓ pOutputSite->Write()
Audio played
```

## COM Threading Model

### Apartment Threading

- **ThreadingModel**: Both (STA/MTA compatible)
- **Registration**: `ForceRemove {A1B2C3D4-E5F6-4A5B-8C9D-1E2F3A4B5C6D}\InprocServer32`
- **Class Factory**: ATL `OBJECT_ENTRY_AUTO`

### Thread Safety

- **TTSEngineManager**: Mutex-protected `m_enginesMutex`
- **Engine instances**: Not thread-safe (use one instance per engine)
- **Config parsing**: Synchronized via mutex

## Registry Structure

### Voice Registration

```
HKLM\SOFTWARE\Microsoft\SPEECH\Voices\Tokens\TestSherpaVoice\
├── (default) = "Test Sherpa Voice"
├── CLSID = "{A1B2C3D4-E5F6-4A5B-8C9D-1E2F3A4B5C6D}"
└── Attributes\
    ├── Language = "409" (en-US)
    ├── Gender = "Female"
    ├── Age = "Adult"
    ├── Vendor = "OpenAssistive"
    ├── Name = "TestSherpaVoice"
    ├── VoiceType = "sherpa-amy"
    └── VoiceName = "amy"
```

### CLSID Registration

```
HKLM\SOFTWARE\Classes\CLSID\{A1B2C3D4-E5F6-4A5B-8C9D-1E2F3A4B5C6D}\
├── InprocServer32\
│   ├── (default) = "C:\path\to\NativeTTSWrapper.dll"
│   └── ThreadingModel = "Both"
└── TypeLib = "{88B9DB87-7A85-4088-AE32-074F55718967}"
```

## Audio Format

### SAPI5 Format

- **Sample Rate**: 22050 Hz (configurable)
- **Channels**: 1 (mono)
- **Bits Per Sample**: 16
- **Format**: PCM (uncompressed)

### SherpaOnnx Format

- **Sample Rate**: 16000 Hz (model-dependent)
- **Channels**: 1 (mono)
- **Bits**: 32-bit float
- **Conversion**: `float → int16` with scaling

### Conversion Process

```cpp
// SherpaOnnx generates: [-1.0, 0.5, 0.3, ...]
// Convert to int16: [-32768, 16384, 9830, ...]

for (float sample : samples) {
    // Clamp to [-1.0, 1.0]
    if (sample > 1.0f) sample = 1.0f;
    if (sample < -1.0f) sample = -1.0f;

    // Scale to int16 range
    int16_t sampleInt16 = static_cast<int16_t>(sample * 32767.0f);

    // Store as bytes (little-endian)
    audioData.push_back(sampleInt16 & 0xFF);
    audioData.push_back((sampleInt16 >> 8) & 0xFF);
}
```

## Error Handling

### HRESULT Return Codes

| Code | Value | Meaning |
|------|-------|---------|
| S_OK | 0x00000000 | Success |
| S_FALSE | 0x00000001 | Success with warning |
| E_INVALIDARG | 0x80070003 | Invalid parameter |
| E_FAIL | 0x80004005 | Unspecified failure |
| E_NOTIMPL | 0x80004001 | Not implemented |
| SPERR_NOT_FOUND | 0x80060050 | Attribute not found |

### Logging

**Locations**:
- Debug log: `C:\OpenSpeech\native_tts_debug.log`
- Engine manager log: `C:\OpenSpeech\engine_manager.log`

**Log Levels**:
- INFO: Normal operation
- ERROR: Failure conditions
- DEBUG: Detailed diagnostics (when debug=true)

## Performance Characteristics

### Generation Speed

| Model | RTF | Latency (per sentence) |
|-------|-----|----------------------|
| amy-low | 0.1x | 0.5s |
| amy-medium | 0.05x | 1.0s |
| amy-high | 0.015x | 3.0s |

RTF = Real-Time Factor (<1.0 means faster than real-time)

### Memory Usage

| Component | Memory |
|-----------|--------|
| SherpaOnnx libraries | ~150 MB (static) |
| Model (amy-low) | ~63 MB |
| Runtime per sentence | ~10-50 MB |
| Total | ~250 MB peak |

## Future Enhancements

### Planned Features

1. **Azure TTS Engine** - Complete Azure Speech SDK integration
2. **Voice Cloning** - Custom voice training support
3. **Streaming** - Real-time audio streaming
4. **SSML Support** - Speech Synthesis Markup Language
5. **GPU Acceleration** - CUDA/ROC support for faster generation

### Extensibility Points

- **ITTSEngine interface** - Add new engines (Coqui, Mimic, etc.)
- **EngineType enum** - Define new engine types
- **TTSEngineFactory** - Create engine instances

## See Also

- [BUILD.md](BUILD.md) - Build instructions
- [SETUP.md](SETUP.md) - Configuration guide
- [TROUBLESHOOTING.md](TROUBLESHOOTING.md) - Common issues
