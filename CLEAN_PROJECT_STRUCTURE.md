# Clean Project Structure

## 📁 Essential Files Only

After cleanup, the project now contains only the essential files needed for the C++ TTS implementation:

### 🔧 **Core C++ Implementation**
```
NativeTTSWrapper/
├── NativeTTSWrapper.cpp        # Main COM wrapper implementation
├── NativeTTSWrapper.h          # COM wrapper header
├── NativeTTSWrapper.idl        # COM interface definition
├── NativeTTSWrapper.vcxproj    # Visual Studio project file
├── ITTSEngine.h/.cpp           # TTS engine interface
├── TTSEngineManager.h/.cpp     # Engine management system
├── SherpaOnnxEngine.h/.cpp     # SherpaOnnx C++ integration
├── AzureTTSEngine.h/.cpp       # Azure TTS integration
├── sherpa-onnx-c-api.h         # SherpaOnnx C API headers
├── engines_config.json         # Engine configuration
├── libs/                       # SherpaOnnx libraries
│   ├── sherpa-onnx-c-api.dll
│   ├── sherpa-onnx-c-api.lib
│   └── onnxruntime.dll
└── x64/Release/                # Build output
    └── NativeTTSWrapper.dll
```

### 🔗 **Supporting .NET Code**
```
OpenSpeechTTS/                  # .NET SAPI implementation (fallback)
├── Sapi5VoiceImpl.cs          # Main SAPI voice implementation
├── SherpaTTS.cs               # SherpaOnnx .NET wrapper
├── AzureTTS.cs                # Azure TTS .NET wrapper
└── OpenSpeechTTS.csproj       # Project file

SherpaWorker/                   # Process bridge worker
├── Program.cs                 # Worker executable
└── SherpaWorker.csproj        # Project file

Installer/                      # Installation utilities
├── Program.cs                 # Main installer
├── Sapi5Registrar.cs          # SAPI voice registration
├── ModelInstaller.cs          # Model download/install
└── AzureVoiceInstaller.cs     # Azure voice setup
```

### 📜 **Essential Scripts**
```
RegisterAmyVoice.ps1           # Voice registration script
BuildNativeOnly.ps1            # Build C++ wrapper only
BuildCompleteInstaller.ps1     # Build full installer
TestAmyVoiceSpecific.ps1       # Essential functionality test
CleanupCodebase.ps1            # Codebase cleanup script
```

### 📦 **Distribution**
```
dist/                          # Built binaries and installer
├── NativeTTSWrapper.dll       # Main C++ COM wrapper
├── OpenSpeechTTS.dll          # .NET fallback implementation
├── SherpaWorker.exe           # Process bridge worker
├── sherpa-onnx.dll            # SherpaOnnx runtime
└── SherpaOnnxSAPIInstaller.exe # Complete installer
```

### 📚 **Documentation**
```
README.md                      # Project overview
CPP_TTS_IMPLEMENTATION_PLAN.md # Detailed implementation plan
LICENSE                        # Project license
CLEAN_PROJECT_STRUCTURE.md     # This file
```

### 🗃️ **Archive**
```
archive/                       # All experimental/old code
├── test-scripts/              # Test scripts and experiments
├── build-experiments/         # Build script experiments
├── old-dotnet/               # Old .NET projects
└── temp-files/               # Temporary and object files
```

## 🎯 **What We Removed**

### Test Scripts (moved to archive/test-scripts/)
- 30+ test scripts (Test*.ps1, Test*.cpp)
- Build experiment scripts
- Debug and diagnostic scripts

### Old .NET Projects (moved to archive/old-dotnet/)
- SherpaNative/ - Old native wrapper experiment
- SignAssembly/ - Assembly signing utilities
- KeyGenerator/ - Key generation utilities

### Temporary Files (moved to archive/temp-files/)
- Object files (*.obj)
- Executable files (*.exe)
- Configuration experiments
- Documentation drafts
- DLL files in root directory

### Build Artifacts (deleted)
- obj/ and bin/ directories
- Temporary build files

## 🚀 **Current Focus**

With the cleaned codebase, we can now focus on:

1. **Rebuilding NativeTTSWrapper.dll** with the latest fallback chain code
2. **Testing the native engine fallback chain** 
3. **Fixing SherpaWorker.exe issues** (exit code 2147516570)
4. **Performance optimization** and deployment

## 📋 **Essential Workflow**

1. **Build**: `.\BuildNativeOnly.ps1`
2. **Register**: `.\RegisterAmyVoice.ps1` (as admin)
3. **Test**: `.\TestAmyVoiceSpecific.ps1`
4. **Deploy**: Copy to `C:\Program Files\OpenAssistive\OpenSpeech\`

## 🎉 **Benefits of Cleanup**

- **Reduced complexity**: Only essential files visible
- **Faster navigation**: No clutter in file explorer
- **Clear purpose**: Each file has a specific role
- **Easier maintenance**: Less confusion about what's current
- **Better focus**: Can concentrate on the core implementation
- **Preserved history**: All experiments safely archived

The codebase is now clean, organized, and ready for the final push to complete the C++ TTS implementation!
