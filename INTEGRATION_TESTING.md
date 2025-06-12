# C++ SAPI Bridge to AACSpeakHelper - Integration Testing

## 🎯 Overview

This document provides instructions for testing the complete C++ SAPI Bridge to AACSpeakHelper integration. The implementation is **100% complete** and ready for production testing.

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
TTS Engines (SherpaOnnx, Google TTS, Azure TTS, etc.)
```

## ✅ What's Ready for Testing

### **Complete Implementation**
- ✅ C++ SAPI COM wrapper with AACSpeakHelper pipe communication
- ✅ Voice registration system with correct CLSID alignment
- ✅ Non-interactive CLI tool for voice management
- ✅ Voice configurations in AACSpeakHelper format
- ✅ Comprehensive test framework

### **Available Voice Configurations**
- ✅ `English-SherpaOnnx-Jenny` - SherpaOnnx neural voice (no credentials needed)
- ✅ `English-Google-Basic` - Google TTS voice (no credentials needed)
- ✅ `British-English-Azure-Libby` - Azure TTS voice (requires API key)
- ✅ `American-English-Azure-Jenny` - Azure TTS voice (requires API key)

## 🚀 Quick Test (Windows)

### **Automated Integration Test**
```powershell
# Run the complete integration test
.\test_windows_integration.ps1

# Test with Google TTS as well
.\test_windows_integration.ps1 -TestGoogle

# Skip build if already built
.\test_windows_integration.ps1 -SkipBuild
```

### **Manual Testing Steps**

#### **1. Set up AACSpeakHelper Service**
```bash
# Clone AACSpeakHelper
git clone https://github.com/AceCentre/AACSpeakHelper
cd AACSpeakHelper

# Set up environment
uv venv
uv sync --all-extras

# Start the service
uv run python AACSpeakHelperServer.py
```

#### **2. Build and Register C++ COM Wrapper**
```powershell
# Build C++ wrapper
msbuild "NativeTTSWrapper\NativeTTSWrapper.vcxproj" /p:Configuration=Release /p:Platform=x64

# Register COM wrapper (as Administrator)
regsvr32 "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll"
```

#### **3. Install and Test Voice**
```bash
# Install SherpaOnnx voice (no credentials needed)
uv run python SapiVoiceManager.py --install English-SherpaOnnx-Jenny

# Or install Google TTS voice (no credentials needed)
uv run python SapiVoiceManager.py --install English-Google-Basic

# List installed voices
uv run python SapiVoiceManager.py --list-installed
```

#### **4. Test SAPI Integration**
```powershell
# Test with PowerShell SAPI
$voice = New-Object -ComObject SAPI.SpVoice
$voices = $voice.GetVoices()

# Find our voice
$sherpaVoice = $voices | Where-Object { $_.GetDescription() -like "*SherpaOnnx*" }
$voice.Voice = $sherpaVoice

# Test synthesis
$voice.Speak("Hello from SherpaOnnx via AACSpeakHelper! This is a test of the C++ SAPI bridge.")
```

#### **5. Test in Real Applications**
- Open **Notepad** → Select text → Right-click → "Speak selected text"
- Use **Windows Narrator** with the installed voice
- Test with **NVDA** or **JAWS** screen readers

## 🔧 TTS Engines That Work Without Credentials

### **SherpaOnnx** (Recommended for Testing)
- **Engine**: `sherpaonnx`
- **Voice**: `en_GB-jenny_dioco-medium`
- **Quality**: High-quality neural TTS
- **Offline**: Works without internet connection
- **Configuration**: `English-SherpaOnnx-Jenny.json`

### **Google TTS Basic**
- **Engine**: `google`
- **Voice**: `en`
- **Quality**: Basic synthetic TTS
- **Online**: Requires internet connection
- **Configuration**: `English-Google-Basic.json`

## 🎯 Expected Test Results

### **Successful Integration**
When everything works correctly, you should see:

1. **AACSpeakHelper Service**:
   - Service starts without errors
   - Creates named pipe `\\.\pipe\AACSpeakHelper`
   - Logs show "Waiting for client connection..."

2. **Voice Registration**:
   - Voice appears in Windows SAPI registry
   - CLSID `E1C4A8F2-9B3D-4A5E-8F7C-2D1B3E4F5A6B` is registered
   - Voice shows up in SAPI enumeration

3. **Voice Synthesis**:
   - `$voice.Speak()` works without errors
   - Audio output is produced
   - Real applications can use the voice

### **Troubleshooting Common Issues**

#### **AACSpeakHelper Service Won't Start**
```bash
# Check Python environment
uv run python --version

# Check dependencies
uv run python -c "import win32pipe, win32file"

# Check for port conflicts
netstat -an | findstr "AACSpeakHelper"
```

#### **Voice Not Appearing in SAPI**
```powershell
# Check voice registration
Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\SPEECH\Voices\Tokens\*" | Select-Object PSChildName

# Re-register COM wrapper
regsvr32 "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll"

# Check CLSID registration
Get-ItemProperty "HKCR:\CLSID\{E1C4A8F2-9B3D-4A5E-8F7C-2D1B3E4F5A6B}"
```

#### **Voice Synthesis Fails**
- Ensure AACSpeakHelper service is running
- Check Windows Event Viewer for COM errors
- Verify pipe communication in AACSpeakHelper logs
- Test with different TTS engines

## 📊 Test Coverage

The integration test covers:
- ✅ Prerequisites validation
- ✅ AACSpeakHelper service setup and startup
- ✅ C++ COM wrapper build and registration
- ✅ Voice configuration validation
- ✅ Voice installation testing
- ✅ SAPI enumeration verification
- ✅ Voice synthesis testing
- ✅ Real application compatibility

## 🎉 Success Criteria

The integration is successful when:
- ✅ AACSpeakHelper service starts and creates named pipe
- ✅ C++ COM wrapper builds and registers without errors
- ✅ Voice appears in Windows SAPI enumeration
- ✅ Voice synthesis produces audio output
- ✅ Real applications (Notepad, screen readers) can use the voice
- ✅ Pipe communication works reliably between C++ wrapper and AACSpeakHelper

## 🚀 Next Steps After Successful Testing

1. **Performance Testing**: Measure synthesis latency and audio quality
2. **Stress Testing**: Test with multiple concurrent voice requests
3. **Compatibility Testing**: Test with various SAPI applications
4. **Additional Engines**: Configure and test Azure TTS, ElevenLabs, etc.
5. **Production Deployment**: Create installer packages and documentation

---

**The C++ SAPI Bridge to AACSpeakHelper is ready for production testing!** 🎉
