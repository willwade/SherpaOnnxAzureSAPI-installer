# Troubleshooting Guide - SherpaOnnx SAPI5 TTS Engine

Common issues and their solutions.

## Build Issues

### "Runtime Library mismatch"

**Error**:
```
error LNK2038: mismatch detected for 'RuntimeLibrary': value 'MT_StaticRelease'
doesn't match value 'MD_DynamicRelease'
```

**Cause**: SherpaOnnx libraries use MT (static runtime), but project uses MD (dynamic).

**Solution**:
1. Open Project Properties
2. Navigate to: C/C++ → Code Generation
3. Set Runtime Library to: `MultiThreaded (/MT)`
4. Rebuild

### "Cannot open include file 'sherpa-onnx/c-api/c-api.h'"

**Cause**: SherpaOnnx Windows binaries not downloaded.

**Solution**:
```powershell
cd NativeTTSWrapper
.\download_sherpa.ps1
```

Or manually download from:
https://huggingface.co/csukuangfj/sherpa-onnx-libs/resolve/main/win64/1.12.10/sherpa-onnx-v1.12.10-win-x64-static.tar.bz2

Extract to: `NativeTTSWrapper\libs-win\`

### "MIDL2270 duplicate UUID error"

**Error**:
```
midl : error MIDL2270 : duplicate UUID
```

**Cause**: Same UUID used for TypeLib and CoClass.

**Solution**: Use different UUIDs for:
- TypeLib: `{88B9DB87-7A85-4088-AE32-074F55718967}`
- CoClass: `{A1B2C3D4-E5F6-4A5B-8C9D-1E2F3A4B5C6D}`

## Registration Issues

### "Class not registered" (0x80040154)

**Symptoms**: Voice not found in enumeration or COM object creation fails.

**Solutions**:

1. **Run as Administrator**
   ```powershell
   Start-Process regsvr32 -ArgumentList '"path\to\NativeTTSWrapper.dll"' -Verb RunAs
   ```

2. **Verify DLL exists**
   ```powershell
   Test-Path "C:\github\SherpaOnnxAzureSAPI-installer\NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll"
   ```

3. **Check dependencies**
   ```powershell
   # Check for MSVC redistributables
   Get-WmiObject Win32_Product | Where-Object { $_.Name -like "*Visual C++*" }
   ```

4. **Use full path**
   ```powershell
   regsvr32 "C:\full\path\to\NativeTTSWrapper.dll"
   ```

### "DllRegisterServer entry point not found"

**Cause**: Using Debug build instead of Release.

**Solution**:
1. Build Release configuration (not Debug)
2. Check exports: `dumpbin /EXPORTS NativeTTSWrapper.dll`

Should see:
```
DllRegisterServer
DllUnregisterServer
```

### Registration succeeds but voice doesn't appear

**Cause**: Registry entry not created or incorrect.

**Solutions**:

1. **Check registry**
   ```powershell
   Get-Item 'HKLM:\SOFTWARE\Microsoft\Speech\Voices\Tokens\TestSherpaVoice' -ErrorAction SilentlyContinue
   ```

2. **Verify CLSID matches**
   - Registry: `HKLM\SOFTWARE\Classes\CLSID\{A1B2C3D4-E5F6-4A5B-8C9D-1E2F3A4B5C6D}`
   - Code: `NativeTTSWrapper.rgs` and `NativeTTSWrapper.idl`

3. **Re-create registry entries**
   ```powershell
   # Remove old entries
   Remove-Item 'HKLM:\SOFTWARE\Microsoft\Speech\Voices\Tokens\TestSherpaVoice' -Recurse -ErrorAction SilentlyContinue
   Remove-Item 'HKLM\SOFTWARE\Classes\CLSID\{A1B2C3D4-E5F6-4A5B-8C9D-1E2F3A4B5C6D}' -Recurse -ErrorAction SilentlyContinue

   # Re-register
   regsvr32 "path\to\NativeTTSWrapper.dll"
   ```

## Runtime Issues

### "Model file not found"

**Error Log**:
```
ERROR: Model file not found: C:\path\to\model.onnx
```

**Solutions**:

1. **Verify paths in engines_config.json**
   - Use double backslashes: `C:\\path\\to\\model.onnx`
   - Use forward slashes: `C:/path/to/model.onnx`
   - NOT: `C:\path\to\model.onnx`

2. **Check file exists**
   ```powershell
   Test-Path "C:\github\SherpaOnnxAzureSAPI-installer\models\amy\vits-piper-en_US-amy-low\en_US-amy-low.onnx"
   ```

3. **Check file permissions**
   - Ensure read access for SYSTEM or current user

4. **Check config file location**
   - Must be in same directory as DLL
   - Default: `x64\Release\engines_config.json`

### "Failed to load configuration"

**Error Log**:
```
ERROR: Failed to open configuration file: engines_config.json
```

**Cause**: Config file not in DLL directory.

**Solutions**:

1. **Copy config to output directory**
   ```powershell
   Copy-Item "NativeTTSWrapper\engines_config.json" "NativeTTSWrapper\x64\Release\"
   ```

2. **Or modify code to use absolute path**
   - Edit `NativeTTSWrapper.cpp:387`

### "Engine initialized successfully but no audio"

**Symptoms**:
- Logs show "Engine initialized successfully"
- SAPI returns success
- No audio output

**Possible Causes**:

1. **Audio format mismatch**
   - Check sample rates match
   - Verify 16-bit PCM format

2. **Buffer too small**
   - SAPI may request larger buffers
   - Check `bytesWritten` vs `audioData.size()`

3. **Muted system**
   - Check system volume
   - Check application volume

**Diagnostics**:
```powershell
# Check debug log
Get-Content "C:\OpenSpeech\native_tts_debug.log" -Tail 50

# Test with simple app
Add-Type -AssemblyName System.Speech
$voice = New-Object System.Speech.Synthesis.SpeechSynthesizer
$voice.SelectVoice("Test Sherpa Voice")
$voice.Speak("Testing 1 2 3")
```

### "Voice sounds like Minnie Mouse"

**Cause**: `lengthScale` parameter too low (speech too fast).

**Solution**: Adjust in `engines_config.json`:
```json
{
  "config": {
    "lengthScale": 1.1,    // Increase for slower speech
    "noiseScale": 0.667,
    "noiseScaleW": 0.8
  }
}
```

**Recommended values**:
- Fast: 0.8
- Normal: 1.0
- Slow: 1.1-1.2
- Very slow: 1.3-1.5

### "Voice sounds robotic/distorted"

**Possible Causes**:

1. **Quality setting too low**
   - Try medium quality model instead of low

2. **CPU overload**
   - Reduce `numThreads` to 1
   - Close other applications

3. **Wrong sample rate**
   - Verify SAPI5 format matches engine output
   - Default: 22050 Hz (SAPI5) vs 16000 Hz (model)

4. **Clipping/distortion**
   - Reduce input volume if applicable
   - Check for signal processing effects

### "Crackling or popping in audio"

**Possible Causes**:

1. **Buffer underrun**
   - Increase audio buffer size
   - Reduce `numThreads`

2. **Format conversion issue**
   - Verify float → int16 conversion
   - Check for clamping: `sample = std::max(-1.0f, std::min(1.0f, sample))`

3. **Driver issues**
   - Update audio drivers
   - Try different audio output device

## Performance Issues

### Slow generation speed

**Symptoms**: Long pause before speech starts.

**Solutions**:

1. **Use smaller model**
   - amy-low instead of amy-medium

2. **Increase threads**
   ```json
   "numThreads": 2
   ```

3. **Use CPU-specific optimizations**
   - Enable AVX2/AVX-512 in compiler settings
   - Use `/O2` optimization

4. **Preload model**
   - Load model at startup
   - Keep engine warm between calls

### High memory usage

**Symptoms**: System slowdown or out of memory errors.

**Solutions**:

1. **Reduce thread count**
   ```json
   "numThreads": 1
   ```

2. **Unload unused models**
   - Use only one voice at a time

3. **Close/reopen application**
   - Clears accumulated memory

4. **Check for memory leaks**
   ```powershell
   # Monitor memory usage
   Get-Process | Where-Object {$_.Name -like "*voice*"} | Select-Object Name, WS
   ```

## Debugging

### Enable Detailed Logging

Edit `engines_config.json`:
```json
{
  "config": {
    "debug": true  // Enable verbose logging
  }
}
```

### Check Logs

**Primary logs**:
```
C:\OpenSpeech\native_tts_debug.log       # Main DLL log
C:\OpenSpeech\engine_manager.log          # Engine manager log
```

**View in real-time**:
```powershell
Get-Content "C:\OpenSpeech\native_tts_debug.log" -Wait -Tail 50
```

### COM Registration Verification

```powershell
# Check COM object registration
Get-ChildItem 'HKLM:\SOFTWARE\Classes\CLSID\{A1B2C3D4-E5F6-4A5B-8C9D-1E2F3A4B5C6D}' -Recurse | Format-List

# Check voice registration
Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Speech\Voices\Tokens\TestSherpaVoice' -Recurse | Format-List

# Test COM object creation
$object = New-Object -ComObject "{A1B2C3D4-E5F6-4A5B-8C9D-1E2F3A4B5C6D}"
if ($object) { Write-Host "COM object created successfully" }
```

### SAPI5 Voice Enumeration

```powershell
Add-Type -AssemblyName System.Speech
$synthesizer = New-Object System.Speech.Synthesis.SpeechSynthesizer

$synthesizer.GetInstalledVoices() | Format-Table Name, Culture, Gender, Age, Description
```

Expected output includes:
```
Test Sherpa Voice    en-US    Female    Adult    Test Sherpa Voice
```

### Export Analysis

Check DLL exports:
```cmd
dumpbin /EXPORTS NativeTTSWrapper.dll
```

Should include:
```
DllRegisterServer
DllUnregisterServer
```

## Common Error Messages

### "SPERR_NOT_FOUND (0x80060050)"

**Cause**: `ISpObjectToken::GetStringValue()` failed for custom attribute.

**Solution**: Use `GetId()` instead to extract voice name from token path.

### "E_NOTIMPL (0x80004001)"

**Cause**: Feature not implemented.

**Solutions**:
- Direct SherpaOnnx fallback: Not needed (use TTSEngineManager)
- Azure TTS: Not yet implemented (stub only)

### "E_FAIL (0x80004005)"

**Generic failure** - Check logs for specific error.

## Getting Help

### Collect Diagnostic Information

```powershell
# Create diagnostic bundle
$diagDir = "C:\OpenSpeech\Diagnostics"
New-Item -ItemType Directory -Force $diagDir | Out-Null

# Copy logs
Copy-Item "C:\OpenSpeech\*.log" $diagDir\

# Registry settings
reg export "HKLM\SOFTWARE\Microsoft\Speech\Voices\Tokens\TestSherpaVoice" "$diagDir\voice_regs.reg"

# System info
Get-ComputerInfo | Out-File "$diagDir\system_info.txt"

# DLL exports
dumpbin /EXPORTS "x64\Release\NativeTTSWrapper.dll" > "$diagDir\dll_exports.txt"

# Create zip
Compress-Archive -Path $diagDir "$diagDir.zip"
```

### Log Locations

| Log | Location | Purpose |
|-----|-----------|---------|
| native_tts_debug.log | `C:\OpenSpeech\native_tts_debug.log` | Main DLL |
| engine_manager.log | `C:\OpenSpeech\engine_manager.log` | Engine manager |
| C:\OpenSpeech | `C:\OpenSpeech\` | Log directory |

### Useful Commands

```powershell
# Re-register DLL (as admin)
regsvr32 /u "path\to\NativeTTSWrapper.dll"
regsvr32 "path\to\NativeTTSWrapper.dll"

# Test voice enumeration
Add-Type -AssemblyName System.Speech
$voice = New-Object System.Speech.Synthesis.SpeechSynthesizer
$voice.GetInstalledVoices() | Where-Object { $_.VoiceInfo.Description -like "*Sherpa*" }

# View config
Get-Content "x64\Release\engines_config.json" | ConvertFrom-Json

# Check model files
Get-ChildItem "models\amy\vits-piper-en_US-amy-low\" | Select-Object Name, Length
```

## See Also

- [BUILD.md](BUILD.md) - Build instructions
- [SETUP.md](SETUP.md) - Configuration guide
- [ARCHITECTURE.md](ARCHITECTURE.md) - System design
