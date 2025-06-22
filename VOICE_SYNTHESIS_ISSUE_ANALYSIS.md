# Voice Synthesis Issue Analysis

## Problem Summary

The SAPI voice installer successfully registers voices in Windows, but **speech synthesis fails** with COM errors. There's a critical disconnect between the updated source code and the compiled DLL.

## Current Status

### ✅ Working Components
- **Voice Installation**: `sudo uv run python sapi_voice_installer.py install English-SherpaOnnx-Jenny` ✅
- **Voice Registration**: Voices appear in Windows SAPI voice list ✅
- **Voice Discovery**: `.\test-voice.ps1` shows voices correctly ✅
- **COM Registration**: `regsvr32` succeeds without errors ✅
- **Dependency Loading**: Native DLL loads successfully with LoadLibrary ✅

### ❌ Failing Components
- **Speech Synthesis**: SAPI `Speak()` calls fail with "Unexpected HRESULT" ❌
- **COM Wrapper Execution**: Runtime behavior doesn't match source code ❌

## Root Cause Analysis

### Evidence of Source/Binary Disconnect

1. **Log Analysis Shows Wrong Code Path**:
   ```
   2025-6-22 18:15:26.922: Starting ProcessBridge audio generation...
   2025-6-22 18:15:26.924: Failed to launch SherpaWorker
   ```

2. **Source Code Shows Different Implementation**:
   ```cpp
   // Current source code (NativeTTSWrapper.cpp line 72-73):
   LogMessage(L"=== USING AACSPEAKHELPER PIPE SERVICE ONLY ===");
   HRESULT hr = GenerateAudioViaPipeService(text, audioData);
   ```

3. **Binary Behavior Contradicts Source**:
   - Source code: Uses AACSpeakHelper pipe service
   - Runtime logs: Uses old ProcessBridge + SherpaWorker
   - **Conclusion**: The registered DLL is NOT compiled from current source

### Build System Issues

1. **MSBuild Not Found**: `msbuild` command not recognized in PATH
2. **Visual Studio Environment**: vcvars64.bat path issues
3. **Dependency Management**: Missing DLLs in output directory (partially fixed)
4. **Path Issues**: COM wrapper uses relative paths instead of absolute paths

### Architecture Mismatch

**Expected Flow (Source Code)**:
```
SAPI → COM Wrapper → AACSpeakHelper Pipe → Python Service → TTS Engines
```

**Actual Flow (Runtime Logs)**:
```
SAPI → COM Wrapper → ProcessBridge → SherpaWorker.exe (FAILS)
```

## Technical Details

### File Timestamps
- `NativeTTSWrapper.dll`: 03/06/2025 06:09:24 (Today, but contains old code)
- Source files: Recently modified with AACSpeakHelper implementation

### Dependencies Status
- ✅ All required DLLs copied to Release directory
- ✅ Visual C++ Redistributables installed
- ✅ Native DLL loads successfully
- ❌ COM wrapper runtime behavior incorrect

### Path Issues Identified
```cpp
// WRONG (current source):
std::wstring configPath = L"C:\\Program Files\\OpenAssistive\\OpenSpeech\\voice_configs\\" + voiceNameStr + L".json";

// Should be absolute path, but DLL wasn't rebuilt with this fix
```

## Error Patterns

### COM Error Details
- **Error**: "Unexpected HRESULT has been returned from a call to a COM component"
- **Context**: During SAPI `Speak()` method execution
- **Timing**: After successful voice discovery and selection

### Log Pattern Analysis
```
✅ CNativeTTSWrapper constructor called
✅ SET OBJECT TOKEN CALLED  
✅ GET OUTPUT FORMAT CALLED
✅ SPEAK METHOD CALLED
❌ Starting ProcessBridge audio generation... (WRONG CODE PATH)
❌ Failed to launch SherpaWorker (OLD ARCHITECTURE)
```

## Debugging Strategy

### Immediate Verification Needed
1. **Confirm DLL Source**: Verify which source code version is actually compiled
2. **Build System Diagnosis**: Fix MSBuild environment setup
3. **Binary Analysis**: Compare expected vs actual DLL behavior
4. **Path Resolution**: Ensure correct file paths at runtime

### Testing Approach
1. **Incremental Testing**: Test each component in isolation
2. **Log Analysis**: Compare source code expectations vs runtime logs
3. **Dependency Verification**: Ensure all required components are present
4. **Environment Testing**: Test in clean environment

## Next Steps Priority

**HIGH PRIORITY**:
1. Fix build system to compile current source code
2. Verify DLL contains expected AACSpeakHelper implementation
3. Test speech synthesis with corrected binary

**MEDIUM PRIORITY**:
1. Implement robust path resolution
2. Add comprehensive error handling
3. Create automated build verification

**LOW PRIORITY**:
1. Optimize performance
2. Add additional TTS engine support
3. Improve logging and diagnostics

## Success Criteria

### Phase 1: Build System Fix
- [ ] MSBuild compiles source code successfully
- [ ] Generated DLL contains AACSpeakHelper pipe service code
- [ ] Runtime logs show correct code path execution

### Phase 2: Speech Synthesis
- [ ] SAPI `Speak()` calls succeed without COM errors
- [ ] Audio output generated successfully
- [ ] Voice synthesis works end-to-end

### Phase 3: Integration
- [ ] All installed voices work correctly
- [ ] AACSpeakHelper service integration functional
- [ ] Robust error handling and recovery

## Impact Assessment

**User Impact**: High - Core functionality (speech synthesis) completely broken
**Development Impact**: Medium - Build system issues affect development workflow  
**Architecture Impact**: Low - Overall design is sound, implementation disconnect only

---

*Analysis Date: 2025-06-22*
*Status: Critical Issue - Speech Synthesis Non-Functional*
