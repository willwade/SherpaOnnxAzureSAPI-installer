# SAPI5 Voice Implementation Status

## Current State (2025-02-10)

### Working Components
- ✅ NativeTTSWrapper DLL builds successfully
- ✅ DLL can be registered with regsvr32 (64-bit)
- ✅ COM object can be instantiated directly via ProgID
- ✅ Voice token appears in SAPI5 voice enumeration
- ✅ CLSID is registered in HKLM\SOFTWARE\Classes\CLSID
- ✅ SherpaOnnx engine integration works (when called directly)

### NOT Working
- ❌ SAPI5 returns `E_ACCESSDENIED (0x80070005)` when attempting to speak
- ❌ DLL is not loaded by SAPI5 when voice is selected
- ❌ No debug log entries from SetObjectToken, Speak, or GetOutputFormat

### Root Cause Analysis

#### Key Finding #1: CLSID Registry Format
**Current (Incorrect):**
```
HKLM\SOFTWARE\Microsoft\Speech\Voices\Tokens\mms_hat\CLSID\
  (default) = {A1B2C3D4-E5F6-4A5B-8C9D-1E2F3A4B5C6D}
```

**Required (Correct):**
```
HKLM\SOFTWARE\Microsoft\Speech\Voices\Tokens\mms_hat\Attributes\
  CLSID = {A1B2C3D4-E5F6-4A5B-8C9D-1E2F3A4B5C6D}
```

**Fix Required in ConfigApp/MainForm.cs:**
```csharp
// REMOVE this line (around line 1118):
voiceKey.CreateSubKey("CLSID").SetValue("", "{A1B2C3D4-E5F6-4A5B-8C9D-1E2F3A4B5C6D}");

// ADD this in the Attributes section (after line 1129):
attrKey.SetValue("CLSID", "{A1B2C3D4-E5F6-4A5B-8C9D-1E2F3A4B5C6D}");
```

#### Key Finding #2: Microsoft vs 3rd Party Voice Architecture
- **Microsoft voices** (Zira, David, etc.) don't have CLSID subkeys - they use a shared Microsoft TTS engine with embedded voice data
- **3rd party voices** MUST use CLSID attribute to point to custom TTS engine
- Our implementation follows the correct 3rd party pattern

#### Key Finding #3: E_ACCESSDENIED Root Cause
- Error occurs BEFORE DLL is loaded (no debug log from our DLL)
- Suggests SAPI5 fails during voice validation, not during TTS engine instantiation
- Possible causes:
  1. Missing or incorrect CLSID attribute (confirmed - see above)
  2. Voice token validation fails before attempting to load DLL
  3. Security/checks fail due to registry format issues

### Registration Script Issues

#### Problem: PowerShell User Context
- Running scripts as Administrator causes `Environment.SpecialFolder.LocalApplicationData` to return admin user's profile
- ConfigApp was writing to wrong user's AppData
- **Current Workaround:** Config writes to DLL directory
- **Proper Fix Needed:** Detect original user when elevated

### File Changes Needed

#### 1. ConfigApp/MainForm.cs
**Lines 1111-1130:**
```csharp
// CURRENT (WRONG):
voiceKey.CreateSubKey("CLSID").SetValue("", "{A1B2C3D4-E5F6-4A5B-8C9D-1E2F3A4B5C6D}");

// CORRECT:
// (Remove the CLSID subkey creation, add CLSID as attribute instead)
```

#### 2. ConfigApp Configuration Paths
**Current:** Config writes to `C:\github\...\NativeTTSWrapper\x64\Release\engines_config.json`
**Should be:** `C:\Program Files\OpenSpeech\engines_config.json` (machine-wide)
**Or:** `C:\Users\<actual-user>\AppData\Local\OpenSpeech\engines_config.json` (user-specific)

### Test Scripts

#### Keep:
- `cleanup.ps1` - Removes all registrations, logs, temp files
- `register-dll.ps1` - Registers DLL with 64-bit regsvr32
- `check-dll-deps.ps1` - Checks for missing DLL dependencies

#### Remove (redundant debugging scripts):
- All `test-*.ps1` scripts created during debugging
- All `compare-*.ps1` scripts
- All `check-*.ps1` scripts except `check-dll-deps.ps1`

### Registry Structure

#### Voice Token (CORRECT FORMAT):
```
HKLM\SOFTWARE\Microsoft\Speech\Voices\Tokens\mms_hat\
  (default) = "Mms_hat (MMS)"
  Attributes\
    Language = 409
    Gender = Female
    Age = Adult
    Name = "Mms_hat (MMS)"
    Vendor = "OpenAssistive"
    Description = "MMS TTS Voice"
    Version = "1.0"
    CLSID = "{A1B2C3D4-E5F6-4A5B-8C9D-1E2F3A4B5C6D}"  ← MUST be attribute, not subkey!
```

#### COM Registration:
```
HKLM\SOFTWARE\Classes\CLSID\{A1B2C3D4-E5F6-4A5B-8C9D-1E2F3A4B5C6D}\
  (default) = "OpenSpeech SAPI5 TTS Engine"
  InprocServer32\
    (default) = "C:\...\NativeTTSWrapper.dll"
    ThreadingModel = "Both"
  ProgID = "NativeTTSWrapper.CNativeTTSWrapper.1"
```

### Next Steps

1. **Fix ConfigApp** - Change CLSID from subkey to attribute
2. **Rebuild & Test** - Uninstall voice, reinstall with fixed ConfigApp
3. **Verify DLL Loading** - Check if debug log shows DLL_PROCESS_ATTACH
4. **If Still Failing** - Check if SAPI5 requires additional attributes or registry entries

### References

- Espeak SAPI5 implementation: `C:\github\espeak-ng\src\windows\com\`
- Key differences: Espeak uses pure COM (no ATL), but architecture is similar
- Microsoft voices use different architecture (shared engine, no CLSID)
