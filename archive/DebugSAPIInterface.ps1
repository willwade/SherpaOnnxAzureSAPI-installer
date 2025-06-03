# SAPI Interface Debugging Script
# This script will help us identify why SAPI doesn't call our methods

Write-Host "SAPI Interface Debugging Script" -ForegroundColor Green
Write-Host "===============================" -ForegroundColor Green
Write-Host ""

# Clear previous logs
$logDir = "C:\OpenSpeech"
if (Test-Path "$logDir\sapi_debug.log") { Clear-Content "$logDir\sapi_debug.log" }
if (Test-Path "$logDir\sherpa_debug.log") { Clear-Content "$logDir\sherpa_debug.log" }

Write-Host "Phase 1: Direct COM Object Testing" -ForegroundColor Yellow
Write-Host "===================================" -ForegroundColor Yellow

try {
    # Test 1: Create COM object directly
    Write-Host "1. Creating COM object directly..." -ForegroundColor Cyan
    $comObject = New-Object -ComObject "OpenSpeechTTS.Sapi5VoiceImpl"
    Write-Host "   ✅ COM object created successfully" -ForegroundColor Green
    
    # Test 2: Check interfaces
    Write-Host "2. Checking COM interfaces..." -ForegroundColor Cyan
    $type = $comObject.GetType()
    Write-Host "   Type: $($type.FullName)" -ForegroundColor White
    
    # Check if it implements the required interfaces
    $interfaces = $type.GetInterfaces()
    Write-Host "   Implemented interfaces:" -ForegroundColor White
    foreach ($interface in $interfaces) {
        Write-Host "     - $($interface.Name)" -ForegroundColor Gray
    }
    
    # Test 3: Call SetObjectToken
    Write-Host "3. Testing SetObjectToken..." -ForegroundColor Cyan
    $setTokenMethod = $type.GetMethod("SetObjectToken")
    if ($setTokenMethod) {
        $result = $setTokenMethod.Invoke($comObject, @([System.IntPtr]::Zero))
        Write-Host "   ✅ SetObjectToken returned: $result" -ForegroundColor Green
    } else {
        Write-Host "   ❌ SetObjectToken method not found" -ForegroundColor Red
    }
    
    # Test 4: Call GetOutputFormat
    Write-Host "4. Testing GetOutputFormat..." -ForegroundColor Cyan
    $getFormatMethod = $type.GetMethod("GetOutputFormat")
    if ($getFormatMethod) {
        Write-Host "   ✅ GetOutputFormat method found" -ForegroundColor Green
    } else {
        Write-Host "   ❌ GetOutputFormat method not found" -ForegroundColor Red
    }
    
    # Release COM object
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) | Out-Null
    $comObject = $null
    
} catch {
    Write-Host "❌ Error in direct COM testing: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "Phase 2: SAPI Voice Enumeration" -ForegroundColor Yellow
Write-Host "===============================" -ForegroundColor Yellow

try {
    # Test SAPI voice enumeration
    Write-Host "1. Creating SAPI Voice object..." -ForegroundColor Cyan
    $voice = New-Object -ComObject SAPI.SpVoice
    
    Write-Host "2. Enumerating voices..." -ForegroundColor Cyan
    $voices = $voice.GetVoices()
    $voiceCount = $voices.Count
    Write-Host "   Found $voiceCount voices:" -ForegroundColor White
    
    $amyVoice = $null
    for ($i = 0; $i -lt $voiceCount; $i++) {
        $voiceItem = $voices.Item($i)
        $voiceName = $voiceItem.GetDescription()
        Write-Host "     Voice $i`: $voiceName" -ForegroundColor Gray
        
        if ($voiceName -eq "amy") {
            $amyVoice = $voiceItem
            Write-Host "       ✅ Found Amy voice!" -ForegroundColor Green
        }
    }
    
    if ($amyVoice) {
        Write-Host "3. Testing Amy voice selection..." -ForegroundColor Cyan
        $currentVoice = $voice.Voice.GetDescription()
        Write-Host "   Current voice: $currentVoice" -ForegroundColor White
        
        # Set Amy voice
        $voice.Voice = $amyVoice
        $newVoice = $voice.Voice.GetDescription()
        Write-Host "   New voice: $newVoice" -ForegroundColor White
        
        if ($newVoice -eq "amy") {
            Write-Host "   ✅ Amy voice selected successfully" -ForegroundColor Green
            
            Write-Host "4. Testing speech synthesis..." -ForegroundColor Cyan
            Write-Host "   About to call voice.Speak() - check logs for method calls..." -ForegroundColor Yellow
            
            try {
                # This should trigger our Speak method
                $result = $voice.Speak("Hello! This is a test of the Amy voice.")
                Write-Host "   ✅ Speak() completed with result: $result" -ForegroundColor Green
            } catch {
                Write-Host "   ❌ Speak() failed: $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "   HRESULT: $($_.Exception.HResult)" -ForegroundColor Red
            }
        } else {
            Write-Host "   ❌ Failed to select Amy voice" -ForegroundColor Red
        }
    } else {
        Write-Host "   ❌ Amy voice not found in enumeration" -ForegroundColor Red
    }
    
} catch {
    Write-Host "❌ Error in SAPI testing: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "Phase 3: Log Analysis" -ForegroundColor Yellow
Write-Host "=====================" -ForegroundColor Yellow

# Check logs
if (Test-Path "$logDir\sapi_debug.log") {
    Write-Host "SAPI Debug Log Contents:" -ForegroundColor Cyan
    $logContent = Get-Content "$logDir\sapi_debug.log"
    if ($logContent) {
        foreach ($line in $logContent) {
            Write-Host "   $line" -ForegroundColor Gray
        }
    } else {
        Write-Host "   ❌ SAPI debug log is empty - methods not called!" -ForegroundColor Red
    }
} else {
    Write-Host "❌ SAPI debug log not found" -ForegroundColor Red
}

if (Test-Path "$logDir\sherpa_debug.log") {
    Write-Host ""
    Write-Host "Sherpa Debug Log Contents:" -ForegroundColor Cyan
    $logContent = Get-Content "$logDir\sherpa_debug.log"
    if ($logContent) {
        foreach ($line in $logContent) {
            Write-Host "   $line" -ForegroundColor Gray
        }
    } else {
        Write-Host "   No Sherpa debug log entries" -ForegroundColor Gray
    }
}

Write-Host ""
Write-Host "Phase 4: Interface Verification" -ForegroundColor Yellow
Write-Host "===============================" -ForegroundColor Yellow

# Check if our interfaces are properly registered
Write-Host "1. Checking ISpTTSEngine interface registration..." -ForegroundColor Cyan
try {
    $interfaceKey = Get-ItemProperty "HKLM:\SOFTWARE\Classes\Interface\{A74D7C8E-4CC5-4F2F-A6EB-804DEE18500E}" -ErrorAction Stop
    Write-Host "   ✅ ISpTTSEngine interface registered" -ForegroundColor Green
} catch {
    Write-Host "   ❌ ISpTTSEngine interface not registered" -ForegroundColor Red
}

Write-Host "2. Checking ISpObjectWithToken interface registration..." -ForegroundColor Cyan
try {
    $interfaceKey = Get-ItemProperty "HKLM:\SOFTWARE\Classes\Interface\{14056581-E16C-11D2-BB90-00C04F8EE6C0}" -ErrorAction Stop
    Write-Host "   ✅ ISpObjectWithToken interface registered" -ForegroundColor Green
} catch {
    Write-Host "   ❌ ISpObjectWithToken interface not registered" -ForegroundColor Red
}

Write-Host "3. Checking COM class registration..." -ForegroundColor Cyan
try {
    $clsidKey = Get-ItemProperty "HKLM:\SOFTWARE\Classes\CLSID\{3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2}" -ErrorAction Stop
    Write-Host "   ✅ COM class registered" -ForegroundColor Green
} catch {
    Write-Host "   ❌ COM class not registered" -ForegroundColor Red
}

Write-Host ""
Write-Host "=== DEBUGGING COMPLETE ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "Key Questions to Answer:" -ForegroundColor Yellow
Write-Host "1. Are our methods being called by SAPI? (Check logs)" -ForegroundColor White
Write-Host "2. Does direct COM object creation work? (Should be ✅)" -ForegroundColor White
Write-Host "3. Can SAPI enumerate and select our voice? (Should be ✅)" -ForegroundColor White
Write-Host "4. Does voice.Speak() fail with E_FAIL? (Current issue)" -ForegroundColor White
Write-Host ""
