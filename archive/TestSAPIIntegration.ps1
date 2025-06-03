# Comprehensive SAPI integration test
Write-Host "COMPREHENSIVE SAPI INTEGRATION TEST" -ForegroundColor Green
Write-Host "====================================" -ForegroundColor Green
Write-Host ""

try {
    # Clear logs
    $logDir = "C:\OpenSpeech"
    if (Test-Path "$logDir\sapi_debug.log") { Clear-Content "$logDir\sapi_debug.log" }

    Write-Host "1. Testing voice enumeration..." -ForegroundColor Yellow
    $voice = New-Object -ComObject SAPI.SpVoice
    $voices = $voice.GetVoices()
    
    Write-Host "   Found $($voices.Count) voices:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $voices.Count; $i++) {
        $voiceToken = $voices.Item($i)
        $voiceName = $voiceToken.GetAttribute("Name")
        $voiceGender = $voiceToken.GetAttribute("Gender")
        $voiceAge = $voiceToken.GetAttribute("Age")
        Write-Host "   Voice $i`: $voiceName (Gender: $voiceGender, Age: $voiceAge)" -ForegroundColor White
    }

    # Find Amy voice
    $amyVoice = $null
    for ($i = 0; $i -lt $voices.Count; $i++) {
        $voiceToken = $voices.Item($i)
        $voiceName = $voiceToken.GetAttribute("Name")
        if ($voiceName -eq "amy") {
            $amyVoice = $voiceToken
            Write-Host "   Found Amy voice at index $i!" -ForegroundColor Green
            break
        }
    }

    if (-not $amyVoice) {
        Write-Host "   Amy voice not found!" -ForegroundColor Red
        return
    }

    Write-Host ""
    Write-Host "2. Testing voice selection..." -ForegroundColor Yellow
    
    # Check current voice
    $currentVoice = $voice.Voice
    Write-Host "   Current voice: $($currentVoice.GetAttribute('Name'))" -ForegroundColor Cyan
    
    # Try to set Amy voice
    Write-Host "   Setting Amy voice..." -ForegroundColor Cyan
    $voice.Voice = $amyVoice
    
    # Verify selection
    $newVoice = $voice.Voice
    Write-Host "   New voice: $($newVoice.GetAttribute('Name'))" -ForegroundColor Cyan
    
    if ($newVoice.GetAttribute('Name') -eq "amy") {
        Write-Host "   Amy voice selected successfully!" -ForegroundColor Green
    } else {
        Write-Host "   Failed to select Amy voice!" -ForegroundColor Red
        return
    }

    Write-Host ""
    Write-Host "3. Testing speech synthesis..." -ForegroundColor Yellow
    
    # Test with different speech flags
    $testTexts = @(
        @{ Text = "Hello world"; Flags = 0; Description = "Synchronous" },
        @{ Text = "This is a test"; Flags = 1; Description = "Asynchronous" },
        @{ Text = "Amy speaking"; Flags = 8; Description = "XML" }
    )
    
    foreach ($test in $testTexts) {
        Write-Host "   Testing: '$($test.Text)' ($($test.Description))..." -ForegroundColor Cyan
        
        try {
            $result = $voice.Speak($test.Text, $test.Flags)
            Write-Host "   Success! Speak returned: $result" -ForegroundColor Green
        } catch {
            Write-Host "   Failed: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "   HRESULT: $($_.Exception.HResult)" -ForegroundColor Red
        }
        
        # Check logs after each attempt
        if (Test-Path "$logDir\sapi_debug.log") {
            $logContent = Get-Content "$logDir\sapi_debug.log" -Tail 3
            if ($logContent) {
                Write-Host "   Recent log entries:" -ForegroundColor Gray
                $logContent | ForEach-Object { Write-Host "     $_" -ForegroundColor Gray }
            }
        }
        
        Write-Host ""
    }

    Write-Host "4. Testing voice properties..." -ForegroundColor Yellow
    
    # Test voice properties
    $properties = @("Name", "Gender", "Age", "Language", "Vendor", "Version")
    foreach ($prop in $properties) {
        try {
            $value = $amyVoice.GetAttribute($prop)
            Write-Host "   $prop`: $value" -ForegroundColor White
        } catch {
            Write-Host "   $prop`: (not available)" -ForegroundColor Gray
        }
    }

    Write-Host ""
    Write-Host "5. Testing output format..." -ForegroundColor Yellow
    
    # Try to get/set output format
    try {
        $audioFormat = $voice.AudioOutput
        if ($audioFormat) {
            Write-Host "   Current audio output: $($audioFormat.GetType().Name)" -ForegroundColor Cyan
        } else {
            Write-Host "   No audio output configured" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "   Error getting audio output: $($_.Exception.Message)" -ForegroundColor Red
    }

    Write-Host ""
    Write-Host "6. Final log check..." -ForegroundColor Yellow
    
    if (Test-Path "$logDir\sapi_debug.log") {
        Write-Host "   Complete SAPI Debug Log:" -ForegroundColor Cyan
        Get-Content "$logDir\sapi_debug.log" | ForEach-Object {
            Write-Host "     $_" -ForegroundColor Gray
        }
    } else {
        Write-Host "   No SAPI debug log found" -ForegroundColor Red
    }

    Write-Host ""
    Write-Host "SAPI INTEGRATION TEST COMPLETED!" -ForegroundColor Green
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "Summary: If SetObjectToken and Speak methods were called, SAPI integration works!" -ForegroundColor Yellow
