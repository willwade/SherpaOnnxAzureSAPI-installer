# Test if Speak method is actually being called by SAPI
Write-Host "Testing if SAPI calls our Speak method" -ForegroundColor Green
Write-Host "=======================================" -ForegroundColor Green
Write-Host ""

# Clear logs
$logDir = "C:\OpenSpeech"
if (Test-Path "$logDir\sapi_debug.log") { Clear-Content "$logDir\sapi_debug.log" }
if (Test-Path "$logDir\sherpa_debug.log") { Clear-Content "$logDir\sherpa_debug.log" }

Write-Host "1. Testing SAPI voice.Speak() call..." -ForegroundColor Cyan

try {
    # Create SAPI voice
    $voice = New-Object -ComObject SAPI.SpVoice
    
    # Find Amy voice
    $voices = $voice.GetVoices()
    $amyVoice = $null
    for ($i = 0; $i -lt $voices.Count; $i++) {
        $voiceItem = $voices.Item($i)
        if ($voiceItem.GetDescription() -eq "amy") {
            $amyVoice = $voiceItem
            break
        }
    }
    
    if ($amyVoice) {
        Write-Host "   Found Amy voice" -ForegroundColor Green
        
        # Set Amy voice
        $voice.Voice = $amyVoice
        Write-Host "   Amy voice selected" -ForegroundColor Green
        
        Write-Host "   About to call voice.Speak() - watch for SPEAK METHOD CALLED in logs..." -ForegroundColor Yellow
        
        try {
            # This should trigger our Speak method
            $result = $voice.Speak("Test message for Amy voice")
            Write-Host "   ✅ voice.Speak() returned: $result" -ForegroundColor Green
        } catch {
            Write-Host "   ❌ voice.Speak() failed: $($_.Exception.Message)" -ForegroundColor Red
        }
    } else {
        Write-Host "   ❌ Amy voice not found" -ForegroundColor Red
    }
    
} catch {
    Write-Host "❌ Error: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "2. Checking logs for method calls..." -ForegroundColor Cyan

if (Test-Path "$logDir\sapi_debug.log") {
    $logContent = Get-Content "$logDir\sapi_debug.log"
    
    $speakCalled = $false
    $getOutputFormatCalled = $false
    
    foreach ($line in $logContent) {
        if ($line -like "*SPEAK METHOD CALLED*") {
            Write-Host "   ✅ SPEAK METHOD WAS CALLED!" -ForegroundColor Green
            Write-Host "     $line" -ForegroundColor White
            $speakCalled = $true
        }
        elseif ($line -like "*GET OUTPUT FORMAT CALLED*") {
            Write-Host "   ✅ GET OUTPUT FORMAT WAS CALLED!" -ForegroundColor Green
            Write-Host "     $line" -ForegroundColor White
            $getOutputFormatCalled = $true
        }
        elseif ($line -like "*auto-initialization*") {
            Write-Host "   ℹ️ Auto-initialization: $line" -ForegroundColor Yellow
        }
        elseif ($line -like "*ERROR*") {
            Write-Host "   ❌ Error: $line" -ForegroundColor Red
        }
    }
    
    if (-not $speakCalled) {
        Write-Host "   ❌ SPEAK METHOD WAS NOT CALLED!" -ForegroundColor Red
        Write-Host "   This means SAPI is not calling our Speak method at all." -ForegroundColor Red
    }
    
    if (-not $getOutputFormatCalled) {
        Write-Host "   ❌ GET OUTPUT FORMAT WAS NOT CALLED!" -ForegroundColor Red
        Write-Host "   This means SAPI is not calling our GetOutputFormat method." -ForegroundColor Red
    }
    
} else {
    Write-Host "   ❌ No debug log found" -ForegroundColor Red
}

Write-Host ""
Write-Host "=== ANALYSIS ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "If SPEAK METHOD CALLED appears above:" -ForegroundColor Yellow
Write-Host "  ✅ SAPI is calling our methods - the issue is in our implementation" -ForegroundColor White
Write-Host ""
Write-Host "If SPEAK METHOD CALLED does NOT appear above:" -ForegroundColor Yellow
Write-Host "  ❌ SAPI is not calling our methods - the issue is in interface registration" -ForegroundColor White
Write-Host ""
