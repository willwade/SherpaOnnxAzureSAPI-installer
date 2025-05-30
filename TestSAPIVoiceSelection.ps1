# Test SAPI Voice Selection with Detailed Debugging
Write-Host "Testing SAPI Voice Selection..." -ForegroundColor Green

try {
    Add-Type -AssemblyName System.Speech
    $synth = New-Object System.Speech.Synthesis.SpeechSynthesizer
    
    # List all available voices with detailed info
    Write-Host "`nDetailed Voice Information:" -ForegroundColor Yellow
    $voices = $synth.GetInstalledVoices()
    foreach ($voice in $voices) {
        $info = $voice.VoiceInfo
        Write-Host "Voice: $($info.Name)" -ForegroundColor Cyan
        Write-Host "  ID: $($info.Id)" -ForegroundColor White
        Write-Host "  Culture: $($info.Culture)" -ForegroundColor White
        Write-Host "  Gender: $($info.Gender)" -ForegroundColor White
        Write-Host "  Age: $($info.Age)" -ForegroundColor White
        Write-Host "  Enabled: $($voice.Enabled)" -ForegroundColor $(if ($voice.Enabled) { "Green" } else { "Red" })
        Write-Host "  Additional Info: $($info.AdditionalInfo | Out-String)" -ForegroundColor Gray
        Write-Host ""
    }
    
    # Try different methods to select Amy voice
    Write-Host "Testing different voice selection methods..." -ForegroundColor Yellow
    
    # Method 1: Direct name selection
    Write-Host "`n1. Testing direct name selection..." -ForegroundColor Cyan
    try {
        $synth.SelectVoice("amy")
        Write-Host "✅ Direct name selection successful!" -ForegroundColor Green
        Write-Host "Current voice: $($synth.Voice.Name)" -ForegroundColor White
    } catch {
        Write-Host "❌ Direct name selection failed: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # Method 2: Try with different casing
    Write-Host "`n2. Testing with different casing..." -ForegroundColor Cyan
    try {
        $synth.SelectVoice("Amy")
        Write-Host "✅ Capitalized name selection successful!" -ForegroundColor Green
        Write-Host "Current voice: $($synth.Voice.Name)" -ForegroundColor White
    } catch {
        Write-Host "❌ Capitalized name selection failed: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # Method 3: Select by voice info
    Write-Host "`n3. Testing selection by VoiceInfo..." -ForegroundColor Cyan
    try {
        $amyVoice = $voices | Where-Object { $_.VoiceInfo.Name -eq "amy" }
        if ($amyVoice) {
            $synth.SelectVoiceByHints($amyVoice.VoiceInfo.Gender, $amyVoice.VoiceInfo.Age, 0, $amyVoice.VoiceInfo.Culture)
            Write-Host "✅ Selection by hints successful!" -ForegroundColor Green
            Write-Host "Current voice: $($synth.Voice.Name)" -ForegroundColor White
        } else {
            Write-Host "❌ Amy voice not found in voice list" -ForegroundColor Red
        }
    } catch {
        Write-Host "❌ Selection by hints failed: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # Method 4: Try to speak with default voice first
    Write-Host "`n4. Testing speech with default voice..." -ForegroundColor Cyan
    try {
        Write-Host "Current default voice: $($synth.Voice.Name)" -ForegroundColor White
        $synth.SetOutputToDefaultAudioDevice()
        $synth.SpeakAsync("Testing default voice")
        Start-Sleep -Seconds 2
        Write-Host "✅ Default voice speech test completed" -ForegroundColor Green
    } catch {
        Write-Host "❌ Default voice speech failed: $($_.Exception.Message)" -ForegroundColor Red
    }
    
} catch {
    Write-Host "❌ Error during SAPI test: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack trace: $($_.Exception.StackTrace)" -ForegroundColor Red
}

Write-Host "`nChecking recent debug logs..." -ForegroundColor Yellow
if (Test-Path "C:\OpenSpeech\sapi_debug.log") {
    Write-Host "Recent log entries:" -ForegroundColor Cyan
    Get-Content "C:\OpenSpeech\sapi_debug.log" | Select-Object -Last 10
} else {
    Write-Host "No debug log found" -ForegroundColor Yellow
}

Write-Host "`nTest completed!" -ForegroundColor Green
