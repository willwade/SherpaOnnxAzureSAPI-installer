# Test Direct Amy Voice Selection
Write-Host "Testing Direct Amy Voice Selection..." -ForegroundColor Green

try {
    Add-Type -AssemblyName System.Speech
    $synth = New-Object System.Speech.Synthesis.SpeechSynthesizer
    
    # List all voices with detailed info
    Write-Host "All available voices:" -ForegroundColor Cyan
    $voices = $synth.GetInstalledVoices()
    foreach ($voice in $voices) {
        $status = if ($voice.Enabled) { "ENABLED" } else { "DISABLED" }
        Write-Host "  - $($voice.VoiceInfo.Name) ($($voice.VoiceInfo.Gender), $status)" -ForegroundColor White
    }
    
    Write-Host "`nTesting direct Amy selection..." -ForegroundColor Yellow
    
    # Method 1: Try exact name match
    try {
        Write-Host "Attempting SelectVoice('amy')..." -ForegroundColor Cyan
        $synth.SelectVoice("amy")
        Write-Host "SUCCESS! Selected voice: $($synth.Voice.Name)" -ForegroundColor Green
        
        # Test speech with Amy
        $synth.SetOutputToDefaultAudioDevice()
        $testText = "Hello! This is Amy speaking directly. The SAPI bridge is working!"
        Write-Host "Speaking with Amy: '$testText'" -ForegroundColor Yellow
        $synth.Speak($testText)
        Write-Host "Amy speech test completed!" -ForegroundColor Green
        
    } catch {
        Write-Host "Direct selection failed: $($_.Exception.Message)" -ForegroundColor Red
        
        # Method 2: Try to find Amy in the voice list and use its exact ID
        Write-Host "`nTrying alternative selection method..." -ForegroundColor Yellow
        $amyVoice = $voices | Where-Object { $_.VoiceInfo.Name -eq "amy" }
        if ($amyVoice) {
            Write-Host "Found Amy voice in list: $($amyVoice.VoiceInfo.Id)" -ForegroundColor Cyan
            Write-Host "Amy voice enabled: $($amyVoice.Enabled)" -ForegroundColor White
            
            if (-not $amyVoice.Enabled) {
                Write-Host "Amy voice is DISABLED - this explains the selection failure!" -ForegroundColor Red
            }
        } else {
            Write-Host "Amy voice not found in voice list!" -ForegroundColor Red
        }
    }
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}

# Check logs for any SetObjectToken calls
Write-Host "`nChecking for SAPI method calls in logs..." -ForegroundColor Yellow
if (Test-Path "C:\OpenSpeech\sapi_debug.log") {
    $logContent = Get-Content "C:\OpenSpeech\sapi_debug.log"
    $recentLogs = $logContent | Select-Object -Last 20
    
    $setTokenCalls = $recentLogs | Where-Object { $_ -like "*SET OBJECT TOKEN*" }
    $speakCalls = $recentLogs | Where-Object { $_ -like "*SPEAK METHOD*" }
    
    if ($setTokenCalls) {
        Write-Host "✅ SetObjectToken calls found:" -ForegroundColor Green
        $setTokenCalls | ForEach-Object { Write-Host "  $_" -ForegroundColor White }
    } else {
        Write-Host "❌ No SetObjectToken calls found" -ForegroundColor Red
    }
    
    if ($speakCalls) {
        Write-Host "✅ Speak method calls found:" -ForegroundColor Green
        $speakCalls | ForEach-Object { Write-Host "  $_" -ForegroundColor White }
    } else {
        Write-Host "❌ No Speak method calls found" -ForegroundColor Red
    }
    
    Write-Host "`nRecent log entries:" -ForegroundColor Cyan
    $recentLogs | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }
    
} else {
    Write-Host "No debug log found" -ForegroundColor Red
}

Write-Host "`nTest completed!" -ForegroundColor Green
