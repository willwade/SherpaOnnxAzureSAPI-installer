# Simple TTS Test
Write-Host "Testing TTS Implementation..." -ForegroundColor Green

try {
    # Test SAPI functionality
    Add-Type -AssemblyName System.Speech
    $synth = New-Object System.Speech.Synthesis.SpeechSynthesizer
    
    # List available voices
    $voices = $synth.GetInstalledVoices()
    Write-Host "Available voices:" -ForegroundColor Cyan
    foreach ($voice in $voices) {
        Write-Host "  - $($voice.VoiceInfo.Name)" -ForegroundColor White
    }
    
    # Test Amy voice
    $amyVoice = $voices | Where-Object { $_.VoiceInfo.Name -eq "amy" }
    if ($amyVoice -and $amyVoice.Enabled) {
        Write-Host "Testing Amy voice..." -ForegroundColor Yellow
        
        $synth.SelectVoice("amy")
        $synth.SetOutputToDefaultAudioDevice()
        
        $testText = "Hello! This is Amy testing the updated TTS implementation."
        Write-Host "Speaking: $testText" -ForegroundColor Cyan
        
        $synth.Speak($testText)
        
        Write-Host "Speech test completed!" -ForegroundColor Green
    } else {
        Write-Host "Amy voice not available" -ForegroundColor Red
    }
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}

# Check logs
Write-Host "Checking debug logs..." -ForegroundColor Yellow
if (Test-Path "C:\OpenSpeech\sapi_debug.log") {
    Write-Host "Recent log entries:" -ForegroundColor Cyan
    Get-Content "C:\OpenSpeech\sapi_debug.log" | Select-Object -Last 10
} else {
    Write-Host "No debug log found" -ForegroundColor Red
}

Write-Host "Test completed!" -ForegroundColor Green
