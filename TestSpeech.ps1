# Test script to verify Amy voice works
Write-Host "Testing Amy Voice Speech Synthesis" -ForegroundColor Green
Write-Host "===================================" -ForegroundColor Green
Write-Host ""

try {
    Add-Type -AssemblyName System.Speech
    $synth = New-Object System.Speech.Synthesis.SpeechSynthesizer
    
    # List available voices
    $voices = $synth.GetInstalledVoices()
    Write-Host "Available voices:" -ForegroundColor Yellow
    foreach ($voice in $voices) {
        Write-Host "  - $($voice.VoiceInfo.Name)" -ForegroundColor Cyan
    }
    Write-Host ""
    
    # Test Amy voice
    Write-Host "Testing Amy voice..." -ForegroundColor Yellow
    $amyVoice = $voices | Where-Object { $_.VoiceInfo.Name -eq "amy" }
    
    if ($amyVoice) {
        Write-Host "Amy voice found! Testing speech..." -ForegroundColor Green
        
        # Select Amy voice
        $synth.SelectVoice("amy")
        
        # Set output to default audio device
        $synth.SetOutputToDefaultAudioDevice()
        
        # Test speech
        $testText = "Hello! This is Amy, a Sherpa ONNX voice speaking through the SAPI bridge. The installation appears to be working correctly."
        Write-Host "Speaking: '$testText'" -ForegroundColor Cyan
        
        $synth.Speak($testText)
        
        Write-Host "Speech test completed successfully!" -ForegroundColor Green
    } else {
        Write-Host "Amy voice not found!" -ForegroundColor Red
    }
    
} catch {
    Write-Host "Error during speech test: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack trace: $($_.Exception.StackTrace)" -ForegroundColor Red
}

Write-Host ""
Write-Host "Test completed!" -ForegroundColor Green
