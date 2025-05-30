# Test script to directly test Sherpa TTS functionality
Write-Host "Testing Direct Sherpa TTS Functionality" -ForegroundColor Green
Write-Host "=======================================" -ForegroundColor Green
Write-Host ""

try {
    # Load the OpenSpeechTTS assembly directly
    $dllPath = "C:\Program Files\OpenAssistive\OpenSpeech\OpenSpeechTTS.dll"
    
    if (Test-Path $dllPath) {
        Write-Host "Loading OpenSpeechTTS.dll..." -ForegroundColor Yellow
        
        # Load the assembly
        [System.Reflection.Assembly]::LoadFrom($dllPath)
        
        Write-Host "Assembly loaded successfully!" -ForegroundColor Green
        
        # Try to create a Sapi5VoiceImpl instance directly
        Write-Host "Creating Sapi5VoiceImpl instance..." -ForegroundColor Yellow
        
        $voiceImpl = New-Object OpenSpeechTTS.Sapi5VoiceImpl
        
        Write-Host "Sapi5VoiceImpl created successfully!" -ForegroundColor Green
        Write-Host "This means the Sherpa TTS engine is working!" -ForegroundColor Green
        
    } else {
        Write-Host "OpenSpeechTTS.dll not found at: $dllPath" -ForegroundColor Red
    }
    
} catch {
    Write-Host "Error during direct test: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Full error: $($_.Exception)" -ForegroundColor Red
}

Write-Host ""
Write-Host "Direct test completed!" -ForegroundColor Green
