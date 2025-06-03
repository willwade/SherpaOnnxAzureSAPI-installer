# Test ElliotNeural Voice
Write-Host "Checking for ElliotNeural voice..." -ForegroundColor Cyan

try {
    $voice = New-Object -ComObject SAPI.SpVoice
    $voices = $voice.GetVoices()
    
    Write-Host "Total voices: $($voices.Count)" -ForegroundColor Green
    Write-Host ""
    
    $elliotFound = $false
    for ($i = 0; $i -lt $voices.Count; $i++) {
        $voiceToken = $voices.Item($i)
        $voiceName = $voiceToken.GetDescription()
        
        Write-Host "Voice $($i + 1): $voiceName" -ForegroundColor Yellow
        
        if ($voiceName -like "*Elliot*") {
            Write-Host "  -> FOUND ElliotNeural!" -ForegroundColor Green
            $elliotFound = $true
            
            # Test the voice
            try {
                $voice.Voice = $voiceToken
                Write-Host "  -> Voice set successfully" -ForegroundColor Green
                
                $result = $voice.Speak("Hello from Elliot Neural voice!", 1)
                Write-Host "  -> Synthesis result: $result" -ForegroundColor Green
                
            } catch {
                Write-Host "  -> Synthesis failed: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
    
    if (-not $elliotFound) {
        Write-Host "ElliotNeural voice NOT found in SAPI!" -ForegroundColor Red
        Write-Host "This suggests a registration issue." -ForegroundColor Yellow
    }
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "Test completed!" -ForegroundColor Cyan
