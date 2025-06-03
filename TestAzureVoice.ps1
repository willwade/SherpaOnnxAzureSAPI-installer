# Test Azure Voice
Write-Host "Testing Azure TTS voice..." -ForegroundColor Cyan

try {
    $voice = New-Object -ComObject SAPI.SpVoice
    $voices = $voice.GetVoices()
    
    Write-Host "Total voices: $($voices.Count)" -ForegroundColor Green
    
    $azureFound = $false
    for ($i = 0; $i -lt $voices.Count; $i++) {
        $voiceToken = $voices.Item($i)
        $voiceName = $voiceToken.GetDescription()
        
        Write-Host "Voice $($i + 1): $voiceName" -ForegroundColor Gray
        
        if ($voiceName -like "*Libby*" -or $voiceName -like "*Azure*") {
            Write-Host "Found Azure voice: $voiceName" -ForegroundColor Blue
            $azureFound = $true
            
            try {
                $voice.Voice = $voiceToken
                Write-Host "Azure voice set successfully!" -ForegroundColor Green
                
                Write-Host "Attempting synthesis..." -ForegroundColor Yellow
                $result = $voice.Speak("Hello! This is the Azure TTS voice through SAPI!", 0)  # Synchronous
                Write-Host "Azure synthesis result: $result" -ForegroundColor Green
                Write-Host "SUCCESS! Azure TTS SAPI integration is working!" -ForegroundColor Green
                
            } catch {
                Write-Host "Azure voice failed: $($_.Exception.Message)" -ForegroundColor Red
            }
            break
        }
    }
    
    if (-not $azureFound) {
        Write-Host "Azure voice not found" -ForegroundColor Yellow
    }
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "Test completed!" -ForegroundColor Cyan
