# Test script to directly test SAPI COM interface
Write-Host "Testing Direct SAPI COM Interface" -ForegroundColor Green
Write-Host "==================================" -ForegroundColor Green
Write-Host ""

try {
    # Create SAPI voice object directly
    Write-Host "Creating SAPI voice object..." -ForegroundColor Yellow
    $voice = New-Object -ComObject SAPI.SpVoice
    
    Write-Host "Getting voice collection..." -ForegroundColor Yellow
    $voices = $voice.GetVoices()
    
    Write-Host "Found $($voices.Count) voices:" -ForegroundColor Cyan
    
    # Find Amy voice
    $amyVoice = $null
    for ($i = 0; $i -lt $voices.Count; $i++) {
        $voiceToken = $voices.Item($i)
        $voiceName = $voiceToken.GetAttribute("Name")
        Write-Host "  Voice $i`: $voiceName" -ForegroundColor White
        
        if ($voiceName -eq "amy") {
            $amyVoice = $voiceToken
            Write-Host "    -> Found Amy voice!" -ForegroundColor Green
        }
    }
    
    if ($amyVoice) {
        Write-Host ""
        Write-Host "Testing Amy voice selection..." -ForegroundColor Yellow
        
        # Try to set the voice
        try {
            $voice.Voice = $amyVoice
            Write-Host "Amy voice selected successfully!" -ForegroundColor Green
            
            # Try to speak
            Write-Host "Attempting to speak..." -ForegroundColor Yellow
            $voice.Speak("Hello, this is a test of the Amy voice using direct SAPI interface.", 0)
            Write-Host "Speech completed!" -ForegroundColor Green
            
        } catch {
            Write-Host "Error using Amy voice: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "HRESULT: $($_.Exception.HResult)" -ForegroundColor Red
        }
    } else {
        Write-Host "Amy voice not found!" -ForegroundColor Red
    }
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "Direct SAPI test completed!" -ForegroundColor Green
