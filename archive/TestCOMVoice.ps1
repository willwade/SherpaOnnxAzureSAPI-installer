# Test script to directly test COM voice functionality
Write-Host "Testing COM Voice Functionality" -ForegroundColor Green
Write-Host "================================" -ForegroundColor Green
Write-Host ""

try {
    # Try to create the COM object directly
    Write-Host "Creating COM object directly..." -ForegroundColor Yellow
    
    $comObject = New-Object -ComObject "OpenSpeechTTS.Sapi5VoiceImpl"
    
    if ($comObject) {
        Write-Host "COM object created successfully!" -ForegroundColor Green
        Write-Host "This means the COM registration is working!" -ForegroundColor Green
        
        # Try to call a method if available
        Write-Host "COM object type: $($comObject.GetType().FullName)" -ForegroundColor Cyan
        
        # Release the COM object
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) | Out-Null
        $comObject = $null
        
    } else {
        Write-Host "Failed to create COM object" -ForegroundColor Red
    }
    
} catch {
    Write-Host "Error creating COM object: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Full error: $($_.Exception)" -ForegroundColor Red
}

Write-Host ""

# Test SAPI voice enumeration in more detail
try {
    Write-Host "Testing detailed SAPI voice enumeration..." -ForegroundColor Yellow
    
    $synth = New-Object -ComObject SAPI.SpVoice
    $voices = $synth.GetVoices()
    
    Write-Host "Found $($voices.Count) voices:" -ForegroundColor Cyan
    
    for ($i = 0; $i -lt $voices.Count; $i++) {
        $voice = $voices.Item($i)
        $voiceInfo = $voice.GetDescription()
        $attributes = $voice.GetAttribute("Name")
        
        Write-Host "  Voice $i`: $voiceInfo (Name: $attributes)" -ForegroundColor White
        
        # Check if this is our Amy voice
        if ($attributes -eq "amy") {
            Write-Host "    -> This is our Amy voice! Attempting to select it..." -ForegroundColor Green
            
            try {
                $synth.Voice = $voice
                Write-Host "    -> Successfully selected Amy voice!" -ForegroundColor Green
                
                # Try to speak something
                Write-Host "    -> Attempting to speak test phrase..." -ForegroundColor Yellow
                $synth.Speak("Hello, this is a test of the Amy voice.", 1) # 1 = SVSFlagsAsync
                Write-Host "    -> Speech command sent successfully!" -ForegroundColor Green
                
            } catch {
                Write-Host "    -> Failed to select or use Amy voice: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
    
} catch {
    Write-Host "Error during SAPI enumeration: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "COM voice test completed!" -ForegroundColor Green
