# Test Direct Interface Calls
Write-Host "Testing Direct Interface Method Calls..." -ForegroundColor Green

try {
    # Create COM object directly
    Write-Host "Creating COM object..." -ForegroundColor Yellow
    $comObject = New-Object -ComObject "OpenSpeechTTS.Sapi5VoiceImpl"
    
    if ($comObject) {
        Write-Host "COM object created successfully!" -ForegroundColor Green
        
        # Try to call SetObjectToken with a real voice token
        Write-Host "Testing SetObjectToken with voice token..." -ForegroundColor Yellow
        
        # Create a SAPI voice object to get a real token
        Add-Type -AssemblyName System.Speech
        $synth = New-Object System.Speech.Synthesis.SpeechSynthesizer
        
        # Get the default voice token (this should work)
        $defaultVoice = $synth.Voice
        Write-Host "Default voice: $($defaultVoice.Name)" -ForegroundColor Cyan
        
        # Try to speak with default voice first to ensure SAPI is working
        Write-Host "Testing default voice..." -ForegroundColor Yellow
        $synth.SetOutputToDefaultAudioDevice()
        $synth.SpeakAsync("Testing default voice")
        Start-Sleep -Seconds 2
        
        Write-Host "Default voice test completed" -ForegroundColor Green
        
        # Now try to manually select Amy voice using a different approach
        Write-Host "Attempting to select Amy voice using SelectVoiceByHints..." -ForegroundColor Yellow
        
        try {
            # Try to select by gender and language
            $synth.SelectVoiceByHints([System.Speech.Synthesis.VoiceGender]::Female, [System.Speech.Synthesis.VoiceAge]::Adult, 0, [System.Globalization.CultureInfo]::GetCultureInfo("en-US"))
            Write-Host "Voice selection by hints: $($synth.Voice.Name)" -ForegroundColor Green
            
            # Test speech with selected voice
            $synth.SpeakAsync("Testing voice selection by hints")
            Start-Sleep -Seconds 3
            
        } catch {
            Write-Host "SelectVoiceByHints failed: $($_.Exception.Message)" -ForegroundColor Red
        }
        
    } else {
        Write-Host "Failed to create COM object" -ForegroundColor Red
    }
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack trace: $($_.Exception.StackTrace)" -ForegroundColor Red
}

# Check logs for any new activity
Write-Host "Checking debug logs..." -ForegroundColor Yellow
if (Test-Path "C:\OpenSpeech\sapi_debug.log") {
    Write-Host "Recent log entries:" -ForegroundColor Cyan
    Get-Content "C:\OpenSpeech\sapi_debug.log" | Select-Object -Last 15
} else {
    Write-Host "No debug log found" -ForegroundColor Red
}

Write-Host "Test completed!" -ForegroundColor Green
