# Test the full TTS pipeline by calling methods directly
Write-Host "Testing Full TTS Pipeline" -ForegroundColor Green
Write-Host "=========================" -ForegroundColor Green
Write-Host ""

try {
    # Clear any existing logs
    if (Test-Path "C:\OpenSpeech\sapi_debug.log") {
        Clear-Content "C:\OpenSpeech\sapi_debug.log"
    }
    if (Test-Path "C:\OpenSpeech\sherpa_debug.log") {
        Clear-Content "C:\OpenSpeech\sherpa_debug.log"
    }
    
    # Create our COM object directly
    Write-Host "Creating OpenSpeechTTS COM object..." -ForegroundColor Yellow
    $comObject = New-Object -ComObject "OpenSpeechTTS.Sapi5VoiceImpl"
    
    if ($comObject) {
        Write-Host "COM object created successfully!" -ForegroundColor Green
        
        # DON'T call SetObjectToken - let the Speak method do auto-initialization
        Write-Host "Skipping SetObjectToken to test auto-initialization..." -ForegroundColor Yellow
        
        # Try to create a simple SAPI voice and use our object
        Write-Host "Testing with SAPI voice object..." -ForegroundColor Yellow
        
        $voice = New-Object -ComObject SAPI.SpVoice
        $voices = $voice.GetVoices()
        
        # Find Amy voice
        $amyVoice = $null
        for ($i = 0; $i -lt $voices.Count; $i++) {
            $voiceToken = $voices.Item($i)
            $voiceName = $voiceToken.GetAttribute("Name")
            if ($voiceName -eq "amy") {
                $amyVoice = $voiceToken
                break
            }
        }
        
        if ($amyVoice) {
            Write-Host "Found Amy voice in SAPI!" -ForegroundColor Green
            
            # Try to set the voice and speak
            try {
                Write-Host "Setting Amy voice..." -ForegroundColor Yellow
                $voice.Voice = $amyVoice
                Write-Host "Amy voice set successfully!" -ForegroundColor Green
                
                Write-Host "Attempting to speak..." -ForegroundColor Yellow
                $voice.Speak("Hello! This is a test of the Amy voice. If you hear this, the TTS bridge is working!", 0)
                Write-Host "Speech command completed!" -ForegroundColor Green
                
            } catch {
                Write-Host "Error during speech: $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "HRESULT: $($_.Exception.HResult)" -ForegroundColor Red
            }
        } else {
            Write-Host "Amy voice not found in SAPI!" -ForegroundColor Red
        }
        
        # Check logs
        Write-Host ""
        Write-Host "Checking logs..." -ForegroundColor Yellow
        
        if (Test-Path "C:\OpenSpeech\sapi_debug.log") {
            Write-Host "SAPI Debug Log:" -ForegroundColor Cyan
            Get-Content "C:\OpenSpeech\sapi_debug.log" | ForEach-Object {
                Write-Host "  $_" -ForegroundColor Gray
            }
        }
        
        if (Test-Path "C:\OpenSpeech\sherpa_debug.log") {
            Write-Host ""
            Write-Host "Sherpa Debug Log:" -ForegroundColor Cyan
            Get-Content "C:\OpenSpeech\sherpa_debug.log" | ForEach-Object {
                Write-Host "  $_" -ForegroundColor Gray
            }
        }
        
        if (Test-Path "C:\OpenSpeech\sherpa_error.log") {
            Write-Host ""
            Write-Host "Sherpa Error Log:" -ForegroundColor Red
            Get-Content "C:\OpenSpeech\sherpa_error.log" | ForEach-Object {
                Write-Host "  $_" -ForegroundColor Gray
            }
        }
        
    } else {
        Write-Host "Failed to create COM object" -ForegroundColor Red
    }
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack trace: $($_.Exception.StackTrace)" -ForegroundColor Red
}

Write-Host ""
Write-Host "Full pipeline test completed!" -ForegroundColor Green
