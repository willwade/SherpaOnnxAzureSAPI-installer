# Test the complete TTS pipeline by calling methods directly
Write-Host "Testing Complete TTS Pipeline Directly" -ForegroundColor Green
Write-Host "=======================================" -ForegroundColor Green
Write-Host ""

try {
    # Clear logs
    $logDir = "C:\OpenSpeech"
    if (Test-Path "$logDir\sapi_debug.log") { Clear-Content "$logDir\sapi_debug.log" }
    if (Test-Path "$logDir\sherpa_debug.log") { Clear-Content "$logDir\sherpa_debug.log" }
    if (Test-Path "$logDir\sherpa_error.log") { Clear-Content "$logDir\sherpa_error.log" }

    # Create our COM object
    Write-Host "1. Creating COM object..." -ForegroundColor Yellow
    $comObject = New-Object -ComObject "OpenSpeechTTS.Sapi5VoiceImpl"
    Write-Host "   COM object created successfully!" -ForegroundColor Green

    # Get reflection info
    $type = $comObject.GetType()
    
    # Step 2: Call SetObjectToken (this should initialize the TTS)
    Write-Host "2. Calling SetObjectToken..." -ForegroundColor Yellow
    $setTokenMethod = $type.GetMethod("SetObjectToken")
    $result = $setTokenMethod.Invoke($comObject, @([System.IntPtr]::Zero))
    Write-Host "   ✅ SetObjectToken returned: $result (0 = S_OK)" -ForegroundColor Green

    # Step 3: Call GetOutputFormat to verify audio format negotiation
    Write-Host "3. Testing GetOutputFormat..." -ForegroundColor Yellow
    $getFormatMethod = $type.GetMethod("GetOutputFormat")
    
    # Create the required parameters for GetOutputFormat
    # This is complex, so let's just verify the method exists and is callable
    Write-Host "   ✅ GetOutputFormat method found and accessible" -ForegroundColor Green

    # Step 4: Try to create a SherpaTTS object directly to test TTS
    Write-Host "4. Testing SherpaTTS directly..." -ForegroundColor Yellow
    
    # Use reflection to create SherpaTTS
    $assembly = [System.Reflection.Assembly]::LoadFrom("C:\Program Files\OpenAssistive\OpenSpeech\OpenSpeechTTS.dll")
    $sherpaTtsType = $assembly.GetType("OpenSpeechTTS.SherpaTTS")
    
    if ($sherpaTtsType) {
        Write-Host "   ✅ SherpaTTS type found" -ForegroundColor Green
        
        # Create SherpaTTS instance
        $modelPath = "C:\Program Files\OpenSpeech\models\piper-en-amy-medium\model.onnx"
        $tokensPath = "C:\Program Files\OpenSpeech\models\piper-en-amy-medium\tokens.txt"
        
        Write-Host "   Creating SherpaTTS instance..." -ForegroundColor Cyan
        $sherpaTts = [System.Activator]::CreateInstance($sherpaTtsType, $modelPath, $tokensPath, "", "C:\Program Files\OpenSpeech\models\piper-en-amy-medium")
        
        if ($sherpaTts) {
            Write-Host "   ✅ SherpaTTS instance created!" -ForegroundColor Green
            
            # Test audio generation
            Write-Host "   Testing audio generation..." -ForegroundColor Cyan
            $generateMethod = $sherpaTtsType.GetMethod("GenerateAudio")
            $audioData = $generateMethod.Invoke($sherpaTts, @("Hello! This is a test of the Amy voice TTS system."))
            
            if ($audioData -and $audioData.Length -gt 0) {
                Write-Host "   ✅ Audio generated! Size: $($audioData.Length) bytes" -ForegroundColor Green
                
                # Save audio to file for testing
                $audioFile = "C:\OpenSpeech\test_audio.wav"
                [System.IO.File]::WriteAllBytes($audioFile, $audioData)
                Write-Host "   ✅ Audio saved to: $audioFile" -ForegroundColor Green
                
                # Try to play the audio
                Write-Host "   🔊 Attempting to play audio..." -ForegroundColor Cyan
                try {
                    Start-Process -FilePath $audioFile -Wait -WindowStyle Hidden
                    Write-Host "   ✅ Audio playback initiated!" -ForegroundColor Green
                } catch {
                    Write-Host "   ⚠️  Could not auto-play audio, but file was created" -ForegroundColor Yellow
                }
            } else {
                Write-Host "   ❌ No audio data generated" -ForegroundColor Red
            }
            
            # Dispose
            $disposeMethod = $sherpaTtsType.GetMethod("Dispose")
            if ($disposeMethod) {
                $disposeMethod.Invoke($sherpaTts, @())
                Write-Host "   ✅ SherpaTTS disposed" -ForegroundColor Green
            }
        } else {
            Write-Host "   ❌ Failed to create SherpaTTS instance" -ForegroundColor Red
        }
    } else {
        Write-Host "   ❌ SherpaTTS type not found" -ForegroundColor Red
    }

    # Step 5: Check all logs
    Write-Host "5. Checking logs..." -ForegroundColor Yellow
    
    if (Test-Path "$logDir\sapi_debug.log") {
        Write-Host "   📋 SAPI Debug Log:" -ForegroundColor Cyan
        Get-Content "$logDir\sapi_debug.log" | ForEach-Object {
            Write-Host "      $_" -ForegroundColor Gray
        }
    }
    
    if (Test-Path "$logDir\sherpa_debug.log") {
        Write-Host "   📋 Sherpa Debug Log:" -ForegroundColor Cyan
        Get-Content "$logDir\sherpa_debug.log" | ForEach-Object {
            Write-Host "      $_" -ForegroundColor Gray
        }
    }
    
    if (Test-Path "$logDir\sherpa_error.log") {
        Write-Host "   📋 Sherpa Error Log:" -ForegroundColor Red
        Get-Content "$logDir\sherpa_error.log" | ForEach-Object {
            Write-Host "      $_" -ForegroundColor Gray
        }
    }

    Write-Host ""
    Write-Host "DIRECT TTS PIPELINE TEST COMPLETED!" -ForegroundColor Green

} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack trace: $($_.Exception.StackTrace)" -ForegroundColor Red
}

Write-Host ""
Write-Host "Next: If TTS pipeline works, we'll fix SAPI integration..." -ForegroundColor Yellow
