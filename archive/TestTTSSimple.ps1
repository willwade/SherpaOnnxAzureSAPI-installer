# Simple test of TTS pipeline
Write-Host "Testing TTS Pipeline" -ForegroundColor Green
Write-Host "====================" -ForegroundColor Green
Write-Host ""

try {
    # Clear logs
    $logDir = "C:\OpenSpeech"
    if (Test-Path "$logDir\sapi_debug.log") { Clear-Content "$logDir\sapi_debug.log" }
    if (Test-Path "$logDir\sherpa_debug.log") { Clear-Content "$logDir\sherpa_debug.log" }

    # Create COM object
    Write-Host "1. Creating COM object..." -ForegroundColor Yellow
    $comObject = New-Object -ComObject "OpenSpeechTTS.Sapi5VoiceImpl"
    Write-Host "   COM object created!" -ForegroundColor Green

    # Call SetObjectToken
    Write-Host "2. Calling SetObjectToken..." -ForegroundColor Yellow
    $type = $comObject.GetType()
    $setTokenMethod = $type.GetMethod("SetObjectToken")
    $result = $setTokenMethod.Invoke($comObject, @([System.IntPtr]::Zero))
    Write-Host "   SetObjectToken returned: $result" -ForegroundColor Green

    # Test SherpaTTS directly
    Write-Host "3. Testing SherpaTTS..." -ForegroundColor Yellow
    $assembly = [System.Reflection.Assembly]::LoadFrom("C:\Program Files\OpenAssistive\OpenSpeech\OpenSpeechTTS.dll")
    $sherpaTtsType = $assembly.GetType("OpenSpeechTTS.SherpaTTS")
    
    if ($sherpaTtsType) {
        Write-Host "   SherpaTTS type found!" -ForegroundColor Green
        
        $modelPath = "C:\Program Files\OpenSpeech\models\piper-en-amy-medium\model.onnx"
        $tokensPath = "C:\Program Files\OpenSpeech\models\piper-en-amy-medium\tokens.txt"
        
        Write-Host "   Creating SherpaTTS instance..." -ForegroundColor Cyan
        $sherpaTts = [System.Activator]::CreateInstance($sherpaTtsType, $modelPath, $tokensPath, "", "C:\Program Files\OpenSpeech\models\piper-en-amy-medium")
        
        if ($sherpaTts) {
            Write-Host "   SherpaTTS created!" -ForegroundColor Green
            
            Write-Host "   Generating audio..." -ForegroundColor Cyan
            $generateMethod = $sherpaTtsType.GetMethod("GenerateAudio")
            $audioData = $generateMethod.Invoke($sherpaTts, @("Hello! This is Amy speaking!"))
            
            if ($audioData -and $audioData.Length -gt 0) {
                Write-Host "   Audio generated! Size: $($audioData.Length) bytes" -ForegroundColor Green
                
                $audioFile = "C:\OpenSpeech\test_audio.wav"
                [System.IO.File]::WriteAllBytes($audioFile, $audioData)
                Write-Host "   Audio saved to: $audioFile" -ForegroundColor Green
                
                Write-Host "   Playing audio..." -ForegroundColor Cyan
                try {
                    Start-Process -FilePath $audioFile -WindowStyle Hidden
                    Write-Host "   Audio playback started!" -ForegroundColor Green
                } catch {
                    Write-Host "   Audio file created but could not auto-play" -ForegroundColor Yellow
                }
            } else {
                Write-Host "   No audio data generated" -ForegroundColor Red
            }
            
            # Dispose
            $disposeMethod = $sherpaTtsType.GetMethod("Dispose")
            if ($disposeMethod) {
                $disposeMethod.Invoke($sherpaTts, @())
            }
        } else {
            Write-Host "   Failed to create SherpaTTS" -ForegroundColor Red
        }
    } else {
        Write-Host "   SherpaTTS type not found" -ForegroundColor Red
    }

    # Check logs
    Write-Host "4. Checking logs..." -ForegroundColor Yellow
    
    if (Test-Path "$logDir\sapi_debug.log") {
        Write-Host "   SAPI Debug Log:" -ForegroundColor Cyan
        Get-Content "$logDir\sapi_debug.log" | ForEach-Object {
            Write-Host "      $_" -ForegroundColor Gray
        }
    }
    
    if (Test-Path "$logDir\sherpa_debug.log") {
        Write-Host "   Sherpa Debug Log:" -ForegroundColor Cyan
        Get-Content "$logDir\sherpa_debug.log" | ForEach-Object {
            Write-Host "      $_" -ForegroundColor Gray
        }
    }
    
    if (Test-Path "$logDir\sherpa_error.log") {
        Write-Host "   Sherpa Error Log:" -ForegroundColor Red
        Get-Content "$logDir\sherpa_error.log" | ForEach-Object {
            Write-Host "      $_" -ForegroundColor Gray
        }
    }

    Write-Host ""
    Write-Host "TTS PIPELINE TEST COMPLETED!" -ForegroundColor Green
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack trace: $($_.Exception.StackTrace)" -ForegroundColor Red
}

Write-Host ""
Write-Host "If audio was generated, the TTS pipeline works!" -ForegroundColor Yellow
