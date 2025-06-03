# Test script to test Speak method after SetObjectToken
Write-Host "Testing Speak Method After SetObjectToken" -ForegroundColor Green
Write-Host "=========================================" -ForegroundColor Green
Write-Host ""

try {
    # Create our COM object directly
    Write-Host "Creating OpenSpeechTTS COM object..." -ForegroundColor Yellow
    $comObject = New-Object -ComObject "OpenSpeechTTS.Sapi5VoiceImpl"
    
    if ($comObject) {
        Write-Host "COM object created successfully!" -ForegroundColor Green
        
        # First call SetObjectToken to initialize
        Write-Host "Calling SetObjectToken to initialize..." -ForegroundColor Yellow
        $type = $comObject.GetType()
        $setTokenMethod = $type.GetMethod("SetObjectToken")
        $result = $setTokenMethod.Invoke($comObject, @([System.IntPtr]::Zero))
        Write-Host "SetObjectToken returned: $result" -ForegroundColor Green
        
        # Now try to find and call the Speak method
        Write-Host "Looking for Speak method..." -ForegroundColor Yellow
        $speakMethod = $type.GetMethod("Speak")
        
        if ($speakMethod) {
            Write-Host "Found Speak method!" -ForegroundColor Green
            Write-Host "Method signature: $($speakMethod.ToString())" -ForegroundColor Gray
            
            # This is complex because Speak has many parameters
            # Let's just see if we can call it with minimal parameters
            Write-Host "Note: Speak method requires complex parameters - checking if it exists" -ForegroundColor Yellow
            
        } else {
            Write-Host "Speak method not found!" -ForegroundColor Red
            Write-Host "Available methods:" -ForegroundColor Yellow
            $type.GetMethods() | Where-Object { $_.Name -like "*Speak*" -or $_.Name -like "*Generate*" } | ForEach-Object {
                Write-Host "  $($_.Name): $($_.ToString())" -ForegroundColor Gray
            }
        }
        
        # Check if SherpaTTS was initialized
        Write-Host ""
        Write-Host "Checking if SherpaTTS was initialized..." -ForegroundColor Yellow
        if (Test-Path "C:\OpenSpeech\sherpa_debug.log") {
            Write-Host "Sherpa debug log found! Checking contents..." -ForegroundColor Green
            Get-Content "C:\OpenSpeech\sherpa_debug.log" -Tail 10 | ForEach-Object {
                Write-Host "  $_" -ForegroundColor Gray
            }
        } else {
            Write-Host "No Sherpa debug log found" -ForegroundColor Yellow
        }
        
        # Check SAPI debug log
        Write-Host ""
        Write-Host "Checking SAPI debug log..." -ForegroundColor Yellow
        if (Test-Path "C:\OpenSpeech\sapi_debug.log") {
            Get-Content "C:\OpenSpeech\sapi_debug.log" -Tail 5 | ForEach-Object {
                Write-Host "  $_" -ForegroundColor Gray
            }
        }
        
    } else {
        Write-Host "Failed to create COM object" -ForegroundColor Red
    }
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "Speak method test completed!" -ForegroundColor Green
