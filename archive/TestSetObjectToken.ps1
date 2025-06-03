# Test script to directly test SetObjectToken method
Write-Host "Testing SetObjectToken Method Directly" -ForegroundColor Green
Write-Host "======================================" -ForegroundColor Green
Write-Host ""

try {
    # Create our COM object directly
    Write-Host "Creating OpenSpeechTTS COM object..." -ForegroundColor Yellow
    $comObject = New-Object -ComObject "OpenSpeechTTS.Sapi5VoiceImpl"
    
    if ($comObject) {
        Write-Host "COM object created successfully!" -ForegroundColor Green
        
        # Try to call SetObjectToken with a null pointer (0)
        Write-Host "Attempting to call SetObjectToken..." -ForegroundColor Yellow
        
        try {
            # Get the type and find the SetObjectToken method
            $type = $comObject.GetType()
            $method = $type.GetMethod("SetObjectToken")
            
            if ($method) {
                Write-Host "Found SetObjectToken method!" -ForegroundColor Green
                
                # Call with IntPtr.Zero (null pointer)
                $result = $method.Invoke($comObject, @([System.IntPtr]::Zero))
                Write-Host "SetObjectToken returned: $result" -ForegroundColor Green
                
                # Check if logs were created
                if (Test-Path "C:\OpenSpeech\sapi_debug.log") {
                    Write-Host "Debug log created! Checking contents..." -ForegroundColor Green
                    Get-Content "C:\OpenSpeech\sapi_debug.log" -Tail 5 | ForEach-Object {
                        Write-Host "  $_" -ForegroundColor Gray
                    }
                } else {
                    Write-Host "No debug log found" -ForegroundColor Yellow
                }
                
            } else {
                Write-Host "SetObjectToken method not found!" -ForegroundColor Red
                Write-Host "Available methods:" -ForegroundColor Yellow
                $type.GetMethods() | ForEach-Object {
                    Write-Host "  $($_.Name)" -ForegroundColor Gray
                }
            }
            
        } catch {
            Write-Host "Error calling SetObjectToken: $($_.Exception.Message)" -ForegroundColor Red
        }
        
    } else {
        Write-Host "Failed to create COM object" -ForegroundColor Red
    }
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "SetObjectToken test completed!" -ForegroundColor Green
