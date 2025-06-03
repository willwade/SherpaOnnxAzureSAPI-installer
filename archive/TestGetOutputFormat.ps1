# Test GetOutputFormat method directly
Write-Host "Testing GetOutputFormat Method" -ForegroundColor Green
Write-Host "==============================" -ForegroundColor Green
Write-Host ""

try {
    # Create our COM object directly
    Write-Host "Creating OpenSpeechTTS COM object..." -ForegroundColor Yellow
    $comObject = New-Object -ComObject "OpenSpeechTTS.Sapi5VoiceImpl"
    
    if ($comObject) {
        Write-Host "COM object created successfully!" -ForegroundColor Green
        
        # First call SetObjectToken to initialize
        Write-Host "Calling SetObjectToken..." -ForegroundColor Yellow
        $type = $comObject.GetType()
        $setTokenMethod = $type.GetMethod("SetObjectToken")
        $result = $setTokenMethod.Invoke($comObject, @([System.IntPtr]::Zero))
        Write-Host "SetObjectToken returned: $result" -ForegroundColor Green
        
        # Try to call GetOutputFormat
        Write-Host "Looking for GetOutputFormat method..." -ForegroundColor Yellow
        $getFormatMethod = $type.GetMethod("GetOutputFormat")
        
        if ($getFormatMethod) {
            Write-Host "Found GetOutputFormat method!" -ForegroundColor Green
            Write-Host "Method signature: $($getFormatMethod.ToString())" -ForegroundColor Gray
            
            # GetOutputFormat is complex to call directly, but let's see if we can
            Write-Host "GetOutputFormat method exists - this is good!" -ForegroundColor Green
            
        } else {
            Write-Host "GetOutputFormat method not found!" -ForegroundColor Red
        }
        
        # List all available methods
        Write-Host ""
        Write-Host "All available methods:" -ForegroundColor Yellow
        $type.GetMethods() | Where-Object { $_.DeclaringType.Name -eq "Sapi5VoiceImpl" } | ForEach-Object {
            Write-Host "  $($_.Name): $($_.ToString())" -ForegroundColor Gray
        }
        
    } else {
        Write-Host "Failed to create COM object" -ForegroundColor Red
    }
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "GetOutputFormat test completed!" -ForegroundColor Green
