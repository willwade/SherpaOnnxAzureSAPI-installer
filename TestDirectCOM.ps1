# Test Direct COM Object Creation and Interface Calls
Write-Host "Testing Direct COM Object Creation..." -ForegroundColor Green

try {
    # Create COM object directly
    Write-Host "Creating COM object directly..." -ForegroundColor Yellow
    $comObject = New-Object -ComObject "OpenSpeechTTS.Sapi5VoiceImpl"
    
    if ($comObject) {
        Write-Host "✅ COM object created successfully!" -ForegroundColor Green
        Write-Host "Object type: $($comObject.GetType().FullName)" -ForegroundColor Cyan
        
        # Try to call SetObjectToken with a dummy token
        Write-Host "`nTesting SetObjectToken method..." -ForegroundColor Yellow
        try {
            # Call SetObjectToken with IntPtr.Zero (null pointer)
            $result = $comObject.SetObjectToken([System.IntPtr]::Zero)
            Write-Host "✅ SetObjectToken called successfully! Result: $result" -ForegroundColor Green
        } catch {
            Write-Host "❌ SetObjectToken failed: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        # Try to call GetObjectToken
        Write-Host "`nTesting GetObjectToken method..." -ForegroundColor Yellow
        try {
            $token = [System.IntPtr]::Zero
            $result = $comObject.GetObjectToken([ref]$token)
            Write-Host "✅ GetObjectToken called successfully! Result: $result, Token: $token" -ForegroundColor Green
        } catch {
            Write-Host "❌ GetObjectToken failed: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        Write-Host "`nCOM object test completed!" -ForegroundColor Green
    } else {
        Write-Host "❌ Failed to create COM object" -ForegroundColor Red
    }
    
} catch {
    Write-Host "❌ Error during COM object test: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack trace: $($_.Exception.StackTrace)" -ForegroundColor Red
}

Write-Host "`nChecking debug logs..." -ForegroundColor Yellow
if (Test-Path "C:\OpenSpeech\sapi_debug.log") {
    Write-Host "Recent log entries:" -ForegroundColor Cyan
    Get-Content "C:\OpenSpeech\sapi_debug.log" | Select-Object -Last 5
} else {
    Write-Host "No debug log found at C:\OpenSpeech\sapi_debug.log" -ForegroundColor Yellow
}

Write-Host "`nTest completed!" -ForegroundColor Green
