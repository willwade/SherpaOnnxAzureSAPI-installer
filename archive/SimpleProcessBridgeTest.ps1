# Simple test for ProcessBridge TTS
Write-Host "Testing ProcessBridge TTS" -ForegroundColor Green

# Clear logs
$logDir = "C:\OpenSpeech"
if (Test-Path "$logDir\sherpa_debug.log") { Clear-Content "$logDir\sherpa_debug.log" }

try {
    # Create COM object and call SetObjectToken
    $comObject = New-Object -ComObject "OpenSpeechTTS.Sapi5VoiceImpl"
    $type = $comObject.GetType()
    $setTokenMethod = $type.GetMethod("SetObjectToken")
    $result = $setTokenMethod.Invoke($comObject, @([System.IntPtr]::Zero))
    
    Write-Host "SetObjectToken result: $result" -ForegroundColor Green
    
    # Try to call SpeakToWaveStream
    $speakMethod = $type.GetMethod("SpeakToWaveStream")
    if ($speakMethod) {
        $memoryStream = New-Object System.IO.MemoryStream
        $speakMethod.Invoke($comObject, @("Test ProcessBridge", $memoryStream))
        Write-Host "Generated audio size: $($memoryStream.Length) bytes" -ForegroundColor Green
        $memoryStream.Dispose()
    }
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}

# Check logs
if (Test-Path "$logDir\sherpa_debug.log") {
    Write-Host "`nDebug Log:" -ForegroundColor Yellow
    Get-Content "$logDir\sherpa_debug.log" | ForEach-Object { Write-Host "  $_" }
} else {
    Write-Host "`nNo debug log found" -ForegroundColor Red
}
