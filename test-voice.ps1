# Test SAPI5 Voice
$voice = New-Object -ComObject SAPI.SpVoice

# List all voices first
Write-Host "Available voices:" -ForegroundColor Cyan
$voices = $voice.GetVoices()
for ($i = 0; $i -lt $voices.Count; $i++) {
    $v = $voices.Item($i)
    $desc = $v.GetDescription()
    if ($desc -like "*Sherpa*") {
        Write-Host "  [$i] $desc <-- SherpaOnnx voice" -ForegroundColor Green
    } else {
        Write-Host "  [$i] $desc" -ForegroundColor Gray
    }
}

# Try to find and select the Sherpa voice
$sherpaVoiceIndex = -1
for ($i = 0; $i -lt $voices.Count; $i++) {
    $desc = $voices.Item($i).GetDescription()
    if ($desc -like "*Sherpa*") {
        $sherpaVoiceIndex = $i
        break
    }
}

if ($sherpaVoiceIndex -ge 0) {
    Write-Host "`nSelecting Sherpa voice at index $sherpaVoiceIndex..." -ForegroundColor Yellow
    $voice.Voice = $voices.Item($sherpaVoiceIndex)
    Write-Host "Selected: $($voice.Voice.GetDescription())" -ForegroundColor Green
    Write-Host "`nSpeaking..." -ForegroundColor Yellow
    $voice.Speak("This is a test of the SherpaOnnx voice engine")
    Write-Host "Done!" -ForegroundColor Green
} else {
    Write-Host "`nERROR: No SherpaOnnx voice found!" -ForegroundColor Red
}
