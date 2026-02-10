# List all SAPI5 voices
$voice = New-Object -ComObject SAPI.SpVoice
$voices = $voice.GetVoices()
Write-Host "Installed SAPI5 Voices:" -ForegroundColor Cyan
for ($i = 0; $i -lt $voices.Count; $i++) {
    $v = $voices.Item($i)
    Write-Host "  [$i] $($v.GetDescription())" -ForegroundColor White
}
