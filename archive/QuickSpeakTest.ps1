# Quick test to see if Speak method is called
Clear-Content "C:\OpenSpeech\sapi_debug.log" -ErrorAction SilentlyContinue

$voice = New-Object -ComObject SAPI.SpVoice
$voices = $voice.GetVoices()
$amy = $null

for($i=0; $i -lt $voices.Count; $i++) {
    if($voices.Item($i).GetDescription() -eq "amy") {
        $amy = $voices.Item($i)
        break
    }
}

if($amy) {
    Write-Host "Found Amy voice, setting it..." -ForegroundColor Green
    $voice.Voice = $amy
    
    Write-Host "Calling voice.Speak()..." -ForegroundColor Yellow
    try {
        $voice.Speak("Test message")
        Write-Host "Speak completed" -ForegroundColor Green
    } catch {
        Write-Host "Speak failed: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "Checking if Speak method was called..." -ForegroundColor Cyan
    if(Test-Path "C:\OpenSpeech\sapi_debug.log") {
        $speakCalls = Get-Content "C:\OpenSpeech\sapi_debug.log" | Where-Object { $_ -like "*SPEAK METHOD CALLED*" }
        if($speakCalls) {
            Write-Host "✅ SPEAK METHOD WAS CALLED!" -ForegroundColor Green
            $speakCalls | ForEach-Object { Write-Host "  $_" -ForegroundColor White }
        } else {
            Write-Host "❌ SPEAK METHOD WAS NOT CALLED" -ForegroundColor Red
        }
    } else {
        Write-Host "❌ No debug log found" -ForegroundColor Red
    }
} else {
    Write-Host "❌ Amy voice not found" -ForegroundColor Red
}
