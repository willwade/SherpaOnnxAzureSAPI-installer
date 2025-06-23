# OpenSpeechSAPI - Test SAPI Voice Installation
# Usage: .\test-voice.ps1 [VoiceName]

param(
    [string]$VoiceName = ""
)

Write-Host "=== SAPI Voice Test ===" -ForegroundColor Cyan

# List all SAPI voices
Write-Host "`n📋 Available SAPI Voices:" -ForegroundColor Yellow
try {
    $voice = New-Object -ComObject SAPI.SpVoice
    $voices = $voice.GetVoices()
    
    for ($i = 0; $i -lt $voices.Count; $i++) {
        $voiceInfo = $voices.Item($i)
        $name = $voiceInfo.GetDescription()
        Write-Host "  $($i + 1). $name" -ForegroundColor Gray
    }
    
    if ($voices.Count -eq 0) {
        Write-Host "  No SAPI voices found!" -ForegroundColor Red
    }
} catch {
    Write-Host "❌ Error listing voices: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Test specific voice if provided
if ($VoiceName) {
    Write-Host "`n🎤 Testing Voice: $VoiceName" -ForegroundColor Cyan
    
    try {
        $voice = New-Object -ComObject SAPI.SpVoice
        $voices = $voice.GetVoices()
        $found = $false
        
        for ($i = 0; $i -lt $voices.Count; $i++) {
            $voiceInfo = $voices.Item($i)
            $name = $voiceInfo.GetDescription()
            
            if ($name -like "*$VoiceName*") {
                Write-Host "✅ Found voice: $name" -ForegroundColor Green
                $voice.Voice = $voiceInfo
                
                Write-Host "🔊 Testing speech synthesis..." -ForegroundColor Yellow
                $voice.Speak("Hello! This is a test of the $VoiceName voice through SAPI.")
                
                Write-Host "✅ Voice test completed!" -ForegroundColor Green
                $found = $true
                break
            }
        }
        
        if (-not $found) {
            Write-Host "❌ Voice not found: $VoiceName" -ForegroundColor Red
            Write-Host "   Make sure the voice is installed correctly." -ForegroundColor Yellow
        }
        
    } catch {
        Write-Host "❌ Error testing voice: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "`n=== Test Complete ===" -ForegroundColor Cyan
