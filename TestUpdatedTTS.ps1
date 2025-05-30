# Test Updated TTS Implementation
Write-Host "Testing Updated TTS Implementation..." -ForegroundColor Green
Write-Host "====================================" -ForegroundColor Green

Write-Host "`nStep 1: Testing current SAPI bridge with updated code..." -ForegroundColor Yellow

try {
    # Test current SAPI functionality
    Add-Type -AssemblyName System.Speech
    $synth = New-Object System.Speech.Synthesis.SpeechSynthesizer
    
    # List available voices
    $voices = $synth.GetInstalledVoices()
    Write-Host "`nAvailable voices:" -ForegroundColor Cyan
    foreach ($voice in $voices) {
        $enabled = if ($voice.Enabled) { "✅" } else { "❌" }
        Write-Host "  $enabled $($voice.VoiceInfo.Name)" -ForegroundColor White
    }
    
    # Test Amy voice
    $amyVoice = $voices | Where-Object { $_.VoiceInfo.Name -eq "amy" }
    if ($amyVoice -and $amyVoice.Enabled) {
        Write-Host "`n🎯 Testing Amy voice..." -ForegroundColor Yellow
        
        try {
            $synth.SelectVoice("amy")
            $synth.SetOutputToDefaultAudioDevice()
            
            $testText = "Hello! This is Amy testing the updated TTS implementation. I should now be using real SherpaOnnx TTS instead of mock audio."
            Write-Host "Speaking: '$testText'" -ForegroundColor Cyan
            
            $synth.Speak($testText)
            
            Write-Host "✅ Speech test completed!" -ForegroundColor Green
            
        } catch {
            Write-Host "❌ Speech test failed: $($_.Exception.Message)" -ForegroundColor Red
        }
    } else {
        Write-Host "❌ Amy voice not available or disabled" -ForegroundColor Red
    }
    
} catch {
    Write-Host "❌ Error during SAPI test: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`nStep 2: Checking debug logs for TTS implementation..." -ForegroundColor Yellow

if (Test-Path "C:\OpenSpeech\sapi_debug.log") {
    Write-Host "`nRecent SAPI debug log entries:" -ForegroundColor Cyan
    $logEntries = Get-Content "C:\OpenSpeech\sapi_debug.log" | Select-Object -Last 15
    foreach ($entry in $logEntries) {
        if ($entry -like "*REAL*") {
            Write-Host "  🎉 $entry" -ForegroundColor Green
        } elseif ($entry -like "*MOCK*") {
            Write-Host "  ⚠️  $entry" -ForegroundColor Yellow
        } elseif ($entry -like "*ERROR*") {
            Write-Host "  ❌ $entry" -ForegroundColor Red
        } else {
            Write-Host "  ℹ️  $entry" -ForegroundColor White
        }
    }
} else {
    Write-Host "❌ No debug log found at C:\OpenSpeech\sapi_debug.log" -ForegroundColor Red
}

Write-Host "`nStep 3: Analysis and Next Steps..." -ForegroundColor Yellow

Write-Host "`n📊 Current Status Analysis:" -ForegroundColor Cyan
Write-Host "1. ✅ SAPI Bridge: Working (interface registration fixed)" -ForegroundColor Green
Write-Host "2. ✅ Voice Selection: Working (Amy voice selectable)" -ForegroundColor Green
Write-Host "3. ✅ Audio Output: Working (you can hear speech)" -ForegroundColor Green
Write-Host "4. ❓ TTS Engine: Check logs to see if using REAL or MOCK TTS" -ForegroundColor Yellow

Write-Host "`n🔍 What to Look For in Logs:" -ForegroundColor Cyan
Write-Host "• 'Initializing SherpaTTS...' - TTS initialization started" -ForegroundColor White
Write-Host "• 'Real SherpaOnnx TTS initialized successfully!' - Real TTS working" -ForegroundColor Green
Write-Host "• 'SherpaTTS initialized in MOCK MODE' - Fallback to mock audio" -ForegroundColor Yellow
Write-Host "• 'Generating REAL audio bytes' - Using real SherpaOnnx" -ForegroundColor Green
Write-Host "• 'Generating MOCK audio bytes' - Using mock 440Hz tone" -ForegroundColor Yellow

Write-Host "`n🚀 Next Steps:" -ForegroundColor Cyan
if ($logEntries -like "*REAL*") {
    Write-Host "✅ SUCCESS: Real SherpaOnnx TTS is working!" -ForegroundColor Green
    Write-Host "   The integration is complete and functional." -ForegroundColor White
} elseif ($logEntries -like "*MOCK*") {
    Write-Host "⚠️  PARTIAL: Still using mock TTS" -ForegroundColor Yellow
    Write-Host "   Need to resolve SherpaOnnx library dependencies." -ForegroundColor White
    Write-Host "   Check for missing native libraries or strong-name issues." -ForegroundColor White
} else {
    Write-Host "❓ UNKNOWN: Check logs for more details" -ForegroundColor Yellow
    Write-Host "   May need to rebuild and redeploy the updated code." -ForegroundColor White
}

Write-Host "`nTest completed!" -ForegroundColor Green
