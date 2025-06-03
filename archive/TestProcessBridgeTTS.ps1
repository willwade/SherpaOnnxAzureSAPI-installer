# Test the new ProcessBridge TTS implementation
Write-Host "Testing ProcessBridge TTS Implementation" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green

# Clear logs
$logDir = "C:\OpenSpeech"
if (Test-Path "$logDir\sherpa_debug.log") { Clear-Content "$logDir\sherpa_debug.log" }
if (Test-Path "$logDir\sherpa_error.log") { Clear-Content "$logDir\sherpa_error.log" }

Write-Host "1. Creating COM object and triggering TTS initialization..." -ForegroundColor Cyan

try {
    # Create our COM object
    $comObject = New-Object -ComObject "OpenSpeechTTS.Sapi5VoiceImpl"
    Write-Host "   ✅ COM object created" -ForegroundColor Green
    
    # Call SetObjectToken to trigger TTS initialization
    $type = $comObject.GetType()
    $setTokenMethod = $type.GetMethod("SetObjectToken")
    $result = $setTokenMethod.Invoke($comObject, @([System.IntPtr]::Zero))
    Write-Host "   ✅ SetObjectToken called, result: $result" -ForegroundColor Green
    
    # Try to call SpeakToWaveStream to trigger audio generation
    Write-Host "   Attempting to call SpeakToWaveStream..." -ForegroundColor Yellow
    
    $speakMethod = $type.GetMethod("SpeakToWaveStream")
    if ($speakMethod) {
        Write-Host "   ✅ SpeakToWaveStream method found" -ForegroundColor Green
        
        # Create a memory stream
        $memoryStream = New-Object System.IO.MemoryStream
        
        try {
            # Call SpeakToWaveStream
            $speakMethod.Invoke($comObject, @("Hello from ProcessBridge TTS!", $memoryStream))
            
            $audioSize = $memoryStream.Length
            Write-Host "   ✅ SpeakToWaveStream completed! Generated $audioSize bytes of audio" -ForegroundColor Green
            
            if ($audioSize -gt 0) {
                Write-Host "   🎵 Audio generation successful!" -ForegroundColor Green
            } else {
                Write-Host "   ⚠️ No audio data generated" -ForegroundColor Yellow
            }
            
        } catch {
            Write-Host "   ❌ SpeakToWaveStream failed: $($_.Exception.Message)" -ForegroundColor Red
        } finally {
            $memoryStream.Dispose()
        }
    } else {
        Write-Host "   ❌ SpeakToWaveStream method not found" -ForegroundColor Red
    }
    
} catch {
    Write-Host "❌ Error: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "2. Checking logs for ProcessBridge activity..." -ForegroundColor Cyan

if (Test-Path "$logDir\sherpa_debug.log") {
    Write-Host "Sherpa Debug Log:" -ForegroundColor Yellow
    $debugContent = Get-Content "$logDir\sherpa_debug.log"
    
    $foundProcessBridge = $false
    $foundRealTTS = $false
    
    foreach ($line in $debugContent) {
        if ($line -like "*ProcessBasedTTS*" -or $line -like "*process bridge*") {
            Write-Host "   ✅ $line" -ForegroundColor Green
            $foundProcessBridge = $true
        } elseif ($line -like "*real TTS*" -or $line -like "*native bridge*") {
            Write-Host "   ℹ️ $line" -ForegroundColor Cyan
            $foundRealTTS = $true
        } elseif ($line -like "*Generated*samples*") {
            Write-Host "   🎵 $line" -ForegroundColor Green
        } elseif ($line -like "*ERROR*" -or $line -like "*failed*") {
            Write-Host "   ❌ $line" -ForegroundColor Red
        } else {
            Write-Host "   $line" -ForegroundColor Gray
        }
    }
    
    Write-Host ""
    if ($foundProcessBridge) {
        Write-Host "✅ ProcessBridge TTS is working!" -ForegroundColor Green
    } elseif ($foundRealTTS) {
        Write-Host "ℹ️ Real TTS attempted but ProcessBridge not reached" -ForegroundColor Cyan
    } else {
        Write-Host "❌ No TTS activity detected in logs" -ForegroundColor Red
    }
    
} else {
    Write-Host "   ❌ No sherpa debug log found" -ForegroundColor Red
}

if (Test-Path "$logDir\sherpa_error.log") {
    Write-Host ""
    Write-Host "Sherpa Error Log:" -ForegroundColor Red
    $errorContent = Get-Content "$logDir\sherpa_error.log"
    foreach ($line in $errorContent) {
        Write-Host "   ❌ $line" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "=== RESULT ===" -ForegroundColor Cyan
Write-Host "If ProcessBridge TTS is working:" -ForegroundColor Yellow
Write-Host "  ✅ Phase 2 (Real TTS Integration) is progressing!" -ForegroundColor Green
Write-Host "  Next: Implement actual SherpaOnnx process bridge" -ForegroundColor White
Write-Host ""
Write-Host "If ProcessBridge TTS is not working:" -ForegroundColor Yellow
Write-Host "  ❌ Need to debug the bridge implementation" -ForegroundColor Red
