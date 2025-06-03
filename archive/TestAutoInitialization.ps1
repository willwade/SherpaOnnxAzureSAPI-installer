# Test Auto-Initialization in Speak Method
Write-Host "Testing Auto-Initialization Logic" -ForegroundColor Green
Write-Host "==================================" -ForegroundColor Green
Write-Host ""

# Clear logs
$logDir = "C:\OpenSpeech"
if (Test-Path "$logDir\sapi_debug.log") { Clear-Content "$logDir\sapi_debug.log" }
if (Test-Path "$logDir\sherpa_debug.log") { Clear-Content "$logDir\sherpa_debug.log" }

try {
    Write-Host "1. Creating COM object (simulating SAPI's second instance)..." -ForegroundColor Cyan
    $comObject = New-Object -ComObject "OpenSpeechTTS.Sapi5VoiceImpl"
    Write-Host "   ✅ COM object created" -ForegroundColor Green
    
    Write-Host "2. Checking model files..." -ForegroundColor Cyan
    $modelPath = "C:\Program Files\OpenSpeech\models\piper-en-amy-medium\model.onnx"
    $tokensPath = "C:\Program Files\OpenSpeech\models\piper-en-amy-medium\tokens.txt"
    
    if (Test-Path $modelPath) {
        Write-Host "   ✅ Model file exists: $modelPath" -ForegroundColor Green
    } else {
        Write-Host "   ❌ Model file missing: $modelPath" -ForegroundColor Red
    }
    
    if (Test-Path $tokensPath) {
        Write-Host "   ✅ Tokens file exists: $tokensPath" -ForegroundColor Green
    } else {
        Write-Host "   ❌ Tokens file missing: $tokensPath" -ForegroundColor Red
    }
    
    Write-Host "3. Attempting to call Speak method directly (without SetObjectToken)..." -ForegroundColor Cyan
    Write-Host "   This should trigger auto-initialization..." -ForegroundColor Yellow
    
    # Get the Speak method
    $type = $comObject.GetType()
    $speakMethod = $type.GetMethod("Speak")
    
    if ($speakMethod) {
        Write-Host "   ✅ Speak method found" -ForegroundColor Green
        
        # Try to call it with minimal parameters
        # This is complex due to the ref parameters, so let's just see if we can invoke it
        Write-Host "   Attempting to invoke Speak method..." -ForegroundColor Yellow
        
        try {
            # Create the required parameters
            $dwSpeakFlags = [uint32]0
            $rguidFormatId = [System.Guid]::Empty
            $pWaveFormatEx = $null
            $pTextFragList = $null
            $pOutputSite = [System.IntPtr]::Zero
            
            # This will likely fail due to parameter complexity, but should trigger auto-init
            $result = $speakMethod.Invoke($comObject, @($dwSpeakFlags, [ref]$rguidFormatId, [ref]$pWaveFormatEx, [ref]$pTextFragList, $pOutputSite))
            Write-Host "   ✅ Speak method returned: $result" -ForegroundColor Green
        } catch {
            Write-Host "   ⚠️ Speak method failed (expected due to parameter complexity): $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host "   But auto-initialization should have been attempted..." -ForegroundColor Yellow
        }
    } else {
        Write-Host "   ❌ Speak method not found" -ForegroundColor Red
    }
    
    # Release COM object
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) | Out-Null
    $comObject = $null
    
} catch {
    Write-Host "❌ Error: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "4. Checking logs for auto-initialization attempts..." -ForegroundColor Cyan

if (Test-Path "$logDir\sapi_debug.log") {
    Write-Host "SAPI Debug Log:" -ForegroundColor Yellow
    $logContent = Get-Content "$logDir\sapi_debug.log"
    if ($logContent) {
        foreach ($line in $logContent) {
            if ($line -like "*auto-initialization*" -or $line -like "*SPEAK METHOD CALLED*") {
                Write-Host "   $line" -ForegroundColor Green
            } else {
                Write-Host "   $line" -ForegroundColor Gray
            }
        }
    } else {
        Write-Host "   No log entries found" -ForegroundColor Gray
    }
} else {
    Write-Host "   No debug log found" -ForegroundColor Gray
}

if (Test-Path "$logDir\sherpa_debug.log") {
    Write-Host ""
    Write-Host "Sherpa Debug Log:" -ForegroundColor Yellow
    $logContent = Get-Content "$logDir\sherpa_debug.log"
    if ($logContent) {
        foreach ($line in $logContent) {
            Write-Host "   $line" -ForegroundColor Gray
        }
    } else {
        Write-Host "   No Sherpa log entries" -ForegroundColor Gray
    }
}

Write-Host ""
Write-Host "=== TEST COMPLETE ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "This test helps us understand:" -ForegroundColor Yellow
Write-Host "1. Whether auto-initialization is being attempted" -ForegroundColor White
Write-Host "2. Whether model files are accessible" -ForegroundColor White
Write-Host "3. What errors occur during auto-initialization" -ForegroundColor White
Write-Host ""
