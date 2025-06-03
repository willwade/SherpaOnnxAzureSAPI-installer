# Test real TTS initialization to see what's failing
Write-Host "Testing Real TTS Initialization" -ForegroundColor Green
Write-Host "===============================" -ForegroundColor Green

# Clear logs
$logDir = "C:\OpenSpeech"
if (Test-Path "$logDir\sherpa_debug.log") { Clear-Content "$logDir\sherpa_debug.log" }
if (Test-Path "$logDir\sherpa_error.log") { Clear-Content "$logDir\sherpa_error.log" }

Write-Host "1. Testing SherpaTTS initialization..." -ForegroundColor Cyan

try {
    # Create our COM object and trigger SherpaTTS initialization
    $comObject = New-Object -ComObject "OpenSpeechTTS.Sapi5VoiceImpl"
    Write-Host "   ✅ COM object created" -ForegroundColor Green
    
    # Call SetObjectToken to trigger TTS initialization
    $type = $comObject.GetType()
    $setTokenMethod = $type.GetMethod("SetObjectToken")
    $result = $setTokenMethod.Invoke($comObject, @([System.IntPtr]::Zero))
    Write-Host "   ✅ SetObjectToken called, result: $result" -ForegroundColor Green
    
    # Try to generate audio to see if real TTS works
    Write-Host "   Testing audio generation..." -ForegroundColor Yellow
    
    # This is complex to call directly, so let's check the logs instead
    
} catch {
    Write-Host "❌ Error: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "2. Checking SherpaTTS initialization logs..." -ForegroundColor Cyan

if (Test-Path "$logDir\sherpa_debug.log") {
    Write-Host "Sherpa Debug Log:" -ForegroundColor Yellow
    $debugContent = Get-Content "$logDir\sherpa_debug.log"
    foreach ($line in $debugContent) {
        if ($line -like "*REAL TTS*") {
            Write-Host "   ✅ $line" -ForegroundColor Green
        } elseif ($line -like "*MOCK*") {
            Write-Host "   ⚠️ $line" -ForegroundColor Yellow
        } elseif ($line -like "*successfully*") {
            Write-Host "   ✅ $line" -ForegroundColor Green
        } else {
            Write-Host "   $line" -ForegroundColor Gray
        }
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
} else {
    Write-Host ""
    Write-Host "   ✅ No sherpa error log (good sign)" -ForegroundColor Green
}

Write-Host ""
Write-Host "3. Checking model files..." -ForegroundColor Cyan

$modelPath = "C:\Program Files\OpenSpeech\models\piper-en-amy-medium\model.onnx"
$tokensPath = "C:\Program Files\OpenSpeech\models\piper-en-amy-medium\tokens.txt"

if (Test-Path $modelPath) {
    $modelSize = (Get-Item $modelPath).Length
    Write-Host "   ✅ Model file exists: $modelPath ($([math]::Round($modelSize/1MB, 1)) MB)" -ForegroundColor Green
} else {
    Write-Host "   ❌ Model file missing: $modelPath" -ForegroundColor Red
}

if (Test-Path $tokensPath) {
    Write-Host "   ✅ Tokens file exists: $tokensPath" -ForegroundColor Green
} else {
    Write-Host "   ❌ Tokens file missing: $tokensPath" -ForegroundColor Red
}

Write-Host ""
Write-Host "4. Checking SherpaOnnx assembly..." -ForegroundColor Cyan

$sherpaPath = "C:\Program Files\OpenAssistive\OpenSpeech\sherpa-onnx.dll"
if (Test-Path $sherpaPath) {
    $sherpaSize = (Get-Item $sherpaPath).Length
    Write-Host "   ✅ SherpaOnnx assembly exists: $sherpaPath ($([math]::Round($sherpaSize/1KB, 1)) KB)" -ForegroundColor Green
    
    # Try to load it to see what happens
    try {
        $assembly = [System.Reflection.Assembly]::LoadFrom($sherpaPath)
        Write-Host "   ✅ Assembly loaded successfully" -ForegroundColor Green
        Write-Host "     Runtime Version: $($assembly.ImageRuntimeVersion)" -ForegroundColor Gray
        Write-Host "     Location: $($assembly.Location)" -ForegroundColor Gray
        
        # Try to find the types we need
        $offlineTtsType = $assembly.GetType("SherpaOnnx.OfflineTts")
        $configType = $assembly.GetType("SherpaOnnx.OfflineTtsConfig")
        
        if ($offlineTtsType) {
            Write-Host "   ✅ OfflineTts type found" -ForegroundColor Green
        } else {
            Write-Host "   ❌ OfflineTts type not found" -ForegroundColor Red
        }
        
        if ($configType) {
            Write-Host "   ✅ OfflineTtsConfig type found" -ForegroundColor Green
        } else {
            Write-Host "   ❌ OfflineTtsConfig type not found" -ForegroundColor Red
        }
        
    } catch {
        Write-Host "   ❌ Failed to load assembly: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "   ❌ SherpaOnnx assembly missing: $sherpaPath" -ForegroundColor Red
}

Write-Host ""
Write-Host "=== DIAGNOSIS ===" -ForegroundColor Cyan
Write-Host "If real TTS initialization failed, the issue is likely:" -ForegroundColor Yellow
Write-Host "1. .NET 6.0 vs .NET Framework 4.7.2 compatibility" -ForegroundColor White
Write-Host "2. Missing dependencies or incorrect assembly loading" -ForegroundColor White
Write-Host "3. Model file issues or incorrect paths" -ForegroundColor White
Write-Host ""
Write-Host "Next step: Implement .NET compatibility solution" -ForegroundColor Green
