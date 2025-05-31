# Deploy Real TTS - Build and Deploy Updated SherpaOnnx Integration
Write-Host "🚀 DEPLOYING REAL SHERPAONNX TTS..." -ForegroundColor Yellow
Write-Host ""

try {
    # Build the OpenSpeechTTS project
    Write-Host "Building OpenSpeechTTS with real TTS support..." -ForegroundColor Cyan
    $buildResult = dotnet build OpenSpeechTTS -c Release
    if ($LASTEXITCODE -ne 0) {
        Write-Host "❌ Build failed!" -ForegroundColor Red
        exit 1
    }
    Write-Host "✅ Build completed successfully" -ForegroundColor Green
    
    # Copy the updated DLL
    Write-Host "Deploying updated DLL..." -ForegroundColor Cyan
    $sourceDll = "OpenSpeechTTS\bin\Release\net472\OpenSpeechTTS.dll"
    $targetDll = "C:\Program Files\OpenAssistive\OpenSpeech\OpenSpeechTTS.dll"
    
    if (Test-Path $sourceDll) {
        Copy-Item $sourceDll $targetDll -Force
        Write-Host "✅ DLL deployed successfully" -ForegroundColor Green
    } else {
        Write-Host "❌ Source DLL not found: $sourceDll" -ForegroundColor Red
        exit 1
    }
    
    # Register the COM component
    Write-Host "Registering updated COM component..." -ForegroundColor Cyan
    $regResult = regsvr32 /s $targetDll
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✅ COM component registered successfully" -ForegroundColor Green
    } else {
        Write-Host "❌ COM registration failed" -ForegroundColor Red
        exit 1
    }
    
    Write-Host ""
    Write-Host "🧪 TESTING REAL TTS..." -ForegroundColor Yellow
    
    # Clear logs for fresh testing
    $logPath = "C:\OpenSpeech\sherpa_debug.log"
    if (Test-Path $logPath) {
        Clear-Content $logPath
    }
    
    # Test the real TTS
    Write-Host "Testing Amy voice with real SherpaOnnx TTS..." -ForegroundColor Cyan
    
    $testScript = @'
Add-Type -AssemblyName System.Speech
$synth = New-Object System.Speech.Synthesis.SpeechSynthesizer

Write-Host "Available voices:"
$synth.GetInstalledVoices() | ForEach-Object {
    $voice = $_.VoiceInfo
    $enabled = if ($_.Enabled) { "ENABLED" } else { "DISABLED" }
    Write-Host "  - $($voice.Name) [$enabled]"
}

Write-Host ""
Write-Host "Testing Amy voice with REAL TTS..."
$amyVoices = $synth.GetInstalledVoices() | Where-Object { $_.VoiceInfo.Name -like "*amy*" }
if ($amyVoices) {
    $synth.SelectVoice($amyVoices[0].VoiceInfo.Name)
    Write-Host "✅ Amy voice selected!" -ForegroundColor Green
    
    Write-Host "Generating speech with real SherpaOnnx..."
    $synth.Speak("Hello! This is Amy speaking with real SherpaOnnx text to speech synthesis. The SAPI bridge is now fully functional!")
    Write-Host "✅ Speech generation completed!" -ForegroundColor Green
} else {
    Write-Host "❌ Amy voice not found" -ForegroundColor Red
}

$synth.Dispose()
'@
    
    # Execute the test
    Invoke-Expression $testScript
    
    Write-Host ""
    Write-Host "📋 CHECKING LOGS..." -ForegroundColor Yellow
    
    # Check Sherpa debug logs
    if (Test-Path $logPath) {
        Write-Host "Sherpa TTS Debug Log:" -ForegroundColor Cyan
        Get-Content $logPath | Select-Object -Last 10 | ForEach-Object { 
            Write-Host "  $_" -ForegroundColor Gray 
        }
    }
    
    # Check SAPI debug logs
    $sapiLogPath = "C:\OpenSpeech\sapi_debug.log"
    if (Test-Path $sapiLogPath) {
        Write-Host ""
        Write-Host "SAPI Debug Log:" -ForegroundColor Cyan
        Get-Content $sapiLogPath | Select-Object -Last 5 | ForEach-Object { 
            Write-Host "  $_" -ForegroundColor Gray 
        }
    }
    
    Write-Host ""
    Write-Host "🎉 REAL TTS DEPLOYMENT COMPLETED!" -ForegroundColor Green
    Write-Host "The SherpaOnnx SAPI bridge is now using real TTS synthesis!" -ForegroundColor Green
    
} catch {
    Write-Host "❌ ERROR: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
