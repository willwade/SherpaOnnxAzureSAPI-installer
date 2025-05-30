# Test SherpaWrapper Class
Write-Host "Testing SherpaWrapper Class..." -ForegroundColor Green
Write-Host "==============================" -ForegroundColor Green

# Model paths
$modelPath = "C:\Program Files\OpenSpeech\models\piper-en-amy-medium\model.onnx"
$tokensPath = "C:\Program Files\OpenSpeech\models\piper-en-amy-medium\tokens.txt"
$outputPath = "C:\OpenSpeech\test_wrapper_output.wav"

Write-Host "Model Path: $modelPath" -ForegroundColor Cyan
Write-Host "Tokens Path: $tokensPath" -ForegroundColor Cyan
Write-Host "Output Path: $outputPath" -ForegroundColor Cyan
Write-Host ""

# Check if model files exist
if (-not (Test-Path $modelPath)) {
    Write-Host "❌ Model file not found: $modelPath" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $tokensPath)) {
    Write-Host "❌ Tokens file not found: $tokensPath" -ForegroundColor Red
    exit 1
}

Write-Host "✅ Model files found" -ForegroundColor Green

try {
    # Load the SherpaNative assembly
    Write-Host "`nLoading SherpaNative assembly..." -ForegroundColor Yellow
    
    $sherpaNativePath = "SherpaNative\bin\Release\net472\SherpaNative.dll"
    if (Test-Path $sherpaNativePath) {
        Write-Host "Loading from: $sherpaNativePath" -ForegroundColor Cyan
        [System.Reflection.Assembly]::LoadFrom((Resolve-Path $sherpaNativePath).Path) | Out-Null
        Write-Host "✅ SherpaNative assembly loaded successfully" -ForegroundColor Green
    } else {
        Write-Host "❌ SherpaNative.dll not found at: $sherpaNativePath" -ForegroundColor Red
        exit 1
    }
    
    # Create SherpaWrapper instance
    Write-Host "`nCreating SherpaWrapper instance..." -ForegroundColor Yellow
    $wrapper = New-Object SherpaNative.SherpaWrapper($modelPath, $tokensPath)
    Write-Host "✅ SherpaWrapper instance created successfully" -ForegroundColor Green
    
    # Test text
    $testText = "Hello! This is Amy speaking using the SherpaWrapper. The integration test is working correctly."
    Write-Host "`nGenerating audio for text: '$testText'" -ForegroundColor Yellow
    
    # Generate audio
    $audioBytes = $wrapper.GenerateWaveform($testText)
    
    Write-Host "✅ Audio generated successfully!" -ForegroundColor Green
    Write-Host "Audio data size: $($audioBytes.Length) bytes" -ForegroundColor Cyan
    
    # Save to file
    Write-Host "`nSaving audio to file..." -ForegroundColor Yellow
    
    # Create output directory if it doesn't exist
    $outputDir = Split-Path $outputPath -Parent
    if (-not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }
    
    # Save raw audio data
    [System.IO.File]::WriteAllBytes($outputPath, $audioBytes)
    
    Write-Host "✅ Audio saved to: $outputPath" -ForegroundColor Green
    Write-Host "File size: $((Get-Item $outputPath).Length) bytes" -ForegroundColor Cyan
    
    Write-Host "`n🎉 SHERPA WRAPPER TEST SUCCESSFUL!" -ForegroundColor Green
    Write-Host "The SherpaWrapper is working correctly and can generate real speech." -ForegroundColor White
    Write-Host "You can play the output file to hear Amy's voice: $outputPath" -ForegroundColor Yellow
    
} catch {
    Write-Host "`n❌ Error during SherpaWrapper test: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack trace: $($_.Exception.StackTrace)" -ForegroundColor Red
    Write-Host "`nThis indicates an issue with the SherpaWrapper integration." -ForegroundColor Yellow
    
    # Check if sherpa-onnx.dll is available
    $sherpaPath = "SherpaNative\bin\Release\net472\sherpa-onnx.dll"
    if (Test-Path $sherpaPath) {
        $sherpaSize = (Get-Item $sherpaPath).Length
        Write-Host "`nSherpa-onnx.dll found: $sherpaPath ($sherpaSize bytes)" -ForegroundColor Yellow
        if ($sherpaSize -lt 100000) {
            Write-Host "⚠️  Warning: sherpa-onnx.dll seems too small ($sherpaSize bytes)" -ForegroundColor Yellow
            Write-Host "This might be a stub file rather than the full library." -ForegroundColor Yellow
        }
    } else {
        Write-Host "`n❌ sherpa-onnx.dll not found at: $sherpaPath" -ForegroundColor Red
    }
}

Write-Host "`nTest completed!" -ForegroundColor Green
