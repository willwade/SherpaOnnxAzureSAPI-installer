# Test Real SherpaOnnx TTS Engine
Write-Host "Testing Real SherpaOnnx TTS Engine..." -ForegroundColor Green
Write-Host "=====================================" -ForegroundColor Green

# Model paths
$modelPath = "C:\Program Files\OpenSpeech\models\piper-en-amy-medium\model.onnx"
$tokensPath = "C:\Program Files\OpenSpeech\models\piper-en-amy-medium\tokens.txt"
$outputPath = "C:\OpenSpeech\test_sherpa_output.wav"

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
    # Load the sherpa-onnx assembly
    Write-Host "`nLoading sherpa-onnx assembly..." -ForegroundColor Yellow
    
    # Try to load from the installed location
    $sherpaPath = "C:\Program Files\OpenAssistive\OpenSpeech\sherpa-onnx.dll"
    if (Test-Path $sherpaPath) {
        Write-Host "Loading from: $sherpaPath" -ForegroundColor Cyan
        [System.Reflection.Assembly]::LoadFrom($sherpaPath) | Out-Null
        Write-Host "✅ sherpa-onnx assembly loaded successfully" -ForegroundColor Green
    } else {
        Write-Host "❌ sherpa-onnx.dll not found at: $sherpaPath" -ForegroundColor Red
        exit 1
    }
    
    # Create TTS configuration
    Write-Host "`nCreating TTS configuration..." -ForegroundColor Yellow
    $config = New-Object SherpaOnnx.OfflineTtsConfig
    $config.Model.Vits.Model = $modelPath
    $config.Model.Vits.Tokens = $tokensPath
    $config.Model.Vits.NoiseScale = 0.667
    $config.Model.Vits.NoiseScaleW = 0.8
    $config.Model.Vits.LengthScale = 1.0
    $config.Model.NumThreads = 1
    $config.Model.Debug = 0
    $config.Model.Provider = "cpu"
    
    Write-Host "✅ Configuration created" -ForegroundColor Green
    
    # Create TTS instance
    Write-Host "`nCreating TTS instance..." -ForegroundColor Yellow
    $tts = New-Object SherpaOnnx.OfflineTts($config)
    Write-Host "✅ TTS instance created successfully" -ForegroundColor Green
    
    # Test text
    $testText = "Hello! This is Amy speaking using the real Sherpa ONNX TTS engine. The integration test is working correctly."
    Write-Host "`nGenerating audio for text: '$testText'" -ForegroundColor Yellow
    
    # Generate audio
    $audio = $tts.Generate($testText, 1.0, 0)
    $samples = $audio.Samples
    
    Write-Host "✅ Audio generated successfully!" -ForegroundColor Green
    Write-Host "Sample count: $($samples.Length)" -ForegroundColor Cyan
    Write-Host "Sample rate: $($audio.SampleRate)" -ForegroundColor Cyan
    
    # Convert to WAV format and save
    Write-Host "`nSaving audio to WAV file..." -ForegroundColor Yellow
    
    # Create output directory if it doesn't exist
    $outputDir = Split-Path $outputPath -Parent
    if (-not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }
    
    # Convert float samples to 16-bit PCM and save as WAV
    $sampleRate = 22050
    $channels = 1
    $bitsPerSample = 16
    
    # Create WAV file
    $fs = [System.IO.File]::Create($outputPath)
    $writer = New-Object System.IO.BinaryWriter($fs)
    
    # WAV header
    $writer.Write([System.Text.Encoding]::ASCII.GetBytes("RIFF"))
    $dataSize = $samples.Length * 2  # 16-bit samples
    $fileSize = 36 + $dataSize
    $writer.Write([uint32]$fileSize)
    $writer.Write([System.Text.Encoding]::ASCII.GetBytes("WAVE"))
    $writer.Write([System.Text.Encoding]::ASCII.GetBytes("fmt "))
    $writer.Write([uint32]16)  # fmt chunk size
    $writer.Write([uint16]1)   # PCM format
    $writer.Write([uint16]$channels)
    $writer.Write([uint32]$sampleRate)
    $writer.Write([uint32]($sampleRate * $channels * $bitsPerSample / 8))
    $writer.Write([uint16]($channels * $bitsPerSample / 8))
    $writer.Write([uint16]$bitsPerSample)
    $writer.Write([System.Text.Encoding]::ASCII.GetBytes("data"))
    $writer.Write([uint32]$dataSize)
    
    # Convert float samples to 16-bit PCM
    foreach ($sample in $samples) {
        $pcmSample = [Math]::Max(-32768, [Math]::Min(32767, [int]($sample * 32767)))
        $writer.Write([int16]$pcmSample)
    }
    
    $writer.Close()
    $fs.Close()
    
    Write-Host "✅ Audio saved to: $outputPath" -ForegroundColor Green
    Write-Host "File size: $((Get-Item $outputPath).Length) bytes" -ForegroundColor Cyan
    
    # Cleanup
    $tts.Dispose()
    
    Write-Host "`n🎉 REAL SHERPA ONNX TTS TEST SUCCESSFUL!" -ForegroundColor Green
    Write-Host "The TTS engine is working correctly and can generate real speech." -ForegroundColor White
    Write-Host "You can play the output file to hear Amy's voice: $outputPath" -ForegroundColor Yellow
    
} catch {
    Write-Host "`n❌ Error during TTS test: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack trace: $($_.Exception.StackTrace)" -ForegroundColor Red
    Write-Host "`nThis indicates an issue with the SherpaOnnx integration." -ForegroundColor Yellow
}

Write-Host "`nTest completed!" -ForegroundColor Green
