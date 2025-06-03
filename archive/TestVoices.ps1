# Test script to check available SAPI voices
Write-Host "Testing SAPI Voice Installation" -ForegroundColor Green
Write-Host "================================" -ForegroundColor Green
Write-Host ""

# Test 1: Check available voices using .NET Speech API
Write-Host "1. Checking available voices using .NET Speech API:" -ForegroundColor Yellow
try {
    Add-Type -AssemblyName System.Speech
    $synth = New-Object System.Speech.Synthesis.SpeechSynthesizer
    $voices = $synth.GetInstalledVoices()
    
    Write-Host "Found $($voices.Count) voices:" -ForegroundColor Cyan
    foreach ($voice in $voices) {
        $info = $voice.VoiceInfo
        Write-Host "  - Name: $($info.Name)" -ForegroundColor White
        Write-Host "    Culture: $($info.Culture)" -ForegroundColor Gray
        Write-Host "    Gender: $($info.Gender)" -ForegroundColor Gray
        Write-Host "    Age: $($info.Age)" -ForegroundColor Gray
        Write-Host "    Description: $($info.Description)" -ForegroundColor Gray
        Write-Host ""
    }
} catch {
    Write-Host "Error checking voices: $($_.Exception.Message)" -ForegroundColor Red
}

# Test 2: Check registry entries for OpenSpeech voices
Write-Host "2. Checking registry for OpenSpeech voices:" -ForegroundColor Yellow
try {
    $voicesPath = "HKLM:\SOFTWARE\Microsoft\Speech\Voices\Tokens"
    if (Test-Path $voicesPath) {
        $voiceKeys = Get-ChildItem $voicesPath
        $openSpeechVoices = $voiceKeys | Where-Object { 
            $clsid = (Get-ItemProperty -Path $_.PSPath -Name "CLSID" -ErrorAction SilentlyContinue).CLSID
            $clsid -eq "{3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2}" -or $clsid -eq "{3d8f5c5e-9d6b-4b92-a12b-1a6dff80b6b3}"
        }
        
        if ($openSpeechVoices.Count -gt 0) {
            Write-Host "Found $($openSpeechVoices.Count) OpenSpeech voices in registry:" -ForegroundColor Cyan
            foreach ($voice in $openSpeechVoices) {
                $voiceName = $voice.PSChildName
                Write-Host "  - $voiceName" -ForegroundColor White
                
                # Check attributes
                $attribPath = Join-Path $voice.PSPath "Attributes"
                if (Test-Path $attribPath) {
                    $modelPath = (Get-ItemProperty -Path $attribPath -Name "ModelPath" -ErrorAction SilentlyContinue).ModelPath
                    $tokensPath = (Get-ItemProperty -Path $attribPath -Name "TokensPath" -ErrorAction SilentlyContinue).TokensPath
                    if ($modelPath) { Write-Host "    Model: $modelPath" -ForegroundColor Gray }
                    if ($tokensPath) { Write-Host "    Tokens: $tokensPath" -ForegroundColor Gray }
                }
                Write-Host ""
            }
        } else {
            Write-Host "No OpenSpeech voices found in registry." -ForegroundColor Yellow
        }
    }
} catch {
    Write-Host "Error checking registry: $($_.Exception.Message)" -ForegroundColor Red
}

# Test 3: Check if model files exist
Write-Host "3. Checking model files:" -ForegroundColor Yellow
$modelsPath = "C:\Program Files\OpenSpeech\models"
if (Test-Path $modelsPath) {
    $modelDirs = Get-ChildItem $modelsPath -Directory
    Write-Host "Found $($modelDirs.Count) model directories:" -ForegroundColor Cyan
    foreach ($dir in $modelDirs) {
        Write-Host "  - $($dir.Name)" -ForegroundColor White
        $modelFile = Join-Path $dir.FullName "model.onnx"
        $tokensFile = Join-Path $dir.FullName "tokens.txt"
        if (Test-Path $modelFile) {
            $size = (Get-Item $modelFile).Length / 1MB
            Write-Host "    model.onnx: $([math]::Round($size, 2)) MB" -ForegroundColor Gray
        }
        if (Test-Path $tokensFile) {
            $size = (Get-Item $tokensFile).Length / 1KB
            Write-Host "    tokens.txt: $([math]::Round($size, 2)) KB" -ForegroundColor Gray
        }
        Write-Host ""
    }
} else {
    Write-Host "Models directory not found: $modelsPath" -ForegroundColor Yellow
}

# Test 4: Check COM registration
Write-Host "4. Checking COM registration:" -ForegroundColor Yellow
$comPath = "HKLM:\SOFTWARE\Classes\CLSID\{3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2}"
if (Test-Path $comPath) {
    Write-Host "Sherpa ONNX COM class is registered" -ForegroundColor Green
    $inprocPath = Join-Path $comPath "InprocServer32"
    if (Test-Path $inprocPath) {
        $dllPath = (Get-ItemProperty -Path $inprocPath -Name "(default)" -ErrorAction SilentlyContinue)."(default)"
        if ($dllPath) {
            Write-Host "  DLL Path: $dllPath" -ForegroundColor Gray
            if (Test-Path $dllPath) {
                Write-Host "  DLL exists: Yes" -ForegroundColor Green
            } else {
                Write-Host "  DLL exists: No" -ForegroundColor Red
            }
        }
    }
} else {
    Write-Host "Sherpa ONNX COM class is NOT registered" -ForegroundColor Red
}

Write-Host ""
Write-Host "Test completed!" -ForegroundColor Green
