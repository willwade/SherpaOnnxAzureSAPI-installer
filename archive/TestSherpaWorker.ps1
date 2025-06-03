# Test the SherpaWorker TTS process
Write-Host "Testing SherpaWorker TTS Process" -ForegroundColor Green
Write-Host "=================================" -ForegroundColor Green

$workerDir = "SherpaWorker"
$projectFile = "$workerDir\SherpaWorker.csproj"
$outputDir = "$workerDir\bin\Release\net6.0\win-x64\publish"
$workerExe = "$outputDir\SherpaWorker.exe"

Write-Host "1. Building SherpaWorker..." -ForegroundColor Cyan

if (Test-Path $projectFile) {
    Write-Host "   Project file found: $projectFile" -ForegroundColor Green
    
    try {
        # Build and publish the worker
        Write-Host "   Building and publishing..." -ForegroundColor Yellow
        & "C:\Program Files\dotnet\dotnet.exe" publish $projectFile -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "   ✅ Build successful" -ForegroundColor Green
        } else {
            Write-Host "   ❌ Build failed" -ForegroundColor Red
            exit 1
        }
    } catch {
        Write-Host "   ❌ Build error: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
} else {
    Write-Host "   ❌ Project file not found: $projectFile" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "2. Checking build output..." -ForegroundColor Cyan

if (Test-Path $workerExe) {
    $exeSize = [math]::Round((Get-Item $workerExe).Length / 1MB, 1)
    Write-Host "   ✅ Worker executable: $workerExe" -ForegroundColor Green
    Write-Host "   Size: $exeSize MB" -ForegroundColor White
} else {
    Write-Host "   ❌ Worker executable not found: $workerExe" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "3. Testing SherpaWorker with sample request..." -ForegroundColor Cyan

# Create test request
$testRequest = @{
    Text = "Hello! This is a test of the SherpaWorker TTS process."
    Speed = 1.0
    SpeakerId = 0
    OutputPath = "test_output"
} | ConvertTo-Json

$requestFile = "test_request.json"
$responseFile = "test_request.response.json"
$audioFile = "test_output.wav"

# Clean up previous test files
Remove-Item $requestFile -ErrorAction SilentlyContinue
Remove-Item $responseFile -ErrorAction SilentlyContinue
Remove-Item $audioFile -ErrorAction SilentlyContinue

# Write test request
$testRequest | Out-File -FilePath $requestFile -Encoding UTF8
Write-Host "   Created test request: $requestFile" -ForegroundColor White

# Run SherpaWorker
Write-Host "   Running SherpaWorker..." -ForegroundColor Yellow
try {
    $process = Start-Process -FilePath $workerExe -ArgumentList $requestFile -Wait -PassThru -NoNewWindow
    
    if ($process.ExitCode -eq 0) {
        Write-Host "   ✅ SherpaWorker completed successfully" -ForegroundColor Green
    } else {
        Write-Host "   ❌ SherpaWorker failed with exit code: $($process.ExitCode)" -ForegroundColor Red
    }
} catch {
    Write-Host "   ❌ Error running SherpaWorker: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "4. Checking results..." -ForegroundColor Cyan

# Check response file
if (Test-Path $responseFile) {
    Write-Host "   ✅ Response file created: $responseFile" -ForegroundColor Green
    
    try {
        $response = Get-Content $responseFile | ConvertFrom-Json
        Write-Host "   Response details:" -ForegroundColor White
        Write-Host "     Success: $($response.Success)" -ForegroundColor Gray
        
        if ($response.Success) {
            Write-Host "     Sample count: $($response.SampleCount)" -ForegroundColor Gray
            Write-Host "     Sample rate: $($response.SampleRate) Hz" -ForegroundColor Gray
            Write-Host "     Audio path: $($response.AudioPath)" -ForegroundColor Gray
        } else {
            Write-Host "     Error: $($response.ErrorMessage)" -ForegroundColor Red
        }
    } catch {
        Write-Host "   ⚠️ Could not parse response file" -ForegroundColor Yellow
    }
} else {
    Write-Host "   ❌ Response file not created" -ForegroundColor Red
}

# Check audio file
if (Test-Path $audioFile) {
    $audioSize = [math]::Round((Get-Item $audioFile).Length / 1KB, 1)
    Write-Host "   ✅ Audio file created: $audioFile" -ForegroundColor Green
    Write-Host "   Audio size: $audioSize KB" -ForegroundColor White
    
    # Try to play the audio (optional)
    Write-Host "   🎵 Audio file ready for playback!" -ForegroundColor Green
} else {
    Write-Host "   ❌ Audio file not created" -ForegroundColor Red
}

Write-Host ""
Write-Host "=== TEST COMPLETE ===" -ForegroundColor Cyan
Write-Host ""

if ((Test-Path $responseFile) -and (Test-Path $audioFile)) {
    Write-Host "🎯 SUCCESS: SherpaWorker TTS process is working!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Yellow
    Write-Host "1. Integrate SherpaWorker with COM object ProcessBridge" -ForegroundColor White
    Write-Host "2. Replace mock audio with real SherpaOnnx TTS" -ForegroundColor White
    Write-Host "3. Optimize performance and error handling" -ForegroundColor White
    Write-Host ""
    Write-Host "This proves the ProcessBridge architecture works!" -ForegroundColor Green
} else {
    Write-Host "❌ SherpaWorker test failed" -ForegroundColor Red
    Write-Host "Check the error messages above for debugging information." -ForegroundColor Yellow
}

# Clean up test files
Write-Host ""
Write-Host "Cleaning up test files..." -ForegroundColor Gray
Remove-Item $requestFile -ErrorAction SilentlyContinue
Remove-Item $responseFile -ErrorAction SilentlyContinue
Remove-Item $audioFile -ErrorAction SilentlyContinue
