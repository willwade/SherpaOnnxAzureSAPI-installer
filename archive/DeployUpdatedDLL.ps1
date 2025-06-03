# Deploy updated DLL - must run as Administrator
Write-Host "Deploying Updated OpenSpeechTTS.dll" -ForegroundColor Green
Write-Host "====================================" -ForegroundColor Green

$sourceDLL = "OpenSpeechTTS\bin\Release\net472\OpenSpeechTTS.dll"
$targetDir = "C:\Program Files\OpenAssistive\OpenSpeech"
$targetDLL = "$targetDir\OpenSpeechTTS.dll"

Write-Host "1. Checking source DLL..." -ForegroundColor Cyan
if (Test-Path $sourceDLL) {
    $sourceTime = (Get-Item $sourceDLL).LastWriteTime
    Write-Host "   Source DLL: $sourceDLL" -ForegroundColor White
    Write-Host "   Last modified: $sourceTime" -ForegroundColor White
} else {
    Write-Host "   ❌ Source DLL not found: $sourceDLL" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "2. Copying DLL..." -ForegroundColor Cyan
try {
    Copy-Item $sourceDLL $targetDLL -Force
    Write-Host "   ✅ DLL copied successfully" -ForegroundColor Green
} catch {
    Write-Host "   ❌ Failed to copy DLL: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "3. Registering COM component..." -ForegroundColor Cyan
try {
    $regasmPath = "${env:ProgramFiles(x86)}\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\RegAsm.exe"
    if (-not (Test-Path $regasmPath)) {
        $regasmPath = "${env:ProgramFiles}\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\x64\RegAsm.exe"
    }
    
    if (Test-Path $regasmPath) {
        & $regasmPath $targetDLL /codebase
        Write-Host "   ✅ COM component registered" -ForegroundColor Green
    } else {
        Write-Host "   ⚠️ RegAsm not found, trying alternative..." -ForegroundColor Yellow
        # Try using .NET Framework regasm
        $frameworkRegasm = "${env:ProgramFiles(x86)}\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
        if (Test-Path $frameworkRegasm) {
            & $frameworkRegasm $targetDLL /codebase
            Write-Host "   ✅ COM component registered with Framework RegAsm" -ForegroundColor Green
        } else {
            Write-Host "   ❌ RegAsm not found" -ForegroundColor Red
        }
    }
} catch {
    Write-Host "   ❌ Failed to register COM: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "4. Verifying deployment..." -ForegroundColor Cyan
if (Test-Path $targetDLL) {
    $targetTime = (Get-Item $targetDLL).LastWriteTime
    Write-Host "   Target DLL: $targetDLL" -ForegroundColor White
    Write-Host "   Last modified: $targetTime" -ForegroundColor White
    
    if ($targetTime -eq $sourceTime) {
        Write-Host "   ✅ Deployment successful!" -ForegroundColor Green
    } else {
        Write-Host "   ⚠️ Timestamps don't match" -ForegroundColor Yellow
    }
} else {
    Write-Host "   ❌ Target DLL not found after copy" -ForegroundColor Red
}

Write-Host ""
Write-Host "Deployment complete. You can now test the updated TTS." -ForegroundColor Green
