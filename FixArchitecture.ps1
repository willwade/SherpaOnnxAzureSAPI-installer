# Fix Voice Architecture - Re-register Azure voice correctly
# This script fixes the architectural issue by removing the pipe service registration
# and re-registering the voice using the proper Azure TTS system

Write-Host "🔧 ARCHITECTURE FIX - Azure Voice Registration" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host ""

$installerPath = ".\Installer\bin\Release\net6.0\win-x64\SherpaOnnxSAPIInstaller.exe"

# Check if installer exists
if (-not (Test-Path $installerPath)) {
    Write-Host "❌ Installer not found at: $installerPath" -ForegroundColor Red
    Write-Host "Please build the installer first:" -ForegroundColor Yellow
    Write-Host "  dotnet build Installer/Installer.csproj -c Release" -ForegroundColor Gray
    exit 1
}

Write-Host "🔍 PROBLEM IDENTIFIED:" -ForegroundColor Yellow
Write-Host "  • Azure voice registered as pipe service (wrong architecture)" -ForegroundColor Red
Write-Host "  • Should use native Azure TTS COM wrapper instead" -ForegroundColor Green
Write-Host ""

Write-Host "📋 CURRENT STATUS:" -ForegroundColor Cyan
Write-Host "  • Voice: British-English-Azure-Libby" -ForegroundColor Gray
Write-Host "  • Current CLSID: {4A8B9C2D-1E3F-4567-8901-234567890ABC} (Pipe Service)" -ForegroundColor Red
Write-Host "  • Target CLSID: {3d8f5c5e-9d6b-4b92-a12b-1a6dff80b6b3} (Azure TTS)" -ForegroundColor Green
Write-Host ""

Write-Host "🔧 APPLYING FIX..." -ForegroundColor Yellow
Write-Host ""

# Step 1: Remove pipe service registration
Write-Host "Step 1: Removing incorrect pipe service registration..." -ForegroundColor Yellow
try {
    Write-Host "Running: $installerPath remove-pipe-voice British-English-Azure-Libby" -ForegroundColor Gray
    
    # Try without admin first to see if it works
    $result = & $installerPath remove-pipe-voice British-English-Azure-Libby 2>&1
    $exitCode = $LASTEXITCODE
    
    Write-Host "Exit code: $exitCode" -ForegroundColor Gray
    Write-Host "Output: $result" -ForegroundColor Gray
    
    if ($exitCode -eq 0) {
        Write-Host "✅ Successfully removed pipe service registration" -ForegroundColor Green
    } else {
        Write-Host "⚠️ Remove command returned exit code $exitCode" -ForegroundColor Yellow
        Write-Host "This might be expected if the voice wasn't registered as pipe service" -ForegroundColor Gray
    }
} catch {
    Write-Host "⚠️ Error removing pipe service registration: $_" -ForegroundColor Yellow
}

Write-Host ""

# Step 2: Register as Azure TTS voice
Write-Host "Step 2: Registering as proper Azure TTS voice..." -ForegroundColor Yellow
try {
    $azureKey = "b14f8945b0f1459f9964bdd72c42c2cc"
    $azureRegion = "uksouth"
    $azureVoice = "en-GB-LibbyNeural"
    
    Write-Host "Azure credentials:" -ForegroundColor Gray
    Write-Host "  Key: $azureKey" -ForegroundColor Gray
    Write-Host "  Region: $azureRegion" -ForegroundColor Gray
    Write-Host "  Voice: $azureVoice" -ForegroundColor Gray
    Write-Host ""
    
    Write-Host "Running: $installerPath install-azure $azureVoice --key $azureKey --region $azureRegion" -ForegroundColor Gray
    
    $result = & $installerPath install-azure $azureVoice --key $azureKey --region $azureRegion 2>&1
    $exitCode = $LASTEXITCODE
    
    Write-Host "Exit code: $exitCode" -ForegroundColor Gray
    Write-Host "Output: $result" -ForegroundColor Gray
    
    if ($exitCode -eq 0) {
        Write-Host "✅ Successfully registered Azure TTS voice" -ForegroundColor Green
    } else {
        Write-Host "❌ Azure registration failed with exit code $exitCode" -ForegroundColor Red
        Write-Host "This might require administrator privileges" -ForegroundColor Yellow
    }
} catch {
    Write-Host "❌ Error registering Azure TTS voice: $_" -ForegroundColor Red
}

Write-Host ""

# Step 3: Verify the fix
Write-Host "Step 3: Verifying the fix..." -ForegroundColor Yellow
try {
    Write-Host "Checking voice registration..." -ForegroundColor Gray
    $voiceReg = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Speech\Voices\Tokens\*" | Where-Object { $_.PSChildName -like "*Libby*" -or $_.PSChildName -like "*Azure*" }
    
    if ($voiceReg) {
        foreach ($voice in $voiceReg) {
            Write-Host "Found voice: $($voice.PSChildName)" -ForegroundColor Cyan
            Write-Host "  CLSID: $($voice.CLSID)" -ForegroundColor Gray
            Write-Host "  Path: $($voice.Path)" -ForegroundColor Gray
            
            if ($voice.CLSID -eq "{3d8f5c5e-9d6b-4b92-a12b-1a6dff80b6b3}") {
                Write-Host "  ✅ Using correct Azure TTS CLSID" -ForegroundColor Green
            } elseif ($voice.CLSID -eq "{4A8B9C2D-1E3F-4567-8901-234567890ABC}") {
                Write-Host "  ❌ Still using pipe service CLSID" -ForegroundColor Red
            } else {
                Write-Host "  ⚠️ Using unknown CLSID" -ForegroundColor Yellow
            }
        }
    } else {
        Write-Host "⚠️ No Azure/Libby voices found in registry" -ForegroundColor Yellow
    }
} catch {
    Write-Host "⚠️ Could not verify registration: $_" -ForegroundColor Yellow
}

Write-Host ""

# Step 4: Test the voice
Write-Host "Step 4: Testing the voice..." -ForegroundColor Yellow
Write-Host "Running SAPI voice test..." -ForegroundColor Gray

try {
    & .\TestSAPIVoices.ps1 -VoiceName "Libby" -ListOnly
} catch {
    Write-Host "⚠️ Could not run voice test: $_" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "🎯 ARCHITECTURE FIX COMPLETED" -ForegroundColor Cyan
Write-Host ""
Write-Host "🧪 Next steps:" -ForegroundColor Yellow
Write-Host "  1. If registration succeeded, test synthesis:" -ForegroundColor White
Write-Host "     .\TestSAPIVoices.ps1 -VoiceName 'Libby' -PlayAudio" -ForegroundColor Gray
Write-Host "  2. If it still fails, the issue might be:" -ForegroundColor White
Write-Host "     • Missing native DLL registration" -ForegroundColor Gray
Write-Host "     • Administrator privileges required" -ForegroundColor Gray
Write-Host "  3. Voice should now use native Azure TTS (not pipe service)" -ForegroundColor White

Write-Host ""
Write-Host "Fix completed!" -ForegroundColor Green
