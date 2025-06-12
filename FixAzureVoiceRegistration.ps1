# Fix Azure Voice Registration
# This script fixes the British English Azure Libby voice registration
# by removing the incorrect pipe service registration and re-registering it as an Azure TTS voice

Write-Host "🔧 AZURE VOICE REGISTRATION FIX" -ForegroundColor Cyan
Write-Host "===============================" -ForegroundColor Cyan
Write-Host ""

# Check if running as administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin) {
    Write-Host "❌ This script requires administrator privileges." -ForegroundColor Red
    Write-Host "Please run PowerShell as Administrator and try again." -ForegroundColor Red
    exit 1
}

$installerPath = ".\Installer\bin\Release\net6.0\win-x64\SherpaOnnxSAPIInstaller.exe"
$voiceName = "British-English-Azure-Libby"

# Check if installer exists
if (-not (Test-Path $installerPath)) {
    Write-Host "❌ Installer not found at: $installerPath" -ForegroundColor Red
    Write-Host "Please build the installer first:" -ForegroundColor Yellow
    Write-Host "  dotnet build Installer/Installer.csproj -c Release" -ForegroundColor Gray
    exit 1
}

Write-Host "Using installer: $installerPath" -ForegroundColor Gray
Write-Host "Voice to fix: $voiceName" -ForegroundColor Gray
Write-Host ""

Write-Host "🗑️ Step 1: Removing incorrect pipe service registration..." -ForegroundColor Yellow

try {
    # Remove the incorrectly registered pipe service voice
    & $installerPath remove-pipe-voice $voiceName
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✅ Removed pipe service registration" -ForegroundColor Green
    } else {
        Write-Host "⚠️ Could not remove pipe service registration (may not exist)" -ForegroundColor Yellow
    }
} catch {
    Write-Host "⚠️ Error removing pipe service registration: $_" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "🔧 Step 2: Registering as proper Azure TTS voice..." -ForegroundColor Yellow

try {
    # Register as Azure TTS voice with the credentials from the config
    $azureKey = "b14f8945b0f1459f9964bdd72c42c2cc"
    $azureRegion = "uksouth"
    $azureVoice = "en-GB-LibbyNeural"
    
    Write-Host "Azure Key: $azureKey" -ForegroundColor Gray
    Write-Host "Azure Region: $azureRegion" -ForegroundColor Gray
    Write-Host "Azure Voice: $azureVoice" -ForegroundColor Gray
    Write-Host ""
    
    & $installerPath install-azure $azureVoice --key $azureKey --region $azureRegion
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✅ Successfully registered Azure TTS voice" -ForegroundColor Green
    } else {
        Write-Host "❌ Failed to register Azure TTS voice" -ForegroundColor Red
        exit 1
    }
} catch {
    Write-Host "❌ Error registering Azure TTS voice: $_" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "🧪 Step 3: Testing the fixed voice..." -ForegroundColor Yellow

# Test the voice
Write-Host "Running SAPI voice test..." -ForegroundColor Gray
& .\TestSAPIVoices.ps1 -VoiceName "Libby" -ListOnly

Write-Host ""
Write-Host "✅ Azure voice registration fix completed!" -ForegroundColor Green
Write-Host ""
Write-Host "🧪 Next steps:" -ForegroundColor Cyan
Write-Host "  1. Test synthesis: .\TestSAPIVoices.ps1 -VoiceName 'Libby' -PlayAudio" -ForegroundColor White
Write-Host "  2. If it works, test in real applications (Notepad, etc.)" -ForegroundColor White
Write-Host "  3. The voice should now use the native Azure TTS COM wrapper" -ForegroundColor White

Write-Host ""
Write-Host "Fix completed!" -ForegroundColor Green
