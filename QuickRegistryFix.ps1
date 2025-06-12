# Quick Registry Fix - Change CLSID for Azure voice
# This script directly fixes the registry to use the correct Azure TTS CLSID

Write-Host "🔧 QUICK REGISTRY FIX" -ForegroundColor Cyan
Write-Host "=====================" -ForegroundColor Cyan
Write-Host ""

# Check if running as administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin) {
    Write-Host "❌ This script requires administrator privileges." -ForegroundColor Red
    Write-Host "Please run PowerShell as Administrator and try again." -ForegroundColor Red
    exit 1
}

$voiceRegistryPath = "HKLM:\SOFTWARE\Microsoft\Speech\Voices\Tokens\British-English-Azure-Libby"
$currentClsid = "{4A8B9C2D-1E3F-4567-8901-234567890ABC}"  # Pipe Service
$targetClsid = "{3d8f5c5e-9d6b-4b92-a12b-1a6dff80b6b3}"   # Azure TTS
$nativeDllPath = "C:\Program Files\OpenAssistive\OpenSpeech\NativeTTSWrapper.dll"

Write-Host "🔍 PROBLEM:" -ForegroundColor Yellow
Write-Host "  Voice is using pipe service CLSID instead of Azure TTS CLSID" -ForegroundColor Red
Write-Host ""

Write-Host "📋 REGISTRY PATHS:" -ForegroundColor Cyan
Write-Host "  Voice: $voiceRegistryPath" -ForegroundColor Gray
Write-Host "  Current CLSID: $currentClsid" -ForegroundColor Red
Write-Host "  Target CLSID: $targetClsid" -ForegroundColor Green
Write-Host "  Native DLL: $nativeDllPath" -ForegroundColor Gray
Write-Host ""

# Check if voice registry entry exists
if (-not (Test-Path $voiceRegistryPath)) {
    Write-Host "❌ Voice registry entry not found: $voiceRegistryPath" -ForegroundColor Red
    Write-Host "The voice may not be registered. Please register it first." -ForegroundColor Yellow
    exit 1
}

Write-Host "✅ Voice registry entry found" -ForegroundColor Green

# Read current CLSID
try {
    $currentClsidValue = Get-ItemProperty -Path $voiceRegistryPath -Name "CLSID" -ErrorAction Stop
    Write-Host "Current CLSID: $($currentClsidValue.CLSID)" -ForegroundColor Cyan
    
    if ($currentClsidValue.CLSID -eq $targetClsid) {
        Write-Host "✅ Voice is already using the correct Azure TTS CLSID" -ForegroundColor Green
        Write-Host "The issue might be elsewhere. Try testing the voice:" -ForegroundColor Yellow
        Write-Host "  .\TestSAPIVoices.ps1 -VoiceName 'Azure Libby' -PlayAudio" -ForegroundColor Gray
        exit 0
    }
} catch {
    Write-Host "⚠️ Could not read current CLSID: $_" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "🔧 APPLYING FIX..." -ForegroundColor Yellow

try {
    # Update the CLSID
    Write-Host "Updating CLSID..." -ForegroundColor Gray
    Set-ItemProperty -Path $voiceRegistryPath -Name "CLSID" -Value $targetClsid
    Write-Host "✅ Updated CLSID to Azure TTS" -ForegroundColor Green
    
    # Update the Path to point to the native DLL
    Write-Host "Updating Path..." -ForegroundColor Gray
    Set-ItemProperty -Path $voiceRegistryPath -Name "Path" -Value $nativeDllPath
    Write-Host "✅ Updated Path to native DLL" -ForegroundColor Green
    
    # Update the VoiceType attribute if it exists
    $attributesPath = "$voiceRegistryPath\Attributes"
    if (Test-Path $attributesPath) {
        Write-Host "Updating VoiceType..." -ForegroundColor Gray
        Set-ItemProperty -Path $attributesPath -Name "VoiceType" -Value "AzureTTS"
        Write-Host "✅ Updated VoiceType to AzureTTS" -ForegroundColor Green
    }
    
} catch {
    Write-Host "❌ Error updating registry: $_" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "🔧 VERIFYING FIX..." -ForegroundColor Yellow

try {
    $updatedClsid = Get-ItemProperty -Path $voiceRegistryPath -Name "CLSID" -ErrorAction Stop
    $updatedPath = Get-ItemProperty -Path $voiceRegistryPath -Name "Path" -ErrorAction Stop
    
    Write-Host "Verification:" -ForegroundColor Cyan
    Write-Host "  CLSID: $($updatedClsid.CLSID)" -ForegroundColor Gray
    Write-Host "  Path: $($updatedPath.Path)" -ForegroundColor Gray
    
    if ($updatedClsid.CLSID -eq $targetClsid) {
        Write-Host "✅ CLSID fix verified" -ForegroundColor Green
    } else {
        Write-Host "❌ CLSID fix failed" -ForegroundColor Red
    }
    
} catch {
    Write-Host "⚠️ Could not verify fix: $_" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "🧪 TESTING VOICE..." -ForegroundColor Yellow

try {
    Write-Host "Running SAPI voice test..." -ForegroundColor Gray
    & .\TestSAPIVoices.ps1 -VoiceName "Libby" -ListOnly
} catch {
    Write-Host "⚠️ Could not run voice test: $_" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "✅ REGISTRY FIX COMPLETED!" -ForegroundColor Green
Write-Host ""
Write-Host "🧪 Next steps:" -ForegroundColor Cyan
Write-Host "  1. Test synthesis: .\TestSAPIVoices.ps1 -VoiceName 'Libby' -PlayAudio" -ForegroundColor White
Write-Host "  2. If it still fails, check if NativeTTSWrapper.dll is registered:" -ForegroundColor White
Write-Host "     regsvr32 'C:\Program Files\OpenAssistive\OpenSpeech\NativeTTSWrapper.dll'" -ForegroundColor Gray
Write-Host "  3. Voice should now use native Azure TTS (not pipe service)" -ForegroundColor White

Write-Host ""
Write-Host "Fix completed!" -ForegroundColor Green
