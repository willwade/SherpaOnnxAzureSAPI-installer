# Fix Voice CLSID Registration
# This script directly fixes the CLSID for the British English Azure Libby voice
# by changing it from the pipe service CLSID to the Azure TTS CLSID

Write-Host "🔧 VOICE CLSID FIX" -ForegroundColor Cyan
Write-Host "==================" -ForegroundColor Cyan
Write-Host ""

# Check if running as administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin) {
    Write-Host "❌ This script requires administrator privileges." -ForegroundColor Red
    Write-Host "Please run PowerShell as Administrator and try again." -ForegroundColor Red
    exit 1
}

$voiceRegistryPath = "HKLM:\SOFTWARE\Microsoft\Speech\Voices\Tokens\British-English-Azure-Libby"
$pipeServiceClsid = "{4A8B9C2D-1E3F-4567-8901-234567890ABC}"
$azureTtsClsid = "{3d8f5c5e-9d6b-4b92-a12b-1a6dff80b6b3}"
$nativeDllPath = "C:\Program Files\OpenAssistive\OpenSpeech\NativeTTSWrapper.dll"

Write-Host "Voice registry path: $voiceRegistryPath" -ForegroundColor Gray
Write-Host "Current CLSID (pipe service): $pipeServiceClsid" -ForegroundColor Gray
Write-Host "Target CLSID (Azure TTS): $azureTtsClsid" -ForegroundColor Gray
Write-Host "Native DLL path: $nativeDllPath" -ForegroundColor Gray
Write-Host ""

# Check if voice registry entry exists
if (-not (Test-Path $voiceRegistryPath)) {
    Write-Host "❌ Voice registry entry not found: $voiceRegistryPath" -ForegroundColor Red
    Write-Host "The voice may not be registered. Please register it first." -ForegroundColor Yellow
    exit 1
}

Write-Host "✅ Voice registry entry found" -ForegroundColor Green

# Check current CLSID
try {
    $currentClsid = Get-ItemProperty -Path $voiceRegistryPath -Name "CLSID" -ErrorAction Stop
    Write-Host "Current CLSID: $($currentClsid.CLSID)" -ForegroundColor Cyan
    
    if ($currentClsid.CLSID -eq $azureTtsClsid) {
        Write-Host "✅ Voice is already using the correct Azure TTS CLSID" -ForegroundColor Green
        Write-Host "The issue might be elsewhere. Try testing the voice:" -ForegroundColor Yellow
        Write-Host "  .\TestSAPIVoices.ps1 -VoiceName 'Azure Libby' -PlayAudio" -ForegroundColor Gray
        exit 0
    }
} catch {
    Write-Host "⚠️ Could not read current CLSID: $_" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "🔧 Fixing CLSID registration..." -ForegroundColor Yellow

try {
    # Update the CLSID to use Azure TTS
    Set-ItemProperty -Path $voiceRegistryPath -Name "CLSID" -Value $azureTtsClsid
    Write-Host "✅ Updated CLSID to Azure TTS" -ForegroundColor Green
    
    # Update the Path to point to the native DLL
    Set-ItemProperty -Path $voiceRegistryPath -Name "Path" -Value $nativeDllPath
    Write-Host "✅ Updated Path to native DLL" -ForegroundColor Green
    
    # Update the VoiceType attribute
    $attributesPath = "$voiceRegistryPath\Attributes"
    if (Test-Path $attributesPath) {
        Set-ItemProperty -Path $attributesPath -Name "VoiceType" -Value "AzureTTS"
        Write-Host "✅ Updated VoiceType to AzureTTS" -ForegroundColor Green
    }
    
} catch {
    Write-Host "❌ Error updating registry: $_" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "🔧 Ensuring Azure TTS CLSID is registered..." -ForegroundColor Yellow

$azureClsidPath = "HKCR:\CLSID\$azureTtsClsid"
try {
    if (-not (Test-Path $azureClsidPath)) {
        Write-Host "Creating Azure TTS CLSID registration..." -ForegroundColor Gray
        New-Item -Path $azureClsidPath -Force | Out-Null
        Set-ItemProperty -Path $azureClsidPath -Name "(Default)" -Value "Azure TTS Engine"
        
        # Create InprocServer32 subkey
        $inprocPath = "$azureClsidPath\InprocServer32"
        New-Item -Path $inprocPath -Force | Out-Null
        Set-ItemProperty -Path $inprocPath -Name "(Default)" -Value $nativeDllPath
        Set-ItemProperty -Path $inprocPath -Name "ThreadingModel" -Value "Apartment"
        
        Write-Host "✅ Created Azure TTS CLSID registration" -ForegroundColor Green
    } else {
        Write-Host "✅ Azure TTS CLSID already registered" -ForegroundColor Green
    }
} catch {
    Write-Host "⚠️ Could not register Azure TTS CLSID: $_" -ForegroundColor Yellow
    Write-Host "This might be okay if the native DLL is already registered" -ForegroundColor Gray
}

Write-Host ""
Write-Host "🔧 Registering native DLL..." -ForegroundColor Yellow

if (Test-Path $nativeDllPath) {
    try {
        # Register the native DLL using regsvr32
        $result = Start-Process -FilePath "regsvr32.exe" -ArgumentList "/s", "`"$nativeDllPath`"" -Wait -PassThru
        if ($result.ExitCode -eq 0) {
            Write-Host "✅ Native DLL registered successfully" -ForegroundColor Green
        } else {
            Write-Host "⚠️ Native DLL registration returned exit code: $($result.ExitCode)" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "⚠️ Error registering native DLL: $_" -ForegroundColor Yellow
    }
} else {
    Write-Host "⚠️ Native DLL not found at: $nativeDllPath" -ForegroundColor Yellow
    Write-Host "You may need to copy it from: .\NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll" -ForegroundColor Gray
}

Write-Host ""
Write-Host "✅ Voice CLSID fix completed!" -ForegroundColor Green
Write-Host ""
Write-Host "🧪 Next steps:" -ForegroundColor Cyan
Write-Host "  1. Test the voice: .\TestSAPIVoices.ps1 -VoiceName 'Azure Libby' -PlayAudio" -ForegroundColor White
Write-Host "  2. If it still fails, check if the native DLL is properly built and copied" -ForegroundColor White
Write-Host "  3. The voice should now use the native Azure TTS COM wrapper" -ForegroundColor White

Write-Host ""
Write-Host "Fix completed!" -ForegroundColor Green
