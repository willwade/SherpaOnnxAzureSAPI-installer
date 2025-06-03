# Register Amy Voice for Native TTS Wrapper
Write-Host "Registering Amy Voice for Native TTS Wrapper..." -ForegroundColor Cyan

$ErrorActionPreference = "Stop"

# Check if running as administrator
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "ERROR: This script must be run as Administrator" -ForegroundColor Red
    Write-Host "Please run: sudo .\RegisterAmyVoice.ps1" -ForegroundColor Yellow
    exit 1
}

try {
    # Define the voice registry path
    $voicesPath = "HKLM:\SOFTWARE\Microsoft\Speech\Voices\Tokens"
    $amyVoicePath = "$voicesPath\amy"  # Use existing amy registration

    Write-Host "Creating Amy voice registry entry..." -ForegroundColor Yellow

    # Create the Amy voice key
    if (Test-Path $amyVoicePath) {
        Write-Host "   Removing existing Amy voice registration..." -ForegroundColor Yellow
        Remove-Item $amyVoicePath -Recurse -Force
    }

    New-Item -Path $amyVoicePath -Force | Out-Null

    # Set voice properties
    Write-Host "   Setting voice properties..." -ForegroundColor Yellow

    # Basic voice information
    Set-ItemProperty -Path $amyVoicePath -Name "(Default)" -Value "Amy (Piper Medium - SherpaOnnx)"
    Set-ItemProperty -Path $amyVoicePath -Name "VoiceName" -Value "amy"
    Set-ItemProperty -Path $amyVoicePath -Name "VoiceDesc" -Value "Amy - High Quality Neural Voice (Piper Medium via SherpaOnnx)"
    Set-ItemProperty -Path $amyVoicePath -Name "VoiceGender" -Value "Female"
    Set-ItemProperty -Path $amyVoicePath -Name "VoiceAge" -Value "Adult"
    Set-ItemProperty -Path $amyVoicePath -Name "VoiceVendor" -Value "OpenAssistive"
    Set-ItemProperty -Path $amyVoicePath -Name "VoiceLanguage" -Value "409" # English US

    # CLSID pointing to our native wrapper (from .idl file)
    $nativeWrapperCLSID = "{E1C4A8F2-9B3D-4A5E-8F7C-2D1B3E4F5A6B}"
    Set-ItemProperty -Path $amyVoicePath -Name "CLSID" -Value $nativeWrapperCLSID

    # CRITICAL: Remove the incorrect Path property - CLSID should be enough
    # The Path property is pointing to the wrong DLL (OpenSpeechTTS.dll instead of NativeTTSWrapper.dll)
    # Remove it so SAPI uses the CLSID registration instead
    try {
        Remove-ItemProperty -Path $amyVoicePath -Name "Path" -ErrorAction SilentlyContinue
        Write-Host "   ✅ Removed incorrect Path property - using CLSID registration" -ForegroundColor Green
    } catch {
        Write-Host "   ⚠️  Path property not found (this is good)" -ForegroundColor Yellow
    }

    # Voice attributes for SAPI (CRITICAL: This is what SAPI uses for voice selection)
    Set-ItemProperty -Path $amyVoicePath -Name "409" -Value "Name=amy;Gender=Female;Age=Adult;Vendor=OpenAssistive;Language=409;VoiceDesc=Amy - High Quality Neural Voice (Piper Medium via SherpaOnnx)"
    
    Write-Host "   Amy voice registered successfully!" -ForegroundColor Green
    
    # Verify registration
    Write-Host "Verifying registration..." -ForegroundColor Yellow
    
    if (Test-Path $amyVoicePath) {
        $voiceName = Get-ItemProperty -Path $amyVoicePath -Name "VoiceName" -ErrorAction SilentlyContinue
        if ($voiceName) {
            Write-Host "   SUCCESS: Amy voice found in registry" -ForegroundColor Green
            Write-Host "   Voice Name: $($voiceName.VoiceName)" -ForegroundColor Green
        } else {
            Write-Host "   ERROR: Voice properties not set correctly" -ForegroundColor Red
        }
    } else {
        Write-Host "   ERROR: Voice registry key not created" -ForegroundColor Red
    }
    
    # List all registered voices
    Write-Host "Current registered voices:" -ForegroundColor Yellow
    $allVoices = Get-ChildItem $voicesPath -ErrorAction SilentlyContinue
    foreach ($voice in $allVoices) {
        $voiceName = Get-ItemProperty -Path $voice.PSPath -Name "VoiceName" -ErrorAction SilentlyContinue
        if ($voiceName) {
            $voiceDesc = Get-ItemProperty -Path $voice.PSPath -Name "VoiceDesc" -ErrorAction SilentlyContinue
            Write-Host "   - $($voiceName.VoiceName): $($voiceDesc.VoiceDesc)" -ForegroundColor Gray
        }
    }
    
    Write-Host "Amy voice registration completed!" -ForegroundColor Cyan
    Write-Host "You can now test with: .\TestCOMWrapper.exe" -ForegroundColor Green
    
} catch {
    Write-Host "ERROR: Failed to register Amy voice: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
