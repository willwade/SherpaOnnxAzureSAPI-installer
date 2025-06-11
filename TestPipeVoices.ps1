# Test script for pipe-based SAPI voices
# This script demonstrates the new configuration-based voice system

Write-Host "Testing Pipe-Based SAPI Voices" -ForegroundColor Green
Write-Host "===============================" -ForegroundColor Green
Write-Host ""

# Check if running as administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin) {
    Write-Host "This script requires administrator privileges." -ForegroundColor Red
    Write-Host "Please run PowerShell as Administrator and try again." -ForegroundColor Red
    exit 1
}

# Build the installer
Write-Host "Building installer..." -ForegroundColor Yellow
try {
    $buildResult = dotnet build Installer/Installer.csproj -c Release
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Build failed!" -ForegroundColor Red
        exit 1
    }
    Write-Host "Build successful!" -ForegroundColor Green
} catch {
    Write-Host "Error building installer: $_" -ForegroundColor Red
    exit 1
}

$installerPath = "Installer\bin\Release\net6.0\Installer.exe"
if (-not (Test-Path $installerPath)) {
    Write-Host "Installer executable not found at: $installerPath" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Testing pipe service connection..." -ForegroundColor Yellow
& $installerPath test-pipe-service

Write-Host ""
Write-Host "Listing available pipe voice configurations..." -ForegroundColor Yellow
if (Test-Path "voice_configs") {
    $configs = Get-ChildItem "voice_configs" -Filter "*.json"
    if ($configs.Count -gt 0) {
        Write-Host "Available configurations:" -ForegroundColor Cyan
        foreach ($config in $configs) {
            $configName = [System.IO.Path]::GetFileNameWithoutExtension($config.Name)
            Write-Host "  - $configName" -ForegroundColor White
        }
    } else {
        Write-Host "No voice configurations found in voice_configs directory" -ForegroundColor Yellow
    }
} else {
    Write-Host "voice_configs directory not found" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Installing pipe-based voices..." -ForegroundColor Yellow

# Install each available voice configuration
if (Test-Path "voice_configs") {
    $configs = Get-ChildItem "voice_configs" -Filter "*.json"
    foreach ($config in $configs) {
        $configName = [System.IO.Path]::GetFileNameWithoutExtension($config.Name)
        Write-Host "Installing voice: $configName" -ForegroundColor Cyan
        
        try {
            & $installerPath install-pipe-voice $configName
            if ($LASTEXITCODE -eq 0) {
                Write-Host "  ✅ Successfully installed $configName" -ForegroundColor Green
            } else {
                Write-Host "  ❌ Failed to install $configName" -ForegroundColor Red
            }
        } catch {
            Write-Host "  ❌ Error installing $configName : $_" -ForegroundColor Red
        }
    }
}

Write-Host ""
Write-Host "Listing installed pipe voices..." -ForegroundColor Yellow
& $installerPath list-pipe-voices

Write-Host ""
Write-Host "Testing SAPI voice enumeration..." -ForegroundColor Yellow
try {
    # Use PowerShell to enumerate SAPI voices
    Add-Type -AssemblyName System.Speech
    $synthesizer = New-Object System.Speech.Synthesis.SpeechSynthesizer
    $voices = $synthesizer.GetInstalledVoices()
    
    Write-Host "Installed SAPI voices:" -ForegroundColor Cyan
    foreach ($voice in $voices) {
        $info = $voice.VoiceInfo
        $name = $info.Name
        $culture = $info.Culture.Name
        $gender = $info.Gender
        $age = $info.Age
        
        # Check if this is one of our pipe-based voices
        if ($name -like "*British-English*" -or $name -like "*American-English*") {
            Write-Host "  🎤 $name ($culture, $gender, $age) [PIPE-BASED]" -ForegroundColor Green
        } else {
            Write-Host "  🔊 $name ($culture, $gender, $age)" -ForegroundColor White
        }
    }
    
    $synthesizer.Dispose()
} catch {
    Write-Host "Error enumerating SAPI voices: $_" -ForegroundColor Red
}

Write-Host ""
Write-Host "Testing voice synthesis (if AACSpeakHelper is running)..." -ForegroundColor Yellow

# Test synthesis with a pipe-based voice
try {
    Add-Type -AssemblyName System.Speech
    $synthesizer = New-Object System.Speech.Synthesis.SpeechSynthesizer
    $voices = $synthesizer.GetInstalledVoices()
    
    # Find a pipe-based voice
    $pipeVoice = $voices | Where-Object { $_.VoiceInfo.Name -like "*British-English*" -or $_.VoiceInfo.Name -like "*American-English*" } | Select-Object -First 1
    
    if ($pipeVoice) {
        $voiceName = $pipeVoice.VoiceInfo.Name
        Write-Host "Testing synthesis with voice: $voiceName" -ForegroundColor Cyan
        
        $synthesizer.SelectVoice($voiceName)
        $synthesizer.Speak("Hello, this is a test of the pipe-based SAPI voice system.")
        
        Write-Host "  ✅ Synthesis test completed" -ForegroundColor Green
    } else {
        Write-Host "  ⚠️  No pipe-based voices found for testing" -ForegroundColor Yellow
    }
    
    $synthesizer.Dispose()
} catch {
    Write-Host "  ❌ Error testing synthesis: $_" -ForegroundColor Red
    Write-Host "  This is expected if AACSpeakHelper pipe service is not running" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Test completed!" -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "1. Start AACSpeakHelper server (AACSpeakHelperServer.py)" -ForegroundColor White
Write-Host "2. Test voice synthesis using Windows Speech API" -ForegroundColor White
Write-Host "3. Use voices in applications that support SAPI" -ForegroundColor White
Write-Host ""
Write-Host "To remove pipe voices, use:" -ForegroundColor Cyan
Write-Host "  .\Installer.exe remove-pipe-voice <voice-name>" -ForegroundColor White
