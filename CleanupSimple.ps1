# Simple Complete Cleanup Script
Write-Host "Complete TTS System Cleanup" -ForegroundColor Red
Write-Host "===========================" -ForegroundColor Red
Write-Host ""

# Step 1: Stop processes
Write-Host "1. Stopping TTS processes..." -ForegroundColor Yellow
try { Stop-Process -Name "SherpaWorker" -Force -ErrorAction SilentlyContinue } catch {}
try { Stop-Process -Name "OpenSpeechTTS" -Force -ErrorAction SilentlyContinue } catch {}
Write-Host "   Processes stopped" -ForegroundColor Green

# Step 2: Unregister COM DLLs
Write-Host ""
Write-Host "2. Unregistering COM objects..." -ForegroundColor Yellow

$dlls = @(
    "C:\Program Files\OpenAssistive\OpenSpeech\OpenSpeechTTS.dll",
    "C:\Program Files\OpenAssistive\OpenSpeech\NativeTTSWrapper.dll"
)

foreach ($dll in $dlls) {
    if (Test-Path $dll) {
        Write-Host "   Unregistering $dll..." -ForegroundColor Gray
        try {
            Start-Process "regsvr32" -ArgumentList "/u", "/s", "`"$dll`"" -Wait -NoNewWindow
            Write-Host "   OK: Unregistered $dll" -ForegroundColor Green
        } catch {
            Write-Host "   Warning: Failed to unregister $dll" -ForegroundColor Yellow
        }
    }
}

# Step 3: Remove voice registrations
Write-Host ""
Write-Host "3. Removing SAPI voice registrations..." -ForegroundColor Yellow

$voicesPath = "HKLM:\SOFTWARE\Microsoft\SPEECH\Voices\Tokens"
if (Test-Path $voicesPath) {
    $voices = Get-ChildItem $voicesPath -ErrorAction SilentlyContinue
    
    foreach ($voice in $voices) {
        $voiceName = $voice.PSChildName
        
        # Remove custom voices
        if ($voiceName -match "amy|northern|elliot|libby|jenny|Server Speech") {
            try {
                Remove-Item $voice.PSPath -Recurse -Force
                Write-Host "   Removed: $voiceName" -ForegroundColor Green
            } catch {
                Write-Host "   Failed to remove: $voiceName" -ForegroundColor Red
            }
        }
    }
}

# Step 4: Remove COM CLSIDs
Write-Host ""
Write-Host "4. Removing COM CLSIDs..." -ForegroundColor Yellow

$clsids = @(
    "{3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2}",
    "{3d8f5c5e-9d6b-4b92-a12b-1a6dff80b6b3}",
    "{E1C4A8F2-9B3D-4A5E-8F7C-2D1B3E4F5A6B}"
)

foreach ($clsid in $clsids) {
    $clsidPath = "HKLM:\SOFTWARE\Classes\CLSID\$clsid"
    if (Test-Path $clsidPath) {
        try {
            Remove-Item $clsidPath -Recurse -Force
            Write-Host "   Removed CLSID: $clsid" -ForegroundColor Green
        } catch {
            Write-Host "   Failed to remove CLSID: $clsid" -ForegroundColor Red
        }
    }
}

# Step 5: Remove directories
Write-Host ""
Write-Host "5. Removing installation directories..." -ForegroundColor Yellow

$dirs = @(
    "C:\Program Files\OpenAssistive\OpenSpeech",
    "C:\Program Files\OpenSpeech"
)

foreach ($dir in $dirs) {
    if (Test-Path $dir) {
        try {
            Remove-Item $dir -Recurse -Force
            Write-Host "   Removed: $dir" -ForegroundColor Green
        } catch {
            Write-Host "   Failed to remove: $dir (files may be in use)" -ForegroundColor Yellow
        }
    }
}

# Step 6: Verify cleanup
Write-Host ""
Write-Host "6. Verifying cleanup..." -ForegroundColor Yellow

try {
    $voice = New-Object -ComObject SAPI.SpVoice
    $voices = $voice.GetVoices()
    
    $customVoices = 0
    for ($i = 0; $i -lt $voices.Count; $i++) {
        $voiceToken = $voices.Item($i)
        $voiceName = $voiceToken.GetDescription()
        
        if ($voiceName -match "amy|northern|elliot|libby|jenny|Server Speech") {
            $customVoices++
            Write-Host "   Still present: $voiceName" -ForegroundColor Yellow
        }
    }
    
    if ($customVoices -eq 0) {
        Write-Host "   All custom voices removed!" -ForegroundColor Green
    } else {
        Write-Host "   $customVoices custom voices still present" -ForegroundColor Yellow
    }
    
    Write-Host "   Total voices remaining: $($voices.Count)" -ForegroundColor Gray
    
} catch {
    Write-Host "   SAPI cleanup successful" -ForegroundColor Green
}

Write-Host ""
Write-Host "CLEANUP COMPLETED!" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "1. Reboot system (recommended)" -ForegroundColor Gray
Write-Host "2. Reinstall voices with installer" -ForegroundColor Gray
Write-Host ""
Write-Host "To reinstall ElliotNeural:" -ForegroundColor Green
Write-Host "sudo .\dist\SherpaOnnxSAPIInstaller.exe install-azure en-GB-ElliotNeural --key b14f8945b0f1459f9964bdd72c42c2cc --region uksouth" -ForegroundColor Gray
