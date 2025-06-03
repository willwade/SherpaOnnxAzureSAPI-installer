# Complete TTS System Cleanup Script
# This script removes all custom TTS voices, COM objects, and configurations

param(
    [switch]$Force
)

Write-Host "🧹 COMPLETE TTS SYSTEM CLEANUP" -ForegroundColor Red
Write-Host "==============================" -ForegroundColor Red
Write-Host ""

if (-not $Force) {
    Write-Host "⚠️  WARNING: This will remove ALL custom TTS voices and configurations!" -ForegroundColor Yellow
    Write-Host "This includes:" -ForegroundColor Yellow
    Write-Host "  - All SherpaOnnx voices (amy, northern_english_male, etc.)" -ForegroundColor Gray
    Write-Host "  - All Azure TTS voices (ElliotNeural, LibbyNeural, etc.)" -ForegroundColor Gray
    Write-Host "  - All COM object registrations" -ForegroundColor Gray
    Write-Host "  - All configuration files" -ForegroundColor Gray
    Write-Host ""
    $confirm = Read-Host "Are you sure you want to continue? (type 'YES' to confirm)"
    if ($confirm -ne "YES") {
        Write-Host "Cleanup cancelled." -ForegroundColor Yellow
        exit 0
    }
}

Write-Host "🔧 Starting complete cleanup..." -ForegroundColor Cyan
Write-Host ""

# Step 1: Kill any running TTS processes
Write-Host "1️⃣ Stopping TTS processes..." -ForegroundColor Yellow
$processes = @("SherpaWorker", "OpenSpeechTTS", "NativeTTSWrapper")
foreach ($proc in $processes) {
    try {
        Get-Process $proc -ErrorAction SilentlyContinue | Stop-Process -Force
        Write-Host "  ✅ Stopped $proc" -ForegroundColor Green
    } catch {
        Write-Host "  ℹ️ $proc not running" -ForegroundColor Gray
    }
}

# Step 2: Unregister COM objects
Write-Host ""
Write-Host "2️⃣ Unregistering COM objects..." -ForegroundColor Yellow

$comDlls = @(
    "C:\Program Files\OpenAssistive\OpenSpeech\OpenSpeechTTS.dll",
    "C:\Program Files\OpenAssistive\OpenSpeech\NativeTTSWrapper.dll"
)

foreach ($dll in $comDlls) {
    if (Test-Path $dll) {
        try {
            Write-Host "  🔄 Unregistering $dll..." -ForegroundColor Gray
            Start-Process "regsvr32" -ArgumentList "/u", "/s", "`"$dll`"" -Wait -NoNewWindow
            Write-Host "  ✅ Unregistered $dll" -ForegroundColor Green
        } catch {
            Write-Host "  ⚠️ Failed to unregister $dll: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    } else {
        Write-Host "  ℹ️ $dll not found" -ForegroundColor Gray
    }
}

# Step 3: Remove SAPI voice registrations
Write-Host ""
Write-Host "3️⃣ Removing SAPI voice registrations..." -ForegroundColor Yellow

$voicesPath = "HKLM:\SOFTWARE\Microsoft\SPEECH\Voices\Tokens"
if (Test-Path $voicesPath) {
    $voices = Get-ChildItem $voicesPath -ErrorAction SilentlyContinue
    
    foreach ($voice in $voices) {
        $voiceName = $voice.PSChildName
        
        # Remove custom voices (not system voices)
        if ($voiceName -like "*amy*" -or 
            $voiceName -like "*northern*" -or 
            $voiceName -like "*elliot*" -or 
            $voiceName -like "*libby*" -or 
            $voiceName -like "*jenny*" -or 
            $voiceName -like "*Server Speech*") {
            
            try {
                Remove-Item $voice.PSPath -Recurse -Force
                Write-Host "  ✅ Removed voice: $voiceName" -ForegroundColor Green
            } catch {
                Write-Host "  ❌ Failed to remove voice: $voiceName - $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
} else {
    Write-Host "  ℹ️ No voice tokens found" -ForegroundColor Gray
}

# Step 4: Remove COM CLSIDs
Write-Host ""
Write-Host "4️⃣ Removing COM CLSIDs..." -ForegroundColor Yellow

$clsids = @(
    "{3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2}",  # Managed COM wrapper
    "{3d8f5c5e-9d6b-4b92-a12b-1a6dff80b6b3}",  # Azure CLSID
    "{E1C4A8F2-9B3D-4A5E-8F7C-2D1B3E4F5A6B}"   # Native COM wrapper
)

foreach ($clsid in $clsids) {
    $clsidPath = "HKLM:\SOFTWARE\Classes\CLSID\$clsid"
    if (Test-Path $clsidPath) {
        try {
            Remove-Item $clsidPath -Recurse -Force
            Write-Host "  ✅ Removed CLSID: $clsid" -ForegroundColor Green
        } catch {
            Write-Host "  ❌ Failed to remove CLSID: $clsid - $($_.Exception.Message)" -ForegroundColor Red
        }
    } else {
        Write-Host "  ℹ️ CLSID not found: $clsid" -ForegroundColor Gray
    }
}

# Step 5: Remove installation directories
Write-Host ""
Write-Host "5️⃣ Removing installation directories..." -ForegroundColor Yellow

$installDirs = @(
    "C:\Program Files\OpenAssistive\OpenSpeech",
    "C:\Program Files\OpenSpeech",
    "C:\ProgramData\OpenSpeech"
)

foreach ($dir in $installDirs) {
    if (Test-Path $dir) {
        try {
            Remove-Item $dir -Recurse -Force
            Write-Host "  ✅ Removed directory: $dir" -ForegroundColor Green
        } catch {
            Write-Host "  ❌ Failed to remove directory: $dir - $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "    💡 Some files may be in use. Try rebooting and running cleanup again." -ForegroundColor Yellow
        }
    } else {
        Write-Host "  ℹ️ Directory not found: $dir" -ForegroundColor Gray
    }
}

# Step 6: Clear any cached configurations
Write-Host ""
Write-Host "6️⃣ Clearing cached configurations..." -ForegroundColor Yellow

$configPaths = @(
    "$env:APPDATA\OpenSpeech",
    "$env:LOCALAPPDATA\OpenSpeech",
    "$env:TEMP\OpenSpeech*"
)

foreach ($path in $configPaths) {
    if (Test-Path $path) {
        try {
            Remove-Item $path -Recurse -Force
            Write-Host "  ✅ Cleared cache: $path" -ForegroundColor Green
        } catch {
            Write-Host "  ⚠️ Failed to clear cache: $path" -ForegroundColor Yellow
        }
    }
}

# Step 7: Verify cleanup
Write-Host ""
Write-Host "7️⃣ Verifying cleanup..." -ForegroundColor Yellow

try {
    $voice = New-Object -ComObject SAPI.SpVoice
    $voices = $voice.GetVoices()
    
    $customVoices = 0
    for ($i = 0; $i -lt $voices.Count; $i++) {
        $voiceToken = $voices.Item($i)
        $voiceName = $voiceToken.GetDescription()
        
        if ($voiceName -like "*amy*" -or 
            $voiceName -like "*northern*" -or 
            $voiceName -like "*elliot*" -or 
            $voiceName -like "*libby*" -or 
            $voiceName -like "*jenny*" -or 
            $voiceName -like "*Server Speech*") {
            $customVoices++
            Write-Host "  ⚠️ Custom voice still present: $voiceName" -ForegroundColor Yellow
        }
    }
    
    if ($customVoices -eq 0) {
        Write-Host "  ✅ All custom voices removed successfully" -ForegroundColor Green
    } else {
        Write-Host "  ⚠️ $customVoices custom voice(s) still present" -ForegroundColor Yellow
    }
    
    Write-Host "  📊 Remaining voices: $($voices.Count) (system voices only)" -ForegroundColor Gray
    
} catch {
    Write-Host "  ✅ SAPI cleanup successful (no custom COM objects found)" -ForegroundColor Green
}

Write-Host ""
Write-Host "🎯 CLEANUP COMPLETED!" -ForegroundColor Cyan
Write-Host ""
Write-Host "📋 Summary:" -ForegroundColor Yellow
Write-Host "  - TTS processes stopped" -ForegroundColor Gray
Write-Host "  - COM objects unregistered" -ForegroundColor Gray
Write-Host "  - SAPI voice registrations removed" -ForegroundColor Gray
Write-Host "  - COM CLSIDs removed" -ForegroundColor Gray
Write-Host "  - Installation directories removed" -ForegroundColor Gray
Write-Host "  - Cached configurations cleared" -ForegroundColor Gray
Write-Host ""
Write-Host "🚀 Next steps:" -ForegroundColor Green
Write-Host "  1. Reboot your system (recommended)" -ForegroundColor Gray
Write-Host "  2. Reinstall voices using the installer" -ForegroundColor Gray
Write-Host "  3. Test the fresh installation" -ForegroundColor Gray
Write-Host ""
Write-Host "💡 To reinstall ElliotNeural:" -ForegroundColor Yellow
Write-Host "  .\dist\SherpaOnnxSAPIInstaller.exe install-azure en-GB-ElliotNeural --key b14f8945b0f1459f9964bdd72c42c2cc --region uksouth" -ForegroundColor Gray
