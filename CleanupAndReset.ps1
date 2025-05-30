# Complete Cleanup and Reset Script for OpenSpeechTTS COM Development
# This script must be run as Administrator
# Use this anytime you need to completely reset the COM registration for testing

Write-Host "=== OpenSpeechTTS Complete Cleanup and Reset ===" -ForegroundColor Cyan
Write-Host "This will completely remove all COM registrations and temp files" -ForegroundColor Yellow
Write-Host ""

# Stop any processes that might be using our DLLs
Write-Host "1. Stopping processes that might be using our DLLs..." -ForegroundColor Green
Get-Process | Where-Object { 
    $_.ProcessName -like "*speech*" -or 
    $_.ProcessName -like "*sapi*" -or 
    $_.ProcessName -like "*narrator*" -or
    $_.ProcessName -like "*cortana*"
} | Stop-Process -Force -ErrorAction SilentlyContinue

# Stop COM-related services
Write-Host "2. Stopping COM-related services..." -ForegroundColor Green
Stop-Service -Name "COMSysApp" -Force -ErrorAction SilentlyContinue
Stop-Service -Name "SENS" -Force -ErrorAction SilentlyContinue
Stop-Service -Name "AudioSrv" -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 3

# Unregister all versions of our COM components
Write-Host "3. Unregistering COM components..." -ForegroundColor Green

# Try to unregister from all possible locations
$locations = @(
    "C:\Program Files\OpenAssistive\OpenSpeech\OpenSpeechTTS.dll",
    "C:\Temp\OpenSpeechTTS.dll",
    "C:\Temp\OpenSpeechTTS_test.dll",
    "C:\Temp\OpenSpeechTTS_updated.dll"
)

foreach ($location in $locations) {
    if (Test-Path $location) {
        Write-Host "  Unregistering from: $location" -ForegroundColor White
        try {
            # Try regsvr32 first
            & regsvr32 /u /s $location 2>$null
            # Try RegAsm
            & "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe" $location /unregister 2>$null
        } catch {
            Write-Host "    Warning: Could not unregister $location" -ForegroundColor Yellow
        }
    }
}

# Remove registry entries completely
Write-Host "4. Removing registry entries..." -ForegroundColor Green

$registryPaths = @(
    "HKLM:\SOFTWARE\Classes\CLSID\{3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2}",
    "HKLM:\SOFTWARE\Classes\CLSID\{3d8f5c5e-9d6b-4b92-a12b-1a6dff80b6b3}",
    "HKCU:\SOFTWARE\Classes\CLSID\{3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2}",
    "HKCU:\SOFTWARE\Classes\CLSID\{3d8f5c5e-9d6b-4b92-a12b-1a6dff80b6b3}"
)

foreach ($path in $registryPaths) {
    if (Test-Path $path) {
        Write-Host "  Removing: $path" -ForegroundColor White
        Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# Remove TypeLib entries
Write-Host "5. Removing TypeLib entries..." -ForegroundColor Green
Get-ChildItem "HKLM:\SOFTWARE\Classes\TypeLib\" -ErrorAction SilentlyContinue | 
    Where-Object { $_.Name -like "*OpenSpeech*" } | 
    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

# Clean up temp files
Write-Host "6. Cleaning up temp files..." -ForegroundColor Green
$tempFiles = @(
    "C:\Temp\OpenSpeechTTS.dll",
    "C:\Temp\OpenSpeechTTS_test.dll", 
    "C:\Temp\OpenSpeechTTS_updated.dll",
    "C:\Temp\SherpaNative.dll",
    "C:\Temp\sherpa-onnx.dll"
)

foreach ($file in $tempFiles) {
    if (Test-Path $file) {
        Write-Host "  Removing: $file" -ForegroundColor White
        Remove-Item $file -Force -ErrorAction SilentlyContinue
    }
}

# Clean up log files
Write-Host "7. Cleaning up log files..." -ForegroundColor Green
$logFiles = @(
    "C:\OpenSpeech\sapi_debug.log",
    "C:\OpenSpeech\sapi_error.log",
    "C:\OpenSpeech\sapi_speak.log"
)

foreach ($file in $logFiles) {
    if (Test-Path $file) {
        Write-Host "  Removing: $file" -ForegroundColor White
        Remove-Item $file -Force -ErrorAction SilentlyContinue
    }
}

# Restart services
Write-Host "8. Restarting services..." -ForegroundColor Green
Start-Service -Name "AudioSrv" -ErrorAction SilentlyContinue
Start-Service -Name "COMSysApp" -ErrorAction SilentlyContinue
Start-Service -Name "SENS" -ErrorAction SilentlyContinue
Start-Sleep -Seconds 3

# Clear .NET assembly cache
Write-Host "9. Clearing .NET assembly cache..." -ForegroundColor Green
try {
    & "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\ngen.exe" uninstall "OpenSpeechTTS" 2>$null
} catch {
    # Ignore errors - assembly might not be in cache
}

Write-Host ""
Write-Host "=== CLEANUP COMPLETE ===" -ForegroundColor Cyan
Write-Host "All COM registrations, temp files, and logs have been removed." -ForegroundColor Green
Write-Host "The system is now in a clean state for fresh testing." -ForegroundColor Green
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "1. Build your project: dotnet build OpenSpeechTTS/OpenSpeechTTS.csproj --configuration Release" -ForegroundColor White
Write-Host "2. Register your DLL: regsvr32 or RegAsm as needed" -ForegroundColor White
Write-Host "3. Test with your test scripts" -ForegroundColor White
Write-Host ""

# Verify cleanup
Write-Host "Verification - checking if cleanup was successful:" -ForegroundColor Yellow
$clsidExists = Test-Path "HKLM:\SOFTWARE\Classes\CLSID\{3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2}"
$tempExists = Test-Path "C:\Temp\OpenSpeechTTS*.dll"
$logsExist = Test-Path "C:\OpenSpeech\sapi_*.log"

if (-not $clsidExists -and -not $tempExists -and -not $logsExist) {
    Write-Host "✅ Cleanup successful - system is clean!" -ForegroundColor Green
} else {
    Write-Host "⚠️  Some items may still exist:" -ForegroundColor Yellow
    if ($clsidExists) { Write-Host "  - Registry entries still present" -ForegroundColor Red }
    if ($tempExists) { Write-Host "  - Temp files still present" -ForegroundColor Red }
    if ($logsExist) { Write-Host "  - Log files still present" -ForegroundColor Red }
}

Write-Host ""
Write-Host "Press any key to continue..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
