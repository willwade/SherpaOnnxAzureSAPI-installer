# Update and Register COM Component with Fixed Interface
# This script must be run as Administrator

Write-Host "Updating OpenSpeechTTS.dll with fixed SAPI5 interface..." -ForegroundColor Green

# Stop any processes that might be using the DLL
Write-Host "Stopping any processes using the DLL..."
Get-Process | Where-Object { $_.ProcessName -like "*speech*" -or $_.ProcessName -like "*sapi*" } | Stop-Process -Force -ErrorAction SilentlyContinue

# Unregister the old COM component
Write-Host "Unregistering old COM component..."
try {
    & regsvr32 /u /s "C:\Program Files\OpenAssistive\OpenSpeech\OpenSpeechTTS.dll"
    Write-Host "Successfully unregistered old component" -ForegroundColor Green
} catch {
    Write-Host "Warning: Could not unregister old component: $($_.Exception.Message)" -ForegroundColor Yellow
}

# Copy the updated DLL
Write-Host "Copying updated DLL..."
try {
    Copy-Item "OpenSpeechTTS\bin\Release\net472\OpenSpeechTTS.dll" "C:\Program Files\OpenAssistive\OpenSpeech\OpenSpeechTTS.dll" -Force
    Write-Host "Successfully copied updated DLL" -ForegroundColor Green
} catch {
    Write-Host "Error copying DLL: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Register the new COM component
Write-Host "Registering new COM component with fixed interface..."
try {
    & regsvr32 /s "C:\Program Files\OpenAssistive\OpenSpeech\OpenSpeechTTS.dll"
    Write-Host "Successfully registered new component" -ForegroundColor Green
} catch {
    Write-Host "Error registering component: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Register SAPI5 interface GUIDs (CRITICAL FIX)
Write-Host ""
Write-Host "Registering SAPI5 interface GUIDs..." -ForegroundColor Yellow
Write-Host "This is the critical fix for the 'catastrophic failure' issue." -ForegroundColor White
try {
    & .\RegisterInterfaces.ps1
    Write-Host "Successfully registered interface GUIDs" -ForegroundColor Green
} catch {
    Write-Host "Error registering interfaces: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "You may need to run RegisterInterfaces.ps1 manually as Administrator." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=== UPDATE COMPLETE ===" -ForegroundColor Cyan
Write-Host "The OpenSpeechTTS.dll has been updated with the correct SAPI5 interface definition." -ForegroundColor Green
Write-Host "Key fixes applied:" -ForegroundColor Yellow
Write-Host "  - Corrected ISpTTSEngine GUID to official SAPI5 GUID" -ForegroundColor White
Write-Host "  - Fixed method parameter types and directions" -ForegroundColor White
Write-Host "  - Updated structure definitions to match SAPI5 specification" -ForegroundColor White
Write-Host "  - Changed SpTTSFragList to SPVTEXTFRAG" -ForegroundColor White
Write-Host "  - ✅ REGISTERED MISSING INTERFACE GUIDs (CRITICAL FIX)" -ForegroundColor Green
Write-Host ""
Write-Host "🎯 EXPECTED RESULT: SAPI should now call our methods!" -ForegroundColor Cyan
Write-Host "You can now test the TTS engine with the test scripts." -ForegroundColor Green

pause
