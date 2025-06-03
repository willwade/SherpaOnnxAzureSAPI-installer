# Test if the issue is with managed COM vs native COM
Write-Host "Testing Managed COM vs Native COM Recognition" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Green

Write-Host "1. Checking if SAPI can work with managed COM objects at all..." -ForegroundColor Cyan

# Test with a simple managed COM object
try {
    # Create a simple .NET object that's COM visible
    $simpleObject = New-Object -ComObject "System.Collections.ArrayList"
    Write-Host "   ✅ Can create managed COM objects (ArrayList)" -ForegroundColor Green
} catch {
    Write-Host "   ❌ Cannot create managed COM objects: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "2. Checking our COM object registration details..." -ForegroundColor Cyan

# Check if our object is properly registered for COM
try {
    $regKey = Get-ItemProperty "HKLM:\SOFTWARE\Classes\CLSID\{3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2}\InprocServer32"
    Write-Host "   InprocServer32: $($regKey.'(default)')" -ForegroundColor White
    Write-Host "   ThreadingModel: $($regKey.ThreadingModel)" -ForegroundColor White
    Write-Host "   Class: $($regKey.Class)" -ForegroundColor White
    Write-Host "   Assembly: $($regKey.Assembly)" -ForegroundColor White
    Write-Host "   RuntimeVersion: $($regKey.RuntimeVersion)" -ForegroundColor White
} catch {
    Write-Host "   ❌ Error reading registry: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "3. Testing if SAPI has specific requirements for TTS engines..." -ForegroundColor Cyan

# Check if there are any SAPI-specific registry requirements
$voiceToken = "HKLM:\SOFTWARE\Microsoft\Speech\Voices\Tokens\amy"
try {
    $tokenProps = Get-ItemProperty $voiceToken
    Write-Host "   Voice token properties:" -ForegroundColor White
    Write-Host "     CLSID: $($tokenProps.CLSID)" -ForegroundColor Gray
    Write-Host "     Path: $($tokenProps.Path)" -ForegroundColor Gray
    
    # Check if there are any missing properties that Microsoft voices have
    $msVoiceToken = "HKLM:\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_EN-US_ZIRA_11.0"
    $msTokenProps = Get-ItemProperty $msVoiceToken
    
    Write-Host ""
    Write-Host "   Microsoft voice token properties:" -ForegroundColor White
    Write-Host "     CLSID: $($msTokenProps.CLSID)" -ForegroundColor Gray
    Write-Host "     LangDataPath: $($msTokenProps.LangDataPath)" -ForegroundColor Gray
    Write-Host "     VoicePath: $($msTokenProps.VoicePath)" -ForegroundColor Gray
    
    # Check for differences
    if($msTokenProps.LangDataPath -and -not $tokenProps.LangDataPath) {
        Write-Host "   ⚠️ Missing LangDataPath in our voice token" -ForegroundColor Yellow
    }
    if($msTokenProps.VoicePath -and -not $tokenProps.VoicePath) {
        Write-Host "   ⚠️ Missing VoicePath in our voice token" -ForegroundColor Yellow
    }
    
} catch {
    Write-Host "   ❌ Error reading voice tokens: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "=== HYPOTHESIS ===" -ForegroundColor Cyan
Write-Host "Based on research, the issue might be:" -ForegroundColor Yellow
Write-Host "1. SAPI doesn't work well with managed COM objects" -ForegroundColor White
Write-Host "2. Missing registry entries (LangDataPath, VoicePath)" -ForegroundColor White
Write-Host "3. Incorrect method signatures or calling conventions" -ForegroundColor White
Write-Host "4. Missing additional interfaces beyond ISpTTSEngine" -ForegroundColor White
Write-Host ""
Write-Host "SOLUTION: We may need to create a native C++ wrapper" -ForegroundColor Green
Write-Host "or fix the managed COM registration." -ForegroundColor Green
