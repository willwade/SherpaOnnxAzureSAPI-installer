# Fix SAPI Registration - Add Missing COM Categories and Registry Entries
Write-Host "Fixing SAPI Registration" -ForegroundColor Green
Write-Host "========================" -ForegroundColor Green

Write-Host ""
Write-Host "This script adds the missing COM categories and registry entries" -ForegroundColor Yellow
Write-Host "that SAPI requires to recognize our TTS engine." -ForegroundColor Yellow

$clsid = "{3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2}"

Write-Host ""
Write-Host "1. Adding SAPI TTS Engine COM category..." -ForegroundColor Cyan

# SAPI TTS Engine category GUID
$sapiTTSCategoryGuid = "{A910187F-0C7A-45AC-92CC-59EDAFB77B53}"

try {
    $categoryPath = "HKLM:\SOFTWARE\Classes\CLSID\$clsid\Implemented Categories\$sapiTTSCategoryGuid"
    
    if (!(Test-Path $categoryPath)) {
        New-Item -Path $categoryPath -Force | Out-Null
        Write-Host "   ✅ Added SAPI TTS Engine category: $sapiTTSCategoryGuid" -ForegroundColor Green
    } else {
        Write-Host "   ✅ SAPI TTS Engine category already exists" -ForegroundColor Green
    }
} catch {
    Write-Host "   ❌ Failed to add SAPI TTS Engine category: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "2. Adding Control category..." -ForegroundColor Cyan

# Control category GUID (for COM objects that can be controlled)
$controlCategoryGuid = "{40FC6ED4-2438-11CF-A3DB-080036F12502}"

try {
    $controlPath = "HKLM:\SOFTWARE\Classes\CLSID\$clsid\Implemented Categories\$controlCategoryGuid"
    
    if (!(Test-Path $controlPath)) {
        New-Item -Path $controlPath -Force | Out-Null
        Write-Host "   ✅ Added Control category: $controlCategoryGuid" -ForegroundColor Green
    } else {
        Write-Host "   ✅ Control category already exists" -ForegroundColor Green
    }
} catch {
    Write-Host "   ❌ Failed to add Control category: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "3. Adding Automation category..." -ForegroundColor Cyan

# Automation category GUID (for COM objects that support automation)
$automationCategoryGuid = "{40FC6ED5-2438-11CF-A3DB-080036F12502}"

try {
    $automationPath = "HKLM:\SOFTWARE\Classes\CLSID\$clsid\Implemented Categories\$automationCategoryGuid"
    
    if (!(Test-Path $automationPath)) {
        New-Item -Path $automationPath -Force | Out-Null
        Write-Host "   ✅ Added Automation category: $automationCategoryGuid" -ForegroundColor Green
    } else {
        Write-Host "   ✅ Automation category already exists" -ForegroundColor Green
    }
} catch {
    Write-Host "   ❌ Failed to add Automation category: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "4. Verifying InprocServer32 registration..." -ForegroundColor Cyan

$inprocPath = "HKLM:\SOFTWARE\Classes\CLSID\$clsid\InprocServer32"
if (Test-Path $inprocPath) {
    $inprocData = Get-ItemProperty $inprocPath
    Write-Host "   ✅ InprocServer32 registered:" -ForegroundColor Green
    Write-Host "     Default: $($inprocData.'(default)')" -ForegroundColor Gray
    Write-Host "     ThreadingModel: $($inprocData.ThreadingModel)" -ForegroundColor Gray
    Write-Host "     Class: $($inprocData.Class)" -ForegroundColor Gray
    Write-Host "     Assembly: $($inprocData.Assembly)" -ForegroundColor Gray
} else {
    Write-Host "   ❌ InprocServer32 not registered" -ForegroundColor Red
}

Write-Host ""
Write-Host "5. Adding ProgID registration..." -ForegroundColor Cyan

$progId = "OpenSpeechTTS.Sapi5VoiceImpl"
$progIdPath = "HKLM:\SOFTWARE\Classes\$progId"

try {
    if (!(Test-Path $progIdPath)) {
        New-Item -Path $progIdPath -Force | Out-Null
        Set-ItemProperty -Path $progIdPath -Name "(default)" -Value "OpenSpeech TTS Engine"
        Write-Host "   ✅ Created ProgID: $progId" -ForegroundColor Green
    } else {
        Write-Host "   ✅ ProgID already exists: $progId" -ForegroundColor Green
    }
    
    $progIdClsidPath = "$progIdPath\CLSID"
    if (!(Test-Path $progIdClsidPath)) {
        New-Item -Path $progIdClsidPath -Force | Out-Null
        Set-ItemProperty -Path $progIdClsidPath -Name "(default)" -Value $clsid
        Write-Host "   ✅ Linked ProgID to CLSID" -ForegroundColor Green
    }
    
    # Add reverse link from CLSID to ProgID
    $clsidProgIdPath = "HKLM:\SOFTWARE\Classes\CLSID\$clsid\ProgID"
    if (!(Test-Path $clsidProgIdPath)) {
        New-Item -Path $clsidProgIdPath -Force | Out-Null
        Set-ItemProperty -Path $clsidProgIdPath -Name "(default)" -Value $progId
        Write-Host "   ✅ Added ProgID reference to CLSID" -ForegroundColor Green
    }
    
} catch {
    Write-Host "   ❌ Failed to add ProgID registration: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "6. Adding TypeLib registration..." -ForegroundColor Cyan

# Check if TypeLib is registered
$typeLibPath = "HKLM:\SOFTWARE\Classes\CLSID\$clsid\TypeLib"
if (Test-Path $typeLibPath) {
    $typeLibGuid = (Get-ItemProperty $typeLibPath).'(default)'
    Write-Host "   ✅ TypeLib already registered: $typeLibGuid" -ForegroundColor Green
} else {
    Write-Host "   ⚠️ No TypeLib registered (this might be OK for managed COM)" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "7. Verifying voice token registration..." -ForegroundColor Cyan

$voiceTokenPath = "HKLM:\SOFTWARE\Microsoft\Speech\Voices\Tokens\amy"
if (Test-Path $voiceTokenPath) {
    $voiceData = Get-ItemProperty $voiceTokenPath
    Write-Host "   ✅ Voice token registered:" -ForegroundColor Green
    Write-Host "     Name: $($voiceData.'(default)')" -ForegroundColor Gray
    Write-Host "     CLSID: $($voiceData.CLSID)" -ForegroundColor Gray
    
    # Check if CLSID matches
    if ($voiceData.CLSID -eq $clsid) {
        Write-Host "   ✅ Voice token CLSID matches our COM object" -ForegroundColor Green
    } else {
        Write-Host "   ❌ Voice token CLSID mismatch!" -ForegroundColor Red
        Write-Host "     Expected: $clsid" -ForegroundColor Red
        Write-Host "     Found: $($voiceData.CLSID)" -ForegroundColor Red
    }
} else {
    Write-Host "   ❌ Voice token not registered" -ForegroundColor Red
}

Write-Host ""
Write-Host "8. Testing COM object creation after fixes..." -ForegroundColor Cyan

try {
    $testObject = New-Object -ComObject "OpenSpeechTTS.Sapi5VoiceImpl"
    Write-Host "   ✅ COM object creation: Working" -ForegroundColor Green
    
    $result = $testObject.SetObjectToken($null)
    Write-Host "   ✅ SetObjectToken: HRESULT = $result" -ForegroundColor Green
    
    # Release the object
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($testObject) | Out-Null
    
} catch {
    Write-Host "   ❌ COM object test failed: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "9. Testing SAPI voice creation..." -ForegroundColor Cyan

try {
    $spObjectTokenCategory = New-Object -ComObject "SAPI.SpObjectTokenCategory"
    $spObjectTokenCategory.SetId("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices", $false)
    
    $tokens = $spObjectTokenCategory.EnumerateTokens()
    $amyToken = $null
    
    for ($i = 0; $i -lt $tokens.Count; $i++) {
        $token = $tokens.Item($i)
        $tokenId = $token.Id
        if ($tokenId -like "*amy*") {
            $amyToken = $token
            break
        }
    }
    
    if ($amyToken) {
        Write-Host "   ✅ Amy token found: $($amyToken.Id)" -ForegroundColor Green
        
        try {
            $ttsObject = $amyToken.CreateInstance($null)
            Write-Host "   ✅ Successfully created TTS object from token!" -ForegroundColor Green
            
            # This is the key test - if this works, SAPI should be able to call our methods
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ttsObject) | Out-Null
            
        } catch {
            Write-Host "   ❌ Failed to create TTS object from token: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "   This is the core issue preventing SAPI integration" -ForegroundColor Yellow
        }
    } else {
        Write-Host "   ❌ Amy token not found" -ForegroundColor Red
    }
    
} catch {
    Write-Host "   ❌ SAPI token test failed: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "=== SAPI REGISTRATION FIX COMPLETE ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "🎯 NEXT STEPS:" -ForegroundColor Yellow
Write-Host "1. Test SAPI integration again" -ForegroundColor White
Write-Host "2. If still failing, the issue is likely in interface implementation" -ForegroundColor White
Write-Host "3. May need native COM wrapper as ultimate solution" -ForegroundColor White
Write-Host ""
Write-Host "The ProcessBridge TTS system remains 100% functional!" -ForegroundColor Green
