# Test COM Interface Methods Directly
Write-Host "Testing COM Interface Methods Directly" -ForegroundColor Green
Write-Host "======================================" -ForegroundColor Green

Write-Host ""
Write-Host "This test calls our COM object methods directly to verify they work" -ForegroundColor Yellow
Write-Host "and identifies why SAPI doesn't call them." -ForegroundColor Yellow

Write-Host ""
Write-Host "1. Creating COM object..." -ForegroundColor Cyan

try {
    $comObject = New-Object -ComObject "OpenSpeechTTS.Sapi5VoiceImpl"
    Write-Host "   ✅ COM object created successfully" -ForegroundColor Green
} catch {
    Write-Host "   ❌ Failed to create COM object: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "2. Testing SetObjectToken..." -ForegroundColor Cyan

try {
    $result = $comObject.SetObjectToken($null)
    Write-Host "   ✅ SetObjectToken: HRESULT = $result" -ForegroundColor Green
} catch {
    Write-Host "   ❌ SetObjectToken failed: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "3. Testing GetObjectToken..." -ForegroundColor Cyan

try {
    $token = $null
    $result = $comObject.GetObjectToken([ref]$token)
    Write-Host "   ✅ GetObjectToken: HRESULT = $result, Token = $token" -ForegroundColor Green
} catch {
    Write-Host "   ❌ GetObjectToken failed: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "4. Testing interface queries..." -ForegroundColor Cyan

# Test if the object supports the required interfaces
$iSpTTSEngineGuid = [System.Guid]::new("A74D7C8E-4CC5-4F2F-A6EB-804DEE18500E")
$iSpObjectWithTokenGuid = [System.Guid]::new("14056581-E16C-11D2-BB90-00C04F8EE6C0")

try {
    # Get the IUnknown interface
    $unknown = [System.Runtime.InteropServices.Marshal]::GetIUnknownForObject($comObject)
    Write-Host "   ✅ Got IUnknown interface: $unknown" -ForegroundColor Green
    
    # Try to query for ISpTTSEngine
    $ttsEnginePtr = [System.IntPtr]::Zero
    $hr = [System.Runtime.InteropServices.Marshal]::QueryInterface($unknown, [ref]$iSpTTSEngineGuid, [ref]$ttsEnginePtr)
    Write-Host "   QueryInterface for ISpTTSEngine: HRESULT = 0x$($hr.ToString('X8'))" -ForegroundColor White
    
    if ($hr -eq 0) {
        Write-Host "   ✅ ISpTTSEngine interface supported" -ForegroundColor Green
        [System.Runtime.InteropServices.Marshal]::Release($ttsEnginePtr)
    } else {
        Write-Host "   ❌ ISpTTSEngine interface NOT supported" -ForegroundColor Red
    }
    
    # Try to query for ISpObjectWithToken
    $tokenPtr = [System.IntPtr]::Zero
    $hr = [System.Runtime.InteropServices.Marshal]::QueryInterface($unknown, [ref]$iSpObjectWithTokenGuid, [ref]$tokenPtr)
    Write-Host "   QueryInterface for ISpObjectWithToken: HRESULT = 0x$($hr.ToString('X8'))" -ForegroundColor White
    
    if ($hr -eq 0) {
        Write-Host "   ✅ ISpObjectWithToken interface supported" -ForegroundColor Green
        [System.Runtime.InteropServices.Marshal]::Release($tokenPtr)
    } else {
        Write-Host "   ❌ ISpObjectWithToken interface NOT supported" -ForegroundColor Red
    }
    
    [System.Runtime.InteropServices.Marshal]::Release($unknown)
    
} catch {
    Write-Host "   ❌ Interface query failed: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "5. Testing SAPI object creation..." -ForegroundColor Cyan

try {
    # Test creating our object through SAPI's mechanism
    $spObjectTokenCategory = New-Object -ComObject "SAPI.SpObjectTokenCategory"
    $spObjectTokenCategory.SetId("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices", $false)
    
    $tokens = $spObjectTokenCategory.EnumerateTokens()
    Write-Host "   Found $($tokens.Count) voice tokens" -ForegroundColor White
    
    $amyToken = $null
    for ($i = 0; $i -lt $tokens.Count; $i++) {
        $token = $tokens.Item($i)
        $tokenId = $token.Id
        if ($tokenId -like "*amy*") {
            $amyToken = $token
            Write-Host "   ✅ Found Amy token: $tokenId" -ForegroundColor Green
            break
        }
    }
    
    if ($amyToken) {
        Write-Host "   Testing object creation from token..." -ForegroundColor Yellow
        
        try {
            $ttsObject = $amyToken.CreateInstance($null)
            Write-Host "   ✅ Created TTS object from token" -ForegroundColor Green
            
            # Test if this object supports the interfaces
            $unknown2 = [System.Runtime.InteropServices.Marshal]::GetIUnknownForObject($ttsObject)
            
            $ttsEnginePtr2 = [System.IntPtr]::Zero
            $hr2 = [System.Runtime.InteropServices.Marshal]::QueryInterface($unknown2, [ref]$iSpTTSEngineGuid, [ref]$ttsEnginePtr2)
            Write-Host "   Token object ISpTTSEngine query: HRESULT = 0x$($hr2.ToString('X8'))" -ForegroundColor White
            
            if ($hr2 -eq 0) {
                Write-Host "   ✅ Token object supports ISpTTSEngine" -ForegroundColor Green
                [System.Runtime.InteropServices.Marshal]::Release($ttsEnginePtr2)
            } else {
                Write-Host "   ❌ Token object does NOT support ISpTTSEngine" -ForegroundColor Red
            }
            
            [System.Runtime.InteropServices.Marshal]::Release($unknown2)
            
        } catch {
            Write-Host "   ❌ Failed to create object from token: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
} catch {
    Write-Host "   ❌ SAPI object creation test failed: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "6. Checking COM registration details..." -ForegroundColor Cyan

$clsid = "{3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2}"
$clsidPath = "HKLM:\SOFTWARE\Classes\CLSID\$clsid"

if (Test-Path $clsidPath) {
    Write-Host "   ✅ CLSID registered: $clsid" -ForegroundColor Green
    
    $inprocPath = "$clsidPath\InprocServer32"
    if (Test-Path $inprocPath) {
        $inprocData = Get-ItemProperty $inprocPath
        Write-Host "   InprocServer32: $($inprocData.'(default)')" -ForegroundColor White
        Write-Host "   ThreadingModel: $($inprocData.ThreadingModel)" -ForegroundColor White
        Write-Host "   Class: $($inprocData.Class)" -ForegroundColor White
        Write-Host "   Assembly: $($inprocData.Assembly)" -ForegroundColor White
    }
    
    # Check for implemented interfaces
    $implementedPath = "$clsidPath\Implemented Categories"
    if (Test-Path $implementedPath) {
        Write-Host "   ✅ Implemented Categories found" -ForegroundColor Green
        Get-ChildItem $implementedPath | ForEach-Object {
            Write-Host "     Category: $($_.PSChildName)" -ForegroundColor Gray
        }
    } else {
        Write-Host "   ⚠️ No Implemented Categories found" -ForegroundColor Yellow
    }
} else {
    Write-Host "   ❌ CLSID not registered: $clsid" -ForegroundColor Red
}

Write-Host ""
Write-Host "7. Checking recent logs..." -ForegroundColor Cyan

$logDir = "C:\OpenSpeech"
if (Test-Path $logDir) {
    $logFiles = Get-ChildItem $logDir -Filter "*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 3
    
    foreach ($logFile in $logFiles) {
        Write-Host "   📄 $($logFile.Name) - $($logFile.LastWriteTime)" -ForegroundColor Gray
        $recentEntries = Get-Content $logFile.FullName -Tail 3 -ErrorAction SilentlyContinue
        if ($recentEntries) {
            foreach ($entry in $recentEntries) {
                Write-Host "     $entry" -ForegroundColor DarkGray
            }
        }
        Write-Host ""
    }
}

Write-Host ""
Write-Host "=== COM INTERFACE TEST COMPLETE ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "This test helps identify why SAPI doesn't call our methods." -ForegroundColor Yellow
Write-Host "Key things to check:" -ForegroundColor White
Write-Host "1. Interface QueryInterface results" -ForegroundColor Gray
Write-Host "2. Object creation from token" -ForegroundColor Gray
Write-Host "3. COM registration completeness" -ForegroundColor Gray
