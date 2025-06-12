# Test Voice Registration and SAPI Integration
# This script comprehensively tests voice registration and SAPI functionality

Write-Host "🔍 Voice Registration and SAPI Integration Test" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host ""

# Function to check registry key
function Test-RegistryKey {
    param(
        [string]$Path,
        [string]$Description
    )
    
    try {
        $key = Get-ItemProperty -Path "HKLM:\$Path" -ErrorAction Stop
        Write-Host "✅ $Description exists" -ForegroundColor Green
        return $key
    } catch {
        Write-Host "❌ $Description not found" -ForegroundColor Red
        return $null
    }
}

# Function to list all SAPI voices
function Get-SAPIVoices {
    Write-Host "🔍 Enumerating SAPI Voices..." -ForegroundColor Cyan
    
    try {
        $voice = New-Object -ComObject SAPI.SpVoice
        $voices = $voice.GetVoices()
        
        Write-Host "Found $($voices.Count) SAPI voices:" -ForegroundColor White
        
        $voiceList = @()
        for ($i = 0; $i -lt $voices.Count; $i++) {
            $voiceInfo = $voices.Item($i)
            $name = $voiceInfo.GetDescription()
            $id = $voiceInfo.Id
            
            # Check if this is our voice
            $isOurVoice = $name -like "*SherpaOnnx*" -or $name -like "*Jenny*" -or $name -like "*AACSpeakHelper*"
            
            if ($isOurVoice) {
                Write-Host "  ✅ $name" -ForegroundColor Green
                Write-Host "     ID: $id" -ForegroundColor Gray
            } else {
                Write-Host "  - $name" -ForegroundColor White
            }
            
            $voiceList += @{
                Name = $name
                Id = $id
                IsOurVoice = $isOurVoice
            }
        }
        
        return $voiceList
    } catch {
        Write-Host "❌ Failed to enumerate SAPI voices: $($_.Exception.Message)" -ForegroundColor Red
        return @()
    }
}

# Function to test voice synthesis
function Test-VoiceSynthesis {
    param(
        [string]$VoiceName,
        [string]$TestText = "Hello from the C++ SAPI Bridge to AACSpeakHelper! This is a test of voice synthesis."
    )
    
    Write-Host "🔊 Testing voice synthesis for: $VoiceName" -ForegroundColor Cyan
    
    try {
        $voice = New-Object -ComObject SAPI.SpVoice
        $voices = $voice.GetVoices()
        
        # Find the voice
        $targetVoice = $null
        for ($i = 0; $i -lt $voices.Count; $i++) {
            $voiceInfo = $voices.Item($i)
            $name = $voiceInfo.GetDescription()
            if ($name -like "*$VoiceName*") {
                $targetVoice = $voiceInfo
                break
            }
        }
        
        if ($targetVoice) {
            Write-Host "✅ Found voice: $($targetVoice.GetDescription())" -ForegroundColor Green
            
            # Set the voice
            $voice.Voice = $targetVoice
            
            # Test synthesis
            Write-Host "🎵 Speaking test text..." -ForegroundColor Yellow
            $voice.Speak($TestText)
            
            Write-Host "✅ Voice synthesis completed successfully!" -ForegroundColor Green
            return $true
        } else {
            Write-Host "❌ Voice not found: $VoiceName" -ForegroundColor Red
            return $false
        }
    } catch {
        Write-Host "❌ Voice synthesis failed: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to check CLSID registration
function Test-CLSIDRegistration {
    param([string]$CLSID, [string]$Description)
    
    Write-Host "🔍 Checking CLSID registration: $Description" -ForegroundColor Cyan
    
    $clsidPath = "HKCR:\CLSID\$CLSID"
    try {
        $clsidKey = Get-ItemProperty -Path $clsidPath -ErrorAction Stop
        Write-Host "✅ CLSID registered: $CLSID" -ForegroundColor Green
        
        # Check InprocServer32
        $inprocPath = "$clsidPath\InprocServer32"
        try {
            $inprocKey = Get-ItemProperty -Path $inprocPath -ErrorAction Stop
            $dllPath = $inprocKey.'(default)'
            Write-Host "   DLL Path: $dllPath" -ForegroundColor Gray
            
            if (Test-Path $dllPath) {
                Write-Host "   ✅ DLL exists" -ForegroundColor Green
            } else {
                Write-Host "   ❌ DLL not found" -ForegroundColor Red
            }
        } catch {
            Write-Host "   ❌ InprocServer32 not found" -ForegroundColor Red
        }
        
        return $true
    } catch {
        Write-Host "❌ CLSID not registered: $CLSID" -ForegroundColor Red
        return $false
    }
}

# Main test execution
Write-Host "📋 Step 1: Check Voice Registry Entries" -ForegroundColor Yellow
Write-Host "=======================================" -ForegroundColor Yellow

# Check for our voice configurations
$voiceTokensPath = "SOFTWARE\Microsoft\SPEECH\Voices\Tokens"
try {
    $voiceTokens = Get-ChildItem -Path "HKLM:\$voiceTokensPath" -ErrorAction Stop
    
    Write-Host "Found $($voiceTokens.Count) voice tokens in registry:" -ForegroundColor White
    
    $ourVoices = @()
    foreach ($token in $voiceTokens) {
        $tokenName = $token.PSChildName
        $isOurVoice = $tokenName -like "*SherpaOnnx*" -or $tokenName -like "*Jenny*" -or $tokenName -like "*AACSpeakHelper*"
        
        if ($isOurVoice) {
            Write-Host "  ✅ $tokenName" -ForegroundColor Green
            $ourVoices += $tokenName
            
            # Get details
            try {
                $tokenDetails = Get-ItemProperty -Path "HKLM:\$voiceTokensPath\$tokenName"
                Write-Host "     CLSID: $($tokenDetails.CLSID)" -ForegroundColor Gray
                Write-Host "     Default: $($tokenDetails.'(default)')" -ForegroundColor Gray
            } catch {
                Write-Host "     ❌ Could not read token details" -ForegroundColor Red
            }
        } else {
            Write-Host "  - $tokenName" -ForegroundColor White
        }
    }
    
    if ($ourVoices.Count -eq 0) {
        Write-Host "⚠️ No SherpaOnnx/AACSpeakHelper voices found in registry" -ForegroundColor Yellow
    }
    
} catch {
    Write-Host "❌ Could not access voice tokens registry" -ForegroundColor Red
}

Write-Host ""

# Step 2: Check CLSID registrations
Write-Host "📋 Step 2: Check CLSID Registrations" -ForegroundColor Yellow
Write-Host "====================================" -ForegroundColor Yellow

# Known CLSIDs from the code
$clsids = @{
    "E1C4A8F2-9B3D-4A5E-8F7C-2D1B3E4F5A6B" = "C++ NativeTTSWrapper (from IDL)"
    "4A8B9C2D-1E3F-4567-8901-234567890ABC" = "PipeService CLSID (from ConfigBasedVoiceManager)"
    "3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2" = "SherpaOnnx CLSID (from Sapi5RegistrarExtended)"
    "3d8f5c5e-9d6b-4b92-a12b-1a6dff80b6b3" = "Azure TTS CLSID (from Sapi5RegistrarExtended)"
}

foreach ($clsid in $clsids.Keys) {
    Test-CLSIDRegistration -CLSID $clsid -Description $clsids[$clsid]
}

Write-Host ""

# Step 3: Enumerate SAPI voices
Write-Host "📋 Step 3: Enumerate SAPI Voices" -ForegroundColor Yellow
Write-Host "================================" -ForegroundColor Yellow

$sapiVoices = Get-SAPIVoices

Write-Host ""

# Step 4: Test voice synthesis
Write-Host "📋 Step 4: Test Voice Synthesis" -ForegroundColor Yellow
Write-Host "===============================" -ForegroundColor Yellow

$ourSapiVoices = $sapiVoices | Where-Object { $_.IsOurVoice }

if ($ourSapiVoices.Count -gt 0) {
    foreach ($voice in $ourSapiVoices) {
        Test-VoiceSynthesis -VoiceName $voice.Name.Split(' ')[0]
        Write-Host ""
    }
} else {
    Write-Host "⚠️ No SherpaOnnx/AACSpeakHelper voices found in SAPI" -ForegroundColor Yellow
    Write-Host "   Attempting to test with any available voice..." -ForegroundColor White
    
    if ($sapiVoices.Count -gt 0) {
        $firstVoice = $sapiVoices[0]
        Test-VoiceSynthesis -VoiceName $firstVoice.Name.Split(' ')[0] -TestText "Testing default voice synthesis."
    }
}

Write-Host ""

# Step 5: Diagnostic information
Write-Host "📋 Step 5: Diagnostic Information" -ForegroundColor Yellow
Write-Host "=================================" -ForegroundColor Yellow

Write-Host "🔍 System Information:" -ForegroundColor Cyan
Write-Host "   OS: $((Get-WmiObject Win32_OperatingSystem).Caption)" -ForegroundColor White
Write-Host "   Architecture: $env:PROCESSOR_ARCHITECTURE" -ForegroundColor White
Write-Host "   PowerShell: $($PSVersionTable.PSVersion)" -ForegroundColor White

Write-Host ""
Write-Host "🔍 File Locations:" -ForegroundColor Cyan

$locations = @{
    "NativeTTSWrapper.dll" = "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll"
    "Voice Configs" = "voice_configs"
    "SapiVoiceManager.py" = "SapiVoiceManager.py"
}

foreach ($item in $locations.Keys) {
    $path = $locations[$item]
    if (Test-Path $path) {
        Write-Host "   ✅ $item`: $path" -ForegroundColor Green
    } else {
        Write-Host "   ❌ $item`: $path (not found)" -ForegroundColor Red
    }
}

Write-Host ""

# Step 6: Recommendations
Write-Host "📋 Step 6: Recommendations" -ForegroundColor Yellow
Write-Host "==========================" -ForegroundColor Yellow

Write-Host "🎯 To fix voice registration issues:" -ForegroundColor Cyan
Write-Host ""
Write-Host "1. Ensure CLSID consistency:" -ForegroundColor White
Write-Host "   - C++ wrapper CLSID: E1C4A8F2-9B3D-4A5E-8F7C-2D1B3E4F5A6B" -ForegroundColor Gray
Write-Host "   - Voice registration should use the same CLSID" -ForegroundColor Gray
Write-Host ""
Write-Host "2. Register COM wrapper:" -ForegroundColor White
Write-Host "   regsvr32 NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll" -ForegroundColor Gray
Write-Host ""
Write-Host "3. Install voice with correct CLSID:" -ForegroundColor White
Write-Host "   uv run python SapiVoiceManager.py --install English-SherpaOnnx-Jenny" -ForegroundColor Gray
Write-Host ""
Write-Host "4. Test AACSpeakHelper service:" -ForegroundColor White
Write-Host "   # Start AACSpeakHelper service first" -ForegroundColor Gray
Write-Host "   uv run python AACSpeakHelperServer.py" -ForegroundColor Gray

Write-Host ""
Write-Host "🎉 Voice registration test completed!" -ForegroundColor Green
