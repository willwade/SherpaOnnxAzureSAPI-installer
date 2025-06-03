# Fix COM Registration Issues
# This script identifies and fixes the COM registration problems

Write-Host "🔧 FIXING COM REGISTRATION ISSUES" -ForegroundColor Cyan
Write-Host "==================================" -ForegroundColor Cyan
Write-Host ""

# Check if running as administrator
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "❌ ERROR: This script must be run as Administrator!" -ForegroundColor Red
    Write-Host "Please run PowerShell as Administrator and try again." -ForegroundColor Yellow
    exit 1
}

# Configuration
$installDir = "C:\Program Files\OpenAssistive\OpenSpeech"
$nativeDll = "$installDir\NativeTTSWrapper.dll"
$managedDll = "$installDir\OpenSpeechTTS.dll"

# CLSIDs that should be registered
$nativeClsid = "{E1C4A8F2-9B3D-4A5E-8F7C-2D1B3E4F5A6B}"
$managedSherpaClsid = "{3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2}"
$azureClsid = "{3d8f5c5e-9d6b-4b92-a12b-1a6dff80b6b3}"

Write-Host "Configuration:" -ForegroundColor Green
Write-Host "  Install Directory: $installDir" -ForegroundColor Gray
Write-Host "  Native DLL: $nativeDll" -ForegroundColor Gray
Write-Host "  Managed DLL: $managedDll" -ForegroundColor Gray
Write-Host ""

# Step 1: Check if DLLs exist
Write-Host "1. Checking DLL Files..." -ForegroundColor Yellow

if (Test-Path $nativeDll) {
    $nativeInfo = Get-Item $nativeDll
    Write-Host "   ✅ Native DLL found: $($nativeInfo.Length) bytes" -ForegroundColor Green
} else {
    Write-Host "   ❌ Native DLL not found: $nativeDll" -ForegroundColor Red
    exit 1
}

if (Test-Path $managedDll) {
    $managedInfo = Get-Item $managedDll
    Write-Host "   ✅ Managed DLL found: $($managedInfo.Length) bytes" -ForegroundColor Green
} else {
    Write-Host "   ❌ Managed DLL not found: $managedDll" -ForegroundColor Red
    exit 1
}

Write-Host ""

# Step 2: Register Native COM DLL
Write-Host "2. Registering Native COM DLL..." -ForegroundColor Yellow

try {
    $regResult = Start-Process -FilePath "regsvr32" -ArgumentList "/s", $nativeDll -Wait -PassThru
    if ($regResult.ExitCode -eq 0) {
        Write-Host "   ✅ Native DLL registered successfully" -ForegroundColor Green
    } else {
        Write-Host "   ❌ Native DLL registration failed (exit code: $($regResult.ExitCode))" -ForegroundColor Red
        
        # Try without /s to see error message
        Write-Host "   Trying without silent mode to see error..." -ForegroundColor Yellow
        $regResult2 = Start-Process -FilePath "regsvr32" -ArgumentList $nativeDll -Wait -PassThru
    }
} catch {
    Write-Host "   ❌ Native DLL registration error: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""

# Step 3: Register Managed COM DLL
Write-Host "3. Registering Managed COM DLL..." -ForegroundColor Yellow

# Find regasm.exe
$regasmPaths = @(
    "${env:ProgramFiles(x86)}\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\RegAsm.exe",
    "${env:ProgramFiles}\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\x64\RegAsm.exe",
    "${env:WINDIR}\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe",
    "${env:WINDIR}\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
)

$regasmPath = $null
foreach ($path in $regasmPaths) {
    if (Test-Path $path) {
        $regasmPath = $path
        break
    }
}

if ($regasmPath) {
    Write-Host "   Found RegAsm: $regasmPath" -ForegroundColor Cyan
    
    try {
        $regasmResult = Start-Process -FilePath $regasmPath -ArgumentList "/codebase", $managedDll -Wait -PassThru -RedirectStandardOutput "regasm_output.txt" -RedirectStandardError "regasm_error.txt"
        
        if ($regasmResult.ExitCode -eq 0) {
            Write-Host "   ✅ Managed DLL registered successfully" -ForegroundColor Green
        } else {
            Write-Host "   ❌ Managed DLL registration failed (exit code: $($regasmResult.ExitCode))" -ForegroundColor Red
            
            if (Test-Path "regasm_error.txt") {
                $errorContent = Get-Content "regasm_error.txt"
                Write-Host "   Error details:" -ForegroundColor Red
                foreach ($line in $errorContent) {
                    Write-Host "     $line" -ForegroundColor Red
                }
            }
        }
    } catch {
        Write-Host "   ❌ Managed DLL registration error: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "   ❌ RegAsm.exe not found" -ForegroundColor Red
}

Write-Host ""

# Step 4: Verify registrations
Write-Host "4. Verifying COM Registrations..." -ForegroundColor Yellow

$clsids = @{
    "Native COM Wrapper" = $nativeClsid
    "Managed SherpaOnnx" = $managedSherpaClsid  
    "Azure TTS" = $azureClsid
}

foreach ($name in $clsids.Keys) {
    $clsid = $clsids[$name]
    $clsidPath = "HKLM:\SOFTWARE\Classes\CLSID\$clsid"
    
    if (Test-Path $clsidPath) {
        Write-Host "   ✅ $name CLSID registered: $clsid" -ForegroundColor Green
    } else {
        Write-Host "   ❌ $name CLSID NOT registered: $clsid" -ForegroundColor Red
    }
}

Write-Host ""

# Step 5: Test voice creation
Write-Host "5. Testing Voice Creation..." -ForegroundColor Yellow

try {
    $voice = New-Object -ComObject SAPI.SpVoice
    $voices = $voice.GetVoices()
    
    Write-Host "   📊 Total voices found: $($voices.Count)" -ForegroundColor Cyan
    
    # Test Amy voice
    $amyFound = $false
    for ($i = 0; $i -lt $voices.Count; $i++) {
        $voiceToken = $voices.Item($i)
        $voiceName = $voiceToken.GetDescription()
        
        if ($voiceName -like "*amy*") {
            Write-Host "   🎯 Found Amy voice: $voiceName" -ForegroundColor Blue
            $amyFound = $true
            
            try {
                $voice.Voice = $voiceToken
                Write-Host "   ✅ Amy voice set successfully" -ForegroundColor Green
                
                # Try to speak
                $result = $voice.Speak("Testing Amy voice", 1) # Async
                Write-Host "   ✅ Amy voice synthesis result: $result" -ForegroundColor Green
                
            } catch {
                Write-Host "   ❌ Amy voice failed: $($_.Exception.Message)" -ForegroundColor Red
            }
            break
        }
    }
    
    if (-not $amyFound) {
        Write-Host "   ⚠️ Amy voice not found" -ForegroundColor Yellow
    }
    
} catch {
    Write-Host "   ❌ Voice testing failed: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "🎯 COM REGISTRATION FIX COMPLETED" -ForegroundColor Cyan
Write-Host "If issues persist, check the error messages above and ensure all dependencies are available." -ForegroundColor Yellow

# Cleanup temp files
if (Test-Path "regasm_output.txt") { Remove-Item "regasm_output.txt" -Force }
if (Test-Path "regasm_error.txt") { Remove-Item "regasm_error.txt" -Force }
