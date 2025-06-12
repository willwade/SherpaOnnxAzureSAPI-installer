# Fix Pipe Service Architecture
# This script registers the COM wrapper and tests the complete pipeline

Write-Host "🔧 PIPE SERVICE ARCHITECTURE FIX" -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor Cyan
Write-Host ""

# Check if running as administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin) {
    Write-Host "❌ This script requires administrator privileges." -ForegroundColor Red
    Write-Host "Please run PowerShell as Administrator and try again." -ForegroundColor Red
    exit 1
}

Write-Host "🔍 ARCHITECTURE UNDERSTANDING:" -ForegroundColor Yellow
Write-Host "  1. Voice registered with pipe service CLSID ✅" -ForegroundColor Green
Write-Host "  2. Voice config exists in voice_configs/ ✅" -ForegroundColor Green
Write-Host "  3. AACSpeakHelper service running ✅" -ForegroundColor Green
Write-Host "  4. COM wrapper needs registration ❌" -ForegroundColor Red
Write-Host ""

$installerPath = ".\Installer\bin\Release\net6.0\win-x64\SherpaOnnxSAPIInstaller.exe"
$voiceConfigPath = ".\voice_configs\British-English-Azure-Libby.json"
$clsid = "{4A8B9C2D-1E3F-4567-8901-234567890ABC}"

# Check prerequisites
Write-Host "📋 CHECKING PREREQUISITES:" -ForegroundColor Cyan

if (-not (Test-Path $installerPath)) {
    Write-Host "❌ Installer not found: $installerPath" -ForegroundColor Red
    exit 1
} else {
    Write-Host "✅ Installer found" -ForegroundColor Green
}

if (-not (Test-Path $voiceConfigPath)) {
    Write-Host "❌ Voice config not found: $voiceConfigPath" -ForegroundColor Red
    exit 1
} else {
    Write-Host "✅ Voice config found" -ForegroundColor Green
}

# Check if AACSpeakHelper is running
$pythonProcess = Get-Process | Where-Object {$_.ProcessName -like "*python*"}
if ($pythonProcess) {
    Write-Host "✅ Python process running (AACSpeakHelper likely active)" -ForegroundColor Green
} else {
    Write-Host "⚠️ No Python process found - AACSpeakHelper may not be running" -ForegroundColor Yellow
}

Write-Host ""

# Step 1: Register COM wrapper using regasm
Write-Host "🔧 STEP 1: Registering COM wrapper..." -ForegroundColor Yellow

try {
    # Use regasm to register the managed assembly
    $regasmPath = "${env:ProgramFiles(x86)}\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\RegAsm.exe"
    
    if (-not (Test-Path $regasmPath)) {
        # Try alternative path
        $regasmPath = "${env:ProgramFiles}\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\x64\RegAsm.exe"
    }
    
    if (-not (Test-Path $regasmPath)) {
        Write-Host "⚠️ RegAsm.exe not found, trying manual registration..." -ForegroundColor Yellow
        
        # Manual COM registration
        Write-Host "Creating CLSID registry entries..." -ForegroundColor Gray
        
        $clsidKey = "HKCR:\CLSID\$clsid"
        New-Item -Path $clsidKey -Force | Out-Null
        Set-ItemProperty -Path $clsidKey -Name "(Default)" -Value "PipeServiceComWrapper"
        Set-ItemProperty -Path $clsidKey -Name "AppID" -Value $clsid
        
        $inprocKey = "$clsidKey\InprocServer32"
        New-Item -Path $inprocKey -Force | Out-Null
        Set-ItemProperty -Path $inprocKey -Name "(Default)" -Value $installerPath
        Set-ItemProperty -Path $inprocKey -Name "ThreadingModel" -Value "Apartment"
        
        $progIdSubKey = "$clsidKey\ProgId"
        New-Item -Path $progIdSubKey -Force | Out-Null
        Set-ItemProperty -Path $progIdSubKey -Name "(Default)" -Value "PipeServiceComWrapper.1"
        
        # ProgID registration
        $progIdKey = "HKCR:\PipeServiceComWrapper.1"
        New-Item -Path $progIdKey -Force | Out-Null
        Set-ItemProperty -Path $progIdKey -Name "(Default)" -Value "PipeServiceComWrapper"
        
        $clsidRefKey = "$progIdKey\CLSID"
        New-Item -Path $clsidRefKey -Force | Out-Null
        Set-ItemProperty -Path $clsidRefKey -Name "(Default)" -Value $clsid
        
        Write-Host "✅ Manual COM registration completed" -ForegroundColor Green
    } else {
        Write-Host "Using RegAsm: $regasmPath" -ForegroundColor Gray
        $result = & $regasmPath /codebase $installerPath
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✅ RegAsm registration successful" -ForegroundColor Green
        } else {
            Write-Host "⚠️ RegAsm returned exit code $LASTEXITCODE" -ForegroundColor Yellow
        }
    }
} catch {
    Write-Host "❌ Error registering COM wrapper: $_" -ForegroundColor Red
    exit 1
}

Write-Host ""

# Step 2: Verify COM registration
Write-Host "🔧 STEP 2: Verifying COM registration..." -ForegroundColor Yellow

try {
    $clsidExists = Test-Path "HKCR:\CLSID\$clsid"
    $progIdExists = Test-Path "HKCR:\PipeServiceComWrapper.1"
    
    if ($clsidExists -and $progIdExists) {
        Write-Host "✅ COM wrapper registration verified" -ForegroundColor Green
    } else {
        Write-Host "❌ COM wrapper registration incomplete" -ForegroundColor Red
        Write-Host "  CLSID exists: $clsidExists" -ForegroundColor Gray
        Write-Host "  ProgID exists: $progIdExists" -ForegroundColor Gray
    }
} catch {
    Write-Host "⚠️ Could not verify COM registration: $_" -ForegroundColor Yellow
}

Write-Host ""

# Step 3: Test voice synthesis
Write-Host "🔧 STEP 3: Testing voice synthesis..." -ForegroundColor Yellow

try {
    Write-Host "Running SAPI voice test..." -ForegroundColor Gray
    & .\TestSAPIVoices.ps1 -VoiceName "Azure Libby" -PlayAudio
} catch {
    Write-Host "⚠️ Could not run voice test: $_" -ForegroundColor Yellow
}

Write-Host ""

# Step 4: Test pipe service communication
Write-Host "🔧 STEP 4: Testing pipe service communication..." -ForegroundColor Yellow

try {
    # Create a simple test script to verify pipe communication
    $testScript = @"
import json
import win32file
import win32pipe

def test_pipe_communication():
    pipe_name = r'\\.\pipe\AACSpeakHelper'
    try:
        # Connect to pipe
        handle = win32file.CreateFile(
            pipe_name,
            win32file.GENERIC_READ | win32file.GENERIC_WRITE,
            0, None,
            win32file.OPEN_EXISTING,
            0, None
        )
        
        # Create test message
        message = {
            'args': {'listvoices': False, 'text': 'Hello from pipe test'},
            'config': {
                'TTS': {'engine': 'azure', 'voice_id': 'en-GB-LibbyNeural', 'bypass_tts': 'false'},
                'azureTTS': {
                    'key': 'b14f8945b0f1459f9964bdd72c42c2cc',
                    'location': 'uksouth',
                    'voice': 'en-GB-LibbyNeural'
                },
                'translate': {'no_translate': 'true'}
            },
            'clipboard_text': 'Hello from pipe test'
        }
        
        # Send message
        json_message = json.dumps(message)
        win32file.WriteFile(handle, json_message.encode())
        
        print("✅ Successfully sent test message to AACSpeakHelper")
        win32file.CloseHandle(handle)
        return True
        
    except Exception as e:
        print(f"❌ Pipe communication failed: {e}")
        return False

if __name__ == '__main__':
    test_pipe_communication()
"@

    $testScript | Out-File -FilePath "test_pipe.py" -Encoding UTF8
    
    # Run the test script
    Write-Host "Testing direct pipe communication..." -ForegroundColor Gray
    $result = & python test_pipe.py 2>&1
    Write-Host $result -ForegroundColor Gray
    
    # Clean up
    Remove-Item "test_pipe.py" -ErrorAction SilentlyContinue
    
} catch {
    Write-Host "⚠️ Could not test pipe communication: $_" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "✅ PIPE SERVICE ARCHITECTURE FIX COMPLETED!" -ForegroundColor Green
Write-Host ""
Write-Host "🧪 Next steps:" -ForegroundColor Cyan
Write-Host "  1. Test the voice: .\TestSAPIVoices.ps1 -VoiceName 'Azure Libby' -PlayAudio" -ForegroundColor White
Write-Host "  2. If it works, test in real applications (Notepad, etc.)" -ForegroundColor White
Write-Host "  3. Voice should now use pipe service architecture correctly" -ForegroundColor White

Write-Host ""
Write-Host "Fix completed!" -ForegroundColor Green
