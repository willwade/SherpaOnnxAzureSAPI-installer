# Register COM Wrapper for Pipe Service Voices
# This script registers the PipeServiceComWrapper COM component needed for SAPI voice synthesis

param(
    [switch]$Unregister
)

Write-Host "🔧 COM WRAPPER REGISTRATION TOOL" -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor Cyan
Write-Host ""

# Check if running as administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin) {
    Write-Host "❌ This script requires administrator privileges." -ForegroundColor Red
    Write-Host "Please run PowerShell as Administrator and try again." -ForegroundColor Red
    exit 1
}

$installerPath = ".\Installer\bin\Release\net6.0\win-x64\SherpaOnnxSAPIInstaller.exe"
$clsid = "{4A8B9C2D-1E3F-4567-8901-234567890ABC}"

# Check if installer exists
if (-not (Test-Path $installerPath)) {
    Write-Host "❌ Installer not found at: $installerPath" -ForegroundColor Red
    Write-Host "Please build the installer first:" -ForegroundColor Yellow
    Write-Host "  dotnet build Installer/Installer.csproj -c Release" -ForegroundColor Gray
    exit 1
}

Write-Host "Using installer: $installerPath" -ForegroundColor Gray
Write-Host "COM CLSID: $clsid" -ForegroundColor Gray
Write-Host ""

if ($Unregister) {
    Write-Host "🗑️ Unregistering COM wrapper..." -ForegroundColor Yellow
    
    # Remove CLSID registration
    try {
        Remove-Item -Path "HKCR:\CLSID\$clsid" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "✅ Removed CLSID registration" -ForegroundColor Green
    } catch {
        Write-Host "⚠️ Could not remove CLSID registration: $_" -ForegroundColor Yellow
    }
    
    # Remove ProgID registration
    try {
        Remove-Item -Path "HKCR:\PipeServiceComWrapper.1" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "✅ Removed ProgID registration" -ForegroundColor Green
    } catch {
        Write-Host "⚠️ Could not remove ProgID registration: $_" -ForegroundColor Yellow
    }
    
    Write-Host ""
    Write-Host "✅ COM wrapper unregistration completed" -ForegroundColor Green
} else {
    Write-Host "📝 Registering COM wrapper..." -ForegroundColor Yellow
    
    # Check current registration status
    $clsidExists = Test-Path "HKCR:\CLSID\$clsid"
    $progIdExists = Test-Path "HKCR:\PipeServiceComWrapper.1"
    
    Write-Host "Current status:" -ForegroundColor Cyan
    Write-Host "  CLSID registered: $clsidExists" -ForegroundColor Gray
    Write-Host "  ProgID registered: $progIdExists" -ForegroundColor Gray
    Write-Host ""
    
    if ($clsidExists -and $progIdExists) {
        Write-Host "✅ COM wrapper appears to be already registered" -ForegroundColor Green
        Write-Host "If you're still getting 'Class not registered' errors, try:" -ForegroundColor Yellow
        Write-Host "  1. .\RegisterComWrapper.ps1 -Unregister" -ForegroundColor Gray
        Write-Host "  2. .\RegisterComWrapper.ps1" -ForegroundColor Gray
    } else {
        Write-Host "🔧 Registering COM wrapper using regasm..." -ForegroundColor Yellow
        
        # Find the managed DLL
        $managedDllPath = ".\Installer\bin\Release\net6.0\win-x64\SherpaOnnxSAPIInstaller.exe"
        
        # Try to register using regasm
        try {
            # Find regasm
            $dotnetPath = (Get-Command dotnet -ErrorAction SilentlyContinue).Source
            if ($dotnetPath) {
                $dotnetDir = Split-Path $dotnetPath -Parent
                $regasmPath = Get-ChildItem -Path $dotnetDir -Recurse -Name "regasm.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($regasmPath) {
                    $regasmFullPath = Join-Path $dotnetDir $regasmPath
                    Write-Host "Found regasm at: $regasmFullPath" -ForegroundColor Gray
                    
                    $result = & $regasmFullPath /codebase $managedDllPath 2>&1
                    if ($LASTEXITCODE -eq 0) {
                        Write-Host "✅ regasm registration successful" -ForegroundColor Green
                    } else {
                        Write-Host "⚠️ regasm failed: $result" -ForegroundColor Yellow
                    }
                }
            }
        } catch {
            Write-Host "⚠️ Could not use regasm: $_" -ForegroundColor Yellow
        }
        
        # Manual registration as fallback
        Write-Host "🔧 Performing manual COM registration..." -ForegroundColor Yellow
        
        try {
            # Create CLSID registration
            $clsidKey = "HKCR:\CLSID\$clsid"
            New-Item -Path $clsidKey -Force | Out-Null
            Set-ItemProperty -Path $clsidKey -Name "(Default)" -Value "PipeServiceComWrapper"
            Set-ItemProperty -Path $clsidKey -Name "AppID" -Value $clsid
            
            # Create InprocServer32 subkey
            $inprocKey = "$clsidKey\InprocServer32"
            New-Item -Path $inprocKey -Force | Out-Null
            Set-ItemProperty -Path $inprocKey -Name "(Default)" -Value $managedDllPath
            Set-ItemProperty -Path $inprocKey -Name "ThreadingModel" -Value "Apartment"
            
            # Create ProgId subkey
            $progIdSubKey = "$clsidKey\ProgId"
            New-Item -Path $progIdSubKey -Force | Out-Null
            Set-ItemProperty -Path $progIdSubKey -Name "(Default)" -Value "PipeServiceComWrapper.1"
            
            Write-Host "✅ CLSID registration completed" -ForegroundColor Green
            
            # Create ProgID registration
            $progIdKey = "HKCR:\PipeServiceComWrapper.1"
            New-Item -Path $progIdKey -Force | Out-Null
            Set-ItemProperty -Path $progIdKey -Name "(Default)" -Value "PipeServiceComWrapper"
            
            # Create CLSID reference
            $clsidRefKey = "$progIdKey\CLSID"
            New-Item -Path $clsidRefKey -Force | Out-Null
            Set-ItemProperty -Path $clsidRefKey -Name "(Default)" -Value $clsid
            
            Write-Host "✅ ProgID registration completed" -ForegroundColor Green
            
        } catch {
            Write-Host "❌ Manual registration failed: $_" -ForegroundColor Red
            exit 1
        }
    }
    
    Write-Host ""
    Write-Host "✅ COM wrapper registration completed!" -ForegroundColor Green
    Write-Host ""
    Write-Host "🧪 Next steps:" -ForegroundColor Cyan
    Write-Host "  1. Test the voice: .\TestSAPIVoices.ps1 -VoiceName 'Azure Libby' -PlayAudio" -ForegroundColor White
    Write-Host "  2. If it still fails, check AACSpeakHelper service is running" -ForegroundColor White
    Write-Host "  3. Try testing in a real application (Notepad, etc.)" -ForegroundColor White
}

Write-Host ""
Write-Host "Registration completed!" -ForegroundColor Green
