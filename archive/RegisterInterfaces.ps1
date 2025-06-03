# RegisterInterfaces.ps1
# Script to register SAPI5 interface GUIDs that are missing from the registry
# This fixes the issue where SAPI cannot recognize our COM object as implementing required interfaces

Write-Host "=== SAPI5 Interface Registration Script ===" -ForegroundColor Cyan
Write-Host "Registering missing SAPI5 interface GUIDs..." -ForegroundColor Yellow

# Check if running as administrator
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "ERROR: This script must be run as Administrator!" -ForegroundColor Red
    Write-Host "Please run PowerShell as Administrator and try again." -ForegroundColor Yellow
    pause
    exit 1
}

# Interface GUIDs to register
$interfaces = @{
    "ISpTTSEngine" = "{A74D7C8E-4CC5-4F2F-A6EB-804DEE18500E}"
    "ISpObjectWithToken" = "{14056581-E16C-11D2-BB90-00C04F8EE6C0}"
}

# Function to register an interface GUID
function Register-InterfaceGuid {
    param(
        [string]$InterfaceName,
        [string]$InterfaceGuid
    )
    
    try {
        Write-Host "Registering $InterfaceName ($InterfaceGuid)..." -ForegroundColor White
        
        # Create the interface registry key
        $interfacePath = "HKLM:\SOFTWARE\Classes\Interface\$InterfaceGuid"
        
        # Check if already exists
        if (Test-Path $interfacePath) {
            Write-Host "  Interface already registered, updating..." -ForegroundColor Yellow
        } else {
            Write-Host "  Creating new interface registration..." -ForegroundColor Green
        }
        
        # Create the main interface key
        $interfaceKey = New-Item -Path $interfacePath -Force
        $interfaceKey.SetValue("", $InterfaceName)
        
        # Create ProxyStubClsid32 subkey (required for COM interfaces)
        $proxyStubPath = "$interfacePath\ProxyStubClsid32"
        $proxyStubKey = New-Item -Path $proxyStubPath -Force
        # Use the standard OLE Automation proxy/stub CLSID
        $proxyStubKey.SetValue("", "{00020424-0000-0000-C000-000000000046}")
        
        # Create TypeLib subkey (optional but recommended)
        $typeLibPath = "$interfacePath\TypeLib"
        $typeLibKey = New-Item -Path $typeLibPath -Force
        # Use SAPI5 type library GUID
        $typeLibKey.SetValue("", "{C866CA3A-32F7-11D2-9602-00C04F8EE628}")
        $typeLibKey.SetValue("Version", "5.4")
        
        Write-Host "  Successfully registered $InterfaceName" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "  ERROR registering $InterfaceName`: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Register each interface
$successCount = 0
$totalCount = $interfaces.Count

foreach ($interface in $interfaces.GetEnumerator()) {
    if (Register-InterfaceGuid -InterfaceName $interface.Key -InterfaceGuid $interface.Value) {
        $successCount++
    }
    Write-Host ""
}

# Summary
Write-Host "=== REGISTRATION SUMMARY ===" -ForegroundColor Cyan
Write-Host "Successfully registered: $successCount/$totalCount interfaces" -ForegroundColor $(if ($successCount -eq $totalCount) { "Green" } else { "Yellow" })

if ($successCount -eq $totalCount) {
    Write-Host ""
    Write-Host "✅ ALL INTERFACES REGISTERED SUCCESSFULLY!" -ForegroundColor Green
    Write-Host "SAPI5 should now be able to recognize our COM object interfaces." -ForegroundColor White
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Yellow
    Write-Host "1. Test the TTS engine with TestSpeech.ps1" -ForegroundColor White
    Write-Host "2. Check logs for 'SET OBJECT TOKEN CALLED' and 'SPEAK METHOD CALLED' messages" -ForegroundColor White
} else {
    Write-Host ""
    Write-Host "⚠️  SOME INTERFACES FAILED TO REGISTER" -ForegroundColor Yellow
    Write-Host "Please check the error messages above and try again." -ForegroundColor White
}

Write-Host ""
Write-Host "Press any key to continue..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
