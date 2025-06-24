# Manually register COM objects by writing registry entries directly
# This bypasses the DllRegisterServer function to test if the issue is with ATL or registry permissions

Write-Host "=== Manual COM Registration Test ===" -ForegroundColor Cyan

# Get the full path to our DLL
$dllPath = (Resolve-Path "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll").Path
Write-Host "DLL Path: $dllPath" -ForegroundColor Gray

# Test 1: Try to create registry entries manually for CNativeTTSWrapper
Write-Host "`n1. Manually registering CNativeTTSWrapper..." -ForegroundColor Yellow
try {
    $clsid = "{E1C4A8F2-9B3D-4A5E-8F7C-2D1B3E4F5A6B}"
    
    # Create CLSID entry
    $clsidPath = "HKLM:\SOFTWARE\Classes\CLSID\$clsid"
    New-Item -Path $clsidPath -Force | Out-Null
    Set-ItemProperty -Path $clsidPath -Name "(Default)" -Value "CNativeTTSWrapper Class"
    
    # Create InprocServer32 entry
    $inprocPath = "$clsidPath\InprocServer32"
    New-Item -Path $inprocPath -Force | Out-Null
    Set-ItemProperty -Path $inprocPath -Name "(Default)" -Value $dllPath
    Set-ItemProperty -Path $inprocPath -Name "ThreadingModel" -Value "Apartment"
    
    # Create ProgID entries
    New-Item -Path "HKLM:\SOFTWARE\Classes\NativeTTSWrapper.CNativeTTSWrapper" -Force | Out-Null
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Classes\NativeTTSWrapper.CNativeTTSWrapper" -Name "(Default)" -Value "CNativeTTSWrapper Class"
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Classes\NativeTTSWrapper.CNativeTTSWrapper" -Name "CLSID" -Value $clsid
    
    Write-Host "SUCCESS: Manual registry entries created" -ForegroundColor Green
    
} catch {
    Write-Host "FAILED: Could not create manual registry entries" -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}

# Test 2: Try to create the COM object
Write-Host "`n2. Testing manually registered COM object..." -ForegroundColor Yellow
try {
    $clsid = New-Object System.Guid("E1C4A8F2-9B3D-4A5E-8F7C-2D1B3E4F5A6B")
    $type = [System.Type]::GetTypeFromCLSID($clsid)
    $nativeWrapper = [System.Activator]::CreateInstance($type)
    Write-Host "SUCCESS: CNativeTTSWrapper created via manual registration!" -ForegroundColor Green
    $nativeWrapper = $null
} catch {
    Write-Host "FAILED: Could not create manually registered COM object" -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}

# Test 3: Check if registry entries exist
Write-Host "`n3. Verifying registry entries..." -ForegroundColor Yellow
$clsidPath = "HKLM:\SOFTWARE\Classes\CLSID\{E1C4A8F2-9B3D-4A5E-8F7C-2D1B3E4F5A6B}"
if (Test-Path $clsidPath) {
    Write-Host "SUCCESS: CLSID registry entry exists" -ForegroundColor Green
    $inprocPath = "$clsidPath\InprocServer32"
    if (Test-Path $inprocPath) {
        $dllReg = Get-ItemProperty -Path $inprocPath -Name "(Default)" -ErrorAction SilentlyContinue
        Write-Host "InprocServer32: $($dllReg.'(Default)')" -ForegroundColor Gray
    }
} else {
    Write-Host "FAILED: CLSID registry entry not found" -ForegroundColor Red
}

Write-Host "`n=== Manual Registration Test Complete ===" -ForegroundColor Cyan
