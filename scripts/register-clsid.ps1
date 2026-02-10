# Manually register the CLSID for NativeTTSWrapper
# Run as Administrator

$ErrorActionPreference = "Stop"

$clsid = "{A1B2C3D4-E5F6-4A5B-8C9D-1E2F3A4B5C6D}"
$dllPath = "C:\temp\NativeTTSWrapper.dll"

Write-Host "Creating CLSID registry entries..." -ForegroundColor Cyan

# Remove existing if present
if (Test-Path "HKLM:\SOFTWARE\Classes\CLSID\$clsid") {
    Write-Host "Removing existing CLSID entry..." -ForegroundColor Yellow
    Remove-Item -Path "HKLM:\SOFTWARE\Classes\CLSID\$clsid" -Recurse -Force
}

# Create CLSID key
New-Item -Path "HKLM:\SOFTWARE\Classes\CLSID\$clsid" -Force | Out-Null
Set-ItemProperty -Path "HKLM:\SOFTWARE\Classes\CLSID\$clsid" -Name "(default)" -Value "NativeTTSWrapper Class"
Write-Host "Created CLSID: $clsid" -ForegroundColor Green

# Create InprocServer32 key
New-Item -Path "HKLM:\SOFTWARE\Classes\CLSID\$clsid\InprocServer32" -Force | Out-Null
Set-ItemProperty -Path "HKLM:\SOFTWARE\Classes\CLSID\$clsid\InprocServer32" -Name "(default)" -Value $dllPath
Set-ItemProperty -Path "HKLM:\SOFTWARE\Classes\CLSID\$clsid\InprocServer32" -Name "ThreadingModel" -Value "Both"
Write-Host "Created InprocServer32: $dllPath" -ForegroundColor Green

# Verify
if (Test-Path "HKLM:\SOFTWARE\Classes\CLSID\$clsid\InprocServer32") {
    $actualPath = (Get-ItemProperty "HKLM:\SOFTWARE\Classes\CLSID\$clsid\InprocServer32")."(default)"
    Write-Host ""
    Write-Host "CLSID registration successful!" -ForegroundColor Green
    Write-Host "  DLL: $actualPath" -ForegroundColor White
    Write-Host ""
    Write-Host "Now run: .\test-voice.ps1" -ForegroundColor Cyan
} else {
    Write-Host "ERROR: Registration failed" -ForegroundColor Red
}
