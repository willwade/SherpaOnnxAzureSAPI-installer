# Properly register NativeTTSWrapper DLL (64-bit)
# Run as Administrator

$ErrorActionPreference = "Stop"

# Use 64-bit regsvr32 explicitly
$regsvr32 = "$env:SystemRoot\System32\regsvr32.exe"

# DLL path
$dllPath = "C:\github\SherpaOnnxAzureSAPI-installer\NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll"

Write-Host "=== Registering NativeTTSWrapper DLL (64-bit) ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "DLL: $dllPath" -ForegroundColor White
Write-Host "regsvr32: $regsvr32" -ForegroundColor Gray
Write-Host ""

# Check if DLL exists
if (-not (Test-Path $dllPath)) {
    Write-Host "ERROR: DLL not found at $dllPath" -ForegroundColor Red
    exit 1
}

# Unregister first (clean up any existing registration)
Write-Host "Step 1: Unregistering existing..." -ForegroundColor Yellow
& $regsvr32 /u /s $dllPath
Start-Sleep -Seconds 1

# Register (64-bit)
Write-Host "Step 2: Registering to 64-bit registry..." -ForegroundColor Yellow
$process = Start-Process -FilePath $regsvr32 -ArgumentList "/s", "`"$dllPath`"" -Wait -PassThru

if ($process.ExitCode -eq 0) {
    Write-Host "  Registration successful (exit code: 0)" -ForegroundColor Green
} else {
    Write-Host "  Registration failed (exit code: $($process.ExitCode))" -ForegroundColor Red
    exit $process.ExitCode
}

# Verify registration
Write-Host ""
Write-Host "Step 3: Verifying registration..." -ForegroundColor Yellow

$clsid = "{A1B2C3D4-E5F6-4A5B-8C9D-1E2F3A4B5C6D}"
$path64 = "Registry::HKLM\SOFTWARE\Classes\CLSID\$clsid\InprocServer32"
$path32 = "Registry::HKLM\SOFTWARE\Wow6432Node\Classes\CLSID\$clsid\InprocServer32"

$registered64 = Test-Path $path64
$registered32 = Test-Path $path32

if ($registered64) {
    Write-Host "  64-bit CLSID: Present" -ForegroundColor Green
} else {
    Write-Host "  64-bit CLSID: Missing" -ForegroundColor Red
}

if ($registered32) {
    Write-Host "  32-bit CLSID: Present (should not exist)" -ForegroundColor Red
} else {
    Write-Host "  32-bit CLSID: Absent (correct)" -ForegroundColor Green
}

if ($registered64) {
    $actualPath = (Get-ItemProperty $path64 -ErrorAction SilentlyContinue)."(default)"
    Write-Host "  DLL path: $actualPath" -ForegroundColor Cyan

    if ($actualPath -eq $dllPath) {
        Write-Host "  Path matches!" -ForegroundColor Green
    }
}

Write-Host ""
Write-Host "Registration Complete" -ForegroundColor Green
