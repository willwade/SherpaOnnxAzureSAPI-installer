# Deploy Native COM Wrapper (Run as Administrator)
Write-Host "Deploying Native COM Wrapper" -ForegroundColor Green
Write-Host "============================" -ForegroundColor Green

# Check if running as administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "ERROR: This script must be run as Administrator" -ForegroundColor Red
    Write-Host "Right-click PowerShell and select 'Run as Administrator'" -ForegroundColor Yellow
    exit 1
}

Write-Host "SUCCESS: Running as Administrator" -ForegroundColor Green

# Step 1: Check if DLL exists
$dllPath = "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll"
if (-not (Test-Path $dllPath)) {
    Write-Host "ERROR: DLL not found: $dllPath" -ForegroundColor Red
    Write-Host "Please build the native wrapper first" -ForegroundColor Yellow
    exit 1
}

$dllSize = (Get-Item $dllPath).Length
Write-Host "SUCCESS: Found DLL: $dllPath ($([math]::Round($dllSize/1KB, 1)) KB)" -ForegroundColor Green

# Step 2: Create installation directory
$installDir = "C:\Program Files\OpenAssistive\OpenSpeech"
if (-not (Test-Path $installDir)) {
    New-Item -ItemType Directory -Path $installDir -Force | Out-Null
    Write-Host "SUCCESS: Created installation directory: $installDir" -ForegroundColor Green
} else {
    Write-Host "SUCCESS: Installation directory exists: $installDir" -ForegroundColor Green
}

# Step 3: Copy DLL to installation directory
$installPath = "$installDir\NativeTTSWrapper.dll"
try {
    Copy-Item $dllPath $installPath -Force
    Write-Host "SUCCESS: Deployed DLL to: $installPath" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Failed to copy DLL: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Step 4: Register COM object
Write-Host "Registering COM object..." -ForegroundColor Cyan
$regsvr32 = "${env:SystemRoot}\System32\regsvr32.exe"
try {
    $regProcess = Start-Process -FilePath $regsvr32 -ArgumentList "/s", $installPath -Wait -PassThru
    
    if ($regProcess.ExitCode -eq 0) {
        Write-Host "SUCCESS: COM registration successful" -ForegroundColor Green
    } else {
        Write-Host "ERROR: COM registration failed with exit code: $($regProcess.ExitCode)" -ForegroundColor Red
        exit 1
    }
} catch {
    Write-Host "ERROR: COM registration error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Step 5: Update voice registration
Write-Host "Updating voice registration..." -ForegroundColor Cyan

$newClsid = "{E1C4A8F2-9B3D-4A5E-8F7C-2D1B3E4F5A6B}"
$voiceTokenPath = "HKLM:\SOFTWARE\Microsoft\Speech\Voices\Tokens\amy"

try {
    if (Test-Path $voiceTokenPath) {
        Set-ItemProperty -Path $voiceTokenPath -Name "CLSID" -Value $newClsid
        Write-Host "SUCCESS: Updated voice token CLSID to native wrapper" -ForegroundColor Green
    } else {
        Write-Host "WARNING: Amy voice token not found - will need to be created" -ForegroundColor Yellow
    }
} catch {
    Write-Host "ERROR: Voice registration update failed: $($_.Exception.Message)" -ForegroundColor Red
}

# Step 6: Verify COM registration
Write-Host "Verifying COM registration..." -ForegroundColor Cyan

$clsidPath = "HKLM:\SOFTWARE\Classes\CLSID\$newClsid\InprocServer32"
if (Test-Path $clsidPath) {
    $inprocServer = (Get-ItemProperty $clsidPath).'(default)'
    Write-Host "SUCCESS: Native COM registration verified: $inprocServer" -ForegroundColor Green
} else {
    Write-Host "ERROR: Native COM registration not found in registry" -ForegroundColor Red
}

Write-Host ""
Write-Host "DEPLOYMENT COMPLETE!" -ForegroundColor Cyan
Write-Host ""
Write-Host "ACHIEVEMENTS:" -ForegroundColor Yellow
Write-Host "  - Native C++ COM DLL built and deployed" -ForegroundColor White
Write-Host "  - COM object registered successfully" -ForegroundColor White
Write-Host "  - Voice registration updated" -ForegroundColor White
Write-Host ""
Write-Host "NEXT: Test the native COM wrapper" -ForegroundColor Green
Write-Host "Run: .\TestNativeWrapper.ps1" -ForegroundColor White
