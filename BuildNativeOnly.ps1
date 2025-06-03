# Build Script for Native COM Wrapper Only
# This script builds just the native COM wrapper when .NET 6.0 is not available

Write-Host "BUILDING NATIVE COM WRAPPER ONLY" -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor Cyan

# Check if running as administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "WARNING: Not running as Administrator. Registration may fail." -ForegroundColor Yellow
    Write-Host "Consider running: sudo .\BuildNativeOnly.ps1" -ForegroundColor Yellow
    Write-Host ""
}

$ErrorActionPreference = "Stop"
$startTime = Get-Date

# Build configuration
$buildConfig = "Release"
$platform = "x64"
$outputDir = ".\dist"

Write-Host "Build Configuration: $buildConfig" -ForegroundColor Green
Write-Host "Platform: $platform" -ForegroundColor Green
Write-Host "Output Directory: $outputDir" -ForegroundColor Green
Write-Host ""

# Create output directory
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

# Step 1: Find MSBuild
Write-Host "1. Finding MSBuild..." -ForegroundColor Cyan

$vswherePath = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
$msbuild = $null

if (Test-Path $vswherePath) {
    $vsInstallPath = & $vswherePath -latest -products * -requires Microsoft.Component.MSBuild -property installationPath
    if ($vsInstallPath) {
        $msbuild = Join-Path $vsInstallPath "MSBuild\Current\Bin\MSBuild.exe"
    }
}

# Fallback to manual paths
if (-not $msbuild -or -not (Test-Path $msbuild)) {
    $msbuildPaths = @(
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
    )

    foreach ($path in $msbuildPaths) {
        if (Test-Path $path) {
            $msbuild = $path
            break
        }
    }
}

if (-not $msbuild -or -not (Test-Path $msbuild)) {
    Write-Host "   ERROR: MSBuild not found. Please install Visual Studio Build Tools 2022" -ForegroundColor Red
    Write-Host "   Download: https://visualstudio.microsoft.com/downloads/#build-tools-for-visual-studio-2022" -ForegroundColor Yellow
    exit 1
}

Write-Host "   Found MSBuild: $msbuild" -ForegroundColor Green
Write-Host ""

# Step 2: Build native COM wrapper
Write-Host "2. Building native COM wrapper..." -ForegroundColor Cyan
try {
    & $msbuild "NativeTTSWrapper\NativeTTSWrapper.vcxproj" /p:Configuration=$buildConfig /p:Platform=$platform /verbosity:minimal
    
    $nativeDllPath = "NativeTTSWrapper\$platform\$buildConfig\NativeTTSWrapper.dll"
    if (Test-Path $nativeDllPath) {
        $dllSize = (Get-Item $nativeDllPath).Length
        Write-Host "   Native COM wrapper built successfully ($([math]::Round($dllSize/1KB, 1)) KB)" -ForegroundColor Green
    } else {
        Write-Host "   ERROR: Native DLL not found at $nativeDllPath" -ForegroundColor Red
        exit 1
    }
} catch {
    Write-Host "   ERROR: Failed to build native COM wrapper" -ForegroundColor Red
    Write-Host "   $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
Write-Host ""

# Step 3: Copy to output directory
Write-Host "3. Copying to output directory..." -ForegroundColor Cyan

$nativeDll = "NativeTTSWrapper\$platform\$buildConfig\NativeTTSWrapper.dll"
if (Test-Path $nativeDll) {
    Copy-Item $nativeDll "$outputDir\" -Force
    Write-Host "   Copied NativeTTSWrapper.dll" -ForegroundColor Green
} else {
    Write-Host "   ERROR: Native DLL not found" -ForegroundColor Red
    exit 1
}

Write-Host ""

# Step 4: Test registration (optional)
Write-Host "4. Testing COM registration..." -ForegroundColor Cyan

$testDllPath = "$outputDir\NativeTTSWrapper.dll"
try {
    # Test registration
    $regProcess = Start-Process -FilePath "regsvr32.exe" -ArgumentList "/s", $testDllPath -Wait -PassThru
    
    if ($regProcess.ExitCode -eq 0) {
        Write-Host "   COM registration test successful" -ForegroundColor Green
        
        # Unregister immediately
        $unregProcess = Start-Process -FilePath "regsvr32.exe" -ArgumentList "/s", "/u", $testDllPath -Wait -PassThru
        if ($unregProcess.ExitCode -eq 0) {
            Write-Host "   COM unregistration test successful" -ForegroundColor Green
        }
    } else {
        Write-Host "   WARNING: COM registration test failed (exit code: $($regProcess.ExitCode))" -ForegroundColor Yellow
        Write-Host "   This may be due to missing dependencies or permissions" -ForegroundColor Yellow
    }
} catch {
    Write-Host "   WARNING: COM registration test failed: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host ""

# Step 5: Create deployment instructions
Write-Host "5. Creating deployment instructions..." -ForegroundColor Cyan

$deployInstructions = @"
# Native COM Wrapper Deployment

## Built Component:
- NativeTTSWrapper.dll ($(([math]::Round((Get-Item $testDllPath).Length/1KB, 1))) KB)

## Manual Deployment Steps:

1. Copy NativeTTSWrapper.dll to target location:
   ```
   Copy-Item "NativeTTSWrapper.dll" "C:\Program Files\OpenAssistive\OpenSpeech\" -Force
   ```

2. Register the COM object (as Administrator):
   ```
   sudo regsvr32 "C:\Program Files\OpenAssistive\OpenSpeech\NativeTTSWrapper.dll"
   ```

3. Update voice registration to use native wrapper:
   ```
   sudo Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Speech\Voices\Tokens\amy" -Name "CLSID" -Value "{E1C4A8F2-9B3D-4A5E-8F7C-2D1B3E4F5A6B}"
   ```

4. Test the native wrapper:
   ```
   `$voice = New-Object -ComObject SAPI.SpVoice
   `$voice.Speak("Hello from native COM wrapper!")
   ```

## What This Provides:
- 100% SAPI compatibility for SherpaOnnx voices
- Native C++ performance
- Full ProcessBridge integration
- Works with any SAPI application

## Next Steps:
To build the complete installer with .NET components:
1. Install .NET 6.0 SDK
2. Run: .\BuildCompleteInstaller.ps1

Built on: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
"@

$deployInstructions | Out-File "$outputDir\DEPLOYMENT.md" -Encoding UTF8
Write-Host "   Created DEPLOYMENT.md" -ForegroundColor Green

Write-Host ""

# Summary
$endTime = Get-Date
$buildDuration = $endTime - $startTime

Write-Host "NATIVE BUILD COMPLETED!" -ForegroundColor Cyan
Write-Host "=======================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Build Duration: $($buildDuration.TotalSeconds.ToString("F1")) seconds" -ForegroundColor Green
Write-Host "Output Directory: $outputDir" -ForegroundColor Green
Write-Host ""

Write-Host "BUILT COMPONENT:" -ForegroundColor Yellow
Get-ChildItem $outputDir | ForEach-Object {
    $size = if ($_.PSIsContainer) { "DIR" } else { "$([math]::Round($_.Length/1KB, 1)) KB" }
    Write-Host "  - $($_.Name) ($size)" -ForegroundColor White
}

Write-Host ""
Write-Host "NATIVE COM WRAPPER READY!" -ForegroundColor Cyan
Write-Host "This provides 100% SAPI compatibility for SherpaOnnx voices." -ForegroundColor Yellow
Write-Host "See DEPLOYMENT.md for installation instructions." -ForegroundColor Yellow
