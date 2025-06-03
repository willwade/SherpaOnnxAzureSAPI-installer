# Build and Deploy ProcessBridge TTS - must run as Administrator
Write-Host "Building and Deploying ProcessBridge TTS" -ForegroundColor Green
Write-Host "=========================================" -ForegroundColor Green

$projectPath = "OpenSpeechTTS\OpenSpeechTTS.csproj"
$sourceDLL = "OpenSpeechTTS\bin\Release\net472\OpenSpeechTTS.dll"
$targetDir = "C:\Program Files\OpenAssistive\OpenSpeech"
$targetDLL = "$targetDir\OpenSpeechTTS.dll"

Write-Host "1. Building project..." -ForegroundColor Cyan

# Try to find MSBuild
$msbuildPaths = @(
    "${env:ProgramFiles}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles}\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles}\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles(x86)}\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe",
    "${env:ProgramFiles(x86)}\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe"
)

$msbuild = $null
foreach ($path in $msbuildPaths) {
    if (Test-Path $path) {
        $msbuild = $path
        Write-Host "   Found MSBuild: $msbuild" -ForegroundColor Green
        break
    }
}

if ($msbuild) {
    try {
        Write-Host "   Building with MSBuild..." -ForegroundColor Yellow
        & $msbuild $projectPath /p:Configuration=Release /p:Platform="Any CPU" /verbosity:minimal
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "   ✅ Build successful" -ForegroundColor Green
        } else {
            Write-Host "   ❌ Build failed with exit code $LASTEXITCODE" -ForegroundColor Red
            exit 1
        }
    } catch {
        Write-Host "   ❌ Build error: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
} else {
    Write-Host "   ❌ MSBuild not found. Trying dotnet build..." -ForegroundColor Yellow
    try {
        dotnet build $projectPath -c Release
        if ($LASTEXITCODE -eq 0) {
            Write-Host "   ✅ Build successful with dotnet" -ForegroundColor Green
        } else {
            Write-Host "   ❌ dotnet build failed" -ForegroundColor Red
            exit 1
        }
    } catch {
        Write-Host "   ❌ dotnet build error: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

Write-Host ""
Write-Host "2. Checking build output..." -ForegroundColor Cyan
if (Test-Path $sourceDLL) {
    $sourceTime = (Get-Item $sourceDLL).LastWriteTime
    $sourceSize = (Get-Item $sourceDLL).Length
    Write-Host "   ✅ Built DLL: $sourceDLL" -ForegroundColor Green
    Write-Host "   Last modified: $sourceTime" -ForegroundColor White
    Write-Host "   Size: $([math]::Round($sourceSize/1KB, 1)) KB" -ForegroundColor White
} else {
    Write-Host "   ❌ Build output not found: $sourceDLL" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "3. Unregistering old COM component..." -ForegroundColor Cyan
try {
    # Find RegAsm
    $regasmPaths = @(
        "${env:ProgramFiles(x86)}\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\RegAsm.exe",
        "${env:ProgramFiles}\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\x64\RegAsm.exe",
        "${env:ProgramFiles(x86)}\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe",
        "${env:ProgramFiles(x86)}\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
    )
    
    $regasm = $null
    foreach ($path in $regasmPaths) {
        if (Test-Path $path) {
            $regasm = $path
            break
        }
    }
    
    if ($regasm) {
        Write-Host "   Using RegAsm: $regasm" -ForegroundColor White
        if (Test-Path $targetDLL) {
            & $regasm $targetDLL /unregister /silent
            Write-Host "   ✅ Unregistered old component" -ForegroundColor Green
        }
    } else {
        Write-Host "   ⚠️ RegAsm not found - skipping unregister" -ForegroundColor Yellow
    }
} catch {
    Write-Host "   ⚠️ Unregister warning: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "4. Deploying new DLL..." -ForegroundColor Cyan
try {
    Copy-Item $sourceDLL $targetDLL -Force
    Write-Host "   ✅ DLL deployed successfully" -ForegroundColor Green
} catch {
    Write-Host "   ❌ Failed to deploy DLL: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "5. Registering new COM component..." -ForegroundColor Cyan
if ($regasm) {
    try {
        & $regasm $targetDLL /codebase /silent
        Write-Host "   ✅ COM component registered" -ForegroundColor Green
    } catch {
        Write-Host "   ❌ Failed to register COM: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
} else {
    Write-Host "   ❌ RegAsm not found - cannot register COM component" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "6. Verifying deployment..." -ForegroundColor Cyan
if (Test-Path $targetDLL) {
    $targetTime = (Get-Item $targetDLL).LastWriteTime
    $targetSize = (Get-Item $targetDLL).Length
    Write-Host "   ✅ Deployed DLL: $targetDLL" -ForegroundColor Green
    Write-Host "   Last modified: $targetTime" -ForegroundColor White
    Write-Host "   Size: $([math]::Round($targetSize/1KB, 1)) KB" -ForegroundColor White
    
    if ($targetTime -ge $sourceTime) {
        Write-Host "   ✅ Deployment timestamp is current" -ForegroundColor Green
    } else {
        Write-Host "   ⚠️ Deployment timestamp is older than source" -ForegroundColor Yellow
    }
} else {
    Write-Host "   ❌ Deployed DLL not found" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "=== DEPLOYMENT COMPLETE ===" -ForegroundColor Cyan
Write-Host "ProcessBridge TTS has been built and deployed!" -ForegroundColor Green
Write-Host "You can now test the ProcessBridge implementation." -ForegroundColor White
