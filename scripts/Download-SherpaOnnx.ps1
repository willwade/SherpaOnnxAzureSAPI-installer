param(
    [string]$Version = "v1.12.20",
    [string]$OutputDir = "NativeTTSWrapper\libs-win"
)

<#
.SYNOPSIS
    SherpaOnnx Library Downloader
.DESCRIPTION
    Downloads and extracts SherpaOnnx v1.12.10 static library for Windows x64.
    Works in GitHub Actions CI and for local development.
.EXAMPLE
    .\scripts\Download-SherpaOnnx.ps1
#>

Write-Host "====================================" -ForegroundColor Cyan
Write-Host "SherpaOnnx Library Downloader" -ForegroundColor Cyan
Write-Host "====================================" -ForegroundColor Cyan
Write-Host ""

# Create directory
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
Write-Host "Created directory: $OutputDir" -ForegroundColor Green

# Try multiple sources if primary fails
$urls = @(
    "https://github.com/willwade/SherpaOnnxAzureSAPI-installer/releases/download/v1.0.0-deps/sherpa-onnx-win-x64-static.tar.bz2",
    "https://github.com/k2-fsa/sherpa-onnx/releases/download/$Version/sherpa-onnx-$Version-win-x64-static.tar.bz2",
    "https://github.com/k2-fsa/sherpa-onnx/releases/download/v1.12.23/sherpa-onnx-v1.12.23-win-x64-static.tar.bz2"
)

$archive = "$OutputDir\archive.tar.bz2"
$downloaded = $false

Write-Host "Attempting to download SherpaOnnx..." -ForegroundColor Yellow
foreach ($url in $urls) {
    Write-Host "  Trying: $url" -ForegroundColor Gray
    try {
        Invoke-WebRequest -Uri $url -OutFile $archive -UseBasicParsing
        $size = (Get-Item $archive).Length / 1MB
        Write-Host "  Downloaded: $([math]::Round($size, 2)) MB" -ForegroundColor Cyan

        if ($size -gt 50) {
            $downloaded = $true
            Write-Host "  SUCCESS: Download complete!" -ForegroundColor Green
            break
        } else {
            Write-Host "  WARNING: File too small, trying next URL..." -ForegroundColor Yellow
            Remove-Item $archive -Force
        }
    } catch {
        Write-Host "  FAILED: $_" -ForegroundColor Red
    }
}

if (-not $downloaded) {
    Write-Host "ERROR: All download attempts failed" -ForegroundColor Red
    exit 1
}

# Extract using available tool
Write-Host ""
Write-Host "Extracting archive..." -ForegroundColor Yellow

# Check what extraction tools are available
$has7zip = Get-Command "7z" -ErrorAction SilentlyContinue
$hasBash = Get-Command "bash" -ErrorAction SilentlyContinue
$hasTar = Get-Command "tar" -ErrorAction SilentlyContinue

$extractionSuccess = $false

# Try 7-Zip first (fastest)
if ($has7zip) {
    Write-Host "  Using 7-Zip..." -ForegroundColor Gray
    try {
        7z x "$archive" -so | 7z x -si -ttar -y "-o$OutputDir" | Out-Null
        $extractionSuccess = $true
        Write-Host "  Extraction with 7-Zip complete" -ForegroundColor Green
    } catch {
        Write-Host "  WARNING: 7-Zip extraction failed: $_" -ForegroundColor Yellow
    }
}

# Try bash + tar (Git Bash, WSL, etc.)
if (-not $extractionSuccess -and $hasBash) {
    Write-Host "  Using bash + tar..." -ForegroundColor Gray
    try {
        # Get full paths
        $fullArchive = (Resolve-Path $archive).Path
        # bash tar needs to run from the output directory
        Push-Location $OutputDir
        bash -c "tar -xf '$fullArchive'"
        Pop-Location
        $extractionSuccess = $true
        Write-Host "  Extraction with bash + tar complete" -ForegroundColor Green
    } catch {
        Write-Host "  WARNING: bash + tar extraction failed: $_" -ForegroundColor Yellow
    }
}

# Try PowerShell tar (Windows 10+)
if (-not $extractionSuccess -and $hasTar) {
    Write-Host "  Using tar..." -ForegroundColor Gray
    try {
        & tar -xf "$archive" -C $OutputDir
        $extractionSuccess = $true
        Write-Host "  Extraction with tar complete" -ForegroundColor Green
    } catch {
        Write-Host "  WARNING: tar extraction failed: $_" -ForegroundColor Yellow
    }
}

# Check if extraction succeeded
$extractedDirs = Get-ChildItem -Path $OutputDir -Directory | Where-Object { $_.Name -match "sherpa-onnx" }
if ($extractedDirs.Count -eq 0) {
    Write-Host ""
    Write-Host "ERROR: Extraction failed. No sherpa-onnx directory found." -ForegroundColor Red
    Write-Host ""
    Write-Host "Please install one of the following:" -ForegroundColor Yellow
    Write-Host "  1. 7-Zip: https://www.7-zip.org/" -ForegroundColor White
    Write-Host "  2. Git for Windows (includes bash): https://git-scm.com/download/win" -ForegroundColor White
    exit 1
}

Write-Host "  Extracted: $($extractedDirs[0].Name)" -ForegroundColor Green

# Find and rename the extracted directory if needed
Write-Host ""
Write-Host "Post-processing extracted files..." -ForegroundColor Yellow
$extractedDir = $extractedDirs[0]
$targetDir = "$OutputDir\sherpa-onnx-v1.12.10-win-x64-static"

if ($extractedDir.Name -ne "sherpa-onnx-v1.12.10-win-x64-static") {
    if (Test-Path $targetDir) {
        Write-Host "  Removing existing target directory..." -ForegroundColor Gray
        Remove-Item $targetDir -Recurse -Force
    }
    Write-Host "  Renaming: $($extractedDir.Name) -> sherpa-onnx-v1.12.10-win-x64-static" -ForegroundColor Gray
    Move-Item $extractedDir.FullName $targetDir -Force
}

# Verify structure
Write-Host ""
Write-Host "Verifying installation..." -ForegroundColor Yellow

# Check headers
$includePath = "$OutputDir\sherpa-onnx-v1.12.10-win-x64-static\include\sherpa-onnx\c-api"
if (Test-Path "$includePath\c-api.h") {
    Write-Host "  [OK] c-api.h found" -ForegroundColor Green
} else {
    Write-Host "  [FAIL] c-api.h NOT found" -ForegroundColor Red
    $found = Get-ChildItem -Path $OutputDir -Recurse -Filter "c-api.h" -ErrorAction SilentlyContinue
    if ($found) {
        Write-Host "    Found at: $($found[0].FullName)" -ForegroundColor Yellow
    }
}

# Check libraries
$libPath = "$OutputDir\sherpa-onnx-v1.12.10-win-x64-static\lib"
$requiredLibs = @("cppinyin_core.lib", "sherpa-onnx-c-api.lib", "sherpa-onnx-core.lib", "onnxruntime.lib")
$missingLibs = @()

foreach ($lib in $requiredLibs) {
    if (Test-Path "$libPath\$lib") {
        $size = (Get-Item "$libPath\$lib").Length / 1KB
        Write-Host "  [OK] $lib ($([math]::Round($size, 0)) KB)" -ForegroundColor Green
    } else {
        Write-Host "  [FAIL] $lib NOT found" -ForegroundColor Red
        $missingLibs += $lib
    }
}

if ($missingLibs.Count -gt 0) {
    Write-Host ""
    Write-Host "Available .lib files:" -ForegroundColor Yellow
    Get-ChildItem -Path $libPath -Filter "*.lib" -ErrorAction SilentlyContinue | Format-Table Name, @{Name="Size(MB)";Expression={[math]::Round($_.Length/1MB,2)}} -AutoSize
}

# Cleanup
Write-Host ""
Remove-Item $archive -Force -ErrorAction SilentlyContinue
Write-Host "Cleanup complete." -ForegroundColor Gray

# Summary
Write-Host ""
Write-Host "====================================" -ForegroundColor Cyan
if ($missingLibs.Count -eq 0) {
    Write-Host "SUCCESS: SherpaOnnx is ready!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Cyan
    Write-Host "  msbuild NativeTTSWrapper\NativeTTSWrapper.sln /p:Configuration=Release /p:Platform=x64" -ForegroundColor White
    Write-Host "====================================" -ForegroundColor Cyan
    exit 0
} else {
    Write-Host "FAILED: Missing $($missingLibs.Count) libraries" -ForegroundColor Red
    Write-Host "====================================" -ForegroundColor Cyan
    exit 1
}
