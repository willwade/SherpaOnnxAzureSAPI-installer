#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Clean up all existing registrations, logs, and temporary files for a fresh start
.DESCRIPTION
    Unregisters DLLs, removes log files, and clears temporary state to avoid confusion
.EXAMPLE
    .\cleanup.ps1
#>

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptDir

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Cleanup Script - Clean Slate" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$CleanupCount = 0

function Cleanup-Step {
    param(
        [string]$Name,
        [scriptblock]$Script
    )

    Write-Host ""
    Write-Host "[$($Name)]" -ForegroundColor Yellow
    try {
        & $Script
        Write-Host "  DONE" -ForegroundColor Green
        $script:ClenaupCount++
        return $true
    }
    catch {
        Write-Host "  SKIPPED: $($_.Exception.Message)" -ForegroundColor Yellow
        return $false
    }
}

# Step 1: Unregister any registered DLLs
Cleanup-Step "Step 1: Unregister DLLs" {
    Write-Host "  Unregistering NativeTTSWrapper DLLs..." -ForegroundColor Gray

    # Try to unregister both possible locations
    $dllPaths = @(
        "C:\github\SherpaOnnxAzureSAPI-installer\NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll",
        "C:\Program Files\OpenAssistive\OpenSpeech\NativeTTSWrapper.dll"
    )

    foreach ($dllPath in $dllPaths) {
        if (Test-Path $dllPath) {
            Write-Host "    Unregistering: $dllPath" -ForegroundColor Gray
            & regsvr32 /u /s "$dllPath" 2>$null
            Start-Sleep -Milliseconds 500
        } else {
            Write-Host "    Not found: $dllPath" -ForegroundColor DarkGray
        }
    }
}

# Step 2: Remove SAPI5 voice registrations
Cleanup-Step "Step 2: Remove SAPI5 Voice Registrations" {
    Write-Host "  Removing SAPI5 voice tokens from registry..." -ForegroundColor Gray

    # Remove mms_hat if it exists
    $voicePath = "HKLM:\SOFTWARE\Microsoft\Speech\Voices\Tokens\mms_hat"
    if (Test-Path $voicePath) {
        Write-Host "    Removing: $voicePath" -ForegroundColor Gray
        Remove-Item -Path $voicePath -Recurse -Force
    } else {
        Write-Host "    Not found: $voicePath" -ForegroundColor DarkGray
    }

    Write-Host "  Windows built-in voices left untouched" -ForegroundColor Cyan
}

# Step 3: Remove log files
Cleanup-Step "Step 3: Remove Log Files" {
    Write-Host "  Removing debug log files..." -ForegroundColor Gray

    $logPaths = @(
        "C:\OpenSpeech",
        "C:\Users\$env:USERNAME\AppData\Local\OpenSpeech",
        "C:\Users\*\AppData\Local\OpenSpeech"  # Wildcard for any user
    )

    foreach ($path in $logPaths) {
        if ($path -like '*\*') {
            # Handle wildcard paths
            $resolved = Resolve-Path $path -ErrorAction SilentlyContinue
            if ($resolved) {
                Write-Host "    Removing: $resolved" -ForegroundColor Gray
                Remove-Item -Path $resolved -Recurse -Force -ErrorAction SilentlyContinue
            }
        } elseif (Test-Path $path) {
            Write-Host "    Removing: $path" -ForegroundColor Gray
            Remove-Item -Path $path -Recurse -Force
        }
    }

    # Also remove log files in DLL directory
    $dllLogPath = "C:\github\SherpaOnnxAzureSAPI-installer\NativeTTSWrapper\x64\Release\native_tts_debug.log"
    if (Test-Path $dllLogPath) {
        Write-Host "    Removing: $dllLogPath" -ForegroundColor Gray
        Remove-Item -Path $dllLogPath -Force
    }

    # Also remove log files in script directory
    $logFiles = Get-ChildItem -Path $ScriptDir -Filter "*.log" -ErrorAction SilentlyContinue
    foreach ($logFile in $logFiles) {
        Write-Host "    Removing: $($logFile.FullName)" -ForegroundColor Gray
        Remove-Item -Path $logFile.FullName -Force
    }
}

# Step 4: Clean up temporary files
Cleanup-Step "Step 4: Remove Temporary Files" {
    Write-Host "  Cleaning temporary files..." -ForegroundColor Gray

    $tempDirs = @(
        "$env:LOCALAPPDATA\Temp\OpenSpeechTTS",
        "$env:TEMP\OpenSpeechTTS"
    )

    foreach ($dir in $tempDirs) {
        if (Test-Path $dir) {
            Write-Host "    Removing: $dir" -ForegroundColor Gray
            Remove-Item -Path $dir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

# Step 5: Clean up test/output files
Cleanup-Step "Step 5: Clean Up Test Files" {
    Write-Host "  Cleaning test and output files..." -ForegroundColor Gray

    $testFiles = @(
        "test-voice.ps1",
        "list-voices.ps1",
        "test-voice.vbs",
        "dump-voice-registry.ps1",
        "debug-paths.ps1",
        "test-com.ps1",
        "add-clsid.ps1",
        "fix-voice-registry.ps1",
        "update-config.ps1",
        "fix-voices-mapping.ps1",
        "add-mms-engine-alias.ps1",
        "update-mms-config.ps1",
        "fix-voice-clsid.ps1"
    )

    foreach ($file in $testFiles) {
        $filePath = Join-Path $ScriptDir $file
        if (Test-Path $filePath) {
            Write-Host "    Removing: $file" -ForegroundColor Gray
            Remove-Item -Path $filePath -Force
        }
    }
}

# Step 6: Verify cleanup
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Cleanup Complete" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "Status: $CleanupCount steps completed" -ForegroundColor Green

Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "  1. Build the DLL" -ForegroundColor White
Write-Host "     cd NativeTTSWrapper" -ForegroundColor Gray
Write-Host "     `"C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe` NativeTTSWrapper.sln /p:Configuration=Release /p:Platform=x64" -ForegroundColor Gray
Write-Host "     Or: dotnet build (for ConfigApp)" -ForegroundColor DarkGray
Write-Host ""
Write-Host "  2. Register the DLL (as Administrator)" -ForegroundColor White
Write-Host "     .\register-dll.ps1" -ForegroundColor Gray
Write-Host ""
Write-Host "  3. Use ConfigApp to install voices (as Administrator)" -ForegroundColor White
Write-Host "     .\ConfigApp\bin\Release\net8.0-windows\SherpaOnnxConfig.exe" -ForegroundColor Gray
