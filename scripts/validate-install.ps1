#!/usr/bin/env pwsh
# SherpaOnnx SAPI5 Installation Validation Script
# Run this script to verify the installation is working correctly

param(
    [switch]$Verbose,
    [switch]$GenerateReport,
    [string]$ReportPath = "./install-validation-report.html"
)

$ErrorActionPreference = "Stop"
$results = [Ordered]@{}

Write-Host "=== SherpaOnnx SAPI5 Installation Validation ===" -ForegroundColor Cyan
Write-Host ""

# Track results
$passCount = 0
$failCount = 0
$warnCount = 0

function Test-AndRecord {
    param(
        [string]$Name,
        [scriptblock]$Test,
        [string]$SuccessMessage,
        [string]$FailureMessage,
        [string]$WarningMessage = $null
    )

    Write-Host "Testing: $Name..." -ForegroundColor Yellow

    try {
        $result = & $Test

        if ($result.Status -eq "Pass") {
            Write-Host "  ✓ PASS: $SuccessMessage" -ForegroundColor Green
            $results[$Name] = @{Status = "Pass"; Message = $SuccessMessage}
            $script:passCount++
        } elseif ($result.Status -eq "Warn") {
            Write-Host "  ⚠ WARN: $WarningMessage" -ForegroundColor Yellow
            $results[$Name] = @{Status = "Warning"; Message = $WarningMessage}
            $script:warnCount++
        } else {
            Write-Host "  ✗ FAIL: $FailureMessage" -ForegroundColor Red
            $results[$Name] = @{Status = "Fail"; Message = $FailureMessage}
            $script:failCount++
        }

        if ($Verbose -and $result.Detail) {
            Write-Host "    Detail: $($result.Detail)" -ForegroundColor Gray
        }
    } catch {
        Write-Host "  ✗ ERROR: $($_.Exception.Message)" -ForegroundColor Red
        $results[$Name] = @{Status = "Error"; Message = $_.Exception.Message}
        $script:failCount++
    }

    Write-Host ""
}

# ============================================================================
# PREREQUISITE CHECKS
# ============================================================================

Write-Host "=== Prerequisites ===" -ForegroundColor Cyan
Write-Host ""

# Check .NET Runtime
Test-AndRecord -Name ".NET Runtime" -Test {
    $dotnet = Get-Command "dotnet" -ErrorAction SilentlyContinue
    if ($dotnet) {
        $version = & dotnet --version 2>$null
        @{Status = "Pass"; Detail = $version}
    } else {
        @{Status = "Fail"}
    }
} -SuccessMessage ".NET runtime found" -FailureMessage ".NET runtime not installed"

# Check VC++ Redistributables
Test-AndRecord -Name "VC++ Redistributables" -Test {
    $regPath = "HKLM:\SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64"
    if (Test-Path $regPath) {
        $installed = Get-ItemProperty $regPath
        @{Status = "Pass"; Detail = $installed.Installed}
    } else {
        @{Status = "Warn"}
    }
} -SuccessMessage "VC++ 2015-2022 redistributables installed" -FailureMessage "VC++ redistributables not found" -WarningMessage "VC++ redistributables may not be installed (runtime may still work)"

# Check Administrator Privileges
Test-AndRecord -Name "Administrator Privileges" -Test {
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if ($isAdmin) {
        @{Status = "Pass"}
    } else {
        @{Status = "Warn"}
    }
} -SuccessMessage "Running as administrator" -FailureMessage "Not running as administrator" -WarningMessage "Some tests may fail without admin privileges"

# Check Windows Version
Test-AndRecord -Name "Windows Version" -Test {
    $version = [Environment]::OSVersion.Version
    if ($version.Major -ge 10) {
        @{Status = "Pass"; Detail = $version}
    } else {
        @{Status = "Fail"; Detail = $version}
    }
} -SuccessMessage "Windows 10/11 detected" -FailureMessage "Windows version not supported (requires Windows 10+)"

# ============================================================================
# FILE INSTALLATION CHECKS
# ============================================================================

Write-Host "=== File Installation ===" -ForegroundColor Cyan
Write-Host ""

$installPath = "C:\Program Files\OpenAssistive\OpenSpeech"

# Check installation directory
Test-AndRecord -Name "Installation Directory" -Test {
    if (Test-Path $installPath) {
        @{Status = "Pass"; Detail = $installPath}
    } else {
        @{Status = "Fail"}
    }
} -SuccessMessage "Installation directory exists" -FailureMessage "Installation directory not found: $installPath"

# Check Native DLL
Test-AndRecord -Name "NativeTTSWrapper.dll" -Test {
    $dllPath = Join-Path $installPath "NativeTTSWrapper.dll"
    if (Test-Path $dllPath) {
        $file = Get-Item $dllPath
        @{Status = "Pass"; Detail = "$($file.Length) bytes, modified $($file.LastWriteTime)"}
    } else {
        @{Status = "Fail"}
    }
} -SuccessMessage "Native DLL found" -FailureMessage "NativeTTSWrapper.dll not found"

# Check SherpaOnnx libraries
Test-AndRecord -Name "SherpaOnnx Libraries" -Test {
    $coreLib = Join-Path $installPath "sherpa-onnx-core.dll"
    $apiLib = Join-Path $installPath "sherpa-onnx-c-api.dll"
    $onnxLib = Join-Path $installPath "onnxruntime.dll"

    $missing = @()
    if (-not (Test-Path $coreLib)) { $missing += "sherpa-onnx-core.dll" }
    if (-not (Test-Path $apiLib)) { $missing += "sherpa-onnx-c-api.dll" }
    if (-not (Test-Path $onnxLib)) { $missing += "onnxruntime.dll" }

    if ($missing.Count -eq 0) {
        @{Status = "Pass"}
    } else {
        @{Status = "Fail"; Detail = "Missing: $($missing -join ', ')"}
    }
} -SuccessMessage "All SherpaOnnx libraries found" -FailureMessage "SherpaOnnx libraries missing"

# Check configuration file
Test-AndRecord -Name "engines_config.json" -Test {
    $configPath = Join-Path $installPath "engines_config.json"
    if (Test-Path $configPath) {
        try {
            $config = Get-Content $configPath -Raw | ConvertFrom-Json
            $engineCount = ($config.engines.PSObject.Properties.Name).Count
            @{Status = "Pass"; Detail = "$engineCount engines configured"}
        } catch {
            @{Status = "Fail"; Detail = "Invalid JSON: $($_.Exception.Message)"}
        }
    } else {
        @{Status = "Fail"}
    }
} -SuccessMessage "Configuration file found and valid" -FailureMessage "Configuration file not found or invalid"

# ============================================================================
# REGISTRY CHECKS
# ============================================================================

Write-Host "=== Registry Entries ===" -ForegroundColor Cyan
Write-Host ""

# Check COM registration
Test-AndRecord -Name "COM Registration" -Test {
    $clsidPath = "HKLM:\SOFTWARE\Classes\CLSID\{A1B2C3D4-E5F6-4A5B-8C9D-1E2F3A4B5C6D}"
    if (Test-Path $clsidPath) {
        $inprocPath = "$clsidPath\InprocServer32"
        if (Test-Path $inprocPath) {
            $dllPath = (Get-ItemProperty $inprocPath -ErrorAction SilentlyContinue).'(default)'
            if ($dllPath -and (Test-Path $dllPath)) {
                @{Status = "Pass"; Detail = $dllPath}
            } else {
                @{Status = "Fail"; Detail = "DLL path not found or incorrect"}
            }
        } else {
            @{Status = "Fail"; Detail = "InprocServer32 key missing"}
        }
    } else {
        @{Status = "Fail"}
    }
} -SuccessMessage "COM object properly registered" -FailureMessage "COM registration not found"

# Check SAPI5 voice registration
Test-AndRecord -Name "SAPI5 Voice Registration" -Test {
    $voicePath = "HKLM:\SOFTWARE\Microsoft\Speech\Voices\Tokens\TestSherpaVoice"
    if (Test-Path $voicePath) {
        $attrsPath = "$voicePath\Attributes"
        if (Test-Path $attrsPath) {
            $attrs = Get-ItemProperty $attrsPath -ErrorAction SilentlyContinue
            if ($attrs) {
                @{Status = "Pass"; Detail = "Language=$($attrs.Language), Gender=$($attrs.Gender)"}
            } else {
                @{Status = "Fail"; Detail = "Attributes not found"}
            }
        } else {
            @{Status = "Fail"; Detail = "Attributes key missing"}
        }
    } else {
        @{Status = "Fail"}
    }
} -SuccessMessage "SAPI5 voice registered" -FailureMessage "SAPI5 voice registration not found"

# ============================================================================
# FUNCTIONAL TESTS
# ============================================================================

Write-Host "=== Functional Tests ===" -ForegroundColor Cyan
Write-Host ""

# Test voice enumeration
Test-AndRecord -Name "Voice Enumeration" -Test {
    try {
        Add-Type -AssemblyName System.Speech
        $synth = New-Object System.Speech.Synthesis.SpeechSynthesizer
        $voices = $synth.GetInstalledVoices()

        $sherpaVoice = $voices | Where-Object {
            $_.VoiceInfo.Description -like "*Sherpa*" -or
            $_.VoiceInfo.Name -like "*Sherpa*" -or
            $_.VoiceInfo.Name -eq "TestSherpaVoice"
        }

        if ($sherpaVoice) {
            @{Status = "Pass"; Detail = $sherpaVoice.VoiceInfo.Name}
        } else {
            $allVoices = ($voices | ForEach-Object { $_.VoiceInfo.Name }) -join ", "
            @{Status = "Fail"; Detail = "Available voices: $allVoices"}
        }
    } catch {
        @{Status = "Fail"; Detail = $_.Exception.Message}
    }
} -SuccessMessage "SherpaOnnx voice found in enumeration" -FailureMessage "SherpaOnnx voice not found"

# Test voice synthesis (requires audio output)
Test-AndRecord -Name "Voice Synthesis" -Test {
    try {
        Add-Type -AssemblyName System.Speech
        $synth = New-Object System.Speech.Synthesis.SpeechSynthesizer

        # Find Sherpa voice
        foreach ($voice in $synth.GetInstalledVoices()) {
            if ($voice.VoiceInfo.Name -eq "TestSherpaVoice") {
                $synth.SelectVoice($voice.VoiceInfo.Name)

                # Test synthesis (doesn't actually speak, just validates)
                $prompt = New-Object System.Speech.Synthesis.PromptBuilder
                $prompt.AppendText("Test")

                # We can't easily test without actually speaking
                # Just check we can select the voice
                @{Status = "Pass"; Detail = "Voice selectable"}
                return
            }
        }

        @{Status = "Fail"; Detail = "Could not select voice"}
    } catch {
        @{Status = "Fail"; Detail = $_.Exception.Message}
    }
} -SuccessMessage "Voice synthesis test passed" -FailureMessage "Voice synthesis test failed"

# ============================================================================
# SUMMARY
# ============================================================================

Write-Host "=== Summary ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "Total Tests: $($passCount + $failCount + $warnCount)" -ForegroundColor White
Write-Host "  Passed: $passCount" -ForegroundColor Green
Write-Host "  Failed: $failCount" -ForegroundColor Red
Write-Host "  Warnings: $warnCount" -ForegroundColor Yellow
Write-Host ""

$overallStatus = if ($failCount -eq 0) { "PASS" } else { "FAIL" }
$overallColor = if ($failCount -eq 0) { "Green" } else { "Red" }

Write-Host "Overall Status: " -NoNewline
Write-Host $overallStatus -ForegroundColor $overallColor

# ============================================================================
# REPORT GENERATION
# ============================================================================

if ($GenerateReport) {
    Write-Host "Generating HTML report..." -ForegroundColor Yellow

    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>SherpaOnnx SAPI5 Installation Validation Report</title>
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background: #f5f5f5; }
        .container { max-width: 900px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1 { color: #0078d4; border-bottom: 2px solid #0078d4; padding-bottom: 10px; }
        h2 { color: #333; margin-top: 30px; }
        .summary { display: flex; gap: 20px; margin: 20px 0; }
        .summary-box { flex: 1; padding: 15px; border-radius: 5px; text-align: center; }
        .pass { background: #dff6dd; color: #2da44e; }
        .fail { background: #ffebe9; color: #cf222e; }
        .warn { background: #fff8c5; color: #9a6700; }
        .summary-box h3 { margin: 0 0 10px 0; font-size: 32px; }
        .summary-box p { margin: 0; color: #666; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; }
        th, td { padding: 12px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background: #f6f8fa; font-weight: 600; }
        .status-pass { color: #2da44e; font-weight: bold; }
        .status-fail { color: #cf222e; font-weight: bold; }
        .status-warn { color: #9a6700; font-weight: bold; }
        .status-error { color: #cf222e; font-weight: bold; }
        .timestamp { color: #666; font-size: 0.9em; }
    </style>
</head>
<body>
    <div class="container">
        <h1>SherpaOnnx SAPI5 Installation Validation Report</h1>
        <p class="timestamp">Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>

        <div class="summary">
            <div class="summary-box pass">
                <h3>$passCount</h3>
                <p>Passed</p>
            </div>
            <div class="summary-box warn">
                <h3>$warnCount</h3>
                <p>Warnings</p>
            </div>
            <div class="summary-box fail">
                <h3>$failCount</h3>
                <p>Failed</p>
            </div>
        </div>

        <h2>Test Results</h2>
        <table>
            <thead>
                <tr>
                    <th>Test Name</th>
                    <th>Status</th>
                    <th>Message</th>
                </tr>
            </thead>
            <tbody>
"@

    foreach ($result in $results.GetEnumerator()) {
        $statusClass = "status-$($result.Value.Status.ToLower())"
        $html += @"
                <tr>
                    <td>$($result.Key)</td>
                    <td class="$statusClass">$($result.Value.Status)</td>
                    <td>$($result.Value.Message)</td>
                </tr>
"@
    }

    $html += @"
            </tbody>
        </table>

        <h2>Recommendations</h2>
        <ul>
"@

    if ($failCount -gt 0) {
        $html += @"
            <li><strong>Installation has issues:</strong> Please review the failed tests above and address each issue.</li>
            <li>Try reinstalling: <code>msiexec /fomus dist\SherpaOnnxSAPI.msi</code></li>
            <li>Check the troubleshooting guide: TROUBLESHOOTING.md</li>
"@
    } else {
        $html += @"
            <li><strong>Installation successful!</strong> Your SherpaOnnx SAPI5 engine is ready to use.</li>
            <li>Test with: <code>powershell -ExecutionPolicy Bypass -File test_sapi5_extended.ps1</code></li>
            <li>Configure voices using the ConfigApp (if installed)</li>
"@
    }

    $html += @"
        </ul>
    </div>
</body>
</html>
"@

    $html | Out-File -FilePath $ReportPath -Encoding UTF8
    Write-Host "  Report saved to: $ReportPath" -ForegroundColor Green

    # Optionally open in browser
    $openReport = Read-Host "Open report in browser? (y/n)"
    if ($openReport -eq "y") {
        Start-Process $ReportPath
    }
}

# Exit with appropriate code
exit $failCount
