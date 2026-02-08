# Extended SAPI5 Test Suite
Add-Type -AssemblyName System.Speech

$ErrorActionPreference = "Stop"

Write-Host "=== SAPI5 Extended Test Suite ===" -ForegroundColor Cyan
Write-Host ""

# Create SAPI5 SpVoice object
$voice = New-Object -ComObject SAPI.SpVoice

# Find TestSherpaVoice
$voices = $voice.GetVoices()
$testVoice = $null
for ($i = 0; $i -lt $voices.Count; $i++) {
    if ($voices.Item($i).GetDescription() -eq "Test Sherpa Voice") {
        $testVoice = $voices.Item($i)
        break
    }
}

if (-not $testVoice) {
    Write-Host "ERROR: Test Sherpa Voice not found!" -ForegroundColor Red
    exit 1
}

$voice.Voice = $testVoice
Write-Host "Selected: $($testVoice.GetDescription())" -ForegroundColor Green
Write-Host ""

# Test cases
$testCases = @(
    @{ Name = "Short sentence"; Text = "Hello world, this is a test." },
    @{ Name = "Longer text"; Text = "The quick brown fox jumps over the lazy dog. This sentence contains all the letters of the alphabet and tests the speech synthesis engine's ability to handle longer text passages." },
    @{ Name = "Numbers"; Text = "I have 123 apples, 456 oranges, and 789 bananas. The total is 1368 fruits. The year is 2026 and the temperature is 23.5 degrees." },
    @{ Name = "Punctuation"; Text = "Hello! How are you? I'm doing great, thanks for asking. This test includes: commas, periods, exclamation marks, question marks, colons, and semi-colons; plus quotes like 'this'." },
    @{ Name = "Technical terms"; Text = "The API uses JSON for configuration. The model is vits-piper-en_US-amy-low with sample rate 16000 Hz. Files include onnxruntime.dll and sherpa-onnx-core.lib." },
    @{ Name = "Multiple sentences"; Text = "Speech synthesis is the artificial production of human speech. A computer system used for this purpose is called a speech synthesizer, and it can be implemented in software or hardware products." }
)

Write-Host "Running $($testCases.Count) test cases..." -ForegroundColor Yellow
Write-Host ""

$passed = 0
$failed = 0

foreach ($test in $testCases) {
    Write-Host "[$($testCases.IndexOf($test) + 1)/$($testCases.Count)] $($test.Name)" -ForegroundColor Cyan
    Write-Host "  Text: $($test.Text.Substring(0, [Math]::Min(80, $test.Text.Length)))..."

    try {
        $result = $voice.Speak($test.Text, 0)  # 0 = SVSFlagsAsync
        if ($result -eq 1) {
            Write-Host "  Result: PASSED" -ForegroundColor Green
            $passed++
        } else {
            Write-Host "  Result: FAILED (code: $result)" -ForegroundColor Red
            $failed++
        }
    } catch {
        Write-Host "  Result: EXCEPTION - $($_.Exception.Message)" -ForegroundColor Red
        $failed++
    }
    Write-Host ""
}

# Summary
Write-Host "=== Test Summary ===" -ForegroundColor Cyan
Write-Host "Total: $($testCases.Count) | Passed: $passed | Failed: $failed" -ForegroundColor $(if ($failed -eq 0) { "Green" } else { "Yellow" })

if ($failed -eq 0) {
    Write-Host ""
    Write-Host "All tests passed!" -ForegroundColor Green
}

# Cleanup
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($voice) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Host ""
Write-Host "Press Enter to exit..."
Read-Host
