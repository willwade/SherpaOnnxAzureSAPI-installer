# Check if DLL dependencies are available
Write-Host "=== Checking DLL Dependencies ===" -ForegroundColor Cyan

$dllPath = "C:\github\SherpaOnnxAzureSAPI-installer\NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll"

Write-Host ""
Write-Host "DLL exists: $(Test-Path $dllPath)"
Write-Host ""

# Check for sherpa-onnx DLL
$sherpaDlls = @(
    "sherpa-onnx.dll",
    "kaldi-native-feat-common.dll",
    "sherpa-onnx-kaldifst-core.dll"
)

Write-Host "Checking for SherpaOnnx dependencies:" -ForegroundColor Yellow
$foundAny = $false
foreach ($dll in $sherpaDlls) {
    $paths = @(
        "C:\github\SherpaOnnxAzureSAPI-installer\NativeTTSWrapper\x64\Release\$dll",
        "C:\Program Files\OpenAssistive\OpenSpeech\$dll",
        "C:\Program Files\OpenSpeech\$dll"
    )

    foreach ($path in $paths) {
        if (Test-Path $path) {
            Write-Host "  Found: $dll at $path" -ForegroundColor Green
            $foundAny = $true
            break
        }
    }

    if (-not $foundAny) {
        Write-Host "  Missing: $dll" -ForegroundColor Red
    }
    $foundAny = $false
}

# Check for vcruntime DLLs
Write-Host ""
Write-Host "Checking for VC++ Runtime:" -ForegroundColor Yellow
$vcruntime = @("vcruntime140.dll", "vcruntime140_1.dll", "msvcp140.dll")
foreach ($dll in $vcruntime) {
    $systemPath = Join-Path $env:SystemRoot "System32\$dll"
    if (Test-Path $systemPath) {
        Write-Host "  Found: $dll" -ForegroundColor Green
    } else {
        Write-Host "  Missing: $dll" -ForegroundColor Red
    }
}

# Try to load the DLL using P/Invoke just to see what happens
Write-Host ""
Write-Host "Attempting to load DLL..." -ForegroundColor Yellow
try {
    $loaded = [System.Reflection.Assembly]::LoadFile($dllPath)
    Write-Host "  DLL loaded successfully" -ForegroundColor Green
    Write-Host "  Types: $($loaded.GetTypes().Count)"

    foreach ($type in $loaded.GetTypes()) {
        Write-Host "    - $type"
    }
}
catch {
    Write-Host "  Failed to load DLL: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "    Inner: $($_.Exception.InnerException.Message)"
}

# Check Windows Event Log for SAPI5 errors
Write-Host ""
Write-Host "Checking Windows Event Log for recent errors..." -ForegroundColor Yellow
try {
    $events = Get-WinEvent -FilterHashtable @{LogName='Application'; Level=2; StartTime=(Get-Date).AddMinutes(-5)} -MaxEvents 10 -ErrorAction SilentlyContinue
    if ($events) {
        Write-Host "  Found $($events.Count) recent errors"
        foreach ($event in $events) {
            Write-Host "    [$($event.TimeCreated)] $($event.Id): $($event.Message)" -ForegroundColor Red
        }
    } else {
        Write-Host "  No recent errors found"
    }
} catch {
    Write-Host "  Could not read event log"
}

Write-Host ""
Write-Host "=== Check Complete ===" -ForegroundColor Cyan
