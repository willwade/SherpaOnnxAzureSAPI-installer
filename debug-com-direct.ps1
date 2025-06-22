#!/usr/bin/env pwsh

Write-Host "=== Direct COM Object Test ===" -ForegroundColor Cyan

try {
    # Try to create the COM object directly using our CORRECT CLSID
    $clsid = "{E1C4A8F2-9B3D-4A5E-8F7C-2D1B3E4F5A6B}"
    Write-Host "🔧 Attempting to create COM object with CLSID: $clsid" -ForegroundColor Yellow
    
    $comObject = New-Object -ComObject $clsid
    if ($comObject) {
        Write-Host "✅ COM object created successfully!" -ForegroundColor Green
        Write-Host "COM object type: $($comObject.GetType().FullName)"
    } else {
        Write-Host "❌ Failed to create COM object" -ForegroundColor Red
    }
}
catch {
    Write-Host "❌ Exception creating COM object: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "=== Registry Check ===" -ForegroundColor Cyan

# Check if our CLSID is registered
$clsidPath = "HKEY_CLASSES_ROOT\CLSID\{E1C4A8F2-9B3D-4A5E-8F7C-2D1B3E4F5A6B}"
try {
    $regKey = Get-ItemProperty -Path "Registry::$clsidPath" -ErrorAction Stop
    Write-Host "✅ CLSID is registered in registry" -ForegroundColor Green
    
    # Check InprocServer32
    $inprocPath = "$clsidPath\InprocServer32"
    $inprocKey = Get-ItemProperty -Path "Registry::$inprocPath" -ErrorAction SilentlyContinue
    if ($inprocKey) {
        Write-Host "✅ InprocServer32 found: $($inprocKey.'(default)')" -ForegroundColor Green
    } else {
        Write-Host "❌ InprocServer32 not found" -ForegroundColor Red
    }
}
catch {
    Write-Host "❌ CLSID not found in registry" -ForegroundColor Red
}

Write-Host ""
Write-Host "=== Check Log File ===" -ForegroundColor Cyan

$logPath = "C:\OpenSpeech\native_tts_debug.log"
if (Test-Path $logPath) {
    Write-Host "✅ Log file exists: $logPath" -ForegroundColor Green
    Write-Host "📄 Last 10 lines:" -ForegroundColor Yellow
    Get-Content $logPath -Tail 10
} else {
    Write-Host "❌ Log file not found: $logPath" -ForegroundColor Red
    Write-Host "This means the COM wrapper is not being instantiated at all."
}

Write-Host ""
Write-Host "=== Test Complete ===" -ForegroundColor Cyan
