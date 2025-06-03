# Test SAPI with ProcessBridge Integration
Write-Host "Testing SAPI with ProcessBridge Integration" -ForegroundColor Green
Write-Host "===========================================" -ForegroundColor Green

Write-Host ""
Write-Host "1. Creating SAPI voice object..." -ForegroundColor Cyan

try {
    $voice = New-Object -ComObject SAPI.SpVoice
    Write-Host "   ✅ SAPI voice object created" -ForegroundColor Green
} catch {
    Write-Host "   ❌ Failed to create SAPI voice: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "2. Finding Amy voice..." -ForegroundColor Cyan

$voices = $voice.GetVoices()
$amyVoice = $null

for ($i = 0; $i -lt $voices.Count; $i++) {
    $voiceToken = $voices.Item($i)
    $voiceName = $voiceToken.GetDescription()
    Write-Host "   Voice $i`: $voiceName" -ForegroundColor White
    
    if ($voiceName -like "*amy*") {
        $amyVoice = $voiceToken
        Write-Host "   ✅ Amy voice found!" -ForegroundColor Green
    }
}

if (!$amyVoice) {
    Write-Host "   ❌ Amy voice not found" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "3. Setting Amy voice..." -ForegroundColor Cyan

try {
    $voice.Voice = $amyVoice
    Write-Host "   ✅ Amy voice set successfully" -ForegroundColor Green
    Write-Host "   Current voice: $($voice.Voice.GetDescription())" -ForegroundColor White
} catch {
    Write-Host "   ❌ Failed to set Amy voice: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "4. Testing speech synthesis..." -ForegroundColor Cyan

$testText = "ProcessBridge test successful"
Write-Host "   Text to speak: '$testText'" -ForegroundColor White

try {
    Write-Host "   Calling Speak method..." -ForegroundColor Yellow
    $startTime = Get-Date
    
    # Call Speak with synchronous flag (0)
    $result = $voice.Speak($testText, 0)
    
    $endTime = Get-Date
    $duration = ($endTime - $startTime).TotalMilliseconds
    
    Write-Host "   ✅ Speak method completed!" -ForegroundColor Green
    Write-Host "   Result: $result" -ForegroundColor White
    Write-Host "   Duration: $([math]::Round($duration, 0))ms" -ForegroundColor White
    
} catch {
    Write-Host "   ❌ Speak method failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "   Error details: $($_.Exception)" -ForegroundColor Red
}

Write-Host ""
Write-Host "5. Checking logs..." -ForegroundColor Cyan

$logDir = "C:\OpenSpeech"
if (Test-Path $logDir) {
    $logFiles = Get-ChildItem $logDir -Filter "*.log" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
    
    if ($logFiles.Count -gt 0) {
        Write-Host "   ✅ Log files found:" -ForegroundColor Green
        
        foreach ($logFile in $logFiles) {
            Write-Host "   📄 $($logFile.Name) - $($logFile.LastWriteTime)" -ForegroundColor Gray
            
            # Show recent entries from this log
            try {
                $logContent = Get-Content $logFile.FullName -Tail 5 -ErrorAction SilentlyContinue
                if ($logContent) {
                    Write-Host "     Recent entries:" -ForegroundColor Yellow
                    foreach ($line in $logContent) {
                        Write-Host "       $line" -ForegroundColor Gray
                    }
                }
            } catch {
                Write-Host "     (Could not read log file)" -ForegroundColor Yellow
            }
            Write-Host ""
        }
    } else {
        Write-Host "   ⚠️ No log files found" -ForegroundColor Yellow
    }
} else {
    Write-Host "   ⚠️ Log directory not found: $logDir" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=== SAPI PROCESSBRIDGE TEST COMPLETE ===" -ForegroundColor Cyan
