# Test Amy Voice Specifically - SAPI to SherpaOnnx Pipeline
Write-Host "🎯 Testing Amy Voice - Complete SAPI to SherpaOnnx Pipeline" -ForegroundColor Cyan
Write-Host "=============================================================" -ForegroundColor Cyan

try {
    # Create SAPI voice object
    Write-Host "1. Creating SAPI voice object..." -ForegroundColor Yellow
    $voice = New-Object -ComObject SAPI.SpVoice
    $voices = $voice.GetVoices()
    
    Write-Host "   Found $($voices.Count) voices total" -ForegroundColor Green
    
    # Find Amy voice specifically
    Write-Host "2. Looking for Amy voice..." -ForegroundColor Yellow
    $amyVoice = $null
    
    for ($i = 0; $i -lt $voices.Count; $i++) {
        $currentVoice = $voices.Item($i)
        $voiceName = $currentVoice.GetDescription()
        Write-Host "   Voice $($i+1): $voiceName" -ForegroundColor Gray
        
        if ($voiceName -like "*amy*" -or $voiceName -eq "amy") {
            $amyVoice = $currentVoice
            Write-Host "   ✅ Found Amy voice: $voiceName" -ForegroundColor Green
            break
        }
    }
    
    if (-not $amyVoice) {
        Write-Host "   ❌ Amy voice not found in SAPI enumeration" -ForegroundColor Red
        Write-Host "   Available voices:" -ForegroundColor Yellow
        for ($i = 0; $i -lt $voices.Count; $i++) {
            $voiceName = $voices.Item($i).GetDescription()
            Write-Host "     - $voiceName" -ForegroundColor Gray
        }
        exit 1
    }
    
    # Test Amy voice selection
    Write-Host "3. Testing Amy voice selection..." -ForegroundColor Yellow
    try {
        $voice.Voice = $amyVoice
        Write-Host "   ✅ Amy voice selected successfully" -ForegroundColor Green
    } catch {
        Write-Host "   ❌ Failed to select Amy voice: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "   HRESULT: 0x$($_.Exception.HResult.ToString('X8'))" -ForegroundColor Red
        exit 1
    }
    
    # Test speech synthesis with Amy
    Write-Host "4. Testing speech synthesis with Amy..." -ForegroundColor Yellow
    
    $testTexts = @(
        @{ Text = "Hello, this is Amy speaking."; Description = "Simple greeting" },
        @{ Text = "The quick brown fox jumps over the lazy dog."; Description = "Pangram test" },
        @{ Text = "Testing SherpaOnnx integration with SAPI."; Description = "Technical test" }
    )
    
    foreach ($test in $testTexts) {
        Write-Host "   Testing: '$($test.Text)' ($($test.Description))..." -ForegroundColor Cyan
        
        try {
            # Test synchronous speech
            $startTime = Get-Date
            $result = $voice.Speak($test.Text, 0) # 0 = synchronous
            $endTime = Get-Date
            $duration = ($endTime - $startTime).TotalMilliseconds
            
            Write-Host "   ✅ SUCCESS: Synthesis completed in $([math]::Round($duration, 0))ms (result: $result)" -ForegroundColor Green
            
            # Check if audio was actually generated (look for temp files or logs)
            $tempFiles = Get-ChildItem $env:TEMP -Filter "*amy*" -ErrorAction SilentlyContinue
            if ($tempFiles) {
                Write-Host "   📁 Found temp files: $($tempFiles.Count) files" -ForegroundColor Gray
            }
            
        } catch {
            Write-Host "   ❌ FAILED: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "   HRESULT: 0x$($_.Exception.HResult.ToString('X8'))" -ForegroundColor Red
            
            # Check for specific error codes
            $hresult = $_.Exception.HResult
            switch ($hresult) {
                0x80004005 { Write-Host "   → Unspecified error (E_FAIL)" -ForegroundColor Yellow }
                0x80070057 { Write-Host "   → Invalid parameter (E_INVALIDARG)" -ForegroundColor Yellow }
                0x8007000E { Write-Host "   → Out of memory (E_OUTOFMEMORY)" -ForegroundColor Yellow }
                0x80040154 { Write-Host "   → Class not registered (REGDB_E_CLASSNOTREG)" -ForegroundColor Yellow }
                default { Write-Host "   → Unknown error code" -ForegroundColor Yellow }
            }
        }
        
        Start-Sleep -Milliseconds 500
    }
    
    # Test asynchronous speech
    Write-Host "5. Testing asynchronous speech..." -ForegroundColor Yellow
    try {
        $result = $voice.Speak("This is an asynchronous test.", 1) # 1 = asynchronous
        Write-Host "   ✅ Asynchronous speech started (result: $result)" -ForegroundColor Green
        
        # Wait a bit for completion
        Start-Sleep -Seconds 2
        
    } catch {
        Write-Host "   ❌ Asynchronous speech failed: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # Check logs
    Write-Host "6. Checking debug logs..." -ForegroundColor Yellow
    $logPaths = @(
        "C:\OpenSpeech\engine_manager.log",
        "C:\OpenSpeech\native_tts_debug.log",
        "C:\OpenSpeech\sapi_debug.log"
    )
    
    foreach ($logPath in $logPaths) {
        if (Test-Path $logPath) {
            Write-Host "   📄 Found log: $logPath" -ForegroundColor Green
            $logContent = Get-Content $logPath -Tail 5 -ErrorAction SilentlyContinue
            if ($logContent) {
                Write-Host "   Recent entries:" -ForegroundColor Gray
                $logContent | ForEach-Object { Write-Host "     $_" -ForegroundColor Gray }
            }
        } else {
            Write-Host "   ⚠️  Log not found: $logPath" -ForegroundColor Yellow
        }
    }
    
    # Summary
    Write-Host ""
    Write-Host "🎉 Amy Voice Test Summary:" -ForegroundColor Cyan
    Write-Host "✅ Amy voice found in SAPI enumeration" -ForegroundColor Green
    Write-Host "✅ Amy voice selection successful" -ForegroundColor Green
    Write-Host "🔄 Speech synthesis results above" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "📋 Next Steps:" -ForegroundColor Cyan
    Write-Host "   1. If synthesis failed, check COM wrapper implementation" -ForegroundColor White
    Write-Host "   2. Verify native engine initialization with real models" -ForegroundColor White
    Write-Host "   3. Test direct SherpaOnnx fallback method" -ForegroundColor White
    Write-Host "   4. Check ProcessBridge as final fallback" -ForegroundColor White
    
} catch {
    Write-Host "❌ CRITICAL ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack trace: $($_.Exception.StackTrace)" -ForegroundColor Gray
    exit 1
}

Write-Host ""
Write-Host "🎯 Amy Voice Test Complete!" -ForegroundColor Cyan
