# Quick test of native COM wrapper
Write-Host "Testing Native COM Wrapper" -ForegroundColor Green

# Test 1: Create native COM object
Write-Host "1. Testing native COM object creation..." -ForegroundColor Cyan
try {
    $nativeObj = New-Object -ComObject "NativeTTSWrapper.CNativeTTSWrapper"
    Write-Host "   SUCCESS: Native COM object created" -ForegroundColor Green
} catch {
    Write-Host "   ERROR: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Test 2: List voices
Write-Host "2. Testing voice enumeration..." -ForegroundColor Cyan
try {
    $voice = New-Object -ComObject SAPI.SpVoice
    $voices = $voice.GetVoices()
    Write-Host "   Found $($voices.Count) voices:" -ForegroundColor White
    
    $amyVoice = $null
    for ($i = 0; $i -lt $voices.Count; $i++) {
        $voiceToken = $voices.Item($i)
        $voiceName = $voiceToken.GetDescription()
        Write-Host "     - $voiceName" -ForegroundColor Gray
        
        if ($voiceName -like "*amy*") {
            $amyVoice = $voiceToken
        }
    }
    
    if ($amyVoice) {
        Write-Host "   SUCCESS: Amy voice found" -ForegroundColor Green
    } else {
        Write-Host "   WARNING: Amy voice not found" -ForegroundColor Yellow
    }
} catch {
    Write-Host "   ERROR: $($_.Exception.Message)" -ForegroundColor Red
}

# Test 3: Try speech synthesis
if ($amyVoice) {
    Write-Host "3. Testing speech synthesis..." -ForegroundColor Cyan
    try {
        $voice.Voice = $amyVoice
        Write-Host "   Voice set to Amy" -ForegroundColor White
        
        $result = $voice.Speak("Hello from native COM wrapper", 1)
        Write-Host "   SUCCESS: Speak returned $result" -ForegroundColor Green
        
        if ($result -eq 1) {
            Write-Host "   AMAZING: Speech synthesis worked!" -ForegroundColor Cyan
        }
    } catch {
        Write-Host "   ERROR: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "Test completed!" -ForegroundColor Green
