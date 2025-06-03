# Test SAPI Voices - Real SAPI Integration Test
# This tests voices exactly as Windows applications would use them

param(
    [string]$VoiceName = "",
    [string]$TestText = "Hello! This is a test of the SAPI voice integration. The quick brown fox jumps over the lazy dog.",
    [switch]$ListOnly,
    [switch]$PlayAudio
)

Write-Host "🎤 SAPI VOICE INTEGRATION TEST" -ForegroundColor Cyan
Write-Host "=============================" -ForegroundColor Cyan
Write-Host ""

try {
    # Create SAPI voice object (same as any Windows app)
    Write-Host "Creating SAPI voice object..." -ForegroundColor Yellow
    $voice = New-Object -ComObject SAPI.SpVoice
    $voices = $voice.GetVoices()
    
    Write-Host "✅ SAPI initialized successfully" -ForegroundColor Green
    Write-Host "📊 Total voices available: $($voices.Count)" -ForegroundColor Cyan
    Write-Host ""
    
    # List all voices with details
    Write-Host "Available SAPI Voices:" -ForegroundColor Yellow
    Write-Host "=====================" -ForegroundColor Yellow
    
    $customVoices = @()
    $systemVoices = @()
    
    for ($i = 0; $i -lt $voices.Count; $i++) {
        $voiceToken = $voices.Item($i)
        $voiceName = $voiceToken.GetDescription()
        $voiceId = $voiceToken.Id
        
        # Get voice attributes
        try {
            $attributes = $voiceToken.GetAttribute("Name")
            $gender = $voiceToken.GetAttribute("Gender")
            $age = $voiceToken.GetAttribute("Age")
            $vendor = $voiceToken.GetAttribute("Vendor")
            $language = $voiceToken.GetAttribute("Language")
        } catch {
            $attributes = "N/A"
            $gender = "N/A"
            $age = "N/A"
            $vendor = "N/A"
            $language = "N/A"
        }
        
        Write-Host "Voice $($i + 1): $voiceName" -ForegroundColor White
        Write-Host "  ID: $voiceId" -ForegroundColor Gray
        if ($gender -ne "N/A") { Write-Host "  Gender: $gender" -ForegroundColor Gray }
        if ($vendor -ne "N/A") { Write-Host "  Vendor: $vendor" -ForegroundColor Gray }
        if ($language -ne "N/A") { Write-Host "  Language: $language" -ForegroundColor Gray }
        
        # Categorize voices
        if ($voiceName -like "*Azure*" -or $voiceName -like "*Libby*" -or $voiceName -like "*Jenny*") {
            Write-Host "  Type: 🌐 Azure TTS Voice" -ForegroundColor Green
            $customVoices += @{Name=$voiceName; Index=$i; Type="Azure"}
        } elseif ($voiceName -like "*amy*" -or $voiceName -like "*northern*" -or $voiceName -like "*sherpa*") {
            Write-Host "  Type: 🤖 SherpaOnnx Voice" -ForegroundColor Blue
            $customVoices += @{Name=$voiceName; Index=$i; Type="SherpaOnnx"}
        } else {
            Write-Host "  Type: 🖥️ System Voice" -ForegroundColor Gray
            $systemVoices += @{Name=$voiceName; Index=$i; Type="System"}
        }
        Write-Host ""
    }
    
    # Summary
    Write-Host "Summary:" -ForegroundColor Cyan
    Write-Host "  🌐 Azure TTS voices: $(($customVoices | Where-Object {$_.Type -eq 'Azure'}).Count)" -ForegroundColor Green
    Write-Host "  🤖 SherpaOnnx voices: $(($customVoices | Where-Object {$_.Type -eq 'SherpaOnnx'}).Count)" -ForegroundColor Blue
    Write-Host "  🖥️ System voices: $($systemVoices.Count)" -ForegroundColor Gray
    Write-Host ""
    
    if ($ListOnly) {
        Write-Host "List-only mode. Exiting." -ForegroundColor Yellow
        return
    }
    
    # Test specific voice or all custom voices
    $voicesToTest = @()
    
    if ($VoiceName) {
        # Test specific voice
        $found = $false
        for ($i = 0; $i -lt $voices.Count; $i++) {
            $voiceToken = $voices.Item($i)
            $voiceName = $voiceToken.GetDescription()
            if ($voiceName -like "*$VoiceName*") {
                $voicesToTest += @{Name=$voiceName; Index=$i; Token=$voiceToken}
                $found = $true
                break
            }
        }
        
        if (-not $found) {
            Write-Host "❌ Voice '$VoiceName' not found!" -ForegroundColor Red
            return
        }
    } else {
        # Test all custom voices
        foreach ($customVoice in $customVoices) {
            $voiceToken = $voices.Item($customVoice.Index)
            $voicesToTest += @{Name=$customVoice.Name; Index=$customVoice.Index; Token=$voiceToken; Type=$customVoice.Type}
        }
    }
    
    if ($voicesToTest.Count -eq 0) {
        Write-Host "⚠️ No custom voices to test. Only system voices found." -ForegroundColor Yellow
        return
    }
    
    # Test each voice
    Write-Host "🧪 TESTING VOICES" -ForegroundColor Cyan
    Write-Host "=================" -ForegroundColor Cyan
    Write-Host ""
    
    foreach ($testVoice in $voicesToTest) {
        Write-Host "Testing: $($testVoice.Name)" -ForegroundColor Yellow
        Write-Host "Text: '$TestText'" -ForegroundColor Gray
        
        try {
            # Set the voice
            $voice.Voice = $testVoice.Token
            Write-Host "✅ Voice set successfully" -ForegroundColor Green
            
            # Test synthesis
            Write-Host "🔊 Starting synthesis..." -ForegroundColor Cyan
            $startTime = Get-Date
            
            if ($PlayAudio) {
                # Synchronous - play audio
                $result = $voice.Speak($TestText, 0)
            } else {
                # Asynchronous - no audio output
                $result = $voice.Speak($TestText, 1)
            }
            
            $endTime = Get-Date
            $duration = ($endTime - $startTime).TotalMilliseconds
            
            Write-Host "✅ SUCCESS! Synthesis completed" -ForegroundColor Green
            Write-Host "   Duration: $([math]::Round($duration, 0))ms" -ForegroundColor Gray
            Write-Host "   Result code: $result" -ForegroundColor Gray
            
            if ($PlayAudio) {
                Write-Host "   🔊 Audio played through speakers" -ForegroundColor Green
            } else {
                Write-Host "   🔇 Silent mode (no audio output)" -ForegroundColor Gray
            }
            
        } catch {
            Write-Host "❌ FAILED: $($_.Exception.Message)" -ForegroundColor Red
            
            # Additional error details
            if ($_.Exception.Message -like "*Class not registered*") {
                Write-Host "   💡 COM object not registered - run installer as Administrator" -ForegroundColor Yellow
            } elseif ($_.Exception.Message -like "*E_FAIL*") {
                Write-Host "   💡 Engine initialization failed - check model files and configuration" -ForegroundColor Yellow
            }
        }
        
        Write-Host ""
    }
    
    Write-Host "🎯 SAPI VOICE TEST COMPLETED" -ForegroundColor Cyan
    
    if ($customVoices.Count -gt 0) {
        Write-Host ""
        Write-Host "💡 Usage Examples:" -ForegroundColor Yellow
        Write-Host "  Test specific voice: .\TestSAPIVoices.ps1 -VoiceName 'northern' -PlayAudio" -ForegroundColor Gray
        Write-Host "  List voices only: .\TestSAPIVoices.ps1 -ListOnly" -ForegroundColor Gray
        Write-Host "  Test all custom voices: .\TestSAPIVoices.ps1 -PlayAudio" -ForegroundColor Gray
    }
    
} catch {
    Write-Host "❌ SAPI Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "This indicates a fundamental SAPI issue." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Test completed!" -ForegroundColor Green
