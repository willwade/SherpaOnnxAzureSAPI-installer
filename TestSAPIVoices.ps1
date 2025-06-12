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
        $currentVoiceName = $voiceToken.GetDescription()
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

        Write-Host "Voice $($i + 1): $currentVoiceName" -ForegroundColor White
        Write-Host "  ID: $voiceId" -ForegroundColor Gray
        if ($gender -ne "N/A") { Write-Host "  Gender: $gender" -ForegroundColor Gray }
        if ($vendor -ne "N/A") { Write-Host "  Vendor: $vendor" -ForegroundColor Gray }
        if ($language -ne "N/A") { Write-Host "  Language: $language" -ForegroundColor Gray }

        # Categorize voices
        if ($currentVoiceName -like "*Azure*" -or $currentVoiceName -like "*Libby*" -or $currentVoiceName -like "*Jenny*") {
            Write-Host "  Type: 🌐 Azure TTS Voice" -ForegroundColor Green
            Write-Host "DEBUG: Adding Azure voice '$currentVoiceName' at index $i" -ForegroundColor Magenta
            $customVoices += @{Name=$currentVoiceName; Index=$i; Type="Azure"}
        } elseif ($currentVoiceName -like "*amy*" -or $currentVoiceName -like "*northern*" -or $currentVoiceName -like "*sherpa*") {
            Write-Host "  Type: 🤖 SherpaOnnx Voice" -ForegroundColor Blue
            Write-Host "DEBUG: Adding SherpaOnnx voice '$currentVoiceName' at index $i" -ForegroundColor Magenta
            $customVoices += @{Name=$currentVoiceName; Index=$i; Type="SherpaOnnx"}
        } else {
            Write-Host "  Type: 🖥️ System Voice" -ForegroundColor Gray
            $systemVoices += @{Name=$currentVoiceName; Index=$i; Type="System"}
        }
        Write-Host ""
    }
    
    # Summary
    Write-Host "Summary:" -ForegroundColor Cyan
    Write-Host "DEBUG: customVoices array count: $($customVoices.Count)" -ForegroundColor Magenta
    Write-Host "DEBUG: customVoices type: $($customVoices.GetType().Name)" -ForegroundColor Magenta
    Write-Host "DEBUG: customVoices content: $($customVoices | ConvertTo-Json -Depth 2)" -ForegroundColor Magenta
    Write-Host "DEBUG: systemVoices array count: $($systemVoices.Count)" -ForegroundColor Magenta
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

    Write-Host "DEBUG: VoiceName parameter: '$VoiceName'" -ForegroundColor Magenta
    if ($VoiceName) {
        Write-Host "DEBUG: Taking specific voice branch" -ForegroundColor Magenta
        # Test specific voice
        $found = $false
        for ($i = 0; $i -lt $voices.Count; $i++) {
            $voiceToken = $voices.Item($i)
            $currentVoiceName = $voiceToken.GetDescription()
            if ($currentVoiceName -like "*$VoiceName*") {
                $voicesToTest += @{Name=$currentVoiceName; Index=$i; Token=$voiceToken}
                $found = $true
                break
            }
        }
        
        if (-not $found) {
            Write-Host "❌ Voice '$VoiceName' not found!" -ForegroundColor Red
            return
        }
    } else {
        Write-Host "DEBUG: Taking all custom voices branch" -ForegroundColor Magenta
        # Test all custom voices
        Write-Host "DEBUG: Found $($customVoices.Count) custom voices to test" -ForegroundColor Magenta
        Write-Host "DEBUG: customVoices is array: $($customVoices -is [array])" -ForegroundColor Magenta

        # Force array iteration
        @($customVoices) | ForEach-Object {
            $customVoice = $_
            Write-Host "DEBUG: Processing custom voice: Name='$($customVoice.Name)', Index=$($customVoice.Index), Type=$($customVoice.Type)" -ForegroundColor Magenta
            $voiceToken = $voices.Item($customVoice.Index)
            $actualVoiceName = $voiceToken.GetDescription()
            Write-Host "DEBUG: Voice token at index $($customVoice.Index) is actually: '$actualVoiceName'" -ForegroundColor Magenta
            $voicesToTest += @{Name=$customVoice.Name; Index=$customVoice.Index; Token=$voiceToken; Type=$customVoice.Type}
        }
        Write-Host "DEBUG: Total voices to test: $($voicesToTest.Count)" -ForegroundColor Magenta
    }
    
    Write-Host "DEBUG: voicesToTest count: $($voicesToTest.Count)" -ForegroundColor Magenta
    if ($voicesToTest.Count -eq 0) {
        Write-Host "⚠️ No custom voices to test. Only system voices found." -ForegroundColor Yellow
        Write-Host "DEBUG: customVoices count was: $($customVoices.Count)" -ForegroundColor Magenta
        return
    }
    
    # Test each voice
    Write-Host "🧪 TESTING VOICES" -ForegroundColor Cyan
    Write-Host "=================" -ForegroundColor Cyan
    Write-Host ""
    
    foreach ($testVoice in $voicesToTest) {
        Write-Host "DEBUG: Testing voice at index $($testVoice.Index): $($testVoice.Name)" -ForegroundColor Magenta
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
