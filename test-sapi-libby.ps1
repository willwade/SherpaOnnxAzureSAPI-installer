# OpenSpeechSAPI - Test SAPI Libby Voice - Comprehensive Test
# This script lists all SAPI voices and specifically tests the Libby voice

Write-Host "🎤 SAPI Voice Test - Libby Focus" -ForegroundColor Cyan
Write-Host "=" * 50 -ForegroundColor Cyan

try {
    # Create SAPI voice object
    Write-Host "🔧 Creating SAPI voice object..." -ForegroundColor Yellow
    $voice = New-Object -ComObject SAPI.SpVoice
    $voices = $voice.GetVoices()
    
    Write-Host "✅ SAPI object created successfully" -ForegroundColor Green
    Write-Host "📊 Found $($voices.Count) total SAPI voices" -ForegroundColor Green
    
    # List all voices with detailed info
    Write-Host "`n📋 All Available SAPI Voices:" -ForegroundColor Yellow
    Write-Host "-" * 40 -ForegroundColor Gray
    
    $libbyVoice = $null
    $libbyIndex = -1
    
    for ($i = 0; $i -lt $voices.Count; $i++) {
        $voiceInfo = $voices.Item($i)
        $name = $voiceInfo.GetDescription()
        $id = $voiceInfo.Id
        
        Write-Host "  $($i + 1). $name" -ForegroundColor White
        Write-Host "      ID: $id" -ForegroundColor Gray
        
        # Check if this is the Libby voice
        if ($name -like "*Libby*" -or $name -like "*British English*Azure*") {
            $libbyVoice = $voiceInfo
            $libbyIndex = $i
            Write-Host "      🎯 THIS IS THE LIBBY VOICE!" -ForegroundColor Green
        }
        Write-Host ""
    }
    
    # Test Libby voice specifically
    if ($libbyVoice) {
        Write-Host "🎯 Testing Libby Voice Specifically" -ForegroundColor Cyan
        Write-Host "-" * 40 -ForegroundColor Gray
        
        $libbyName = $libbyVoice.GetDescription()
        Write-Host "✅ Found Libby voice: $libbyName" -ForegroundColor Green
        Write-Host "📍 Voice index: $($libbyIndex + 1)" -ForegroundColor Green
        Write-Host "🆔 Voice ID: $($libbyVoice.Id)" -ForegroundColor Green
        
        # Set the voice
        Write-Host "`n🔧 Setting voice to Libby..." -ForegroundColor Yellow
        $voice.Voice = $libbyVoice
        
        # Verify the voice was set
        $currentVoice = $voice.Voice.GetDescription()
        Write-Host "✅ Current voice set to: $currentVoice" -ForegroundColor Green
        
        # Test speech synthesis
        Write-Host "`n🔊 Testing speech synthesis with Libby..." -ForegroundColor Yellow
        $testText = "Hello! This is a test of the Libby voice through SAPI. I am speaking using the British English Azure Libby voice. This should connect to AACSpeakHelper and return real audio bytes."
        
        Write-Host "📝 Text to speak: $testText" -ForegroundColor Gray
        Write-Host "⏱️  Starting synthesis..." -ForegroundColor Yellow
        
        $startTime = Get-Date
        
        # Speak the text (this should trigger our C++ COM wrapper)
        $voice.Speak($testText, 0)  # 0 = synchronous
        
        $endTime = Get-Date
        $duration = ($endTime - $startTime).TotalSeconds
        
        Write-Host "✅ Speech synthesis completed!" -ForegroundColor Green
        Write-Host "⏱️  Duration: $([math]::Round($duration, 2)) seconds" -ForegroundColor Green
        Write-Host "🎉 If you heard audio, the synth-to-bytestream pipeline is working!" -ForegroundColor Green
        
    } else {
        Write-Host "❌ Libby voice not found!" -ForegroundColor Red
        Write-Host "🔍 Looking for voices containing 'Libby', 'British English', or 'Azure'..." -ForegroundColor Yellow
        
        $foundSimilar = $false
        for ($i = 0; $i -lt $voices.Count; $i++) {
            $voiceInfo = $voices.Item($i)
            $name = $voiceInfo.GetDescription()
            
            if ($name -like "*British*" -or $name -like "*Azure*" -or $name -like "*Libby*") {
                Write-Host "🔍 Similar voice found: $name" -ForegroundColor Yellow
                $foundSimilar = $true
            }
        }
        
        if (-not $foundSimilar) {
            Write-Host "❌ No similar voices found. The Libby voice may not be properly registered." -ForegroundColor Red
            Write-Host "💡 Try running: uv run python sapi_voice_installer.py list" -ForegroundColor Yellow
        }
    }
    
    # Additional diagnostics
    Write-Host "`n🔍 Additional Diagnostics:" -ForegroundColor Cyan
    Write-Host "-" * 40 -ForegroundColor Gray
    
    # Check current voice
    $currentVoiceDesc = $voice.Voice.GetDescription()
    Write-Host "📍 Current SAPI voice: $currentVoiceDesc" -ForegroundColor White
    
    # Check voice rate and volume
    Write-Host "🔊 Current rate: $($voice.Rate)" -ForegroundColor White
    Write-Host "🔊 Current volume: $($voice.Volume)" -ForegroundColor White
    
} catch {
    Write-Host "❌ Error during SAPI test: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "🔧 Full error details:" -ForegroundColor Yellow
    Write-Host $_.Exception.ToString() -ForegroundColor Gray
}

Write-Host "`n🏁 Test Complete" -ForegroundColor Cyan
Write-Host "=" * 50 -ForegroundColor Cyan
