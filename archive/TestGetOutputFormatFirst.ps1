# Test if SAPI calls GetOutputFormat before Speak
Write-Host "Testing GetOutputFormat Method Call" -ForegroundColor Green
Write-Host "===================================" -ForegroundColor Green

# Clear logs
Clear-Content "C:\OpenSpeech\sapi_debug.log" -ErrorAction SilentlyContinue

Write-Host "1. Testing direct GetOutputFormat call..." -ForegroundColor Cyan

try {
    # Create our COM object directly
    $comObject = New-Object -ComObject "OpenSpeechTTS.Sapi5VoiceImpl"
    Write-Host "   ✅ COM object created" -ForegroundColor Green
    
    # Try to call GetOutputFormat directly
    $type = $comObject.GetType()
    $getFormatMethod = $type.GetMethod("GetOutputFormat")
    
    if($getFormatMethod) {
        Write-Host "   Attempting to call GetOutputFormat directly..." -ForegroundColor Yellow
        
        try {
            # Create parameters for GetOutputFormat
            # This is complex due to ref/out parameters, but let's try
            $targetFormatId = [System.Guid]::Empty
            $targetWaveFormat = $null
            $outputFormatId = [System.Guid]::Empty
            $outputWaveFormat = [System.IntPtr]::Zero
            
            # This will likely fail due to parameter complexity, but should log the call
            $result = $getFormatMethod.Invoke($comObject, @([ref]$targetFormatId, [ref]$targetWaveFormat, [ref]$outputFormatId, [ref]$outputWaveFormat))
            Write-Host "   ✅ GetOutputFormat returned: $result" -ForegroundColor Green
        } catch {
            Write-Host "   ⚠️ GetOutputFormat failed (expected): $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    
} catch {
    Write-Host "❌ Error: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "2. Testing SAPI voice.Speak() to see method call order..." -ForegroundColor Cyan

try {
    # Create SAPI voice and try to speak
    $voice = New-Object -ComObject SAPI.SpVoice
    $voices = $voice.GetVoices()
    
    # Find Amy voice
    $amy = $null
    for($i=0; $i -lt $voices.Count; $i++) {
        if($voices.Item($i).GetDescription() -eq "amy") {
            $amy = $voices.Item($i)
            break
        }
    }
    
    if($amy) {
        Write-Host "   Setting Amy voice..." -ForegroundColor Yellow
        $voice.Voice = $amy
        
        Write-Host "   Calling voice.Speak() - watch for method call order..." -ForegroundColor Yellow
        try {
            $voice.Speak("Test")
        } catch {
            Write-Host "   Speak failed (expected)" -ForegroundColor Yellow
        }
    }
    
} catch {
    Write-Host "❌ SAPI Error: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "3. Analyzing method call order..." -ForegroundColor Cyan

if(Test-Path "C:\OpenSpeech\sapi_debug.log") {
    $logContent = Get-Content "C:\OpenSpeech\sapi_debug.log"
    
    $getOutputFormatCalled = $false
    $speakCalled = $false
    $callOrder = @()
    
    foreach($line in $logContent) {
        if($line -like "*GET OUTPUT FORMAT CALLED*") {
            $getOutputFormatCalled = $true
            $callOrder += "GetOutputFormat"
            Write-Host "   ✅ GetOutputFormat was called" -ForegroundColor Green
        }
        elseif($line -like "*SPEAK METHOD CALLED*") {
            $speakCalled = $true
            $callOrder += "Speak"
            Write-Host "   ✅ Speak was called" -ForegroundColor Green
        }
        elseif($line -like "*SET OBJECT TOKEN CALLED*") {
            $callOrder += "SetObjectToken"
            Write-Host "   ✅ SetObjectToken was called" -ForegroundColor Green
        }
        elseif($line -like "*ERROR*") {
            Write-Host "   ❌ Error: $line" -ForegroundColor Red
        }
    }
    
    if($callOrder.Count -gt 0) {
        Write-Host ""
        Write-Host "Method call order:" -ForegroundColor Yellow
        for($i=0; $i -lt $callOrder.Count; $i++) {
            Write-Host "   $($i+1). $($callOrder[$i])" -ForegroundColor White
        }
    } else {
        Write-Host "   ❌ NO METHODS WERE CALLED!" -ForegroundColor Red
    }
    
} else {
    Write-Host "   ❌ No debug log found" -ForegroundColor Red
}

Write-Host ""
Write-Host "=== CONCLUSION ===" -ForegroundColor Cyan
Write-Host "If GetOutputFormat is called but Speak is not:" -ForegroundColor Yellow
Write-Host "  → SAPI is failing in GetOutputFormat and not proceeding to Speak" -ForegroundColor White
Write-Host "If neither method is called:" -ForegroundColor Yellow
Write-Host "  → SAPI is not recognizing our object as a valid TTS engine" -ForegroundColor White
