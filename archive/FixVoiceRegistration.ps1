# Fix voice registration by adding missing registry entries
# This script must be run as Administrator

Write-Host "Fixing Voice Registration with Missing Registry Entries" -ForegroundColor Green
Write-Host "======================================================" -ForegroundColor Green

$voiceToken = "HKLM:\SOFTWARE\Microsoft\Speech\Voices\Tokens\amy"

Write-Host "1. Adding missing LangDataPath and VoicePath entries..." -ForegroundColor Cyan

try {
    # Add LangDataPath (we'll create a dummy path since we don't use language data files)
    $langDataPath = "C:\Program Files\OpenAssistive\OpenSpeech\LangData.dat"
    Set-ItemProperty -Path $voiceToken -Name "LangDataPath" -Value $langDataPath
    Write-Host "   ✅ Added LangDataPath: $langDataPath" -ForegroundColor Green
    
    # Add VoicePath (we'll create a dummy path since we don't use voice data files)
    $voicePath = "C:\Program Files\OpenAssistive\OpenSpeech\VoiceData"
    Set-ItemProperty -Path $voiceToken -Name "VoicePath" -Value $voicePath
    Write-Host "   ✅ Added VoicePath: $voicePath" -ForegroundColor Green
    
    # Create dummy files so the paths exist
    $langDataDir = Split-Path $langDataPath
    if(-not (Test-Path $langDataDir)) {
        New-Item -Path $langDataDir -ItemType Directory -Force | Out-Null
    }
    if(-not (Test-Path $langDataPath)) {
        "# Dummy language data file for Amy voice" | Out-File -FilePath $langDataPath -Encoding UTF8
        Write-Host "   ✅ Created dummy LangData file" -ForegroundColor Green
    }
    
    if(-not (Test-Path $voicePath)) {
        New-Item -Path $voicePath -ItemType Directory -Force | Out-Null
        "# Dummy voice data directory for Amy voice" | Out-File -FilePath "$voicePath\VoiceInfo.txt" -Encoding UTF8
        Write-Host "   ✅ Created dummy VoiceData directory" -ForegroundColor Green
    }
    
} catch {
    Write-Host "❌ Error adding registry entries: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Make sure you're running as Administrator!" -ForegroundColor Yellow
    exit 1
}

Write-Host ""
Write-Host "2. Verifying updated voice registration..." -ForegroundColor Cyan

try {
    $tokenProps = Get-ItemProperty $voiceToken
    Write-Host "   Updated voice token properties:" -ForegroundColor White
    Write-Host "     CLSID: $($tokenProps.CLSID)" -ForegroundColor Gray
    Write-Host "     Path: $($tokenProps.Path)" -ForegroundColor Gray
    Write-Host "     LangDataPath: $($tokenProps.LangDataPath)" -ForegroundColor Gray
    Write-Host "     VoicePath: $($tokenProps.VoicePath)" -ForegroundColor Gray
    
} catch {
    Write-Host "❌ Error reading updated registry: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "3. Testing if SAPI now recognizes our voice..." -ForegroundColor Cyan

# Clear logs
Clear-Content "C:\OpenSpeech\sapi_debug.log" -ErrorAction SilentlyContinue

try {
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
        Write-Host "   ✅ Amy voice still found in enumeration" -ForegroundColor Green
        
        $voice.Voice = $amy
        Write-Host "   ✅ Amy voice selected successfully" -ForegroundColor Green
        
        Write-Host "   Testing speech synthesis..." -ForegroundColor Yellow
        try {
            $voice.Speak("Testing with updated registry entries")
            Write-Host "   ✅ Speech synthesis succeeded!" -ForegroundColor Green
        } catch {
            Write-Host "   ❌ Speech synthesis still fails: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        # Check if methods were called
        if(Test-Path "C:\OpenSpeech\sapi_debug.log") {
            $logContent = Get-Content "C:\OpenSpeech\sapi_debug.log"
            $methodsCalled = $logContent | Where-Object { $_ -like "*METHOD CALLED*" }
            if($methodsCalled) {
                Write-Host "   ✅ METHODS WERE CALLED!" -ForegroundColor Green
                $methodsCalled | ForEach-Object { Write-Host "     $_" -ForegroundColor White }
            } else {
                Write-Host "   ❌ Methods still not called" -ForegroundColor Red
            }
        }
        
    } else {
        Write-Host "   ❌ Amy voice not found after registry update" -ForegroundColor Red
    }
    
} catch {
    Write-Host "❌ Error testing SAPI: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "=== RESULT ===" -ForegroundColor Cyan
Write-Host "If methods are now being called:" -ForegroundColor Yellow
Write-Host "  ✅ The missing registry entries were the issue!" -ForegroundColor Green
Write-Host "If methods are still not called:" -ForegroundColor Yellow
Write-Host "  ❌ The issue is deeper - likely managed COM vs native COM" -ForegroundColor Red
