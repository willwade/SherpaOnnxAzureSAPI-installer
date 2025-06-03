# Investigate Microsoft TTS engine to understand interface requirements
Write-Host "Investigating Microsoft TTS Engine Interface" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Green

$voice = New-Object -ComObject SAPI.SpVoice
$voices = $voice.GetVoices()

Write-Host "Available voices:" -ForegroundColor Cyan
for($i = 0; $i -lt $voices.Count; $i++) {
    $voiceItem = $voices.Item($i)
    $name = $voiceItem.GetDescription()
    $clsid = $voiceItem.GetAttribute("CLSID")
    Write-Host "  $i`: $name (CLSID: $clsid)" -ForegroundColor White
}

# Get Microsoft voice (not Amy)
$msVoice = $null
for($i = 0; $i -lt $voices.Count; $i++) {
    $voiceItem = $voices.Item($i)
    if($voiceItem.GetDescription() -like "*Microsoft*") {
        $msVoice = $voiceItem
        break
    }
}

if($msVoice) {
    $msClsid = $msVoice.GetAttribute("CLSID")
    Write-Host ""
    Write-Host "Microsoft Voice CLSID: $msClsid" -ForegroundColor Yellow
    
    # Check registry for this CLSID
    Write-Host "Checking Microsoft TTS registry entries..." -ForegroundColor Cyan
    
    try {
        $clsidPath = "HKLM:\SOFTWARE\Classes\CLSID\$msClsid"
        if(Test-Path $clsidPath) {
            Write-Host "Microsoft TTS CLSID registry:" -ForegroundColor Green
            Get-ItemProperty $clsidPath | Format-List
            
            $inprocPath = "$clsidPath\InprocServer32"
            if(Test-Path $inprocPath) {
                Write-Host "Microsoft TTS InprocServer32:" -ForegroundColor Green
                Get-ItemProperty $inprocPath | Format-List
            }
        }
    } catch {
        Write-Host "Error reading Microsoft TTS registry: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "Comparing with our Amy voice..." -ForegroundColor Cyan

# Get Amy voice
$amyVoice = $null
for($i = 0; $i -lt $voices.Count; $i++) {
    $voiceItem = $voices.Item($i)
    if($voiceItem.GetDescription() -eq "amy") {
        $amyVoice = $voiceItem
        break
    }
}

if($amyVoice) {
    $amyClsid = $amyVoice.GetAttribute("CLSID")
    Write-Host "Amy Voice CLSID: $amyClsid" -ForegroundColor Yellow
    
    try {
        $clsidPath = "HKLM:\SOFTWARE\Classes\CLSID\$amyClsid"
        if(Test-Path $clsidPath) {
            Write-Host "Amy TTS CLSID registry:" -ForegroundColor Green
            Get-ItemProperty $clsidPath | Format-List
            
            $inprocPath = "$clsidPath\InprocServer32"
            if(Test-Path $inprocPath) {
                Write-Host "Amy TTS InprocServer32:" -ForegroundColor Green
                Get-ItemProperty $inprocPath | Format-List
            }
        }
    } catch {
        Write-Host "Error reading Amy TTS registry: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "Key Differences to Investigate:" -ForegroundColor Yellow
Write-Host "1. InprocServer32 path (native DLL vs managed DLL)" -ForegroundColor White
Write-Host "2. Threading model" -ForegroundColor White
Write-Host "3. Missing registry entries" -ForegroundColor White
Write-Host "4. Interface registration differences" -ForegroundColor White
