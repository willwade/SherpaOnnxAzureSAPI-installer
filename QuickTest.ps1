# Quick Test - No Admin Check
Write-Host "=== SherpaOnnx Installer Test ===" -ForegroundColor Cyan

# Test 1: Check if installer exists
$installerPath = "dist\SherpaOnnxSAPIInstaller.exe"
if (Test-Path $installerPath) {
    Write-Host "SUCCESS: Found installer at $installerPath" -ForegroundColor Green
} else {
    Write-Host "ERROR: Installer not found" -ForegroundColor Red
    exit 1
}

# Test 2: Check if models file exists
$modelsPath = "dist\merged_models.json"
if (Test-Path $modelsPath) {
    Write-Host "SUCCESS: Found models database" -ForegroundColor Green
} else {
    Write-Host "ERROR: Models database not found" -ForegroundColor Red
    exit 1
}

# Test 3: Load and show some voices
try {
    Write-Host "Loading voice models..." -ForegroundColor Yellow
    $models = Get-Content $modelsPath -Raw | ConvertFrom-Json
    
    Write-Host "SUCCESS: Loaded models database" -ForegroundColor Green
    Write-Host "Total voices available: $($models.PSObject.Properties.Name.Count)" -ForegroundColor Gray
    
    # Find English voices
    $englishVoices = @()
    $counter = 1
    
    foreach ($voiceId in $models.PSObject.Properties.Name) {
        if ($voiceId -like "piper-en-*") {
            $voice = $models.$voiceId
            if ($voice.language -and $voice.language[0].lang_code -eq "en") {
                $size = [math]::Round($voice.filesize_mb, 1)
                if ($size -lt 100) {  # Show smaller voices
                    $englishVoices += [PSCustomObject]@{
                        Number = $counter++
                        ID = $voiceId
                        Name = $voice.name
                        Quality = $voice.quality
                        Size = "$size MB"
                    }
                    if ($counter -gt 5) { break }  # Just show first 5
                }
            }
        }
    }
    
    Write-Host ""
    Write-Host "Sample English Voices (first 5 small ones):" -ForegroundColor Yellow
    $englishVoices | Format-Table -AutoSize
    
    Write-Host ""
    Write-Host "NEXT STEPS:" -ForegroundColor Green
    Write-Host "1. Run as administrator to test actual installation" -ForegroundColor Yellow
    Write-Host "2. Use: .\TestInstaller.ps1 -Voice 1" -ForegroundColor Yellow
    Write-Host "3. Or use: .\TestInstaller.ps1 -Voice amy" -ForegroundColor Yellow
    
} catch {
    Write-Host "ERROR: Could not load models: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "Basic test completed!" -ForegroundColor Green
