# Simple Test Installer - Clean Version
param(
    [string]$Voice = ""
)

# Check admin privileges
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-Administrator)) {
    Write-Host "ERROR: This installer requires administrator privileges." -ForegroundColor Red
    exit 1
}

# Configuration
$DistPath = Join-Path $PSScriptRoot "dist"
$InstallerPath = Join-Path $DistPath "SherpaOnnxSAPIInstaller.exe"
$ModelsPath = Join-Path $DistPath "merged_models.json"

Write-Host "Enhanced SherpaOnnx SAPI Installer - Test Version" -ForegroundColor Cyan
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host ""

# Check if installer exists
if (-not (Test-Path $InstallerPath)) {
    Write-Host "ERROR: SherpaOnnxSAPIInstaller.exe not found" -ForegroundColor Red
    Write-Host "Expected at: $InstallerPath" -ForegroundColor Yellow
    exit 1
}

Write-Host "SUCCESS: Found existing installer at: $InstallerPath" -ForegroundColor Green

# Check if models file exists
if (Test-Path $ModelsPath) {
    Write-Host "SUCCESS: Found models database at: $ModelsPath" -ForegroundColor Green
    
    # Load and show some voices
    try {
        $models = Get-Content $ModelsPath -Raw | ConvertFrom-Json
        $englishVoices = @()
        $counter = 1
        
        foreach ($voiceId in $models.PSObject.Properties.Name | Sort-Object) {
            $voice = $models.$voiceId
            if ($voiceId -like "piper-en-*" -and $voice.language -and $voice.language[0].lang_code -eq "en") {
                $size = [math]::Round($voice.filesize_mb, 1)
                if ($size -lt 100) {  # Show smaller voices first
                    $englishVoices += [PSCustomObject]@{
                        Number = $counter++
                        ID = $voiceId
                        Name = $voice.name
                        Quality = $voice.quality
                        Size = "$size MB"
                    }
                    if ($counter -gt 10) { break }  # Limit to first 10
                }
            }
        }
        
        Write-Host ""
        Write-Host "Available English Voices (first 10):" -ForegroundColor Yellow
        $englishVoices | Format-Table Number, Name, Quality, Size, ID -AutoSize
        
        if ($Voice) {
            Write-Host "You requested voice: $Voice" -ForegroundColor Cyan
            
            # Try to find the voice
            $selectedVoice = $null
            
            # Try as number
            if ($Voice -match '^\d+$') {
                $number = [int]$Voice
                $voiceObj = $englishVoices | Where-Object { $_.Number -eq $number }
                if ($voiceObj) { 
                    $selectedVoice = $voiceObj.ID 
                    Write-Host "Found by number: $selectedVoice" -ForegroundColor Green
                }
            }
            
            # Try as exact ID
            if (-not $selectedVoice -and $models.PSObject.Properties.Name -contains $Voice) {
                $selectedVoice = $Voice
                Write-Host "Found by exact ID: $selectedVoice" -ForegroundColor Green
            }
            
            # Try partial matching
            if (-not $selectedVoice) {
                $partialMatches = $englishVoices | Where-Object { $_.Name -like "*$Voice*" -or $_.ID -like "*$Voice*" }
                if ($partialMatches.Count -eq 1) {
                    $selectedVoice = $partialMatches[0].ID
                    Write-Host "Found by partial match: $selectedVoice" -ForegroundColor Green
                } elseif ($partialMatches.Count -gt 1) {
                    Write-Host "Multiple matches found for '$Voice':" -ForegroundColor Yellow
                    $partialMatches | Format-Table Number, Name, Quality, Size, ID -AutoSize
                }
            }
            
            if ($selectedVoice) {
                Write-Host ""
                Write-Host "TESTING: Would install voice: $selectedVoice" -ForegroundColor Cyan
                Write-Host "Command would be: $InstallerPath install $selectedVoice" -ForegroundColor Gray
                
                # Ask if user wants to proceed
                $proceed = Read-Host "Do you want to actually install this voice? (y/n)"
                if ($proceed -eq "y" -or $proceed -eq "yes") {
                    Write-Host "Installing voice..." -ForegroundColor Yellow
                    $result = Start-Process -FilePath $InstallerPath -ArgumentList "install", $selectedVoice -Wait -PassThru -NoNewWindow
                    
                    if ($result.ExitCode -eq 0) {
                        Write-Host "SUCCESS: Voice installed!" -ForegroundColor Green
                        Write-Host "NOTE: The engine config bug fix would be applied here in the full version" -ForegroundColor Yellow
                    } else {
                        Write-Host "ERROR: Installation failed with exit code: $($result.ExitCode)" -ForegroundColor Red
                    }
                }
            } else {
                Write-Host "ERROR: Could not find voice '$Voice'" -ForegroundColor Red
            }
        } else {
            Write-Host ""
            Write-Host "USAGE EXAMPLES:" -ForegroundColor Green
            Write-Host "  .\TestInstaller.ps1 -Voice 1          # Install first voice"
            Write-Host "  .\TestInstaller.ps1 -Voice amy        # Find Amy voices"
            Write-Host "  .\TestInstaller.ps1 -Voice piper-en-amy-low  # Exact ID"
        }
        
    } catch {
        Write-Host "ERROR: Could not parse models file: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "WARNING: Models database not found at: $ModelsPath" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Test completed!" -ForegroundColor Green
