# Register mms_hat voice to SAPI5
# Run as Administrator

$ErrorActionPreference = "Stop"

$voiceTokenPath = "HKLM:\SOFTWARE\Microsoft\Speech\Voices\Tokens\mms_hat"
$clsid = "{A1B2C3D4-E5F6-4A5B-8C9D-1E2F3A4B5C6D}"

Write-Host "Registering mms_hat SAPI5 voice..." -ForegroundColor Cyan

# Remove existing if present
if (Test-Path $voiceTokenPath) {
    Write-Host "Removing existing voice token..." -ForegroundColor Yellow
    Remove-Item -Path $voiceTokenPath -Recurse -Force
}

# Create voice token
Write-Host "Creating voice token..." -ForegroundColor Gray
New-Item -Path $voiceTokenPath -Force | Out-Null
Set-ItemProperty -Path $voiceTokenPath -Name "(default)" -Value "Mms_hat (MMS)"
Set-ItemProperty -Path $voiceTokenPath -Name "CLSID" -Value $clsid
Set-ItemProperty -Path $voiceTokenPath -Name "411" -Value "Mms_hat (MMS)"

# Create Attributes subkey
$attributesPath = "$voiceTokenPath\Attributes"
New-Item -Path $attributesPath -Force | Out-Null
Set-ItemProperty -Path $attributesPath -Name "Language" -Value "409"
Set-ItemProperty -Path $attributesPath -Name "Gender" -Value "Female"
Set-ItemProperty -Path $attributesPath -Name "Age" -Value "Adult"
Set-ItemProperty -Path $attributesPath -Name "Name" -Value "Mms_hat (MMS)"
Set-ItemProperty -Path $attributesPath -Name "Vendor" -Value "OpenAssistive"

Write-Host ""
Write-Host "Voice registered successfully!" -ForegroundColor Green
Write-Host "You can now test it with: .\test-voice.ps1" -ForegroundColor White
