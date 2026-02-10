# Download MMS Hausa (mms_hat) model files from HuggingFace
# Run as normal user (no admin required)

$ErrorActionPreference = "Stop"

$ModelDir = "$env:LOCALAPPDATA\OpenSpeech\models\mms_hat"
$BaseUrl = "https://huggingface.co/willwade/mms-tts-multilingual-models-onnx/resolve/main/hat"

Write-Host "=== Downloading MMS Hausa Model ===" -ForegroundColor Cyan
Write-Host ""

# Create directory
New-Item -ItemType Directory -Path $ModelDir -Force | Out-Null
Write-Host "Created directory: $ModelDir" -ForegroundColor Green

# Download files using a helper function
function Download-File {
    param(
        [string]$Url,
        [string]$OutputPath
    )

    Write-Host "Downloading: $($Url | Split-Path -Leaf)" -ForegroundColor Gray
    try {
        # Use TLS 1.2
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($Url, $OutputPath)
        $size = (Get-Item $OutputPath).Length
        Write-Host "  ✓ Downloaded ($size bytes)" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "  ✗ Failed: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Download model files
$files = @(
    @{ Url = "$BaseUrl/tokens.txt"; Output = "$ModelDir\tokens.txt" },
    @{ Url = "$BaseUrl/model.onnx"; Output = "$ModelDir\model.onnx" },
    @{ Url = "$BaseUrl/phontab"; Output = "$ModelDir\phontab" }  # Required for MMS
)

$success = $true
foreach ($file in $files) {
    if (-not (Download-File -Url $file.Url -OutputPath $file.Output)) {
        # int8 model is optional
        if ($file.Url -notlike "*int8*") {
            $success = $false
        }
    }
}

Write-Host ""
if ($success) {
    Write-Host "✓ Model download complete!" -ForegroundColor Green
    Write-Host "  Location: $ModelDir" -ForegroundColor White
    Write-Host ""
    Write-Host "Now run: .\test-voice.ps1" -ForegroundColor Cyan
} else {
    Write-Host "✗ Some files failed to download" -ForegroundColor Red
}
