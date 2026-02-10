$modelDir = "C:\Users\WillWade\AppData\Local\OpenSpeech\models\mms_hat"
New-Item -ItemType Directory -Path $modelDir -Force | Out-Null
Write-Host "Created directory: $modelDir"

$baseUrl = "https://huggingface.co/willwade/mms-tts-multilingual-models-onnx/resolve/main/hat"

Write-Host "Downloading tokens.txt..."
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $wc = New-Object System.Net.WebClient
    $wc.DownloadFile("$baseUrl/tokens.txt", "$modelDir\tokens.txt")
    Write-Host "  Downloaded tokens.txt"
} catch {
    Write-Host "  Failed: $($_.Exception.Message)"
}

Write-Host "Downloading model.onnx..."
try {
    $wc = New-Object System.Net.WebClient
    $wc.DownloadFile("$baseUrl/model.onnx", "$modelDir\model.onnx")
    $size = (Get-Item "$modelDir\model.onnx").Length
    Write-Host "  Downloaded model.onnx ($size bytes)"
} catch {
    Write-Host "  Failed: $($_.Exception.Message)"
}

Write-Host ""
Write-Host "Checking files..."
Get-ChildItem $modelDir
