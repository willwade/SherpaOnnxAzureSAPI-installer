# Azure TTS Bridge - Direct HTTP API Integration
# This script provides a working Azure TTS solution that bypasses SAPI COM issues

param(
    [string]$Text = "Hello from Azure TTS!",
    [string]$Voice = "en-GB-LibbyNeural",
    [string]$OutputFile = "azure_output.wav",
    [string]$AzureKey = "b14f8945b0f1459f9964bdd72c42c2cc",
    [string]$AzureRegion = "uksouth",
    [switch]$Play
)

Write-Host "🌐 AZURE TTS BRIDGE" -ForegroundColor Cyan
Write-Host "===================" -ForegroundColor Cyan
Write-Host ""

Write-Host "Configuration:" -ForegroundColor Green
Write-Host "   Text: $Text" -ForegroundColor Gray
Write-Host "   Voice: $Voice" -ForegroundColor Gray
Write-Host "   Output: $OutputFile" -ForegroundColor Gray
Write-Host "   Region: $AzureRegion" -ForegroundColor Gray
Write-Host ""

# Function to synthesize speech using Azure TTS
function Invoke-AzureTTS {
    param(
        [string]$Text,
        [string]$Voice,
        [string]$SubscriptionKey,
        [string]$Region
    )
    
    try {
        $headers = @{
            "Ocp-Apim-Subscription-Key" = $SubscriptionKey
            "Content-Type" = "application/ssml+xml"
            "X-Microsoft-OutputFormat" = "riff-24khz-16bit-mono-pcm"
        }
        
        $endpoint = "https://$Region.tts.speech.microsoft.com/cognitiveservices/v1"
        
        # Create SSML
        $ssml = @"
<speak version='1.0' xml:lang='en-GB'>
    <voice name='$Voice'>$Text</voice>
</speak>
"@
        
        Write-Host "🌐 Calling Azure TTS API..." -ForegroundColor Yellow
        $response = Invoke-WebRequest -Uri $endpoint -Method POST -Headers $headers -Body $ssml -TimeoutSec 15
        
        if ($response.StatusCode -eq 200) {
            Write-Host "✅ Azure TTS synthesis successful - $($response.Content.Length) bytes" -ForegroundColor Green
            return $response.Content
        } else {
            Write-Host "❌ Azure TTS failed with status: $($response.StatusCode)" -ForegroundColor Red
            return $null
        }
        
    } catch {
        Write-Host "❌ Azure TTS error: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

# Function to play audio using Windows Media Player
function Play-Audio {
    param([string]$AudioFile)
    
    try {
        $mediaPlayer = New-Object -ComObject "WMPlayer.OCX"
        $mediaPlayer.URL = (Resolve-Path $AudioFile).Path
        $mediaPlayer.controls.play()
        
        Write-Host "🔊 Playing audio..." -ForegroundColor Cyan
        
        # Wait for playback to start
        Start-Sleep -Seconds 1
        
        # Wait for playback to finish
        while ($mediaPlayer.playState -eq 3) { # 3 = playing
            Start-Sleep -Milliseconds 100
        }
        
        Write-Host "✅ Playback completed" -ForegroundColor Green
        
    } catch {
        Write-Host "⚠️ Could not play audio: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "💡 Try opening the file manually: $AudioFile" -ForegroundColor Gray
    }
}

# Main synthesis
Write-Host "🗣️ Synthesizing speech..." -ForegroundColor Yellow

$audioData = Invoke-AzureTTS -Text $Text -Voice $Voice -SubscriptionKey $AzureKey -Region $AzureRegion

if ($audioData) {
    # Save audio file
    try {
        [System.IO.File]::WriteAllBytes($OutputFile, $audioData)
        Write-Host "💾 Audio saved: $OutputFile" -ForegroundColor Green

        $fileInfo = Get-Item $OutputFile
        Write-Host "📊 File size: $([math]::Round($fileInfo.Length / 1KB, 1)) KB" -ForegroundColor Gray
        Write-Host "⏱️ Duration: ~$([math]::Round($fileInfo.Length / 48000, 1)) seconds" -ForegroundColor Gray

        # Play audio if requested
        if ($Play) {
            Play-Audio -AudioFile $OutputFile
        }

        Write-Host ""
        Write-Host "🎉 Azure TTS synthesis completed successfully!" -ForegroundColor Green

    } catch {
        Write-Host "❌ Failed to save audio: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "❌ Azure TTS synthesis failed" -ForegroundColor Red
}

Write-Host ""
Write-Host "Bridge completed!" -ForegroundColor Green
