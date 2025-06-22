# Update voice config with listvoices parameter
Copy-Item "temp_voice_config.json" "C:\Program Files\OpenAssistive\OpenSpeech\voice_configs\English-SherpaOnnx-Jenny.json" -Force
Write-Host "Voice config updated successfully"
