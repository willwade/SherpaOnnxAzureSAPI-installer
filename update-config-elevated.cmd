@echo off
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$config = Get-Content 'C:\Program Files\OpenAssistive\OpenSpeech\engines_config.json' | ConvertFrom-Json; ^
   $config.engines | Add-Member -NotePropertyName 'mms_hat' -NotePropertyValue $config.engines.'sherpa-mms_hat' -Force; ^
   $config | ConvertTo-Json -Depth 10 | Set-Content 'C:\Program Files\OpenAssistive\OpenSpeech\engines_config.json'; ^
   Write-Host 'Done'"
pause
