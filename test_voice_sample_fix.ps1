Add-Type -AssemblyName System.Speech
$voice = New-Object System.Speech.Synthesis.SpeechSynthesizer
$voice.SelectVoice("TestSherpaVoice")
Write-Host "Testing voice with sample rate fix (16000Hz output)..."
$voice.Speak("The quick brown fox jumps over the lazy dog.")
Write-Host "Test complete!"
