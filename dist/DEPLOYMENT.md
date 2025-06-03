# Native COM Wrapper Deployment

## Built Component:
- NativeTTSWrapper.dll (109 KB)

## Manual Deployment Steps:

1. Copy NativeTTSWrapper.dll to target location:
   `
   Copy-Item "NativeTTSWrapper.dll" "C:\Program Files\OpenAssistive\OpenSpeech\" -Force
   `

2. Register the COM object (as Administrator):
   `
   sudo regsvr32 "C:\Program Files\OpenAssistive\OpenSpeech\NativeTTSWrapper.dll"
   `

3. Update voice registration to use native wrapper:
   `
   sudo Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Speech\Voices\Tokens\amy" -Name "CLSID" -Value "{E1C4A8F2-9B3D-4A5E-8F7C-2D1B3E4F5A6B}"
   `

4. Test the native wrapper:
   `
   $voice = New-Object -ComObject SAPI.SpVoice
   $voice.Speak("Hello from native COM wrapper!")
   `

## What This Provides:
- 100% SAPI compatibility for SherpaOnnx voices
- Native C++ performance
- Full ProcessBridge integration
- Works with any SAPI application

## Next Steps:
To build the complete installer with .NET components:
1. Install .NET 6.0 SDK
2. Run: .\BuildCompleteInstaller.ps1

Built on: 2025-06-03 16:13:02
