# SherpaOnnx SAPI Installer

This package contains the complete SherpaOnnx SAPI installer with Azure TTS support.

## Contents:
- SherpaOnnxSAPIInstaller.exe - Main installer (single executable)
- OpenSpeechTTS.dll - Managed COM objects for Azure TTS
- NativeTTSWrapper.dll - Native COM wrapper for SherpaOnnx (100% SAPI compatibility)
- SherpaWorker.exe - ProcessBridge worker for SherpaOnnx
- Dependencies - Required DLLs for SherpaOnnx functionality

## Usage:

### Install SherpaOnnx voice:
`
sudo .\SherpaOnnxSAPIInstaller.exe install amy
`

### Install Azure TTS voice:
`
sudo .\SherpaOnnxSAPIInstaller.exe install-azure en-US-JennyNeural --key YOUR_KEY --region eastus
`

### Interactive mode:
`
sudo .\SherpaOnnxSAPIInstaller.exe
`

### Uninstall all voices:
`
sudo .\SherpaOnnxSAPIInstaller.exe uninstall all
`

## Requirements:
- Windows 10/11
- Administrator privileges
- .NET Framework 4.7.2+ (for Azure TTS)
- .NET 6.0 Runtime (included in installer)

Built on: 2025-06-03 05:37:04
