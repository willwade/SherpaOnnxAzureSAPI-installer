# Configuration-Based SAPI Voices with Pipe Service Integration

This document describes the new configuration-based SAPI voice system that integrates with the AACSpeakHelper pipe service.

## Overview

The new system allows you to create SAPI voices that are essentially configuration files. Each voice corresponds to a specific TTS engine configuration that gets sent to the AACSpeakHelper pipe service for synthesis.

### Key Benefits

- **Flexible Configuration**: Each voice can have different TTS engines, settings, and parameters
- **Easy Management**: Add/remove voices without rebuilding COM components
- **Centralized Processing**: All TTS processing happens through the AACSpeakHelper pipe service
- **Multiple Engine Support**: Support for Azure TTS, SherpaOnnx, Google TTS, and more

## Architecture

```
SAPI Application
       ↓
Windows SAPI System
       ↓
Configuration-Based Voice (Registry Entry)
       ↓
PipeServiceComWrapper (COM Component)
       ↓
AACSpeakHelper Pipe Service
       ↓
TTS Engine (Azure/SherpaOnnx/Google/etc.)
```

## Voice Configuration Format

Each voice is defined by a JSON configuration file in the `voice_configs/` directory:

```json
{
  "name": "British-English-Azure-Libby",
  "displayName": "British English (Azure Libby)",
  "description": "British English voice using Azure TTS Libby Neural voice",
  "language": "English",
  "locale": "en-GB",
  "gender": "Female",
  "age": "Adult",
  "vendor": "Microsoft Azure",
  "ttsConfig": {
    "engine": "azure",
    "voice_id": "en-GB-LibbyNeural",
    "azureTTS": {
      "key": "your-azure-key",
      "location": "uksouth",
      "voice": "en-GB-LibbyNeural",
      "style": "",
      "role": ""
    },
    "TTS": {
      "engine": "azure",
      "voice_id": "en-GB-LibbyNeural",
      "bypass_tts": false
    },
    "translate": {
      "no_translate": true,
      "provider": "",
      "start_lang": "auto",
      "end_lang": "en",
      "replace_pb": false
    }
  }
}
```

## Installation Commands

### Install a Pipe-Based Voice
```bash
Installer.exe install-pipe-voice <config-name>
```

### List Available Configurations
```bash
Installer.exe list-pipe-voices
```

### Remove a Pipe-Based Voice
```bash
Installer.exe remove-pipe-voice <voice-name>
```

### Test Pipe Service Connection
```bash
Installer.exe test-pipe-service
```

## Interactive Mode

Run the installer without arguments to access the interactive menu:

```bash
Installer.exe
```

The menu now includes options for:
- Install Pipe-based voice
- List pipe-based voices  
- Test pipe service connection

## Prerequisites

1. **AACSpeakHelper Server**: The pipe service must be running
   - Download from: https://github.com/AceCentre/AACSpeakHelper
   - Run `AACSpeakHelperServer.py` to start the service

2. **Voice Configurations**: Create JSON configuration files in `voice_configs/` directory

3. **TTS Engine Credentials**: Configure appropriate API keys and settings in voice configurations

## Example Voice Configurations

### Azure TTS Voice
```json
{
  "name": "American-English-Azure-Jenny",
  "displayName": "American English (Azure Jenny)",
  "description": "American English voice using Azure TTS Jenny Neural voice",
  "locale": "en-US",
  "gender": "Female",
  "ttsConfig": {
    "engine": "azure",
    "azureTTS": {
      "key": "your-azure-subscription-key",
      "location": "uksouth",
      "voice": "en-US-JennyNeural"
    }
  }
}
```

### SherpaOnnx Voice
```json
{
  "name": "British-English-SherpaOnnx-Amy",
  "displayName": "British English (SherpaOnnx Amy)",
  "description": "British English voice using SherpaOnnx Amy model",
  "locale": "en-GB",
  "gender": "Female",
  "ttsConfig": {
    "engine": "sherpaonnx",
    "voice_id": "sherpa-piper-amy-normal"
  }
}
```

## Testing

Use the provided test script to verify the installation:

```powershell
.\TestPipeVoices.ps1
```

This script will:
1. Build the installer
2. Test pipe service connection
3. Install all available voice configurations
4. List installed voices
5. Test SAPI voice enumeration
6. Attempt voice synthesis (if pipe service is running)

## Troubleshooting

### Voice Not Appearing in SAPI
- Check if voice is registered: `Installer.exe list-pipe-voices`
- Verify registry entries in `HKLM\SOFTWARE\Microsoft\SPEECH\Voices\Tokens`
- Ensure COM component is registered

### Synthesis Not Working
- Verify AACSpeakHelper pipe service is running
- Test connection: `Installer.exe test-pipe-service`
- Check voice configuration file syntax
- Verify TTS engine credentials (Azure keys, etc.)

### Configuration Errors
- Validate JSON syntax in configuration files
- Check file paths and permissions
- Ensure all required fields are present

## Advanced Usage

### Creating Custom Configurations

1. Copy an existing configuration file from `voice_configs/`
2. Modify the settings for your specific needs
3. Save with a descriptive filename
4. Install using `install-pipe-voice` command

### Multiple Engine Support

You can create voices that use different TTS engines by modifying the `ttsConfig` section:

- **Azure TTS**: Set `azureTTS` configuration
- **Google TTS**: Set `googleTTS` configuration  
- **SherpaOnnx**: Set engine to "sherpaonnx" and appropriate voice_id
- **Translation**: Configure `translate` section for multi-language support

### Batch Installation

Create a script to install multiple voices:

```powershell
$voices = @("British-English-Azure-Libby", "American-English-Azure-Jenny", "British-English-SherpaOnnx-Amy")
foreach ($voice in $voices) {
    .\Installer.exe install-pipe-voice $voice
}
```

## Integration with Applications

Once installed, pipe-based voices appear as standard SAPI voices and can be used in:

- Windows Speech Recognition
- Screen readers (NVDA, JAWS)
- Text-to-speech applications
- Custom applications using System.Speech.Synthesis
- Web browsers with speech synthesis APIs

## Future Enhancements

- GUI configuration tool
- Voice preview functionality
- Automatic voice discovery
- Cloud-based voice configurations
- Performance monitoring and logging
