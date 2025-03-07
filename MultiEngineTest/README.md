# MultiEngineTest

A comprehensive test application for the OpenSpeech TTS SAPI Installer with multi-engine support.

## Overview

This project provides a testing framework for the OpenSpeech TTS SAPI Installer, allowing you to test multiple text-to-speech engines in different modes. The supported engines include:

- Sherpa ONNX
- Azure TTS
- ElevenLabs
- PlayHT

## Testing Modes

The application supports three modes of testing:

1. **Direct Testing**: Tests each engine directly using their respective APIs.
2. **SAPI Testing**: Tests the integration with the Windows Speech API (SAPI).
3. **Plugin System Testing**: Tests the plugin system for loading and using TTS engines.

## Configuration

The application uses a configuration file named `test-config.json` to specify test parameters for each engine. A sample configuration file (`test-config-example.json`) is provided as a template.

The configuration file has the following structure:

```json
{
  "engines": {
    "SherpaOnnx": {
      "enabled": true,
      "parameters": {
        "modelsDirectory": "C:\\OpenSpeech\\models"
      },
      "testVoiceId": "vits-ljspeech",
      "testText": "This is a test of the Sherpa ONNX text-to-speech engine."
    },
    "AzureTTS": {
      "enabled": true,
      "parameters": {
        "subscriptionKey": "your-subscription-key-here",
        "region": "eastus"
      },
      "testVoiceId": "en-US-JennyNeural",
      "testText": "This is a test of the Azure text-to-speech engine."
    },
    "ElevenLabs": {
      "enabled": true,
      "parameters": {
        "apiKey": "your-api-key-here",
        "modelId": "eleven_monolingual_v1"
      },
      "testVoiceId": "21m00Tcm4TlvDq8ikWAM",
      "testText": "This is a test of the ElevenLabs text-to-speech engine."
    },
    "PlayHT": {
      "enabled": true,
      "parameters": {
        "apiKey": "your-api-key-here",
        "userId": "your-user-id-here",
        "quality": "premium"
      },
      "testVoiceId": "s3://voice-cloning-zero-shot/d9ff78ba-d016-47f6-b0ef-dd630f59414e/female-voice/manifest.json",
      "testText": "This is a test of the PlayHT text-to-speech engine."
    }
  },
  "testOptions": {
    "runDirectTests": true,
    "runSapiTests": true,
    "runPluginTests": true,
    "outputDirectory": "C:\\OpenSpeech\\TestOutput",
    "sapiDllPath": "C:\\Program Files\\OpenAssistive\\OpenSpeech\\OpenSpeechTTS.dll",
    "pluginDirectory": "C:\\Program Files\\OpenAssistive\\OpenSpeech\\plugins"
  }
}
```

### Configuration Options

#### Engine Configuration

Each engine has the following configuration options:

- `enabled`: Whether to test this engine.
- `parameters`: Engine-specific parameters:
  - **SherpaOnnx**: `modelsDirectory` - Directory containing the Sherpa ONNX models.
  - **AzureTTS**: `subscriptionKey` - Azure Cognitive Services subscription key, `region` - Azure region.
  - **ElevenLabs**: `apiKey` - ElevenLabs API key, `modelId` - Model ID to use.
  - **PlayHT**: `apiKey` - PlayHT API key, `userId` - PlayHT user ID, `quality` - Audio quality.
- `testVoiceId`: The voice ID to use for testing.
- `testText`: The text to synthesize during testing.

#### Test Options

- `runDirectTests`: Whether to run direct API tests.
- `runSapiTests`: Whether to run SAPI integration tests.
- `runPluginTests`: Whether to run plugin system tests.
- `outputDirectory`: Directory to save generated audio files.
- `sapiDllPath`: Path to the OpenSpeechTTS.dll file.
- `pluginDirectory`: Directory containing TTS engine plugins.

## Usage

1. Copy `test-config-example.json` to `test-config.json`.
2. Edit `test-config.json` to add your API keys and adjust settings as needed.
3. Run the application using the command:

```
dotnet run
```

## Test Output

The application will generate audio files in the specified output directory with the following naming conventions:

- Direct tests: `{engine-name}-direct-{voice-id}.wav` or `.mp3`
- SAPI tests: `sapi-{voice-name}.wav`
- Plugin tests: `plugin-{engine-name}-{voice-id}.wav`

## Requirements

- .NET 6.0 or later
- Windows OS (for SAPI tests)
- NAudio library
- Newtonsoft.Json library

## Adding New Engines

To add tests for a new engine:

1. Update the configuration file to include the new engine.
2. Implement the engine in the `Models.cs` file by creating a new class that implements the `ITtsEngine` interface.
3. Register the engine in the `RunPluginSystemTests` method in `Program.cs`.

## Troubleshooting

### API Keys

If you encounter authentication errors, ensure that your API keys are correctly set in the configuration file.

### SAPI Tests

If SAPI tests fail, ensure that:
- The OpenSpeechTTS.dll is correctly installed and registered.
- The path to the DLL in the configuration file is correct.
- You have administrative privileges to access the SAPI.

### Plugin System Tests

If plugin system tests fail, ensure that:
- The plugin directory exists and is accessible.
- The plugins are correctly implemented and compatible with the plugin system.
- The engine configurations are valid. 