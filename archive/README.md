# TTS Installer

This project contains a complete installer for TTS voices, integrating with Sherpa ONNX and Azure TTS, downloading models dynamically, and registering them with SAPI5.

## Features
- Downloads the latest `merged_models.json` dynamically from the web
- Handles MMS and non-MMS models differently during installation
- Registers the TTS voices with correct language codes (LCIDs) for SAPI5
- Uses Sherpa ONNX for offline TTS audio generation
- Supports Azure TTS for cloud-based high-quality voices with styles and roles
- Provides interactive voice selection for both Sherpa ONNX and Azure TTS
- Supports configuration file for storing Azure credentials securely

## Prerequisites
1. Install [.NET 6.0](https://dotnet.microsoft.com/en-us/download) or later
2. Install [WiX Toolset v4](https://wixtoolset.org/docs/intro/)
3. Install WiX .NET CLI Tool:
   ```bash
   dotnet tool install --global wix
   ```
4. Ensure the `dotnet` CLI is available in your system PATH
5. For Azure TTS: An Azure subscription with Speech Service resource

## Building the Project

**Note**: Building the WiX installer requires administrative privileges. Run your terminal/PowerShell as Administrator.

1. **Clean and Restore**:
   ```bash
   dotnet clean TTSInstaller.csproj
   dotnet restore TTSInstaller.csproj
   ```

2. **Build the TTS Project**:
   ```bash
   dotnet build TTSInstaller.csproj -c Release
   ```

3. **Sign the Sherpa-ONNX Assembly (if needed)**:
   If you encounter strong name signing errors, you'll need to sign the sherpa-onnx assembly:
   
   a. Generate a strong name key:
   ```bash
   dotnet build KeyGenerator/KeyGenerator.csproj -t:GenerateKeyFile
   ```
   
   b. Sign the assembly:
   ```bash
   dotnet run --project SignAssembly/SignAssembly.csproj
   ```

## Installing Voices

### Interactive Mode
Run the installer without arguments to enter interactive mode:
```bash
dotnet run --project TTSInstaller.csproj
```

This will present a menu with options to:
1. Install Sherpa ONNX voice
2. Install Azure TTS voice
3. Uninstall all voices
4. Exit

### Command Line Installation - Sherpa ONNX
```bash
dotnet run --project TTSInstaller.csproj -- install <model-id>
```
   
Available model IDs include:
- `piper-en-joe-medium` - Male English voice
- `piper-en-amy-medium` - Female English voice

### Command Line Installation - Azure TTS
```bash
dotnet run --project TTSInstaller.csproj -- install-azure <voice-name> --key <subscription-key> --region <region> [--style <style>] [--role <role>]
```

Example:
```bash
dotnet run --project TTSInstaller.csproj -- install-azure en-US-GuyNeural --key YOUR_SUBSCRIPTION_KEY --region eastus
```

### Listing Available Azure Voices
To see all available Azure voices for your subscription:
```bash
dotnet run --project TTSInstaller.csproj -- list-azure-voices --key <subscription-key> --region <region>
```

### Saving Azure Configuration
To save your Azure credentials for future use:
```bash
dotnet run --project TTSInstaller.csproj -- save-azure-config --key <subscription-key> --region <region> [--secure <true|false>]
```

This will save your Azure credentials to a configuration file, which will be used automatically when no key or region is provided in subsequent commands.

### Uninstalling Voices
To uninstall all voices:
```bash
dotnet run --project TTSInstaller.csproj -- uninstall all
```

To uninstall a specific voice:
```bash
dotnet run --project TTSInstaller.csproj -- uninstall <voice-name>
```

**Note**: Installation requires administrative privileges to register COM components.

## Testing the Installation
```bash
dotnet run --project SimpleTest/SimpleTest.csproj
```
   
This will:
- List all available voices (Sherpa ONNX, Azure TTS, and standard voices)
- Test speech synthesis with installed voices
- For Azure voices, test style and role features if available

## Azure TTS Integration

### Azure Key Management
Azure subscription keys are managed in the following ways:

1. **Configuration file**:
   ```bash
   dotnet run --project TTSInstaller.csproj -- save-azure-config --key <subscription-key> --region <region>
   ```
   This saves your Azure credentials to a configuration file at `%APPDATA%\OpenSpeech\azure_config.json`.
   
   By default, the subscription key is encrypted for security. You can disable encryption with:
   ```bash
   dotnet run --project TTSInstaller.csproj -- save-azure-config --key <subscription-key> --region <region> --secure false
   ```

2. **Command-line parameters**:
   ```bash
   dotnet run --project TTSInstaller.csproj -- install-azure <voice-name> --key <subscription-key> --region <region>
   ```
   Command-line parameters override any settings in the configuration file.

3. **Interactive prompts**:
   When using interactive mode, you'll be prompted to enter your subscription key and region.
   If a configuration file exists, you can press Enter to use the saved credentials.

4. **Registry storage**:
   Keys are stored in the registry during voice installation, allowing the voices to work without re-entering the key.

### Azure Voice Features

#### Styles and Roles
Many Azure voices support styles and roles for more expressive speech:

- **Styles**: Different speaking styles like cheerful, sad, excited, etc.
- **Roles**: Different character roles the voice can portray

To use a style or role:
```bash
dotnet run --project TTSInstaller.csproj -- install-azure <voice-name> --key <subscription-key> --region <region> --style <style> --role <role>
```

Example with style:
```bash
dotnet run --project TTSInstaller.csproj -- install-azure en-US-AriaNeural --key YOUR_KEY --region eastus --style cheerful
```

### SSML Requirement for Azure Voices

**Important Note**: Azure TTS voices require SSML (Speech Synthesis Markup Language) to work properly. This means:

1. In applications that support SSML, Azure voices will work correctly.
2. In applications that only use the standard `SelectVoice()` method without SSML, Azure voices may not work.

This is a limitation of how Azure TTS integrates with SAPI, as Azure voices require additional parameters (subscription key, region, etc.) that are only accessible through SSML.

For developers integrating with these voices, we recommend using SSML for all Azure voice interactions:

```csharp
// Example SSML for Azure voice
string ssml = $@"<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xml:lang='en-US'>
    <voice name='Microsoft Server Speech Text to Speech Voice (en-GB, OliverNeural)'>
        This is a test of the Azure voice.
    </voice>
</speak>";

// Use SSML with the speech synthesizer
speechSynthesizer.SpeakSsml(ssml);
```

### Security Considerations
- Azure subscription keys are stored in the Windows Registry and/or configuration file
- Keys in the configuration file are encrypted by default using Windows DPAPI
- Keys are only accessible to administrators and the SYSTEM account
- Consider using environment variables for keys in sensitive environments
- The installer validates keys before installation to ensure they work

## Publishing and Distribution

### Publish the TTS Application
```bash
dotnet publish TTSInstaller.csproj -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true
```

### Build the WiX Installer (requires admin privileges)
```bash
wix build TTSInstaller.wxs -o TTSInstaller.msi -b .
```

## Installation

### For End Users
- Double-click the generated `TTSInstaller.msi`
- Or run from command line:
  ```bash
  msiexec /i TTSInstaller.msi
  ```

### For Development/Testing
Manual COM registration (if needed):
```bash
& "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe" "publish\OpenSpeechTTS.dll" /codebase
```

### Uninstallation
To uninstall:
- Use Windows Control Panel
- Or run from command line:
  ```bash
  msiexec /x TTSInstaller.msi
  ```

## Troubleshooting

### Common Issues and Solutions

#### Strong Name Signing Errors
If you see errors like:
```
Could not load file or assembly 'sherpa-onnx, Version=1.9.12.0, Culture=neutral, PublicKeyToken=null' or one of its dependencies. A strongly-named assembly is required.
```

Follow the signing steps in the "Building the Project" section to sign the sherpa-onnx assembly.

#### Voice Not Appearing in SAPI Applications
1. Verify the COM registration:
   ```bash
   reg query "HKLM\SOFTWARE\Microsoft\SPEECH\Voices\Tokens\<voice-name>" /s
   ```
   
2. Check the registry entries for the voice attributes:
   ```bash
   reg query "HKLM\SOFTWARE\Microsoft\SPEECH\Voices\Tokens\<voice-name>\Attributes" /s
   ```

3. Ensure the COM DLL is properly registered:
   ```bash
   & "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe" "C:\Program Files\OpenAssistive\OpenSpeech\OpenSpeechTTS.dll" /codebase
   ```

4. For Azure voices, remember they require SSML to work properly. Applications that don't use SSML may not be able to use Azure voices.

#### Azure TTS Specific Issues

##### Invalid Subscription Key or Region
If you see "Error: Invalid Azure subscription key or region":
1. Verify your subscription key is correct
2. Ensure the region is valid (e.g., eastus, westus, etc.)
3. Check that your Azure subscription is active
4. Verify the Speech Service resource is properly set up in Azure

##### Azure Voice Not Found
If you see "Error: Azure voice 'X' not found":
1. Use the `list-azure-voices` command to see all available voices
2. Ensure you're using the correct voice name format (e.g., en-US-GuyNeural)
3. Check if the voice is available in your selected region

##### Style or Role Not Available
If you see "Warning: Style/Role 'X' not available for this voice":
1. Not all voices support all styles and roles
2. Use the `list-azure-voices` command to see available styles and roles
3. Check the Azure documentation for voice capabilities

##### Azure Voice Not Working in Applications
If Azure voices don't work in certain applications:
1. Check if the application supports SSML - Azure voices require SSML to work properly
2. Try using the voice with SSML in our SimpleTest application to verify it works
3. For applications that don't support SSML, Azure voices may not be usable

##### Configuration File Issues
If you're having issues with the configuration file:
1. Check if the file exists at `%APPDATA%\OpenSpeech\azure_config.json`
2. Try recreating the configuration file with the `save-azure-config` command
3. If encryption issues occur, try saving without encryption using `--secure false`

##### Network Issues
If you experience timeouts or connection errors:
1. Check your internet connection
2. Verify your firewall isn't blocking connections to Azure
3. Try a different Azure region that might be closer to your location

#### Debugging Logs
Check the log files at:
- `C:\OpenSpeech\sherpa_debug.log` - General debug information
- `C:\OpenSpeech\sherpa_error.log` - Error details
- `C:\OpenSpeech\azure_debug.log` - Azure TTS debug information

## License
This project is licensed under Apache 2.0.
