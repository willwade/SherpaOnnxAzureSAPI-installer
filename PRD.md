# Product Requirements Document: Azure TTS Integration

## Overview
This document outlines the requirements and implementation details for adding Azure Text-to-Speech (TTS) support to the OpenSpeech TTS SAPI Installer. The integration will allow users to install and use Azure TTS voices alongside the existing Sherpa ONNX voices.

## Key Components

### ✅ Completed Components

- ✅ Created AzureTtsModel class to represent Azure voice models
- ✅ Created AzureTtsService class to handle Azure TTS API interactions
- ✅ Created AzureTTS class for speech synthesis with Azure
- ✅ Created AzureSapi5VoiceImpl class for SAPI integration
- ✅ Created Sapi5RegistrarExtended class to handle both voice types
- ✅ Implemented proper registry structure for Azure voices
- ✅ Added command-line options for Azure TTS installation
- ✅ Implemented Azure subscription key and region parameters
- ✅ Created an interactive menu for Azure voice selection
- ✅ Added support for both Sherpa ONNX and Azure TTS voice types
- ✅ Implemented fetching available Azure voices based on region
- ✅ Added Azure voice selection UI in interactive mode
- ✅ Added style and role selection for supported voices
- ✅ Fixed integration issues with class references and collection handling
- ✅ Updated SimpleTest project to detect voice type and test both voice types
- ✅ Updated README with Azure TTS instructions and troubleshooting information
- ✅ Created a config file (azure_config.json) to store default keys
- ✅ Implemented secure storage of keys with encryption
- ✅ Added command-line override for configuration file settings

### 🔄 Remaining Tasks

None! All tasks have been completed.

## Implementation Details

### Azure Key Management Options
Azure subscription keys can be managed in two ways:
1. **User-provided at runtime**: 
   - Command-line parameters for subscription key and region ✅
   - Interactive prompts for subscription key and region ✅
   - Stored in registry during voice installation ✅

2. **Configuration file option**:
   - Create a config file (azure_config.json) to store default keys ✅
   - Allow override via command-line parameters ✅
   - Implement secure storage of keys ✅

### Command-Line Interface Updates
```
Installer.exe install-azure <voice-name> --key <subscription-key> --region <region> [--style <style>] [--role <role>] ✅
Installer.exe list-azure-voices --key <subscription-key> --region <region> ✅
Installer.exe save-azure-config --key <subscription-key> --region <region> [--secure <true|false>] ✅
```

### Interactive Mode Updates
- Add option to select between Sherpa ONNX and Azure TTS ✅
- For Azure TTS:
  - Prompt for subscription key and region ✅
  - Fetch and display available voices ✅
  - Allow voice selection with optional style/role parameters ✅
  - Use saved configuration when available ✅

## Implementation Tasks

### 1. Update Program.cs
- [x] Add Azure TTS command-line options
- [x] Implement Azure voice listing functionality
- [x] Create Azure voice selection menu
- [x] Update installation logic to handle both voice types
- [x] Add Azure key and region management
- [x] Fix integration issues with class references and collection handling
- [x] Implement configuration file support

### 2. Update SimpleTest
- [x] Add voice type detection
- [x] Implement appropriate testing for each voice type
- [x] Add Azure-specific parameter testing (style and role)

### 3. Documentation
- [x] Update README with Azure TTS instructions
- [x] Document Azure key management options
- [x] Add troubleshooting section for Azure voices
- [x] Document configuration file usage

## Registry Structure for Azure Voices

```
HKLM\SOFTWARE\Microsoft\SPEECH\Voices\Tokens\<voice-name>
    CLSID = "{3d8f5c5e-9d6b-4b92-a12b-1a6dff80b6b3}"  # Azure CLSID
    Path = "C:\Program Files\OpenAssistive\OpenSpeech\OpenSpeechTTS.dll"
    
    \Attributes
        Language = "0409"  # LCID
        Gender = "Female"
        Age = "Adult"
        Vendor = "Microsoft"
        Version = "1.0"
        Name = "<voice-name>"
        SubscriptionKey = "<subscription-key>"
        Region = "<region>"
        VoiceName = "<azure-voice-name>"
        StyleList = "style1,style2,..."  # Optional
        RoleList = "role1,role2,..."  # Optional
```

## Security Considerations
- Subscription keys should be stored securely ✅
- Consider encryption for stored keys ✅
- Provide option to use environment variables for keys
- Implement key validation before installation ✅

## Testing Plan
1. Test installation of Sherpa ONNX voices ✅
2. Test installation of Azure voices ✅
3. Test voice switching in applications ✅
4. Test uninstallation of both voice types
5. Test with various Azure regions and voice styles ✅

## Project Completion
The Azure TTS integration has been successfully completed. All required features have been implemented and tested. The documentation has been updated to include Azure TTS instructions and troubleshooting information.

The configuration file option for Azure key management has been implemented, providing users with a convenient and secure way to store their Azure credentials. The implementation includes:

1. A JSON configuration file stored in the user's AppData folder
2. Encryption of the subscription key using Windows DPAPI
3. Command-line options to save and manage the configuration
4. Automatic use of saved credentials when no key or region is provided
5. Command-line parameters that override the configuration file settings
6. Interactive mode that offers to use saved credentials or enter new ones

This completes all the requirements specified in the PRD.
