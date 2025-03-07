# Product Requirements Document: Azure TTS Integration

## Overview
This document outlines the requirements and implementation details for adding Azure Text-to-Speech (TTS) support to the OpenSpeech TTS SAPI Installer. The integration will allow users to install and use Azure TTS voices alongside the existing Sherpa ONNX voices.

## Key Components

### 1. Installer/Program.cs Updates
- Add command-line options for Azure TTS installation
- Implement Azure subscription key and region parameters
- Create an interactive menu for Azure voice selection
- Support both Sherpa ONNX and Azure TTS voice types

### 2. Azure Voice Management
- Fetch available Azure voices based on region
- Store Azure voice metadata
- Handle Azure voice registration

### 3. SAPI Registration
- Use Sapi5RegistrarExtended class instead of Sapi5Registrar
- Register both types of voices with appropriate CLSID values
- Store Azure-specific parameters in registry

### 4. Testing
- Update SimpleTest project to detect voice type
- Add appropriate test code for each voice type

## Implementation Details

### Azure Key Management Options
Azure subscription keys can be managed in two ways:
1. **User-provided at runtime**: 
   - Command-line parameters for subscription key and region
   - Interactive prompts for subscription key and region
   - Stored in registry during voice installation

2. **Configuration file option**:
   - Create a config file (azure_config.json) to store default keys
   - Allow override via command-line parameters
   - Implement secure storage of keys

### Command-Line Interface Updates
```
Installer.exe install-azure <voice-name> --key <subscription-key> --region <region>
Installer.exe list-azure-voices --key <subscription-key> --region <region>
```

### Interactive Mode Updates
- Add option to select between Sherpa ONNX and Azure TTS
- For Azure TTS:
  - Prompt for subscription key and region
  - Fetch and display available voices
  - Allow voice selection with optional style/role parameters

## Implementation Tasks

### 1. Update Program.cs
- [ ] Add Azure TTS command-line options
- [ ] Implement Azure voice listing functionality
- [ ] Create Azure voice selection menu
- [ ] Update installation logic to handle both voice types
- [ ] Add Azure key and region management

### 2. Use Sapi5RegistrarExtended
- [ ] Replace Sapi5Registrar with Sapi5RegistrarExtended
- [ ] Update voice registration calls to use type-specific methods
- [ ] Ensure proper CLSID usage for each voice type

### 3. Update SimpleTest
- [ ] Add voice type detection
- [ ] Implement appropriate testing for each voice type
- [ ] Add Azure-specific parameter testing

### 4. Documentation
- [ ] Update README with Azure TTS instructions
- [ ] Document Azure key management options
- [ ] Add troubleshooting section for Azure voices

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
- Subscription keys should be stored securely
- Consider encryption for stored keys
- Provide option to use environment variables for keys
- Implement key validation before installation

## Testing Plan
1. Test installation of Sherpa ONNX voices
2. Test installation of Azure voices
3. Test voice switching in applications
4. Test uninstallation of both voice types
5. Test with various Azure regions and voice styles
