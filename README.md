# TTS Installer

This project contains a complete installer for TTS voices, integrating with Sherpa ONNX, downloading models dynamically, and registering them with SAPI5.

## Features
- Downloads the latest `merged_models.json` dynamically from the web
- Handles MMS and non-MMS models differently during installation
- Registers the TTS voices with correct language codes (LCIDs) for SAPI5
- Uses Sherpa ONNX for TTS audio generation

## Prerequisites
1. Install [.NET 6.0](https://dotnet.microsoft.com/en-us/download) or later
2. Install [WiX Toolset v4](https://wixtoolset.org/docs/intro/)
3. Install WiX .NET CLI Tool:
   ```bash
   dotnet tool install --global wix
   ```
4. Ensure the `dotnet` CLI is available in your system PATH

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

4. **Install a Voice**:
   ```bash
   dotnet run --project TTSInstaller.csproj -- install <model-id>
   ```
   
   Available model IDs include:
   - `piper-en-joe-medium` - Male English voice
   - `piper-en-amy-medium` - Female English voice
   
   **Note**: Installation requires administrative privileges to register COM components.

5. **Test the Installation**:
   ```bash
   dotnet run --project SimpleTest/SimpleTest.csproj
   ```
   
   This will list all available voices and test speech synthesis with the installed voice.

6. **Publish the TTS Application**:
   ```bash
   dotnet publish TTSInstaller.csproj -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true
   ```

7. **Build the WiX Installer** (requires admin privileges):
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
   reg query "HKLM\SOFTWARE\Microsoft\Speech\Voices\Tokens\<voice-name>" /s
   ```
   
2. Check the registry entries for the voice attributes:
   ```bash
   reg query "HKLM\SOFTWARE\Microsoft\Speech\Voices\Tokens\<voice-name>\Attributes" /s
   ```

3. Ensure the COM DLL is properly registered:
   ```bash
   & "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe" "C:\Program Files\OpenAssistive\OpenSpeech\OpenSpeechTTS.dll" /codebase
   ```

#### Debugging Logs
Check the log files at:
- `C:\OpenSpeech\sherpa_debug.log` - General debug information
- `C:\OpenSpeech\sherpa_error.log` - Error details

## License
This project is licensed under Apache 2.0.
