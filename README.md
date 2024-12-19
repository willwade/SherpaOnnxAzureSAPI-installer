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

3. **Publish the TTS Application**:
   ```bash
   dotnet publish TTSInstaller.csproj -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true
   ```

4. **Build the WiX Installer** (requires admin privileges):
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
- If you encounter build errors, ensure all prerequisites are installed
- Check the Windows Event Viewer for installation errors
- Review the install.log file in the project directory for detailed logs
- For permission errors during build, ensure you're running as Administrator


## whats not working?

- voice isnt showing up in sapi applications eg balabolka
- basic test synth not working. should do
- we have a error about strong tyoes in nuget sherpa onnx. ive dine wacky workarounds but not sure its needed



## License
This project is licensed under Apache 2.0.
