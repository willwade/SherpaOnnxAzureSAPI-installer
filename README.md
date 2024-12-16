# TTS Installer

This project contains a complete installer for TTS voices, integrating with Sherpa ONNX, downloading models dynamically, and registering them with SAPI5.

## Features
- Downloads the latest `merged_models.json` dynamically from the web.
- Handles MMS and non-MMS models differently during installation.
- Registers the TTS voices with correct language codes (LCIDs) for SAPI5.
- Uses Sherpa ONNX for TTS audio generation.


### Example Usage

- **Install Models**:
   ```bash
   TTSInstaller.exe
   ```

- **Uninstall Models**:
   ```bash
   TTSInstaller.exe uninstall
   ```

### Uninstallation
To uninstall all voices and unregister the DLL, run:
```bash
TTSInstaller.exe uninstall
```
This removes voices and unregisters the DLL if no models are left installed.



### Developer Details


### Prerequisites
1. Install [.NET 6.0](https://dotnet.microsoft.com/en-us/download) or later.
2. Ensure the `dotnet` CLI is available in your system PATH.

---

1. **Clone the Repository**:
   ```bash
   git clone https://your-repo-url
   cd TTSInstaller
   ```

2. **Restore and Build**:
   Restore dependencies and build the installer:
   ```bash
   dotnet restore
   dotnet build
   ```

3. **Publish the Executable**:
   Create a single EXE for distribution:
   ```bash
   dotnet publish -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true
   ```
   The EXE will be in:
   ```
   bin\Release\net6.0-windows\win-x64\publish\TTSInstaller.exe
   ```

4. **Run the Installer**:
   Execute the installer to download and register voices:
   ```bash
   TTSInstaller.exe
   ```

---


### License
This project is licensed under Apache 2.0.

