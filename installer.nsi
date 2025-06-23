; OpenSpeechSAPI - Universal SAPI Bridge Installer - NSIS Script
; ============================================================================

!define PRODUCT_NAME "OpenSpeechSAPI"
!define PRODUCT_VERSION "1.0.0"
!define PRODUCT_PUBLISHER "AceCentre"
!define PRODUCT_WEB_SITE "https://github.com/willwade/OpenSpeechSAPI"
!define PRODUCT_DIR_REGKEY "Software\Microsoft\Windows\CurrentVersion\App Paths\sapi_voice_installer.exe"
!define PRODUCT_UNINST_KEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}"
!define PRODUCT_UNINST_ROOT_KEY "HKLM"

; Request administrator privileges
RequestExecutionLevel admin

; Modern UI
!include "MUI2.nsh"

; General
Name "${PRODUCT_NAME} ${PRODUCT_VERSION}"
OutFile "OpenSpeechSAPI-installer.exe"
InstallDir "$PROGRAMFILES64\${PRODUCT_NAME}"
InstallDirRegKey HKLM "${PRODUCT_DIR_REGKEY}" ""
ShowInstDetails show
ShowUnInstDetails show

; Interface Settings
!define MUI_ABORTWARNING
!define MUI_ICON "${NSISDIR}\Contrib\Graphics\Icons\modern-install.ico"
!define MUI_UNICON "${NSISDIR}\Contrib\Graphics\Icons\modern-uninstall.ico"

; Pages
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE "LICENSE"
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_WELCOME
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_UNPAGE_FINISH

; Languages
!insertmacro MUI_LANGUAGE "English"

; Installer sections
Section "Main Application" SEC01
  SetOutPath "$INSTDIR"
  SetOverwrite ifnewer
  
  ; Copy main executables
  File "dist\sapi_voice_installer.exe"
  File "dist\AACSpeakHelperServer.exe"

  ; Copy C++ COM Wrapper and all dependencies
  File "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll"
  File "NativeTTSWrapper\x64\Release\fmt.dll"
  File "NativeTTSWrapper\x64\Release\onnxruntime.dll"
  File "NativeTTSWrapper\x64\Release\onnxruntime_providers_shared.dll"
  File "NativeTTSWrapper\x64\Release\sherpa-onnx-c-api.dll"

  ; Copy voice configurations
  SetOutPath "$INSTDIR\voice_configs"
  File /r "voice_configs\*.*"

  ; Copy configuration files
  SetOutPath "$INSTDIR"
  File /nonfatal "settings.cfg.example"
  
  ; Create shortcuts
  CreateDirectory "$SMPROGRAMS\${PRODUCT_NAME}"
  CreateShortCut "$SMPROGRAMS\${PRODUCT_NAME}\SAPI Voice Installer.lnk" "$INSTDIR\sapi_voice_installer.exe"
  CreateShortCut "$SMPROGRAMS\${PRODUCT_NAME}\AACSpeakHelper Server.lnk" "$INSTDIR\AACSpeakHelperServer.exe"
  CreateShortCut "$DESKTOP\SAPI Voice Installer.lnk" "$INSTDIR\sapi_voice_installer.exe"

  ; Create startup shortcut for AACSpeakHelper Server (auto-start with Windows)
  CreateShortCut "$SMSTARTUP\AACSpeakHelperServer.lnk" "$INSTDIR\AACSpeakHelperServer.exe" "" "$INSTDIR\AACSpeakHelperServer.exe" 0 SW_SHOWMINIMIZED
  
  ; Register COM wrapper using SAPI Voice Installer
  DetailPrint "Registering COM wrapper using SAPI Voice Installer..."
  DetailPrint "Executing: $\"$INSTDIR\sapi_voice_installer.exe$\" register-com"
  ExecWait '"$INSTDIR\sapi_voice_installer.exe" register-com' $0
  DetailPrint "SAPI Voice Installer COM registration completed with exit code: $0"

  ${If} $0 != 0
    DetailPrint "ERROR: COM registration failed with exit code $0"
    DetailPrint "The SAPI Voice Installer includes sophisticated error handling and cleanup"
    MessageBox MB_ICONEXCLAMATION|MB_OK "COM registration failed!$\r$\n$\r$\nExit code: $0$\r$\n$\r$\nThe SAPI Voice Installer attempted automatic cleanup and retry.$\r$\nIf this persists:$\r$\n1. Restart your computer$\r$\n2. Run the installer again as Administrator$\r$\n3. Check Windows Event Viewer for COM errors"
  ${Else}
    DetailPrint "SUCCESS: COM wrapper registered successfully using SAPI Voice Installer"
    DetailPrint "The installer includes automatic verification and cleanup on failure"
  ${EndIf}
  
SectionEnd

Section "Documentation" SEC02
  SetOutPath "$INSTDIR\docs"
  File /nonfatal "README.md"
  File /nonfatal "CHANGELOG.md"
  File /nonfatal "docs\*.*"
SectionEnd

Section -AdditionalIcons
  WriteIniStr "$INSTDIR\${PRODUCT_NAME}.url" "InternetShortcut" "URL" "${PRODUCT_WEB_SITE}"
  CreateShortCut "$SMPROGRAMS\${PRODUCT_NAME}\Website.lnk" "$INSTDIR\${PRODUCT_NAME}.url"
  CreateShortCut "$SMPROGRAMS\${PRODUCT_NAME}\Uninstall.lnk" "$INSTDIR\uninst.exe"
SectionEnd

Section -Post
  WriteUninstaller "$INSTDIR\uninst.exe"
  WriteRegStr HKLM "${PRODUCT_DIR_REGKEY}" "" "$INSTDIR\sapi_voice_installer.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayName" "$(^Name)"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "UninstallString" "$INSTDIR\uninst.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayIcon" "$INSTDIR\sapi_voice_installer.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayVersion" "${PRODUCT_VERSION}"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "URLInfoAbout" "${PRODUCT_WEB_SITE}"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "Publisher" "${PRODUCT_PUBLISHER}"
SectionEnd

; Post-installation function
Function .onInstSuccess
  ; Ask user if they want to start AACSpeakHelper Server now
  MessageBox MB_ICONQUESTION|MB_YESNO "Installation completed successfully!$\r$\n$\r$\nWould you like to start AACSpeakHelper Server now?$\r$\n$\r$\n(It will also start automatically with Windows)" IDYES StartServer IDNO SkipStart

  StartServer:
    DetailPrint "Starting AACSpeakHelper Server..."
    ExecShell "open" "$INSTDIR\AACSpeakHelperServer.exe"
    Goto ShowFinalMessage

  SkipStart:
    DetailPrint "AACSpeakHelper Server will start automatically with Windows"

  ShowFinalMessage:
    MessageBox MB_ICONINFORMATION|MB_OK "Setup complete!$\r$\n$\r$\nNext steps:$\r$\n1. AACSpeakHelper Server is running (or will start with Windows)$\r$\n2. Use SAPI Voice Installer to install voices$\r$\n3. Test with any SAPI application$\r$\n$\r$\nNote: If COM registration failed, you may need to restart your computer."
FunctionEnd

; Uninstaller sections
Section Uninstall
  ; Unregister COM wrapper using SAPI Voice Installer
  DetailPrint "Unregistering COM wrapper using SAPI Voice Installer..."
  DetailPrint "Executing: $\"$INSTDIR\sapi_voice_installer.exe$\" unregister-com"
  ExecWait '"$INSTDIR\sapi_voice_installer.exe" unregister-com' $0
  ${If} $0 != 0
    DetailPrint "COM unregistration failed with exit code $0"
    DetailPrint "The SAPI Voice Installer includes comprehensive cleanup - this may still be effective"
  ${Else}
    DetailPrint "COM wrapper unregistered successfully using SAPI Voice Installer"
    DetailPrint "All SAPI voice registrations and registry entries have been cleaned"
  ${EndIf}
  
  ; Remove files
  Delete "$INSTDIR\${PRODUCT_NAME}.url"
  Delete "$INSTDIR\uninst.exe"
  Delete "$INSTDIR\sapi_voice_installer.exe"
  Delete "$INSTDIR\AACSpeakHelperServer.exe"
  Delete "$INSTDIR\NativeTTSWrapper.dll"
  Delete "$INSTDIR\fmt.dll"
  Delete "$INSTDIR\onnxruntime.dll"
  Delete "$INSTDIR\onnxruntime_providers_shared.dll"
  Delete "$INSTDIR\sherpa-onnx-c-api.dll"
  Delete "$INSTDIR\settings.cfg.example"
  
  ; Remove directories
  RMDir /r "$INSTDIR\voice_configs"
  RMDir /r "$INSTDIR\docs"
  
  ; Remove shortcuts
  Delete "$SMPROGRAMS\${PRODUCT_NAME}\Uninstall.lnk"
  Delete "$SMPROGRAMS\${PRODUCT_NAME}\Website.lnk"
  Delete "$SMPROGRAMS\${PRODUCT_NAME}\SAPI Voice Installer.lnk"
  Delete "$SMPROGRAMS\${PRODUCT_NAME}\AACSpeakHelper Server.lnk"
  Delete "$DESKTOP\SAPI Voice Installer.lnk"
  Delete "$SMSTARTUP\AACSpeakHelperServer.lnk"
  
  RMDir "$SMPROGRAMS\${PRODUCT_NAME}"
  RMDir "$INSTDIR"
  
  ; Remove registry keys
  DeleteRegKey ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}"
  DeleteRegKey HKLM "${PRODUCT_DIR_REGKEY}"
  
  SetAutoClose true
SectionEnd

; Section descriptions
!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
  !insertmacro MUI_DESCRIPTION_TEXT ${SEC01} "Main application files including SAPI voice installer and TTS server"
  !insertmacro MUI_DESCRIPTION_TEXT ${SEC02} "Documentation and help files"
!insertmacro MUI_FUNCTION_DESCRIPTION_END

; Functions
Function .onInit
  ; Check if running as administrator
  UserInfo::GetAccountType
  pop $0
  ${If} $0 != "admin"
    MessageBox MB_ICONSTOP "Administrator rights required!$\r$\n$\r$\nThis installer needs admin privileges to register COM components.$\r$\n$\r$\nPlease right-click the installer and select 'Run as administrator'."
    SetErrorLevel 740 ; ERROR_ELEVATION_REQUIRED
    Quit
  ${EndIf}

  ; Show admin confirmation
  DetailPrint "Running with administrator privileges - COM registration should work"
FunctionEnd

Function un.onUninstSuccess
  HideWindow
  MessageBox MB_ICONINFORMATION|MB_OK "$(^Name) was successfully removed from your computer."
FunctionEnd

Function un.onInit
  MessageBox MB_ICONQUESTION|MB_YESNO|MB_DEFBUTTON2 "Are you sure you want to completely remove $(^Name) and all of its components?" IDYES +2
  Abort
FunctionEnd
