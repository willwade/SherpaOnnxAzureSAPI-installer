; SherpaOnnx Azure SAPI Installer - NSIS Script
; ============================================================================

!define PRODUCT_NAME "SherpaOnnx Azure SAPI Bridge"
!define PRODUCT_VERSION "1.0.0"
!define PRODUCT_PUBLISHER "AceCentre"
!define PRODUCT_WEB_SITE "https://github.com/willwade/SherpaOnnxAzureSAPI-installer"
!define PRODUCT_DIR_REGKEY "Software\Microsoft\Windows\CurrentVersion\App Paths\sapi_voice_installer.exe"
!define PRODUCT_UNINST_KEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}"
!define PRODUCT_UNINST_ROOT_KEY "HKLM"

; Request administrator privileges
RequestExecutionLevel admin

; Modern UI
!include "MUI2.nsh"

; General
Name "${PRODUCT_NAME} ${PRODUCT_VERSION}"
OutFile "SherpaOnnxAzureSAPI-installer.exe"
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
  File "AACSpeakHelper\dist\AACSpeakHelperServer.exe"
  
  ; Copy C++ COM Wrapper
  File "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll"
  File "NativeTTSWrapper\x64\Release\*.dll"
  
  ; Copy voice configurations
  SetOutPath "$INSTDIR\voice_configs"
  File /r "voice_configs\*.*"
  
  ; Copy AACSpeakHelper files
  SetOutPath "$INSTDIR\AACSpeakHelper"
  File "AACSpeakHelper\*.py"
  File "AACSpeakHelper\*.cfg"
  File /nonfatal "AACSpeakHelper\requirements.txt"
  
  ; Create shortcuts
  CreateDirectory "$SMPROGRAMS\${PRODUCT_NAME}"
  CreateShortCut "$SMPROGRAMS\${PRODUCT_NAME}\SAPI Voice Installer.lnk" "$INSTDIR\sapi_voice_installer.exe"
  CreateShortCut "$SMPROGRAMS\${PRODUCT_NAME}\AACSpeakHelper Server.lnk" "$INSTDIR\AACSpeakHelperServer.exe"
  CreateShortCut "$DESKTOP\SAPI Voice Installer.lnk" "$INSTDIR\sapi_voice_installer.exe"
  
  ; Register COM wrapper
  ExecWait 'regsvr32 /s "$INSTDIR\NativeTTSWrapper.dll"'
  
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

; Uninstaller sections
Section Uninstall
  ; Unregister COM wrapper
  ExecWait 'regsvr32 /u /s "$INSTDIR\NativeTTSWrapper.dll"'
  
  ; Remove files
  Delete "$INSTDIR\${PRODUCT_NAME}.url"
  Delete "$INSTDIR\uninst.exe"
  Delete "$INSTDIR\sapi_voice_installer.exe"
  Delete "$INSTDIR\AACSpeakHelperServer.exe"
  Delete "$INSTDIR\NativeTTSWrapper.dll"
  Delete "$INSTDIR\*.dll"
  
  ; Remove directories
  RMDir /r "$INSTDIR\voice_configs"
  RMDir /r "$INSTDIR\AACSpeakHelper"
  RMDir /r "$INSTDIR\docs"
  
  ; Remove shortcuts
  Delete "$SMPROGRAMS\${PRODUCT_NAME}\Uninstall.lnk"
  Delete "$SMPROGRAMS\${PRODUCT_NAME}\Website.lnk"
  Delete "$SMPROGRAMS\${PRODUCT_NAME}\SAPI Voice Installer.lnk"
  Delete "$SMPROGRAMS\${PRODUCT_NAME}\AACSpeakHelper Server.lnk"
  Delete "$DESKTOP\SAPI Voice Installer.lnk"
  
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
    MessageBox MB_ICONSTOP "Administrator rights required!"
    SetErrorLevel 740 ; ERROR_ELEVATION_REQUIRED
    Quit
  ${EndIf}
FunctionEnd

Function un.onUninstSuccess
  HideWindow
  MessageBox MB_ICONINFORMATION|MB_OK "$(^Name) was successfully removed from your computer."
FunctionEnd

Function un.onInit
  MessageBox MB_ICONQUESTION|MB_YESNO|MB_DEFBUTTON2 "Are you sure you want to completely remove $(^Name) and all of its components?" IDYES +2
  Abort
FunctionEnd
