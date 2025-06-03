@echo off
echo Registering SAPI5 Interface GUIDs...
echo This script must be run as Administrator!
echo.

REM Register ISpTTSEngine interface
echo Registering ISpTTSEngine interface...
reg add "HKLM\SOFTWARE\Classes\Interface\{A74D7C8E-4CC5-4F2F-A6EB-804DEE18500E}" /ve /d "ISpTTSEngine" /f
reg add "HKLM\SOFTWARE\Classes\Interface\{A74D7C8E-4CC5-4F2F-A6EB-804DEE18500E}\ProxyStubClsid32" /ve /d "{00020424-0000-0000-C000-000000000046}" /f
reg add "HKLM\SOFTWARE\Classes\Interface\{A74D7C8E-4CC5-4F2F-A6EB-804DEE18500E}\TypeLib" /ve /d "{C866CA3A-32F7-11D2-9602-00C04F8EE628}" /f
reg add "HKLM\SOFTWARE\Classes\Interface\{A74D7C8E-4CC5-4F2F-A6EB-804DEE18500E}\TypeLib" /v "Version" /d "5.4" /f

echo ISpTTSEngine interface registered successfully!
echo.

REM Register ISpObjectWithToken interface
echo Registering ISpObjectWithToken interface...
reg add "HKLM\SOFTWARE\Classes\Interface\{14056581-E16C-11D2-BB90-00C04F8EE6C0}" /ve /d "ISpObjectWithToken" /f
reg add "HKLM\SOFTWARE\Classes\Interface\{14056581-E16C-11D2-BB90-00C04F8EE6C0}\ProxyStubClsid32" /ve /d "{00020424-0000-0000-C000-000000000046}" /f
reg add "HKLM\SOFTWARE\Classes\Interface\{14056581-E16C-11D2-BB90-00C04F8EE6C0}\TypeLib" /ve /d "{C866CA3A-32F7-11D2-9602-00C04F8EE628}" /f
reg add "HKLM\SOFTWARE\Classes\Interface\{14056581-E16C-11D2-BB90-00C04F8EE6C0}\TypeLib" /v "Version" /d "5.4" /f

echo ISpObjectWithToken interface registered successfully!
echo.

echo ===================================
echo INTERFACE REGISTRATION COMPLETE!
echo ===================================
echo Both SAPI5 interfaces are now registered.
echo SAPI should now be able to call our methods!
echo.
echo Next steps:
echo 1. Test with TestSpeech.ps1
echo 2. Look for "SET OBJECT TOKEN CALLED" in logs
echo 3. Look for "SPEAK METHOD CALLED" in logs
echo.
pause
