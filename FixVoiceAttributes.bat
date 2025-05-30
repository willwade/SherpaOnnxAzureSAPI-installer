@echo off
echo Fixing Amy Voice Attributes...
echo This script must be run as Administrator!
echo.

echo Updating Gender from Male to Female...
reg add "HKLM\SOFTWARE\Microsoft\Speech\Voices\Tokens\amy\Attributes" /v Gender /d "Female" /f

echo.
echo Verifying the change...
reg query "HKLM\SOFTWARE\Microsoft\Speech\Voices\Tokens\amy\Attributes" /v Gender

echo.
echo Voice attributes updated successfully!
echo Amy voice should now be properly recognized as Female.
echo.
pause
