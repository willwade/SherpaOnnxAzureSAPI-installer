@echo off
setlocal
cd /d C:\github\SherpaOnnxAzureSAPI-installer\NativeTTSWrapper
"C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\amd64\MSBuild.exe" NativeTTSWrapper.sln /p:Configuration=Release /p:Platform=x64 /v:minimal /nologo
echo Build exit code: %ERRORLEVEL%
if %ERRORLEVEL% EQU 0 (
    echo Build succeeded!
    dir x64\Release\NativeTTSWrapper.dll
)
endlocal
