@echo off
echo Building COM Wrapper...

REM Try multiple Visual Studio paths
set "VS_PATH="
if exist "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvars64.bat" (
    set "VS_PATH=C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvars64.bat"
) else if exist "C:\Program Files\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvars64.bat" (
    set "VS_PATH=C:\Program Files\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvars64.bat"
) else if exist "C:\Program Files (x86)\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat" (
    set "VS_PATH=C:\Program Files (x86)\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat"
) else if exist "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat" (
    set "VS_PATH=C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat"
) else if exist "C:\Program Files (x86)\Microsoft Visual Studio\2022\Professional\VC\Auxiliary\Build\vcvars64.bat" (
    set "VS_PATH=C:\Program Files (x86)\Microsoft Visual Studio\2022\Professional\VC\Auxiliary\Build\vcvars64.bat"
) else if exist "C:\Program Files\Microsoft Visual Studio\2022\Professional\VC\Auxiliary\Build\vcvars64.bat" (
    set "VS_PATH=C:\Program Files\Microsoft Visual Studio\2022\Professional\VC\Auxiliary\Build\vcvars64.bat"
) else if exist "C:\Program Files (x86)\Microsoft Visual Studio\2022\Enterprise\VC\Auxiliary\Build\vcvars64.bat" (
    set "VS_PATH=C:\Program Files (x86)\Microsoft Visual Studio\2022\Enterprise\VC\Auxiliary\Build\vcvars64.bat"
) else if exist "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\VC\Auxiliary\Build\vcvars64.bat" (
    set "VS_PATH=C:\Program Files\Microsoft Visual Studio\2022\Enterprise\VC\Auxiliary\Build\vcvars64.bat"
)

if "%VS_PATH%"=="" (
    echo ERROR: Visual Studio 2022 not found
    echo Please install Visual Studio 2022 Build Tools
    pause
    exit /b 1
)

echo Found Visual Studio at: %VS_PATH%

REM Set up Visual Studio environment
call "%VS_PATH%"
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Failed to set up Visual Studio environment
    pause
    exit /b 1
)

echo Visual Studio environment set up successfully

REM Clean build directory
echo Cleaning build directory...
if exist "NativeTTSWrapper\x64\Release\*.obj" del "NativeTTSWrapper\x64\Release\*.obj" /Q
if exist "NativeTTSWrapper\x64\Release\NativeTTSWrapper.tlog" rmdir "NativeTTSWrapper\x64\Release\NativeTTSWrapper.tlog" /S /Q

REM Try to rename old DLL if it exists and is locked
if exist "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll" (
    echo Attempting to rename old DLL...
    ren "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll" "NativeTTSWrapper_old.dll" 2>nul
    if exist "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll" (
        echo Warning: Could not rename old DLL - it may be in use
    )
)

REM Build the project
echo Building project...
msbuild "NativeTTSWrapper\NativeTTSWrapper.vcxproj" /p:Configuration=Release /p:Platform=x64 /p:VerbosityLevel=minimal
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Build failed
    pause
    exit /b 1
)

echo Build successful!

REM Copy dependencies
echo Copying dependencies...
copy "NativeTTSWrapper\libs\*.dll" "NativeTTSWrapper\x64\Release\" /Y >nul
copy "NativeTTSWrapper\azure-speech-sdk\bin\*.dll" "NativeTTSWrapper\x64\Release\" /Y >nul

echo Build complete successfully!
