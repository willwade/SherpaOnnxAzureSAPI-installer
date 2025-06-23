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

REM Clean build directories
echo Cleaning build directories...
if exist "NativeTTSWrapper\x64\Release\*.obj" del "NativeTTSWrapper\x64\Release\*.obj" /Q
if exist "NativeTTSWrapper\x64\Release\NativeTTSWrapper.tlog" rmdir "NativeTTSWrapper\x64\Release\NativeTTSWrapper.tlog" /S /Q
if exist "NativeTTSWrapper\Win32\Release\*.obj" del "NativeTTSWrapper\Win32\Release\*.obj" /Q
if exist "NativeTTSWrapper\Win32\Release\NativeTTSWrapper.tlog" rmdir "NativeTTSWrapper\Win32\Release\NativeTTSWrapper.tlog" /S /Q

REM Try to rename old DLLs if they exist and are locked
if exist "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll" (
    echo Attempting to rename old x64 DLL...
    ren "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll" "NativeTTSWrapper_old.dll" 2>nul
    if exist "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll" (
        echo Warning: Could not rename old x64 DLL - it may be in use
    )
)
if exist "NativeTTSWrapper\Win32\Release\NativeTTSWrapper.dll" (
    echo Attempting to rename old x86 DLL...
    ren "NativeTTSWrapper\Win32\Release\NativeTTSWrapper.dll" "NativeTTSWrapper_old.dll" 2>nul
    if exist "NativeTTSWrapper\Win32\Release\NativeTTSWrapper.dll" (
        echo Warning: Could not rename old x86 DLL - it may be in use
    )
)

REM Build both x86 and x64 versions
echo Building x64 version...
msbuild "NativeTTSWrapper\NativeTTSWrapper.vcxproj" /p:Configuration=Release /p:Platform=x64 /p:VerbosityLevel=minimal
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: x64 Build failed
    pause
    exit /b 1
)

echo Building x86 version...
msbuild "NativeTTSWrapper\NativeTTSWrapper.vcxproj" /p:Configuration=Release /p:Platform=Win32 /p:VerbosityLevel=minimal
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: x86 Build failed
    pause
    exit /b 1
)

echo Both builds successful!

echo Verifying output files...
if exist "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll" (
    echo ✅ x64 DLL created successfully
) else (
    echo ❌ x64 DLL not found
)

if exist "NativeTTSWrapper\Win32\Release\NativeTTSWrapper.dll" (
    echo ✅ x86 DLL created successfully
) else (
    echo ❌ x86 DLL not found
)

echo Build complete successfully!
