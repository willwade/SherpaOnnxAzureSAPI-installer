@echo off
setlocal enabledelayedexpansion

REM ============================================================================
REM SherpaOnnx Azure SAPI Installer - Complete Build Script
REM ============================================================================
REM This script builds the entire project:
REM 1. C++ COM Wrapper (NativeTTSWrapper)
REM 2. Python Installer Executable (sapi_voice_installer.exe)
REM 3. AACSpeakHelper Server Executable
REM 4. NSIS Installer Package
REM ============================================================================

echo.
echo ============================================================================
echo  SherpaOnnx Azure SAPI Installer - Complete Build Script
echo ============================================================================
echo.

REM Check if running as administrator
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo ❌ ERROR: This script must be run as Administrator
    echo    Right-click and select "Run as administrator"
    pause
    exit /b 1
)

echo ✅ Running as Administrator

REM Set build variables
set "PROJECT_ROOT=%~dp0"
set "MSBUILD_PATH="
set "PYTHON_EXE="
set "PYINSTALLER_EXE="
set "NSIS_EXE="

echo.
echo 🔍 Checking build dependencies...
echo ----------------------------------------

REM Find MSBuild
if exist "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe" (
    set "MSBUILD_PATH=C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
    echo ✅ Found MSBuild 2022: !MSBUILD_PATH!
) else if exist "C:\Program Files (x86)\Microsoft Visual Studio\2019\BuildTools\MSBuild\Current\Bin\MSBuild.exe" (
    set "MSBUILD_PATH=C:\Program Files (x86)\Microsoft Visual Studio\2019\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
    echo ✅ Found MSBuild 2019: !MSBUILD_PATH!
) else if exist "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" (
    set "MSBUILD_PATH=C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
    echo ✅ Found MSBuild 2022 Community: !MSBUILD_PATH!
) else (
    echo ❌ ERROR: MSBuild not found
    echo    Please install Visual Studio Build Tools or Visual Studio
    echo    Download from: https://visualstudio.microsoft.com/downloads/
    pause
    exit /b 1
)

REM Find Python
where python >nul 2>&1
if %errorLevel% equ 0 (
    set "PYTHON_EXE=python"
    echo ✅ Found Python: !PYTHON_EXE!
) else (
    echo ❌ ERROR: Python not found in PATH
    echo    Please install Python and add it to PATH
    pause
    exit /b 1
)

REM Check for uv (Python package manager)
where uv >nul 2>&1
if %errorLevel% equ 0 (
    echo ✅ Found uv package manager
) else (
    echo ❌ ERROR: uv package manager not found
    echo    Please install uv: pip install uv
    pause
    exit /b 1
)

REM Find PyInstaller
where pyinstaller >nul 2>&1
if %errorLevel% equ 0 (
    set "PYINSTALLER_EXE=pyinstaller"
    echo ✅ Found PyInstaller: !PYINSTALLER_EXE!
) else (
    echo ⚠️  PyInstaller not found - will install it
)

REM Find NSIS
if exist "C:\Program Files (x86)\NSIS\makensis.exe" (
    set "NSIS_EXE=C:\Program Files (x86)\NSIS\makensis.exe"
    echo ✅ Found NSIS: !NSIS_EXE!
) else if exist "C:\Program Files\NSIS\makensis.exe" (
    set "NSIS_EXE=C:\Program Files\NSIS\makensis.exe"
    echo ✅ Found NSIS: !NSIS_EXE!
) else (
    echo ⚠️  NSIS not found - installer package will not be created
    echo    Download from: https://nsis.sourceforge.io/Download
)

echo.
echo 🏗️  Starting build process...
echo ============================================================================

REM Step 1: Build C++ COM Wrapper
echo.
echo 📦 Step 1: Building C++ COM Wrapper...
echo ----------------------------------------
cd /d "%PROJECT_ROOT%NativeTTSWrapper"

echo Cleaning previous build...
"!MSBUILD_PATH!" NativeTTSWrapper.vcxproj /p:Configuration=Release /p:Platform=x64 /t:Clean
if %errorLevel% neq 0 (
    echo ❌ ERROR: Failed to clean C++ project
    pause
    exit /b 1
)

echo Building C++ COM Wrapper...
"!MSBUILD_PATH!" NativeTTSWrapper.vcxproj /p:Configuration=Release /p:Platform=x64
if %errorLevel% neq 0 (
    echo ❌ ERROR: Failed to build C++ COM Wrapper
    pause
    exit /b 1
)

if not exist "x64\Release\NativeTTSWrapper.dll" (
    echo ❌ ERROR: NativeTTSWrapper.dll was not created
    pause
    exit /b 1
)

echo ✅ C++ COM Wrapper built successfully

REM Step 2: Register COM Wrapper
echo.
echo 📝 Step 2: Registering COM Wrapper...
echo ----------------------------------------
cd /d "%PROJECT_ROOT%"

call register-com-wrapper.bat
if %errorLevel% neq 0 (
    echo ❌ ERROR: Failed to register COM wrapper
    pause
    exit /b 1
)

echo ✅ COM Wrapper registered successfully

REM Step 3: Install PyInstaller if needed
if "!PYINSTALLER_EXE!"=="" (
    echo.
    echo 📦 Step 3a: Installing PyInstaller...
    echo ----------------------------------------
    pip install pyinstaller
    if %errorLevel% neq 0 (
        echo ❌ ERROR: Failed to install PyInstaller
        pause
        exit /b 1
    )
    set "PYINSTALLER_EXE=pyinstaller"
    echo ✅ PyInstaller installed successfully
)

REM Step 4: Build Python Installer Executable
echo.
echo 🐍 Step 4: Building Python Installer Executable...
echo ----------------------------------------

echo Building sapi_voice_installer.exe...
pyinstaller --onefile --console --name sapi_voice_installer sapi_voice_installer.py
if %errorLevel% neq 0 (
    echo ❌ ERROR: Failed to build Python installer executable
    pause
    exit /b 1
)

if not exist "dist\sapi_voice_installer.exe" (
    echo ❌ ERROR: sapi_voice_installer.exe was not created
    pause
    exit /b 1
)

echo ✅ Python installer executable built successfully

REM Step 5: Build AACSpeakHelper Server Executable
echo.
echo 🗣️  Step 5: Building AACSpeakHelper Server Executable...
echo ----------------------------------------
cd /d "%PROJECT_ROOT%AACSpeakHelper"

echo Building AACSpeakHelperServer.exe...
uv run pyinstaller --onefile --console --name AACSpeakHelperServer AACSpeakHelperServer.py
if %errorLevel% neq 0 (
    echo ❌ ERROR: Failed to build AACSpeakHelper server executable
    pause
    exit /b 1
)

if not exist "dist\AACSpeakHelperServer.exe" (
    echo ❌ ERROR: AACSpeakHelperServer.exe was not created
    pause
    exit /b 1
)

echo ✅ AACSpeakHelper server executable built successfully

cd /d "%PROJECT_ROOT%"

echo.
echo 🎉 BUILD COMPLETED SUCCESSFULLY!
echo ============================================================================
echo.
echo Built components:
echo   ✅ NativeTTSWrapper.dll (C++ COM Wrapper)
echo   ✅ sapi_voice_installer.exe (Python Installer)
echo   ✅ AACSpeakHelperServer.exe (TTS Server)
echo.

if not "!NSIS_EXE!"=="" (
    echo 📦 Step 6: Creating NSIS Installer Package...
    echo ----------------------------------------
    if exist "installer.nsi" (
        "!NSIS_EXE!" installer.nsi
        if %errorLevel% equ 0 (
            echo ✅ NSIS installer package created successfully
        ) else (
            echo ⚠️  NSIS installer creation failed
        )
    ) else (
        echo ⚠️  installer.nsi not found - skipping NSIS package creation
    )
) else (
    echo ⚠️  NSIS not found - installer package not created
    echo    Install NSIS to create installer packages
)

echo.
echo 🚀 Ready for distribution!
echo.
pause
