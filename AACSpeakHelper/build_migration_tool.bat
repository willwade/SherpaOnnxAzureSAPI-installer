@echo off
echo Building AACSpeakHelper Settings Migration Tool...
echo.

REM Build the migration script as a standalone executable
uv run python -m PyInstaller migrate_settings.py --onefile --console --name "AACSpeakHelper-Settings-Migration" --clean -y

echo.
if exist "dist\AACSpeakHelper-Settings-Migration.exe" (
    echo ✅ Migration tool built successfully!
    echo 📁 Location: dist\AACSpeakHelper-Settings-Migration.exe
    echo.
    echo 💡 Usage:
    echo    1. Copy the executable and your settings.cfg to the same folder
    echo    2. Run the executable
    echo    3. It will automatically migrate the settings to the correct location
) else (
    echo ❌ Build failed!
)

pause
