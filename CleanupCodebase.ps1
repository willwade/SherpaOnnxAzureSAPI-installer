#!/usr/bin/env pwsh
# Cleanup Codebase - Remove clutter and organize essential files
# This script will clean up the codebase and keep only essential files

Write-Host "🧹 Cleaning up codebase..." -ForegroundColor Cyan
Write-Host "==============================" -ForegroundColor Cyan

# Create archive directory if it doesn't exist
if (-not (Test-Path "archive")) {
    New-Item -ItemType Directory -Path "archive" | Out-Null
}

# Create archive subdirectories
$archiveDirs = @("test-scripts", "old-dotnet", "build-experiments", "temp-files")
foreach ($dir in $archiveDirs) {
    $archivePath = "archive\$dir"
    if (-not (Test-Path $archivePath)) {
        New-Item -ItemType Directory -Path $archivePath | Out-Null
        Write-Host "   Created: $archivePath" -ForegroundColor Green
    }
}

Write-Host "`n1. Moving test scripts to archive..." -ForegroundColor Yellow

# Test scripts to archive
$testScripts = @(
    "Test*.ps1",
    "Test*.cpp", 
    "Test*.exe",
    "Test*.obj",
    "Build*Test*.ps1",
    "Build*Test*.bat",
    "Check*.ps1",
    "Debug*.ps1",
    "Fix*.ps1"
)

foreach ($pattern in $testScripts) {
    $files = Get-ChildItem -Filter $pattern -ErrorAction SilentlyContinue
    foreach ($file in $files) {
        if ($file.Name -notin @("TestAmyVoiceSpecific.ps1")) { # Keep essential test
            Move-Item $file.FullName "archive\test-scripts\" -Force
            Write-Host "   Moved: $($file.Name)" -ForegroundColor Gray
        }
    }
}

Write-Host "`n2. Moving build experiments to archive..." -ForegroundColor Yellow

# Build experiment scripts
$buildExperiments = @(
    "Build*.bat",
    "Build*.ps1"
)

foreach ($pattern in $buildExperiments) {
    $files = Get-ChildItem -Filter $pattern -ErrorAction SilentlyContinue
    foreach ($file in $files) {
        if ($file.Name -notin @("BuildNativeOnly.ps1", "BuildCompleteInstaller.ps1")) { # Keep essential builds
            Move-Item $file.FullName "archive\build-experiments\" -Force
            Write-Host "   Moved: $($file.Name)" -ForegroundColor Gray
        }
    }
}

Write-Host "`n3. Moving temporary and object files..." -ForegroundColor Yellow

# Temporary and object files
$tempFiles = @(
    "*.obj",
    "*.exe",
    "real_engine_test_config.json",
    "test_engine_config.json",
    "engines_config.json"
)

foreach ($pattern in $tempFiles) {
    $files = Get-ChildItem -Filter $pattern -ErrorAction SilentlyContinue
    foreach ($file in $files) {
        Move-Item $file.FullName "archive\temp-files\" -Force
        Write-Host "   Moved: $($file.Name)" -ForegroundColor Gray
    }
}

Write-Host "`n4. Moving old .NET experimental code..." -ForegroundColor Yellow

# Old .NET projects that are no longer needed
$oldDotNetDirs = @(
    "SherpaNative",
    "SignAssembly",
    "KeyGenerator"
)

foreach ($dir in $oldDotNetDirs) {
    if (Test-Path $dir) {
        Move-Item $dir "archive\old-dotnet\" -Force
        Write-Host "   Moved: $dir" -ForegroundColor Gray
    }
}

Write-Host "`n5. Cleaning up root directory files..." -ForegroundColor Yellow

# Root directory files to archive
$rootFilesToArchive = @(
    "*.md",
    "*.dll",
    "*.tar.bz2",
    "app.manifest",
    "Directory.build.props"
)

foreach ($pattern in $rootFilesToArchive) {
    $files = Get-ChildItem -Filter $pattern -ErrorAction SilentlyContinue
    foreach ($file in $files) {
        # Keep essential documentation
        if ($file.Name -notin @("README.md", "CPP_TTS_IMPLEMENTATION_PLAN.md")) {
            Move-Item $file.FullName "archive\temp-files\" -Force
            Write-Host "   Moved: $($file.Name)" -ForegroundColor Gray
        }
    }
}

Write-Host "`n6. Cleaning up build artifacts..." -ForegroundColor Yellow

# Clean build artifacts
$buildArtifacts = @("obj", "bin")
foreach ($dir in $buildArtifacts) {
    if (Test-Path $dir) {
        Remove-Item $dir -Recurse -Force
        Write-Host "   Removed: $dir" -ForegroundColor Gray
    }
}

Write-Host "`n7. Creating clean project structure..." -ForegroundColor Yellow

# Create a clean project structure summary
$cleanStructure = @"
📁 Clean Project Structure:
==========================

Essential C++ Code:
├── NativeTTSWrapper/          # Main C++ COM wrapper
│   ├── *.cpp, *.h            # Core implementation
│   ├── libs/                 # SherpaOnnx libraries
│   └── engines_config.json   # Engine configuration

Essential .NET Code:
├── OpenSpeechTTS/            # .NET SAPI implementation
├── SherpaWorker/             # Process bridge worker
└── Installer/                # Installation utilities

Essential Scripts:
├── RegisterAmyVoice.ps1      # Voice registration
├── BuildNativeOnly.ps1       # Build C++ wrapper
├── BuildCompleteInstaller.ps1 # Build full installer
└── TestAmyVoiceSpecific.ps1  # Essential test

Distribution:
└── dist/                     # Built binaries and installer

Documentation:
├── README.md                 # Project overview
└── CPP_TTS_IMPLEMENTATION_PLAN.md # Implementation plan

Archive:
└── archive/                  # All old/experimental code
    ├── test-scripts/         # Test scripts
    ├── build-experiments/    # Build experiments
    ├── old-dotnet/          # Old .NET projects
    └── temp-files/          # Temporary files
"@

Write-Host $cleanStructure -ForegroundColor Green

Write-Host "`n✅ Codebase cleanup complete!" -ForegroundColor Green
Write-Host "📊 Summary:" -ForegroundColor Cyan
Write-Host "   - Moved test scripts to archive/test-scripts/" -ForegroundColor White
Write-Host "   - Moved build experiments to archive/build-experiments/" -ForegroundColor White
Write-Host "   - Moved old .NET projects to archive/old-dotnet/" -ForegroundColor White
Write-Host "   - Moved temporary files to archive/temp-files/" -ForegroundColor White
Write-Host "   - Cleaned up build artifacts" -ForegroundColor White
Write-Host "   - Kept only essential code and scripts" -ForegroundColor White

Write-Host "`n🎯 Next Steps:" -ForegroundColor Cyan
Write-Host "   1. Review the cleaned structure" -ForegroundColor White
Write-Host "   2. Rebuild NativeTTSWrapper.dll with latest code" -ForegroundColor White
Write-Host "   3. Test the essential functionality" -ForegroundColor White
Write-Host "   4. Deploy and verify the fallback chain" -ForegroundColor White
