# Build Complete Fixed Installer with All Source Files
Write-Host "🔧 BUILDING COMPLETE FIXED INSTALLER" -ForegroundColor Cyan
Write-Host "====================================" -ForegroundColor Cyan
Write-Host ""

# Find compiler
$csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (-not (Test-Path $csc)) {
    $csc = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"
}

if (-not (Test-Path $csc)) {
    Write-Host "❌ C# compiler not found!" -ForegroundColor Red
    exit 1
}

Write-Host "✅ Found compiler: $csc" -ForegroundColor Green

# Check for Newtonsoft.Json
if (-not (Test-Path "Newtonsoft.Json.dll")) {
    Write-Host "⚠️ Newtonsoft.Json.dll not found. Downloading..." -ForegroundColor Yellow
    try {
        $url = "https://www.nuget.org/api/v2/package/Newtonsoft.Json/13.0.3"
        Invoke-WebRequest -Uri $url -OutFile "newtonsoft.zip"
        Expand-Archive -Path "newtonsoft.zip" -DestinationPath "temp" -Force
        Copy-Item "temp\lib\net45\Newtonsoft.Json.dll" "Newtonsoft.Json.dll"
        Remove-Item "newtonsoft.zip" -Force
        Remove-Item "temp" -Recurse -Force
        Write-Host "✅ Downloaded Newtonsoft.Json.dll" -ForegroundColor Green
    } catch {
        Write-Host "❌ Failed to download Newtonsoft.Json: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

# List all source files
$sourceFiles = @(
    "Installer\Program.cs",
    "Installer\ModelInstaller.cs",
    "Installer\Sapi5RegistrarExtended.cs",
    "Installer\AzureConfigManager.cs",
    "Installer\AzureTtsService.cs",
    "Installer\AzureVoiceInstaller.cs",
    "Installer\Shared\TtsModel.cs",
    "Installer\Shared\AzureConfig.cs",
    "Installer\Shared\AzureTtsModel.cs",
    "Installer\Shared\LanguageCodeConverter.cs",
    "Installer\Shared\LanguageInfo.cs"
)

Write-Host "📁 Source files to compile:" -ForegroundColor Yellow
foreach ($file in $sourceFiles) {
    if (Test-Path $file) {
        Write-Host "  ✅ $file" -ForegroundColor Green
    } else {
        Write-Host "  ❌ $file (MISSING)" -ForegroundColor Red
    }
}

# Check if all files exist
$missingFiles = $sourceFiles | Where-Object { -not (Test-Path $_) }
if ($missingFiles.Count -gt 0) {
    Write-Host ""
    Write-Host "❌ Missing source files! Cannot compile." -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "🔨 Compiling complete installer..." -ForegroundColor Yellow

# Compiler arguments
$outputFile = "SherpaOnnxSAPIInstaller_Fixed.exe"
$args = @(
    "/out:$outputFile",
    "/target:exe",
    "/platform:x64",
    "/optimize+",
    "/reference:System.dll",
    "/reference:System.Core.dll",
    "/reference:System.Net.Http.dll",
    "/reference:System.Threading.Tasks.dll",
    "/reference:Microsoft.Win32.Registry.dll",
    "/reference:Newtonsoft.Json.dll"
)

# Add all source files
$args += $sourceFiles

Write-Host "Compiler command:" -ForegroundColor Gray
Write-Host "  $csc" -ForegroundColor Gray
foreach ($arg in $args) {
    Write-Host "    $arg" -ForegroundColor Gray
}
Write-Host ""

try {
    # Run compiler
    $process = Start-Process -FilePath $csc -ArgumentList $args -Wait -PassThru -NoNewWindow -RedirectStandardOutput "compile_output.txt" -RedirectStandardError "compile_errors.txt"
    
    if ($process.ExitCode -eq 0) {
        Write-Host "✅ COMPILATION SUCCESSFUL!" -ForegroundColor Green
        Write-Host "📦 Created: $outputFile" -ForegroundColor Green
        
        # Copy dependencies
        if (Test-Path "Newtonsoft.Json.dll") {
            Write-Host "📋 Newtonsoft.Json.dll is available" -ForegroundColor Gray
        }
        
        Write-Host ""
        Write-Host "🎯 FIXED INSTALLER IS READY!" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "🧪 Test the fixed installer:" -ForegroundColor Yellow
        Write-Host "  # Uninstall current ElliotNeural voice" -ForegroundColor Gray
        Write-Host "  .\$outputFile uninstall 'Microsoft Server Speech Text to Speech Voice (en-GB, ElliotNeural)'" -ForegroundColor Gray
        Write-Host ""
        Write-Host "  # Reinstall with fixed installer" -ForegroundColor Gray
        Write-Host "  .\$outputFile install-azure en-GB-ElliotNeural --key b14f8945b0f1459f9964bdd72c42c2cc --region uksouth" -ForegroundColor Gray
        Write-Host ""
        Write-Host "  # Test the voice" -ForegroundColor Gray
        Write-Host "  powershell -File TestElliotVoice.ps1" -ForegroundColor Gray
        
    } else {
        Write-Host "❌ COMPILATION FAILED!" -ForegroundColor Red
        
        if (Test-Path "compile_errors.txt") {
            Write-Host ""
            Write-Host "Compilation Errors:" -ForegroundColor Red
            Get-Content "compile_errors.txt" | ForEach-Object {
                Write-Host "  $_" -ForegroundColor Red
            }
        }
        
        if (Test-Path "compile_output.txt") {
            Write-Host ""
            Write-Host "Compilation Output:" -ForegroundColor Yellow
            Get-Content "compile_output.txt" | ForEach-Object {
                Write-Host "  $_" -ForegroundColor Yellow
            }
        }
    }
    
} catch {
    Write-Host "❌ Compilation error: $($_.Exception.Message)" -ForegroundColor Red
}

# Cleanup temp files
Remove-Item "compile_output.txt" -ErrorAction SilentlyContinue
Remove-Item "compile_errors.txt" -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "🎯 BUILD COMPLETED!" -ForegroundColor Cyan
