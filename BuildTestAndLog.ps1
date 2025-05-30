# Complete Build, Register, Test and Log Script for OpenSpeechTTS
# This script must be run as Administrator
# All output will be logged to a file for analysis

param(
    [string]$LogFile = "C:\OpenSpeech\build_test_log.txt"
)

# Ensure log directory exists
$logDir = Split-Path $LogFile -Parent
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

# Function to write to both console and log file
function Write-Log {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $logMessage = "$timestamp : $Message"
    
    # Write to console with color
    Write-Host $logMessage -ForegroundColor $Color
    
    # Write to log file
    Add-Content -Path $LogFile -Value $logMessage
}

# Start logging
Write-Log "=== OpenSpeechTTS Build, Register, Test and Log Script ===" "Cyan"
Write-Log "Log file: $LogFile" "Yellow"
Write-Log ""

try {
    # Step 1: Build the project
    Write-Log "STEP 1: Building the project..." "Green"
    Set-Location "C:\GitHub\SherpaOnnxAzureSAPI-installer"
    
    $buildOutput = & dotnet build OpenSpeechTTS/OpenSpeechTTS.csproj --configuration Release 2>&1
    Write-Log "Build output:"
    $buildOutput | ForEach-Object { Write-Log "  $_" }
    
    if ($LASTEXITCODE -ne 0) {
        Write-Log "BUILD FAILED!" "Red"
        exit 1
    }
    Write-Log "Build completed successfully" "Green"
    Write-Log ""

    # Step 2: Register the COM component
    Write-Log "STEP 2: Registering COM component..." "Green"
    
    $dllPath = "OpenSpeechTTS\bin\Release\net472\OpenSpeechTTS.dll"
    if (-not (Test-Path $dllPath)) {
        Write-Log "ERROR: DLL not found at $dllPath" "Red"
        exit 1
    }
    
    Write-Log "Registering DLL: $dllPath"
    $regOutput = & "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe" $dllPath /codebase 2>&1
    Write-Log "Registration output:"
    $regOutput | ForEach-Object { Write-Log "  $_" }
    
    if ($LASTEXITCODE -ne 0) {
        Write-Log "REGISTRATION FAILED!" "Red"
        exit 1
    }
    Write-Log "Registration completed successfully" "Green"
    Write-Log ""

    # Step 3: Verify registration
    Write-Log "STEP 3: Verifying registration..." "Green"
    try {
        $regInfo = Get-ItemProperty -Path "HKLM:\SOFTWARE\Classes\CLSID\{3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2}\InprocServer32" -ErrorAction Stop
        Write-Log "Registry verification successful:"
        Write-Log "  Default: $($regInfo.'(default)')"
        Write-Log "  Class: $($regInfo.Class)"
        Write-Log "  CodeBase: $($regInfo.CodeBase)"
        Write-Log "  Assembly: $($regInfo.Assembly)"
    } catch {
        Write-Log "ERROR: Registry verification failed: $($_.Exception.Message)" "Red"
        exit 1
    }
    Write-Log ""

    # Step 4: Clear old logs
    Write-Log "STEP 4: Clearing old SAPI logs..." "Green"
    $sapiLogs = @(
        "C:\OpenSpeech\sapi_debug.log",
        "C:\OpenSpeech\sapi_error.log",
        "C:\OpenSpeech\sapi_speak.log"
    )
    
    foreach ($log in $sapiLogs) {
        if (Test-Path $log) {
            Remove-Item $log -Force
            Write-Log "  Removed: $log"
        }
    }
    Write-Log ""

    # Step 5: Run the test
    Write-Log "STEP 5: Running COM voice test..." "Green"
    
    if (-not (Test-Path ".\TestCOMVoice.ps1")) {
        Write-Log "ERROR: TestCOMVoice.ps1 not found in current directory" "Red"
        exit 1
    }
    
    Write-Log "Executing TestCOMVoice.ps1..."
    $testOutput = & ".\TestCOMVoice.ps1" 2>&1
    Write-Log "Test output:"
    $testOutput | ForEach-Object { Write-Log "  $_" }
    Write-Log ""

    # Step 6: Analyze SAPI debug logs
    Write-Log "STEP 6: Analyzing SAPI debug logs..." "Green"
    
    $debugLogPath = "C:\OpenSpeech\sapi_debug.log"
    if (Test-Path $debugLogPath) {
        Write-Log "SAPI Debug Log Contents:"
        $debugContent = Get-Content $debugLogPath
        $debugContent | ForEach-Object { Write-Log "  $_" }
        
        # Look for key indicators
        Write-Log ""
        Write-Log "KEY INDICATORS ANALYSIS:" "Yellow"
        
        $constructorCalls = ($debugContent | Where-Object { $_ -like "*constructor*" }).Count
        $setTokenCalls = ($debugContent | Where-Object { $_ -like "*SET OBJECT TOKEN*" }).Count
        $getFormatCalls = ($debugContent | Where-Object { $_ -like "*GET OUTPUT FORMAT*" }).Count
        $speakCalls = ($debugContent | Where-Object { $_ -like "*SPEAK METHOD CALLED*" }).Count
        
        Write-Log "  Constructor calls: $constructorCalls"
        Write-Log "  SetObjectToken calls: $setTokenCalls"
        Write-Log "  GetOutputFormat calls: $getFormatCalls"
        Write-Log "  Speak method calls: $speakCalls"
        
        if ($constructorCalls -gt 0) {
            Write-Log "  ✅ Constructor is being called with new code" "Green"
        } else {
            Write-Log "  ❌ Constructor calls not detected - old code may still be running" "Red"
        }
        
        if ($setTokenCalls -gt 0) {
            Write-Log "  ✅ SetObjectToken is being called - ISpObjectWithToken interface working!" "Green"
        } else {
            Write-Log "  ❌ SetObjectToken not called - interface issue" "Red"
        }
        
        if ($speakCalls -gt 0) {
            Write-Log "  🎉 BREAKTHROUGH: Speak method is being called!" "Green"
        } else {
            Write-Log "  ⏳ Speak method not yet called - still working on interface" "Yellow"
        }
        
    } else {
        Write-Log "WARNING: No SAPI debug log found at $debugLogPath" "Yellow"
    }
    Write-Log ""

    # Step 7: Analyze SAPI error logs
    Write-Log "STEP 7: Analyzing SAPI error logs..." "Green"
    
    $errorLogPath = "C:\OpenSpeech\sapi_error.log"
    if (Test-Path $errorLogPath) {
        Write-Log "SAPI Error Log Contents:"
        $errorContent = Get-Content $errorLogPath
        $errorContent | ForEach-Object { Write-Log "  $_" "Red" }
    } else {
        Write-Log "✅ No SAPI error log found - no errors detected!" "Green"
    }
    Write-Log ""

    # Step 8: Summary and next steps
    Write-Log "STEP 8: Summary and recommendations..." "Green"
    
    if ($speakCalls -gt 0) {
        Write-Log "🎉 SUCCESS: SAPI is calling our Speak method!" "Green"
        Write-Log "Next steps: Enable real TTS processing and test audio output" "Yellow"
    } elseif ($setTokenCalls -gt 0) {
        Write-Log "✅ PROGRESS: SetObjectToken working, need to debug why Speak isn't called" "Yellow"
        Write-Log "Next steps: Check GetOutputFormat return values and error handling" "Yellow"
    } elseif ($getFormatCalls -gt 0) {
        Write-Log "⚠️  PARTIAL: GetOutputFormat working, but SetObjectToken missing" "Yellow"
        Write-Log "Next steps: Debug ISpObjectWithToken interface implementation" "Yellow"
    } else {
        Write-Log "❌ ISSUE: No method calls detected - COM interface problem" "Red"
        Write-Log "Next steps: Check COM registration and interface definitions" "Yellow"
    }

} catch {
    Write-Log "SCRIPT ERROR: $($_.Exception.Message)" "Red"
    Write-Log "Stack trace: $($_.ScriptStackTrace)" "Red"
    exit 1
}

Write-Log ""
Write-Log "=== BUILD, TEST AND LOG COMPLETE ===" "Cyan"
Write-Log "All results logged to: $LogFile" "Green"
Write-Log "You can now analyze the log file for detailed results." "Green"
