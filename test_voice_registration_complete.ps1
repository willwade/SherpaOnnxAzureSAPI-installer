# Complete Voice Registration Test
# Tests the entire workflow from building to SAPI synthesis

param(
    [string]$VoiceName = "English-SherpaOnnx-Jenny",
    [switch]$SkipBuild,
    [switch]$SkipRegistration,
    [switch]$Verbose
)

Write-Host "🎯 Complete Voice Registration Test" -ForegroundColor Cyan
Write-Host "Voice: $VoiceName" -ForegroundColor White
Write-Host "=================================" -ForegroundColor Cyan
Write-Host ""

$ErrorActionPreference = "Continue"

# Function to run command with error handling
function Invoke-SafeCommand {
    param(
        [string]$Command,
        [string]$Arguments,
        [string]$Description,
        [switch]$RequireSuccess
    )
    
    Write-Host "🔧 $Description..." -ForegroundColor Yellow
    
    try {
        if ($Arguments) {
            $result = Start-Process -FilePath $Command -ArgumentList $Arguments -Wait -PassThru -NoNewWindow
        } else {
            $result = Start-Process -FilePath $Command -Wait -PassThru -NoNewWindow
        }
        
        if ($result.ExitCode -eq 0) {
            Write-Host "✅ $Description completed successfully" -ForegroundColor Green
            return $true
        } else {
            Write-Host "❌ $Description failed with exit code: $($result.ExitCode)" -ForegroundColor Red
            if ($RequireSuccess) {
                throw "$Description failed"
            }
            return $false
        }
    } catch {
        Write-Host "❌ $Description failed: $($_.Exception.Message)" -ForegroundColor Red
        if ($RequireSuccess) {
            throw
        }
        return $false
    }
}

# Step 1: Check Prerequisites
Write-Host "📋 Step 1: Prerequisites Check" -ForegroundColor Yellow
Write-Host "==============================" -ForegroundColor Yellow

$prerequisites = @{
    "Python" = { python --version }
    "uv" = { uv --version }
    "MSBuild" = { Test-Path "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe" }
    ".NET SDK" = { dotnet --version }
}

$allPrereqsMet = $true
foreach ($prereq in $prerequisites.Keys) {
    try {
        $result = & $prerequisites[$prereq]
        if ($result) {
            Write-Host "✅ $prereq`: $result" -ForegroundColor Green
        } else {
            Write-Host "✅ $prereq`: Available" -ForegroundColor Green
        }
    } catch {
        Write-Host "❌ $prereq`: Not found" -ForegroundColor Red
        $allPrereqsMet = $false
    }
}

if (-not $allPrereqsMet) {
    Write-Host "❌ Prerequisites not met. Exiting." -ForegroundColor Red
    exit 1
}

Write-Host ""

# Step 2: Build Components (if not skipped)
if (-not $SkipBuild) {
    Write-Host "📋 Step 2: Build Components" -ForegroundColor Yellow
    Write-Host "===========================" -ForegroundColor Yellow

    # Build C++ COM Wrapper
    $msbuildPath = "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
    $cppProject = "NativeTTSWrapper\NativeTTSWrapper.vcxproj"
    
    if (Test-Path $cppProject) {
        $success = Invoke-SafeCommand -Command $msbuildPath -Arguments "`"$cppProject`" /p:Configuration=Release /p:Platform=x64" -Description "Building C++ COM Wrapper" -RequireSuccess
        
        $dllPath = "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll"
        if (Test-Path $dllPath) {
            Write-Host "✅ C++ COM wrapper built: $dllPath" -ForegroundColor Green
        } else {
            Write-Host "❌ C++ COM wrapper DLL not found" -ForegroundColor Red
            exit 1
        }
    } else {
        Write-Host "❌ C++ project not found: $cppProject" -ForegroundColor Red
        exit 1
    }

    # Build .NET Installer
    $dotnetProject = "Installer\Installer.csproj"
    if (Test-Path $dotnetProject) {
        Set-Location "Installer"
        $success = Invoke-SafeCommand -Command "dotnet" -Arguments "build -c Release" -Description "Building .NET Installer" -RequireSuccess
        Set-Location ".."
        
        $installerPath = "Installer\bin\Release\net6.0\SherpaOnnxSAPIInstaller.exe"
        if (Test-Path $installerPath) {
            Write-Host "✅ .NET installer built: $installerPath" -ForegroundColor Green
        } else {
            Write-Host "❌ .NET installer executable not found" -ForegroundColor Red
            exit 1
        }
    } else {
        Write-Host "❌ .NET project not found: $dotnetProject" -ForegroundColor Red
        exit 1
    }

    Write-Host ""
}

# Step 3: Register COM Wrapper (if not skipped)
if (-not $SkipRegistration) {
    Write-Host "📋 Step 3: Register COM Wrapper" -ForegroundColor Yellow
    Write-Host "===============================" -ForegroundColor Yellow

    $dllPath = "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll"
    if (Test-Path $dllPath) {
        $fullDllPath = (Resolve-Path $dllPath).Path
        $success = Invoke-SafeCommand -Command "regsvr32" -Arguments "/s `"$fullDllPath`"" -Description "Registering COM Wrapper" -RequireSuccess
    } else {
        Write-Host "❌ COM wrapper DLL not found: $dllPath" -ForegroundColor Red
        exit 1
    }

    Write-Host ""
}

# Step 4: Check Voice Configuration
Write-Host "📋 Step 4: Check Voice Configuration" -ForegroundColor Yellow
Write-Host "====================================" -ForegroundColor Yellow

$configPath = "voice_configs\$VoiceName.json"
if (Test-Path $configPath) {
    Write-Host "✅ Voice configuration found: $configPath" -ForegroundColor Green
    
    # Display configuration
    try {
        $config = Get-Content $configPath | ConvertFrom-Json
        Write-Host "   Name: $($config.name)" -ForegroundColor Gray
        Write-Host "   Display Name: $($config.displayName)" -ForegroundColor Gray
        Write-Host "   Engine: $($config.ttsConfig.args.engine)" -ForegroundColor Gray
        Write-Host "   Voice: $($config.ttsConfig.args.voice)" -ForegroundColor Gray
    } catch {
        Write-Host "⚠️ Could not parse configuration file" -ForegroundColor Yellow
    }
} else {
    Write-Host "❌ Voice configuration not found: $configPath" -ForegroundColor Red
    Write-Host "Available configurations:" -ForegroundColor White
    if (Test-Path "voice_configs") {
        Get-ChildItem "voice_configs" -Filter "*.json" | ForEach-Object {
            Write-Host "  - $($_.BaseName)" -ForegroundColor Gray
        }
    }
    exit 1
}

Write-Host ""

# Step 5: Install Voice
Write-Host "📋 Step 5: Install Voice" -ForegroundColor Yellow
Write-Host "========================" -ForegroundColor Yellow

try {
    $result = uv run python SapiVoiceManager.py --install $VoiceName
    Write-Host "✅ Voice installation command completed" -ForegroundColor Green
} catch {
    Write-Host "❌ Voice installation failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Trying alternative installation method..." -ForegroundColor Yellow
    
    # Try using .NET installer directly
    $installerPath = "Installer\bin\Release\net6.0\SherpaOnnxSAPIInstaller.exe"
    if (Test-Path $installerPath) {
        try {
            $result = & $installerPath "install-pipe-voice" $VoiceName
            Write-Host "✅ Voice installed using .NET installer" -ForegroundColor Green
        } catch {
            Write-Host "❌ .NET installer also failed: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

Write-Host ""

# Step 6: Verify Registry Registration
Write-Host "📋 Step 6: Verify Registry Registration" -ForegroundColor Yellow
Write-Host "=======================================" -ForegroundColor Yellow

$voiceRegistryPath = "HKLM:\SOFTWARE\Microsoft\SPEECH\Voices\Tokens\$VoiceName"
try {
    $voiceKey = Get-ItemProperty -Path $voiceRegistryPath -ErrorAction Stop
    Write-Host "✅ Voice registered in SAPI registry" -ForegroundColor Green
    Write-Host "   CLSID: $($voiceKey.CLSID)" -ForegroundColor Gray
    Write-Host "   Default: $($voiceKey.'(default)')" -ForegroundColor Gray
    
    # Check CLSID registration
    $clsid = $voiceKey.CLSID
    if ($clsid) {
        $clsidPath = "HKCR:\CLSID\$clsid"
        try {
            $clsidKey = Get-ItemProperty -Path $clsidPath -ErrorAction Stop
            Write-Host "✅ CLSID registered: $clsid" -ForegroundColor Green
            
            # Check InprocServer32
            $inprocPath = "$clsidPath\InprocServer32"
            try {
                $inprocKey = Get-ItemProperty -Path $inprocPath -ErrorAction Stop
                $dllPath = $inprocKey.'(default)'
                Write-Host "   DLL Path: $dllPath" -ForegroundColor Gray
                
                if (Test-Path $dllPath) {
                    Write-Host "   ✅ DLL exists" -ForegroundColor Green
                } else {
                    Write-Host "   ❌ DLL not found" -ForegroundColor Red
                }
            } catch {
                Write-Host "   ❌ InprocServer32 not found" -ForegroundColor Red
            }
        } catch {
            Write-Host "❌ CLSID not registered: $clsid" -ForegroundColor Red
        }
    }
} catch {
    Write-Host "❌ Voice not found in SAPI registry" -ForegroundColor Red
}

Write-Host ""

# Step 7: Test SAPI Enumeration
Write-Host "📋 Step 7: Test SAPI Enumeration" -ForegroundColor Yellow
Write-Host "=================================" -ForegroundColor Yellow

try {
    $voice = New-Object -ComObject SAPI.SpVoice
    $voices = $voice.GetVoices()
    
    Write-Host "Found $($voices.Count) SAPI voices:" -ForegroundColor White
    
    $ourVoiceFound = $false
    for ($i = 0; $i -lt $voices.Count; $i++) {
        $voiceInfo = $voices.Item($i)
        $name = $voiceInfo.GetDescription()
        
        if ($name -like "*$VoiceName*" -or $name -like "*SherpaOnnx*" -or $name -like "*Jenny*") {
            Write-Host "  ✅ $name" -ForegroundColor Green
            $ourVoiceFound = $true
        } else {
            Write-Host "  - $name" -ForegroundColor White
        }
    }
    
    if (-not $ourVoiceFound) {
        Write-Host "⚠️ Our voice not found in SAPI enumeration" -ForegroundColor Yellow
    }
} catch {
    Write-Host "❌ Failed to enumerate SAPI voices: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""

# Step 8: Test Voice Synthesis
Write-Host "📋 Step 8: Test Voice Synthesis" -ForegroundColor Yellow
Write-Host "===============================" -ForegroundColor Yellow

try {
    $voice = New-Object -ComObject SAPI.SpVoice
    $voices = $voice.GetVoices()
    
    # Find our voice
    $targetVoice = $null
    for ($i = 0; $i -lt $voices.Count; $i++) {
        $voiceInfo = $voices.Item($i)
        $name = $voiceInfo.GetDescription()
        if ($name -like "*$VoiceName*" -or $name -like "*SherpaOnnx*" -or $name -like "*Jenny*") {
            $targetVoice = $voiceInfo
            break
        }
    }
    
    if ($targetVoice) {
        Write-Host "✅ Found voice: $($targetVoice.GetDescription())" -ForegroundColor Green
        
        # Set the voice
        $voice.Voice = $targetVoice
        
        # Test synthesis
        $testText = "Hello from the C++ SAPI Bridge to AACSpeakHelper! This is a test of voice synthesis using $VoiceName."
        Write-Host "🎵 Speaking test text..." -ForegroundColor Yellow
        
        try {
            $voice.Speak($testText)
            Write-Host "✅ Voice synthesis completed successfully!" -ForegroundColor Green
        } catch {
            Write-Host "❌ Voice synthesis failed: $($_.Exception.Message)" -ForegroundColor Red
        }
    } else {
        Write-Host "❌ Voice not found for synthesis test" -ForegroundColor Red
        
        # Try with any available voice as fallback
        if ($voices.Count -gt 0) {
            Write-Host "🔄 Testing with default voice..." -ForegroundColor Yellow
            $voice.Speak("Testing default voice synthesis.")
            Write-Host "✅ Default voice synthesis works" -ForegroundColor Green
        }
    }
} catch {
    Write-Host "❌ Voice synthesis test failed: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""

# Step 9: Summary
Write-Host "📋 Step 9: Test Summary" -ForegroundColor Yellow
Write-Host "=======================" -ForegroundColor Yellow

Write-Host "🎯 Voice Registration Test Results:" -ForegroundColor Cyan
Write-Host "  Voice: $VoiceName" -ForegroundColor White
Write-Host "  Configuration: $(if (Test-Path $configPath) { '✅ Found' } else { '❌ Missing' })" -ForegroundColor White
Write-Host "  Registry: $(if (Test-Path "HKLM:\SOFTWARE\Microsoft\SPEECH\Voices\Tokens\$VoiceName") { '✅ Registered' } else { '❌ Not Registered' })" -ForegroundColor White
Write-Host "  SAPI: $(if ($ourVoiceFound) { '✅ Enumerable' } else { '❌ Not Found' })" -ForegroundColor White

Write-Host ""
Write-Host "🚀 Next Steps:" -ForegroundColor Cyan
Write-Host "  1. Ensure AACSpeakHelper service is running:" -ForegroundColor White
Write-Host "     uv run python AACSpeakHelperServer.py" -ForegroundColor Gray
Write-Host "  2. Test in real applications (Notepad, etc.)" -ForegroundColor White
Write-Host "  3. Check logs for any errors:" -ForegroundColor White
Write-Host "     C:\OpenSpeech\native_tts_debug.log" -ForegroundColor Gray

Write-Host ""
Write-Host "🎉 Voice registration test completed!" -ForegroundColor Green
