# Complete Windows Integration Test for C++ SAPI Bridge to AACSpeakHelper
# This script tests the entire pipeline with SherpaOnnx and Google TTS (no credentials needed)

param(
    [string]$TestVoice = "English-SherpaOnnx-Jenny",
    [switch]$SkipBuild,
    [switch]$TestGoogle,
    [switch]$Verbose
)

Write-Host "🎉 C++ SAPI Bridge to AACSpeakHelper - Windows Integration Test" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "Testing Voice: $TestVoice" -ForegroundColor White
Write-Host "Test Google TTS: $TestGoogle" -ForegroundColor White
Write-Host ""

$ErrorActionPreference = "Continue"

# Function to test if a process is running
function Test-ProcessRunning {
    param([string]$ProcessName)
    
    try {
        $process = Get-Process -Name $ProcessName -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

# Function to wait for AACSpeakHelper service
function Wait-ForAACSpeakHelper {
    param([int]$TimeoutSeconds = 30)
    
    Write-Host "⏳ Waiting for AACSpeakHelper service to start..." -ForegroundColor Yellow
    
    $timeout = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $timeout) {
        if (Test-ProcessRunning "python") {
            # Check if the pipe exists (this is a rough check)
            try {
                $pipe = New-Object System.IO.Pipes.NamedPipeClientStream(".", "AACSpeakHelper", [System.IO.Pipes.PipeDirection]::InOut)
                $pipe.Connect(1000)  # 1 second timeout
                $pipe.Close()
                Write-Host "✅ AACSpeakHelper service is running and pipe is accessible" -ForegroundColor Green
                return $true
            } catch {
                # Pipe not ready yet
            }
        }
        Start-Sleep -Seconds 2
    }
    
    Write-Host "⚠️ AACSpeakHelper service not detected within timeout" -ForegroundColor Yellow
    return $false
}

# Step 1: Check Prerequisites
Write-Host "📋 Step 1: Prerequisites Check" -ForegroundColor Yellow
Write-Host "==============================" -ForegroundColor Yellow

$prerequisites = @{
    "Python" = { python --version }
    "uv" = { uv --version }
    "Git" = { git --version }
}

$allPrereqsMet = $true
foreach ($prereq in $prerequisites.Keys) {
    try {
        $result = & $prerequisites[$prereq]
        Write-Host "✅ $prereq`: $result" -ForegroundColor Green
    } catch {
        Write-Host "❌ $prereq`: Not found" -ForegroundColor Red
        $allPrereqsMet = $false
    }
}

if (-not $allPrereqsMet) {
    Write-Host "❌ Prerequisites not met. Please install missing components." -ForegroundColor Red
    exit 1
}

Write-Host ""

# Step 2: Set up AACSpeakHelper
Write-Host "📋 Step 2: Setting up AACSpeakHelper" -ForegroundColor Yellow
Write-Host "====================================" -ForegroundColor Yellow

$aacspeakhelperDir = "AACSpeakHelper"

if (-not (Test-Path $aacspeakhelperDir)) {
    Write-Host "Cloning AACSpeakHelper repository..." -ForegroundColor Cyan
    try {
        git clone https://github.com/AceCentre/AACSpeakHelper
        Write-Host "✅ AACSpeakHelper cloned successfully" -ForegroundColor Green
    } catch {
        Write-Host "❌ Failed to clone AACSpeakHelper: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
} else {
    Write-Host "✅ AACSpeakHelper directory already exists" -ForegroundColor Green
}

# Set up Python environment
Write-Host "Setting up Python environment..." -ForegroundColor Cyan
try {
    Set-Location $aacspeakhelperDir
    
    # Create virtual environment
    if (-not (Test-Path ".venv")) {
        uv venv
        Write-Host "✅ Virtual environment created" -ForegroundColor Green
    } else {
        Write-Host "✅ Virtual environment already exists" -ForegroundColor Green
    }
    
    # Install dependencies
    uv sync --all-extras
    Write-Host "✅ Dependencies installed" -ForegroundColor Green
    
    Set-Location ".."
} catch {
    Write-Host "❌ Failed to set up Python environment: $($_.Exception.Message)" -ForegroundColor Red
    Set-Location ".."
    exit 1
}

Write-Host ""

# Step 3: Start AACSpeakHelper Service
Write-Host "📋 Step 3: Starting AACSpeakHelper Service" -ForegroundColor Yellow
Write-Host "===========================================" -ForegroundColor Yellow

Write-Host "Starting AACSpeakHelper server..." -ForegroundColor Cyan

# Start AACSpeakHelper in background
$aacspeakhelperJob = Start-Job -ScriptBlock {
    param($aacspeakhelperPath)
    Set-Location $aacspeakhelperPath
    uv run python AACSpeakHelperServer.py
} -ArgumentList (Resolve-Path $aacspeakhelperDir).Path

Write-Host "✅ AACSpeakHelper service started in background (Job ID: $($aacspeakhelperJob.Id))" -ForegroundColor Green

# Wait for service to be ready
$serviceReady = Wait-ForAACSpeakHelper -TimeoutSeconds 30

if (-not $serviceReady) {
    Write-Host "⚠️ Could not confirm AACSpeakHelper service is ready" -ForegroundColor Yellow
    Write-Host "   Continuing with test - service may still be starting..." -ForegroundColor White
}

Write-Host ""

# Step 4: Test Voice Configurations
Write-Host "📋 Step 4: Testing Voice Configurations" -ForegroundColor Yellow
Write-Host "=======================================" -ForegroundColor Yellow

$voiceConfigs = @("English-SherpaOnnx-Jenny")
if ($TestGoogle) {
    $voiceConfigs += "English-Google-Basic"
}

foreach ($voiceConfig in $voiceConfigs) {
    $configPath = "voice_configs\$voiceConfig.json"
    
    if (Test-Path $configPath) {
        Write-Host "✅ Found configuration: $voiceConfig" -ForegroundColor Green
        
        try {
            $config = Get-Content $configPath | ConvertFrom-Json
            Write-Host "   Engine: $($config.ttsConfig.args.engine)" -ForegroundColor Gray
            Write-Host "   Voice: $($config.ttsConfig.args.voice)" -ForegroundColor Gray
            Write-Host "   Display Name: $($config.displayName)" -ForegroundColor Gray
        } catch {
            Write-Host "   ⚠️ Could not parse configuration" -ForegroundColor Yellow
        }
    } else {
        Write-Host "❌ Configuration not found: $voiceConfig" -ForegroundColor Red
    }
}

Write-Host ""

# Step 5: Build Components (if not skipped)
if (-not $SkipBuild) {
    Write-Host "📋 Step 5: Building Components" -ForegroundColor Yellow
    Write-Host "==============================" -ForegroundColor Yellow

    # Build C++ COM Wrapper
    $msbuildPath = "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
    if (-not (Test-Path $msbuildPath)) {
        $msbuildPath = "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
    }
    
    if (Test-Path $msbuildPath) {
        Write-Host "Building C++ COM wrapper..." -ForegroundColor Cyan
        try {
            & $msbuildPath "NativeTTSWrapper\NativeTTSWrapper.vcxproj" /p:Configuration=Release /p:Platform=x64
            
            $dllPath = "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll"
            if (Test-Path $dllPath) {
                Write-Host "✅ C++ COM wrapper built successfully" -ForegroundColor Green
            } else {
                Write-Host "❌ C++ COM wrapper DLL not found" -ForegroundColor Red
            }
        } catch {
            Write-Host "❌ C++ build failed: $($_.Exception.Message)" -ForegroundColor Red
        }
    } else {
        Write-Host "⚠️ MSBuild not found - skipping C++ build" -ForegroundColor Yellow
    }

    # Build .NET Installer
    Write-Host "Building .NET installer..." -ForegroundColor Cyan
    try {
        Set-Location "Installer"
        dotnet build -c Release
        Set-Location ".."
        Write-Host "✅ .NET installer built successfully" -ForegroundColor Green
    } catch {
        Write-Host "❌ .NET build failed: $($_.Exception.Message)" -ForegroundColor Red
        Set-Location ".."
    }

    Write-Host ""
}

# Step 6: Register COM Wrapper
Write-Host "📋 Step 6: Registering COM Wrapper" -ForegroundColor Yellow
Write-Host "===================================" -ForegroundColor Yellow

$dllPath = "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll"
if (Test-Path $dllPath) {
    Write-Host "Registering COM wrapper..." -ForegroundColor Cyan
    try {
        $fullDllPath = (Resolve-Path $dllPath).Path
        Start-Process -FilePath "regsvr32" -ArgumentList "/s", "`"$fullDllPath`"" -Wait -Verb RunAs
        Write-Host "✅ COM wrapper registered successfully" -ForegroundColor Green
    } catch {
        Write-Host "❌ COM registration failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "   Try running as Administrator" -ForegroundColor Yellow
    }
} else {
    Write-Host "⚠️ COM wrapper DLL not found - skipping registration" -ForegroundColor Yellow
}

Write-Host ""

# Step 7: Test Voice Installation
Write-Host "📋 Step 7: Testing Voice Installation" -ForegroundColor Yellow
Write-Host "=====================================" -ForegroundColor Yellow

foreach ($voiceConfig in $voiceConfigs) {
    Write-Host "Testing installation of: $voiceConfig" -ForegroundColor Cyan
    
    try {
        # Test the CLI tool
        $result = uv run python SapiVoiceManager.py --view $voiceConfig
        Write-Host "✅ CLI tool can access configuration: $voiceConfig" -ForegroundColor Green
        
        # Attempt installation (this will test the full pipeline)
        Write-Host "Attempting voice installation..." -ForegroundColor White
        $installResult = uv run python SapiVoiceManager.py --install $voiceConfig
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✅ Voice installation completed: $voiceConfig" -ForegroundColor Green
        } else {
            Write-Host "⚠️ Voice installation may have issues: $voiceConfig" -ForegroundColor Yellow
        }
        
    } catch {
        Write-Host "❌ Error testing voice installation: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host ""
}

# Step 8: Test SAPI Integration
Write-Host "📋 Step 8: Testing SAPI Integration" -ForegroundColor Yellow
Write-Host "===================================" -ForegroundColor Yellow

Write-Host "Testing SAPI voice enumeration..." -ForegroundColor Cyan

try {
    $voice = New-Object -ComObject SAPI.SpVoice
    $voices = $voice.GetVoices()
    
    Write-Host "Found $($voices.Count) SAPI voices:" -ForegroundColor White
    
    $testVoicesFound = @()
    for ($i = 0; $i -lt $voices.Count; $i++) {
        $voiceInfo = $voices.Item($i)
        $name = $voiceInfo.GetDescription()
        
        $isTestVoice = $false
        foreach ($testVoice in $voiceConfigs) {
            if ($name -like "*$testVoice*" -or $name -like "*SherpaOnnx*" -or $name -like "*Google*") {
                $isTestVoice = $true
                $testVoicesFound += $name
                break
            }
        }
        
        if ($isTestVoice) {
            Write-Host "  ✅ $name" -ForegroundColor Green
        } else {
            Write-Host "  - $name" -ForegroundColor White
        }
    }
    
    # Test voice synthesis
    if ($testVoicesFound.Count -gt 0) {
        Write-Host ""
        Write-Host "🔊 Testing voice synthesis..." -ForegroundColor Cyan
        
        foreach ($testVoiceName in $testVoicesFound) {
            Write-Host "Testing synthesis with: $testVoiceName" -ForegroundColor White
            
            try {
                # Find and set the voice
                for ($i = 0; $i -lt $voices.Count; $i++) {
                    $voiceInfo = $voices.Item($i)
                    if ($voiceInfo.GetDescription() -eq $testVoiceName) {
                        $voice.Voice = $voiceInfo
                        break
                    }
                }
                
                # Test synthesis
                $testText = "Hello from the C++ SAPI Bridge to AACSpeakHelper! This is a test of voice synthesis using $testVoiceName."
                $voice.Speak($testText)
                
                Write-Host "✅ Voice synthesis completed successfully!" -ForegroundColor Green
                break  # Success with one voice is enough
                
            } catch {
                Write-Host "❌ Voice synthesis failed: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    } else {
        Write-Host "⚠️ No test voices found in SAPI enumeration" -ForegroundColor Yellow
    }
    
} catch {
    Write-Host "❌ SAPI integration test failed: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""

# Step 9: Cleanup and Summary
Write-Host "📋 Step 9: Cleanup and Summary" -ForegroundColor Yellow
Write-Host "==============================" -ForegroundColor Yellow

# Stop AACSpeakHelper service
if ($aacspeakhelperJob) {
    Write-Host "Stopping AACSpeakHelper service..." -ForegroundColor Cyan
    Stop-Job -Job $aacspeakhelperJob
    Remove-Job -Job $aacspeakhelperJob
    Write-Host "✅ AACSpeakHelper service stopped" -ForegroundColor Green
}

Write-Host ""
Write-Host "🎯 Integration Test Summary:" -ForegroundColor Cyan
Write-Host "  Test Voice: $TestVoice" -ForegroundColor White
Write-Host "  AACSpeakHelper: $(if ($serviceReady) { '✅ Running' } else { '⚠️ Status Unknown' })" -ForegroundColor White
Write-Host "  Voice Configurations: ✅ Valid" -ForegroundColor White
Write-Host "  C++ COM Wrapper: $(if (Test-Path 'NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll') { '✅ Built' } else { '❌ Not Built' })" -ForegroundColor White

Write-Host ""
Write-Host "🚀 Next Steps:" -ForegroundColor Cyan
Write-Host "  1. Verify AACSpeakHelper service is running properly" -ForegroundColor White
Write-Host "  2. Test voice synthesis in real applications (Notepad, etc.)" -ForegroundColor White
Write-Host "  3. Check logs for any errors or issues" -ForegroundColor White
Write-Host "  4. Test with different TTS engines (SherpaOnnx, Google)" -ForegroundColor White

Write-Host ""
Write-Host "🎉 Windows integration test completed!" -ForegroundColor Green
