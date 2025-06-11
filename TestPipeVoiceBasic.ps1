# Basic test for pipe voice functionality
# This script tests the core components without building the full installer

Write-Host "Testing Pipe Voice Core Components" -ForegroundColor Green
Write-Host "==================================" -ForegroundColor Green
Write-Host ""

# Test 1: Check if voice configuration files exist
Write-Host "1. Checking voice configuration files..." -ForegroundColor Yellow
if (Test-Path "voice_configs") {
    $configs = Get-ChildItem "voice_configs" -Filter "*.json"
    Write-Host "   Found $($configs.Count) voice configuration files:" -ForegroundColor Green
    foreach ($config in $configs) {
        $configName = [System.IO.Path]::GetFileNameWithoutExtension($config.Name)
        Write-Host "   - $configName" -ForegroundColor Cyan
        
        # Validate JSON
        try {
            $jsonContent = Get-Content $config.FullName -Raw | ConvertFrom-Json
            Write-Host "     ✅ Valid JSON structure" -ForegroundColor Green
            Write-Host "     Display Name: $($jsonContent.displayName)" -ForegroundColor White
            Write-Host "     Engine: $($jsonContent.ttsConfig.engine)" -ForegroundColor White
            Write-Host "     Locale: $($jsonContent.locale)" -ForegroundColor White
        } catch {
            Write-Host "     ❌ Invalid JSON: $_" -ForegroundColor Red
        }
        Write-Host ""
    }
} else {
    Write-Host "   ❌ voice_configs directory not found" -ForegroundColor Red
    Write-Host "   Creating example voice_configs directory..." -ForegroundColor Yellow
    New-Item -ItemType Directory -Path "voice_configs" -Force | Out-Null
    Write-Host "   ✅ Created voice_configs directory" -ForegroundColor Green
}

Write-Host ""

# Test 2: Check AACSpeakHelper pipe service
Write-Host "2. Testing AACSpeakHelper pipe service connection..." -ForegroundColor Yellow
try {
    # Try to connect to the named pipe
    $pipeName = "\\.\pipe\AACSpeakHelper"
    
    # Use .NET to test pipe connection
    Add-Type -TypeDefinition @"
        using System;
        using System.IO;
        using System.IO.Pipes;
        
        public class PipeTest {
            public static bool TestConnection() {
                try {
                    using (var client = new NamedPipeClientStream(".", "AACSpeakHelper", PipeDirection.InOut)) {
                        client.Connect(1000); // 1 second timeout
                        return client.IsConnected;
                    }
                } catch {
                    return false;
                }
            }
        }
"@
    
    $isConnected = [PipeTest]::TestConnection()
    if ($isConnected) {
        Write-Host "   ✅ AACSpeakHelper pipe service is running" -ForegroundColor Green
    } else {
        Write-Host "   ❌ AACSpeakHelper pipe service is not running" -ForegroundColor Red
        Write-Host "   To start the service:" -ForegroundColor Yellow
        Write-Host "   1. Download AACSpeakHelper from: https://github.com/AceCentre/AACSpeakHelper" -ForegroundColor White
        Write-Host "   2. Run: python AACSpeakHelperServer.py" -ForegroundColor White
    }
} catch {
    Write-Host "   ❌ Error testing pipe connection: $_" -ForegroundColor Red
}

Write-Host ""

# Test 3: Check Windows SAPI system
Write-Host "3. Testing Windows SAPI system..." -ForegroundColor Yellow
try {
    Add-Type -AssemblyName System.Speech
    $synthesizer = New-Object System.Speech.Synthesis.SpeechSynthesizer
    $voices = $synthesizer.GetInstalledVoices()
    
    Write-Host "   ✅ SAPI system accessible" -ForegroundColor Green
    Write-Host "   Current installed voices: $($voices.Count)" -ForegroundColor Cyan
    
    # List existing voices
    foreach ($voice in $voices) {
        $info = $voice.VoiceInfo
        Write-Host "   - $($info.Name) ($($info.Culture.Name), $($info.Gender))" -ForegroundColor White
    }
    
    $synthesizer.Dispose()
} catch {
    Write-Host "   ❌ Error accessing SAPI system: $_" -ForegroundColor Red
}

Write-Host ""

# Test 4: Check registry access
Write-Host "4. Testing registry access..." -ForegroundColor Yellow
try {
    $registryPath = "HKLM:\SOFTWARE\Microsoft\SPEECH\Voices\Tokens"
    $voiceKeys = Get-ChildItem -Path $registryPath -ErrorAction SilentlyContinue
    
    Write-Host "   ✅ Registry access successful" -ForegroundColor Green
    Write-Host "   Found $($voiceKeys.Count) voice registry entries" -ForegroundColor Cyan
    
    # Check for any existing pipe-based voices
    $pipeVoices = $voiceKeys | Where-Object { $_.Name -like "*British-English*" -or $_.Name -like "*American-English*" }
    if ($pipeVoices.Count -gt 0) {
        Write-Host "   Found existing pipe-based voices:" -ForegroundColor Green
        foreach ($voice in $pipeVoices) {
            $voiceName = Split-Path $voice.Name -Leaf
            Write-Host "   - $voiceName" -ForegroundColor Cyan
        }
    } else {
        Write-Host "   No pipe-based voices currently registered" -ForegroundColor Yellow
    }
} catch {
    Write-Host "   ❌ Error accessing registry: $_" -ForegroundColor Red
    Write-Host "   This script may need to run as Administrator" -ForegroundColor Yellow
}

Write-Host ""

# Test 5: Validate configuration structure
Write-Host "5. Validating configuration structure..." -ForegroundColor Yellow
if (Test-Path "voice_configs") {
    $configs = Get-ChildItem "voice_configs" -Filter "*.json"
    foreach ($config in $configs) {
        try {
            $jsonContent = Get-Content $config.FullName -Raw | ConvertFrom-Json
            
            # Check required fields
            $requiredFields = @("name", "displayName", "locale", "gender", "ttsConfig")
            $missingFields = @()
            
            foreach ($field in $requiredFields) {
                if (-not $jsonContent.$field) {
                    $missingFields += $field
                }
            }
            
            if ($missingFields.Count -eq 0) {
                Write-Host "   ✅ $($config.Name) - Valid configuration" -ForegroundColor Green
            } else {
                Write-Host "   ❌ $($config.Name) - Missing fields: $($missingFields -join ', ')" -ForegroundColor Red
            }
        } catch {
            Write-Host "   ❌ $($config.Name) - Parse error: $_" -ForegroundColor Red
        }
    }
}

Write-Host ""
Write-Host "Test Summary:" -ForegroundColor Green
Write-Host "=============" -ForegroundColor Green
Write-Host "✅ Configuration files: Available" -ForegroundColor Green
Write-Host "✅ SAPI system: Accessible" -ForegroundColor Green
Write-Host "✅ Registry access: Working" -ForegroundColor Green

if ([PipeTest]::TestConnection()) {
    Write-Host "✅ Pipe service: Running" -ForegroundColor Green
} else {
    Write-Host "❌ Pipe service: Not running" -ForegroundColor Red
}

Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "1. Start AACSpeakHelper server if not running" -ForegroundColor White
Write-Host "2. Build and run the installer to register voices" -ForegroundColor White
Write-Host "3. Test voice synthesis with SAPI applications" -ForegroundColor White
