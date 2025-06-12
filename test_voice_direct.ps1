# Test Voice Direct - Bypass COM and test pipe service directly
# This script tests if the AACSpeakHelper can handle Azure TTS requests

Write-Host "🧪 DIRECT VOICE TEST (Bypassing COM)" -ForegroundColor Cyan
Write-Host "====================================" -ForegroundColor Cyan
Write-Host ""

# Check if AACSpeakHelper is running
$pythonProcess = Get-Process | Where-Object {$_.ProcessName -like "*python*"}
if (-not $pythonProcess) {
    Write-Host "❌ AACSpeakHelper service not running" -ForegroundColor Red
    Write-Host "Please start it with: cd AACSpeakHelper && uv run AACSpeakHelperServer.py" -ForegroundColor Yellow
    exit 1
}

Write-Host "✅ AACSpeakHelper service is running" -ForegroundColor Green
Write-Host ""

# Test using the AACSpeakHelper client directly
Write-Host "🔧 Testing AACSpeakHelper with Azure TTS..." -ForegroundColor Yellow

try {
    # Set environment variables for the test
    $env:MICROSOFT_TOKEN = "b14f8945b0f1459f9964bdd72c42c2cc"
    $env:MICROSOFT_REGION = "uksouth"
    
    # Change to AACSpeakHelper directory
    Push-Location AACSpeakHelper
    
    # Test with Azure TTS
    Write-Host "Running: uv run client.py --engine azure --voice en-GB-LibbyNeural --text 'Hello from Azure TTS test'" -ForegroundColor Gray
    
    $result = uv run client.py --engine azure --voice en-GB-LibbyNeural --text "Hello from Azure TTS test" 2>&1
    $exitCode = $LASTEXITCODE
    
    Pop-Location
    
    Write-Host "Exit code: $exitCode" -ForegroundColor Gray
    Write-Host "Output:" -ForegroundColor Gray
    Write-Host $result -ForegroundColor Gray
    
    if ($exitCode -eq 0) {
        Write-Host ""
        Write-Host "✅ DIRECT TEST SUCCESSFUL!" -ForegroundColor Green
        Write-Host "🎯 The pipe service can handle Azure TTS requests" -ForegroundColor Green
        Write-Host "🎯 Issue is with COM wrapper registration, not TTS engine" -ForegroundColor Yellow
    } else {
        Write-Host ""
        Write-Host "❌ DIRECT TEST FAILED" -ForegroundColor Red
        Write-Host "🎯 Issue might be with Azure credentials or AACSpeakHelper configuration" -ForegroundColor Yellow
    }
    
} catch {
    Write-Host "❌ Error running direct test: $_" -ForegroundColor Red
}

Write-Host ""
Write-Host "🔍 ANALYSIS:" -ForegroundColor Cyan
Write-Host "  • If direct test works: COM wrapper registration is the issue" -ForegroundColor White
Write-Host "  • If direct test fails: Azure TTS configuration is the issue" -ForegroundColor White
Write-Host ""
Write-Host "💡 NEXT STEPS:" -ForegroundColor Cyan
Write-Host "  • If direct test works: Need to fix .NET 6 COM registration" -ForegroundColor White
Write-Host "  • If direct test fails: Need to fix Azure TTS credentials/config" -ForegroundColor White

Write-Host ""
Write-Host "Test completed!" -ForegroundColor Green
