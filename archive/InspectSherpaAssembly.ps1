# Inspect the sherpa-onnx.dll to see what types are available
Write-Host "Inspecting sherpa-onnx.dll Assembly" -ForegroundColor Green
Write-Host "===================================" -ForegroundColor Green
Write-Host ""

try {
    $sherpaPath = "C:\Program Files\OpenAssistive\OpenSpeech\sherpa-onnx.dll"
    
    Write-Host "Loading assembly from: $sherpaPath" -ForegroundColor Yellow
    
    if (Test-Path $sherpaPath) {
        Write-Host "File exists!" -ForegroundColor Green
        
        # Get file info
        $fileInfo = Get-Item $sherpaPath
        Write-Host "File size: $($fileInfo.Length) bytes" -ForegroundColor Cyan
        Write-Host "Last modified: $($fileInfo.LastWriteTime)" -ForegroundColor Cyan
        
        # Try to load the assembly
        $assembly = [System.Reflection.Assembly]::UnsafeLoadFrom($sherpaPath)
        
        if ($assembly) {
            Write-Host "Assembly loaded successfully!" -ForegroundColor Green
            Write-Host "Assembly full name: $($assembly.FullName)" -ForegroundColor Cyan
            
            # Get all types
            Write-Host ""
            Write-Host "Available types in assembly:" -ForegroundColor Yellow
            $types = $assembly.GetTypes()
            
            if ($types.Length -eq 0) {
                Write-Host "No types found in assembly!" -ForegroundColor Red
            } else {
                Write-Host "Found $($types.Length) types:" -ForegroundColor Green
                
                foreach ($type in $types) {
                    Write-Host "  $($type.FullName)" -ForegroundColor White
                    
                    # Check if this looks like a TTS-related type
                    if ($type.Name -like "*Tts*" -or $type.Name -like "*Config*" -or $type.Name -like "*Offline*") {
                        Write-Host "    ^ This looks TTS-related!" -ForegroundColor Cyan
                    }
                }
                
                # Look specifically for SherpaOnnx namespace
                Write-Host ""
                Write-Host "Types in SherpaOnnx namespace:" -ForegroundColor Yellow
                $sherpaTypes = $types | Where-Object { $_.Namespace -eq "SherpaOnnx" }
                
                if ($sherpaTypes) {
                    foreach ($type in $sherpaTypes) {
                        Write-Host "  $($type.FullName)" -ForegroundColor Green
                    }
                } else {
                    Write-Host "No types found in SherpaOnnx namespace" -ForegroundColor Red
                    
                    # Show all namespaces
                    Write-Host ""
                    Write-Host "Available namespaces:" -ForegroundColor Yellow
                    $namespaces = $types | Where-Object { $_.Namespace } | Select-Object -ExpandProperty Namespace | Sort-Object | Get-Unique
                    foreach ($ns in $namespaces) {
                        Write-Host "  $ns" -ForegroundColor White
                    }
                }
            }
            
        } else {
            Write-Host "Failed to load assembly" -ForegroundColor Red
        }
        
    } else {
        Write-Host "File does not exist: $sherpaPath" -ForegroundColor Red
    }
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Exception type: $($_.Exception.GetType().Name)" -ForegroundColor Red
    
    if ($_.Exception.InnerException) {
        Write-Host "Inner exception: $($_.Exception.InnerException.Message)" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "Assembly inspection completed!" -ForegroundColor Green
