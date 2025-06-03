# Check loader exceptions for sherpa-onnx.dll
Write-Host "Checking Loader Exceptions" -ForegroundColor Green
Write-Host "==========================" -ForegroundColor Green
Write-Host ""

try {
    $sherpaPath = "C:\Program Files\OpenAssistive\OpenSpeech\sherpa-onnx.dll"
    
    Write-Host "Loading assembly: $sherpaPath" -ForegroundColor Yellow
    $assembly = [System.Reflection.Assembly]::UnsafeLoadFrom($sherpaPath)
    
    Write-Host "Attempting to get types..." -ForegroundColor Yellow
    
    try {
        $types = $assembly.GetTypes()
        Write-Host "Success! Found $($types.Length) types" -ForegroundColor Green
    } catch [System.Reflection.ReflectionTypeLoadException] {
        Write-Host "ReflectionTypeLoadException caught!" -ForegroundColor Red
        
        $loadException = $_.Exception
        
        Write-Host ""
        Write-Host "Successfully loaded types:" -ForegroundColor Green
        if ($loadException.Types) {
            foreach ($type in $loadException.Types) {
                if ($type) {
                    Write-Host "  $($type.FullName)" -ForegroundColor White
                }
            }
        }
        
        Write-Host ""
        Write-Host "Loader exceptions:" -ForegroundColor Red
        if ($loadException.LoaderExceptions) {
            foreach ($ex in $loadException.LoaderExceptions) {
                Write-Host "  Exception: $($ex.GetType().Name)" -ForegroundColor Yellow
                Write-Host "  Message: $($ex.Message)" -ForegroundColor Red
                
                if ($ex -is [System.IO.FileNotFoundException]) {
                    Write-Host "  Missing file: $($ex.FileName)" -ForegroundColor Red
                }
                
                Write-Host ""
            }
        }
    }
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "Also checking what files are in the OpenSpeech directory:" -ForegroundColor Yellow
Get-ChildItem "C:\Program Files\OpenAssistive\OpenSpeech\" -Name | ForEach-Object {
    Write-Host "  $_" -ForegroundColor Gray
}

Write-Host ""
Write-Host "Loader exception check completed!" -ForegroundColor Green
