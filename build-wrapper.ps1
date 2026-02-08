$vsDevCmd = "C:\Program Files\Microsoft Visual Studio\18\Community\VC\Auxiliary\Build\vcvars64.bat"
$projectDir = "C:\github\SherpaOnnxAzureSAPI-installer\NativeTTSWrapper"

Write-Host "Setting up Visual Studio environment..."
cmd /c "`"$vsDevCmd`" && set" | ForEach-Object {
    if ($_ -match "=(.+)") {
        [Environment]::SetEnvironmentVariable($matches[1], $matches[2], "Process")
    }
}

Write-Host "Building NativeTTSWrapper..."
Push-Location $projectDir

$output = cmd /c "msbuild NativeTTSWrapper.vcxproj /p:Configuration=Release /p:Platform=x64 /v:m 2>&1"
Write-Host $output

Pop-Location

Write-Host "Build completed."
