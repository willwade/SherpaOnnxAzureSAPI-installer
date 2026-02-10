$ErrorActionPreference = "Stop"
$msbuild = "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\amd64\MSBuild.exe"
$sln = "C:\github\SherpaOnnxAzureSAPI-installer\NativeTTSWrapper\NativeTTSWrapper.sln"

Write-Host "Building DLL..."
$process = Start-Process -FilePath $msbuild -ArgumentList "`"$sln`"", "/p:Configuration=Release", "/p:Platform=x64", "/v:minimal" -Wait -PassThru -NoNewWindow

Write-Host "Build exit code: $($process.ExitCode)"
if ($process.ExitCode -eq 0) {
    $dllPath = "C:\github\SherpaOnnxAzureSAPI-installer\NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll"
    if (Test-Path $dllPath) {
        $timestamp = (Get-Item $dllPath).LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
        Write-Host "DLL built: $timestamp"
    }
}
exit $process.ExitCode
