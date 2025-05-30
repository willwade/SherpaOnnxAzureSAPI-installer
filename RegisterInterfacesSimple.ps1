# Simple Interface Registration Script
# Register the missing SAPI5 interface GUIDs

Write-Host "Registering SAPI5 Interface GUIDs..." -ForegroundColor Yellow

# ISpTTSEngine interface
$guid1 = "A74D7C8E-4CC5-4F2F-A6EB-804DEE18500E"
$path1 = "HKLM:\SOFTWARE\Classes\Interface\{$guid1}"

Write-Host "Registering ISpTTSEngine ($guid1)..."
New-Item -Path $path1 -Force | Out-Null
Set-ItemProperty -Path $path1 -Name "(Default)" -Value "ISpTTSEngine"

New-Item -Path "$path1\ProxyStubClsid32" -Force | Out-Null
Set-ItemProperty -Path "$path1\ProxyStubClsid32" -Name "(Default)" -Value "{00020424-0000-0000-C000-000000000046}"

New-Item -Path "$path1\TypeLib" -Force | Out-Null
Set-ItemProperty -Path "$path1\TypeLib" -Name "(Default)" -Value "{C866CA3A-32F7-11D2-9602-00C04F8EE628}"
Set-ItemProperty -Path "$path1\TypeLib" -Name "Version" -Value "5.4"

Write-Host "ISpTTSEngine registered successfully" -ForegroundColor Green

# ISpObjectWithToken interface
$guid2 = "14056581-E16C-11D2-BB90-00C04F8EE6C0"
$path2 = "HKLM:\SOFTWARE\Classes\Interface\{$guid2}"

Write-Host "Registering ISpObjectWithToken ($guid2)..."
New-Item -Path $path2 -Force | Out-Null
Set-ItemProperty -Path $path2 -Name "(Default)" -Value "ISpObjectWithToken"

New-Item -Path "$path2\ProxyStubClsid32" -Force | Out-Null
Set-ItemProperty -Path "$path2\ProxyStubClsid32" -Name "(Default)" -Value "{00020424-0000-0000-C000-000000000046}"

New-Item -Path "$path2\TypeLib" -Force | Out-Null
Set-ItemProperty -Path "$path2\TypeLib" -Name "(Default)" -Value "{C866CA3A-32F7-11D2-9602-00C04F8EE628}"
Set-ItemProperty -Path "$path2\TypeLib" -Name "Version" -Value "5.4"

Write-Host "ISpObjectWithToken registered successfully" -ForegroundColor Green

Write-Host ""
Write-Host "✅ INTERFACE REGISTRATION COMPLETE!" -ForegroundColor Green
Write-Host "Both SAPI5 interfaces are now registered in the Windows registry." -ForegroundColor White
