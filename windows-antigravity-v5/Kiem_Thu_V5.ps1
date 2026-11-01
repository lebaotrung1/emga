$ErrorActionPreference = 'Stop'
Push-Location $PSScriptRoot
try {
    & (Join-Path $PSScriptRoot 'Build.ps1')
    $csc = "$env:WINDIR/Microsoft.NET/Framework64/v4.0.30319/csc.exe"
    & $csc /nologo /target:exe /out:ManagementPayrollTests.exe /reference:CiscoLDTLApp_v5.exe tests/ManagementPayrollTests.cs
    if ($LASTEXITCODE -ne 0) { throw 'Payroll test build failed' }
    & .\ManagementPayrollTests.exe
    if ($LASTEXITCODE -ne 0) { throw 'Payroll tests failed' }
    & $csc /nologo /target:exe /out:ImportTests_v5.exe /reference:CiscoLDTLApp_v5.exe /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll tests/ImportTests.cs
    if ($LASTEXITCODE -ne 0) { throw 'Import test build failed' }
    & .\ImportTests_v5.exe
    if ($LASTEXITCODE -ne 0) { throw 'Import tests failed' }
} finally { Pop-Location }
