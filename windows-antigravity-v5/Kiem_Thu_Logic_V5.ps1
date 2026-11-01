$ErrorActionPreference = 'Stop'
Push-Location $PSScriptRoot
try {
    & (Join-Path $PSScriptRoot 'Build.ps1')
    $csc = "$env:WINDIR/Microsoft.NET/Framework64/v4.0.30319/csc.exe"
    & $csc /nologo /target:exe /out:ManagementPayrollTests.exe /reference:CiscoLDTLApp_v5.exe tests/ManagementPayrollTests.cs
    if ($LASTEXITCODE -ne 0) { throw 'Test build failed' }
    & .\ManagementPayrollTests.exe --logic-only
    if ($LASTEXITCODE -ne 0) { throw 'Tests failed' }
    & $csc /nologo /target:exe /out:FrameworkCompatibilityTests.exe /reference:CiscoLDTLApp_v5.exe /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll tests/FrameworkCompatibilityTests.cs
    if ($LASTEXITCODE -ne 0) { throw 'Framework compatibility test build failed' }
    & .\FrameworkCompatibilityTests.exe
    if ($LASTEXITCODE -ne 0) { throw 'Framework compatibility tests failed' }
} finally { Pop-Location }
