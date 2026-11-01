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
    & $csc /nologo /target:exe /out:FrameworkCompatibilityTests.exe /reference:CiscoLDTLApp_v5.exe /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll tests/FrameworkCompatibilityTests.cs
    if ($LASTEXITCODE -ne 0) { throw 'Framework compatibility test build failed' }
    & .\FrameworkCompatibilityTests.exe
    if ($LASTEXITCODE -ne 0) { throw 'Framework compatibility tests failed' }
    & $csc /nologo /target:exe /out:ProductionPayrollTests.exe /reference:CiscoLDTLApp_v5.exe /reference:System.Windows.Forms.dll tests/ProductionPayrollTests.cs
    if ($LASTEXITCODE -ne 0) { throw 'Production payroll test build failed' }
    & .\ProductionPayrollTests.exe --workbook
    if ($LASTEXITCODE -ne 0) { throw 'Production payroll tests failed' }
    & $csc /nologo /target:exe /out:ManagementProductionAndWordTests.exe /reference:CiscoLDTLApp_v5.exe /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll tests/ManagementProductionAndWordTests.cs
    if ($LASTEXITCODE -ne 0) { throw 'Management production and Word test build failed' }
    & .\ManagementProductionAndWordTests.exe --workbook
    if ($LASTEXITCODE -ne 0) { throw 'Management production and Word tests failed' }
} finally { Pop-Location }
