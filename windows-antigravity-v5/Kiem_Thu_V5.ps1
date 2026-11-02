$ErrorActionPreference = 'Stop'
Push-Location $PSScriptRoot
try {
    & (Join-Path $PSScriptRoot 'Build.ps1')
    $csc = "$env:WINDIR/Microsoft.NET/Framework64/v4.0.30319/csc.exe"
    & $csc /nologo /target:exe /out:AllWorkshopGnvExportTests.exe /reference:CiscoLDTLApp_v5.exe /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll tests/AllWorkshopGnvExportTests.cs
    if ($LASTEXITCODE -ne 0) { throw 'All-workshop GNV test build failed' }
    & .\AllWorkshopGnvExportTests.exe (Join-Path $PSScriptRoot 'Data/03.9.2026_NT_T9_TCHC.xlsx')
    if ($LASTEXITCODE -ne 0) { throw 'All-workshop GNV tests failed' }
    & $csc /nologo /target:exe /out:SteelGnvExportTests.exe /reference:CiscoLDTLApp_v5.exe /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll tests/SteelGnvExportTests.cs
    if ($LASTEXITCODE -ne 0) { throw 'Steel GNV export test build failed' }
    & .\SteelGnvExportTests.exe (Join-Path $PSScriptRoot 'Data/03.9.2026_NT_T9_TCHC.xlsx')
    if ($LASTEXITCODE -ne 0) { throw 'Steel GNV export tests failed' }
    & $csc /nologo /target:exe /out:ClassificationAndSaveTests.exe /reference:CiscoLDTLApp_v5.exe /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll tests/ClassificationAndSaveTests.cs
    if ($LASTEXITCODE -ne 0) { throw 'Classification/save test build failed' }
    & .\ClassificationAndSaveTests.exe (Join-Path $PSScriptRoot 'Data/03.9.2026_NT_T9_TCHC.xlsx')
    if ($LASTEXITCODE -ne 0) { throw 'Classification/save tests failed' }
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
    & $csc /nologo /target:exe /out:GnvCalculationTests.exe /reference:CiscoLDTLApp_v5.exe /reference:System.Windows.Forms.dll tests/GnvCalculationTests.cs
    if ($LASTEXITCODE -ne 0) { throw 'GNV test build failed' }
    & .\GnvCalculationTests.exe --workbook
    if ($LASTEXITCODE -ne 0) { throw 'GNV tests failed' }
} finally { Pop-Location }
