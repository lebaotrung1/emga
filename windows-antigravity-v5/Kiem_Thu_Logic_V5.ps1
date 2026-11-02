$ErrorActionPreference = 'Stop'
Push-Location $PSScriptRoot
try {
    & (Join-Path $PSScriptRoot 'Build.ps1')
    $csc = "$env:WINDIR/Microsoft.NET/Framework64/v4.0.30319/csc.exe"
    & $csc /nologo /target:exe /out:TimesheetLayoutTests.exe /reference:CiscoLDTLApp_v5.exe /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll tests/TimesheetLayoutTests.cs
    if ($LASTEXITCODE -ne 0) { throw 'Timesheet layout test build failed' }
    & .\TimesheetLayoutTests.exe
    if ($LASTEXITCODE -ne 0) { throw 'Timesheet layout tests failed' }
    & $csc /nologo /target:exe /out:ImportNameMatchingTests.exe /reference:CiscoLDTLApp_v5.exe /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll tests/ImportNameMatchingTests.cs
    if ($LASTEXITCODE -ne 0) { throw 'Name matching test build failed' }
    & .\ImportNameMatchingTests.exe
    if ($LASTEXITCODE -ne 0) { throw 'Name matching tests failed' }
    & $csc /nologo /target:exe /out:AllWorkshopGnvExportTests.exe /reference:CiscoLDTLApp_v5.exe /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll tests/AllWorkshopGnvExportTests.cs
    if ($LASTEXITCODE -ne 0) { throw 'All-workshop GNV test build failed' }
    & .\AllWorkshopGnvExportTests.exe
    if ($LASTEXITCODE -ne 0) { throw 'All-workshop GNV tests failed' }
    & $csc /nologo /target:exe /out:SteelGnvExportTests.exe /reference:CiscoLDTLApp_v5.exe /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll tests/SteelGnvExportTests.cs
    if ($LASTEXITCODE -ne 0) { throw 'Steel GNV export test build failed' }
    & .\SteelGnvExportTests.exe
    if ($LASTEXITCODE -ne 0) { throw 'Steel GNV export tests failed' }
    & $csc /nologo /target:exe /out:ClassificationAndSaveTests.exe /reference:CiscoLDTLApp_v5.exe /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll tests/ClassificationAndSaveTests.cs
    if ($LASTEXITCODE -ne 0) { throw 'Classification/save test build failed' }
    & .\ClassificationAndSaveTests.exe
    if ($LASTEXITCODE -ne 0) { throw 'Classification/save tests failed' }
    & $csc /nologo /target:exe /out:ManagementPayrollTests.exe /reference:CiscoLDTLApp_v5.exe tests/ManagementPayrollTests.cs
    if ($LASTEXITCODE -ne 0) { throw 'Test build failed' }
    & .\ManagementPayrollTests.exe --logic-only
    if ($LASTEXITCODE -ne 0) { throw 'Tests failed' }
    & $csc /nologo /target:exe /out:FrameworkCompatibilityTests.exe /reference:CiscoLDTLApp_v5.exe /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll tests/FrameworkCompatibilityTests.cs
    if ($LASTEXITCODE -ne 0) { throw 'Framework compatibility test build failed' }
    & .\FrameworkCompatibilityTests.exe
    if ($LASTEXITCODE -ne 0) { throw 'Framework compatibility tests failed' }
    & $csc /nologo /target:exe /out:ProductionPayrollTests.exe /reference:CiscoLDTLApp_v5.exe /reference:System.Windows.Forms.dll tests/ProductionPayrollTests.cs
    if ($LASTEXITCODE -ne 0) { throw 'Production payroll test build failed' }
    & .\ProductionPayrollTests.exe
    if ($LASTEXITCODE -ne 0) { throw 'Production payroll tests failed' }
    & $csc /nologo /target:exe /out:ManagementProductionAndWordTests.exe /reference:CiscoLDTLApp_v5.exe /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll tests/ManagementProductionAndWordTests.cs
    if ($LASTEXITCODE -ne 0) { throw 'Management production and Word test build failed' }
    & .\ManagementProductionAndWordTests.exe
    if ($LASTEXITCODE -ne 0) { throw 'Management production and Word tests failed' }
    & $csc /nologo /target:exe /out:GnvCalculationTests.exe /reference:CiscoLDTLApp_v5.exe /reference:System.Windows.Forms.dll tests/GnvCalculationTests.cs
    if ($LASTEXITCODE -ne 0) { throw 'GNV test build failed' }
    & .\GnvCalculationTests.exe
    if ($LASTEXITCODE -ne 0) { throw 'GNV tests failed' }
    & $csc /nologo /target:exe /out:CranePayrollTests.exe /reference:CiscoLDTLApp_v5.exe /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll /reference:System.Windows.Forms.dll tests/CranePayrollTests.cs
    if ($LASTEXITCODE -ne 0) { throw 'Crane payroll test build failed' }
    & .\CranePayrollTests.exe
    if ($LASTEXITCODE -ne 0) { throw 'Crane payroll tests failed' }
    & $csc /nologo /target:exe /out:CdnlPayrollTests.exe /reference:CiscoLDTLApp_v5.exe /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll /reference:System.Windows.Forms.dll tests/CdnlPayrollTests.cs
    if ($LASTEXITCODE -ne 0) { throw 'CDNL test build failed' }
    & .\CdnlPayrollTests.exe
    if ($LASTEXITCODE -ne 0) { throw 'CDNL payroll tests failed' }
} finally { Pop-Location }
