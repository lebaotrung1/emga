param([switch]$Ui)
$ErrorActionPreference = 'Stop'
Push-Location $PSScriptRoot
try {
    & (Join-Path $PSScriptRoot 'Build.ps1')
    $csc = @("$env:WINDIR/Microsoft.NET/Framework64/v4.0.30319/csc.exe", "$env:WINDIR/Microsoft.NET/Framework/v4.0.30319/csc.exe") | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    & $csc /nologo /target:exe /out:StudioTests.exe /reference:NghiemThuStudio.exe /reference:System.Core.dll /reference:System.Xml.dll /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll tests/StudioTests.cs
    if ($LASTEXITCODE -ne 0) { throw 'Test build failed' }
    & .\StudioTests.exe
    if ($LASTEXITCODE -ne 0) { throw 'Logic/workflow tests failed' }
    if ($Ui) {
        & $csc /nologo /target:exe /main:StudioUiTests /out:StudioUiTests.exe /reference:NghiemThuStudio.exe /reference:System.Core.dll /reference:System.Xml.dll /reference:System.Windows.Forms.dll /reference:System.Drawing.dll /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll tests/StudioTests.cs tests/StudioUiTests.cs
        if ($LASTEXITCODE -ne 0) { throw 'UI test build failed' }
        & .\StudioUiTests.exe
        if ($LASTEXITCODE -ne 0) { throw 'UI tests failed' }
    }
} finally { Pop-Location }
