$ErrorActionPreference = 'Stop'
$compilers = @("$env:WINDIR/Microsoft.NET/Framework64/v4.0.30319/csc.exe", "$env:WINDIR/Microsoft.NET/Framework/v4.0.30319/csc.exe")
$csc = $compilers | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $csc) { throw 'Không tìm thấy trình biên dịch .NET Framework. Cần .NET Framework 4.5 trở lên.' }
$files = @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'src') -Filter *.cs | ForEach-Object { $_.FullName })
& $csc /nologo /target:winexe "/out:$PSScriptRoot/NghiemThuStudio.exe" /reference:System.Core.dll /reference:System.Xml.dll /reference:System.Runtime.Serialization.dll /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll /reference:System.Windows.Forms.dll /reference:System.Drawing.dll $files
if ($LASTEXITCODE -ne 0) { throw 'Biên dịch không thành công.' }
