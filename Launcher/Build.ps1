#Requires -Version 5.1
[CmdletBinding()]
param([switch]$Test, [string]$OutputPath)
$ErrorActionPreference = 'Stop'
$output = Join-Path (Split-Path $PSScriptRoot -Parent) 'Office Studio 1.2.2.exe'
if ($OutputPath) { $output = [IO.Path]::GetFullPath($OutputPath) }
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path -LiteralPath $compiler)) { $compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe' }
$framework = Join-Path ${env:ProgramFiles(x86)} 'Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8'
if (-not (Test-Path -LiteralPath $framework)) { throw 'Building requires the .NET Framework 4.8 Developer Pack. Running the EXE only needs .NET Framework 4.8.' }

# Original, code-drawn KAREER app icon; no external asset dependencies.
Add-Type -AssemblyName System.Drawing
$bitmap = New-Object Drawing.Bitmap(64, 64)
$graphics = [Drawing.Graphics]::FromImage($bitmap)
$graphics.SmoothingMode = [Drawing.Drawing2D.SmoothingMode]::AntiAlias
$graphics.Clear([Drawing.Color]::FromArgb(18, 42, 49))
$brush = New-Object Drawing.SolidBrush([Drawing.Color]::FromArgb(125, 226, 187))
$font = New-Object Drawing.Font('Segoe UI', 38, [Drawing.FontStyle]::Bold, [Drawing.GraphicsUnit]::Pixel)
$format = New-Object Drawing.StringFormat
$format.Alignment = [Drawing.StringAlignment]::Center
$format.LineAlignment = [Drawing.StringAlignment]::Center
$graphics.DrawString('K', $font, $brush, [Drawing.RectangleF]::new(0, -1, 64, 64), $format)
$iconPath = Join-Path $PSScriptRoot 'Kareer.ico'
$icon = [Drawing.Icon]::FromHandle($bitmap.GetHicon())
$stream = [IO.File]::Create($iconPath)
try { $icon.Save($stream) } finally { $stream.Dispose(); $icon.Dispose(); $graphics.Dispose(); $bitmap.Dispose(); $brush.Dispose(); $font.Dispose(); $format.Dispose() }

$compileArgs = @('/nologo', '/noconfig', '/target:winexe', '/platform:anycpu', '/optimize+', '/utf8output', '/nostdlib+', ('/out:' + $output), ('/win32manifest:' + (Join-Path $PSScriptRoot 'app.manifest')), ('/win32icon:' + $iconPath))
foreach ($assembly in @('mscorlib', 'System', 'System.Core', 'System.Xml', 'System.Xaml', 'System.Web.Extensions', 'System.Drawing', 'System.Windows.Forms', 'WindowsBase', 'PresentationCore', 'PresentationFramework', 'UIAutomationTypes')) {
    $compileArgs += '/reference:' + (Join-Path $framework ($assembly + '.dll'))
}
$compileArgs += '/resource:' + (Join-Path $PSScriptRoot 'MainWindow.xaml') + ',Kareer.MainWindow.xaml'
$compileArgs += '/resource:' + (Join-Path $PSScriptRoot 'DeploymentWorker.ps1') + ',Kareer.DeploymentWorker.ps1'
$compileArgs += '/resource:' + (Join-Path $PSScriptRoot 'OfficeMaintenance.ps1') + ',Kareer.OfficeMaintenance.ps1'
$compileArgs += '/resource:' + (Join-Path $PSScriptRoot 'MaintenanceWindow.xaml') + ',Kareer.MaintenanceWindow.xaml'
$compileArgs += '/resource:' + $iconPath + ',Kareer.ico'
$compileArgs += Join-Path $PSScriptRoot 'OfficeStudio.cs'
$compileArgs += Join-Path $PSScriptRoot 'MaintenanceUi.cs'
$compileArgs += Join-Path $PSScriptRoot 'DownloadJob.cs'
& $compiler $compileArgs
if ($LASTEXITCODE -ne 0) { throw "C# compilation failed with exit code $LASTEXITCODE." }
Write-Host "Built: $output"
if ($Test) {
    $testOutput = Join-Path $PSScriptRoot 'Verification'
    [void][IO.Directory]::CreateDirectory($testOutput)
    $run = Start-Process -FilePath $output -ArgumentList ('--self-test "' + $testOutput + '"') -WindowStyle Hidden -PassThru -Wait
    Get-Content -LiteralPath (Join-Path $testOutput 'ui-test-results.txt')
    if ($run.ExitCode -ne 0) { throw 'UI self-test failed.' }
    $startup = Start-Process -FilePath $output -ArgumentList ('--startup-test "' + $testOutput + '"') -WindowStyle Hidden -PassThru -Wait
    Get-Content -LiteralPath (Join-Path $testOutput 'startup-test-results.txt')
    if ($startup.ExitCode -ne 0) { throw 'Asynchronous startup test failed.' }
    & (Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe') -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-Deployment.ps1')
    if ($LASTEXITCODE -ne 0) { throw 'Deployment tests failed.' }
    & (Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe') -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-Maintenance.ps1')
    if ($LASTEXITCODE -ne 0) { throw 'Maintenance tests failed.' }
}
Get-FileHash -LiteralPath $output -Algorithm SHA256
