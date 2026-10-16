#Requires -Version 5.1
# Exercises the packaged EXE and its real hidden PowerShell protocol without
# network calls or running any Office/OneDrive/Teams installer.
param([string]$ExecutablePath)
$ErrorActionPreference = 'Stop'
$executable = Join-Path (Split-Path $PSScriptRoot -Parent) 'Office Studio 1.3.0.exe'
if ($ExecutablePath) { $executable = [IO.Path]::GetFullPath($ExecutablePath) }
$assembly = [Reflection.Assembly]::LoadFile($executable)
$resource = $assembly.GetManifestResourceStream('Kareer.DeploymentWorker.ps1')
$memory = New-Object IO.MemoryStream
$resource.CopyTo($memory)
$embedded = $memory.ToArray()
$resource.Dispose()
$memory.Dispose()
$sha = [Security.Cryptography.SHA256]::Create()
$embeddedHash = [BitConverter]::ToString($sha.ComputeHash($embedded))
$localHash = [BitConverter]::ToString($sha.ComputeHash([IO.File]::ReadAllBytes((Join-Path $PSScriptRoot 'DeploymentWorker.ps1'))))
$sha.Dispose()
if ($embeddedHash -ne $localHash) { throw 'Embedded worker differs from tested source.' }
Write-Output 'PASS: Embedded worker matches final tested source'
$maintenanceResource = $assembly.GetManifestResourceStream('Kareer.OfficeMaintenance.ps1')
$maintenanceBuffer = New-Object IO.MemoryStream
$maintenanceResource.CopyTo($maintenanceBuffer)
$maintenanceResource.Dispose()
$maintenanceBytes = $maintenanceBuffer.ToArray()
$maintenanceBuffer.Dispose()
$maintenanceHasher = [Security.Cryptography.SHA256]::Create()
try {
    $packedHash = [BitConverter]::ToString($maintenanceHasher.ComputeHash($maintenanceBytes))
    $sourceHash = [BitConverter]::ToString($maintenanceHasher.ComputeHash([IO.File]::ReadAllBytes((Join-Path $PSScriptRoot 'OfficeMaintenance.ps1'))))
    if ($packedHash -ne $sourceHash) { throw 'Embedded maintenance worker differs from the tested source.' }
} finally { $maintenanceHasher.Dispose() }
Write-Output 'PASS: Embedded maintenance worker matches final tested source'

$scratch = Join-Path ([IO.Path]::GetTempPath()) ('OS & ' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($scratch)
$scriptPath = Join-Path $scratch 'embedded worker.ps1'
$planPath = Join-Path $scratch 'invalid plan.json'
$portable = Join-Path $scratch 'Office Studio 1.3.0.exe'
$testOutput = Join-Path $scratch 'Results'
try {
    [IO.File]::WriteAllBytes($scriptPath, $embedded)
    [IO.File]::WriteAllText($planPath, '{}')
    $info = New-Object Diagnostics.ProcessStartInfo
    $info.FileName = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $info.Arguments = '-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' + $scriptPath + '" -Action Validate -PlanPath "' + $planPath + '"'
    $info.UseShellExecute = $false
    $info.CreateNoWindow = $true
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    $child = [Diagnostics.Process]::Start($info)
    try {
        $stdout = $child.StandardOutput.ReadToEnd()
        $stderr = $child.StandardError.ReadToEnd()
        $child.WaitForExit()
        if ($child.ExitCode -ne 1 -or $stdout -notmatch '"kind":"error"' -or $stdout -notmatch 'missing PlanId') {
            throw ('Unexpected worker response: ' + $stdout + $stderr)
        }
    } finally { $child.Dispose() }
    Write-Output 'PASS: Hidden worker handles paths with spaces and ampersands and returns a structured validation error'
    [IO.File]::Delete($scriptPath)
    [IO.File]::Delete($planPath)
    [IO.File]::Copy($executable, $portable)
    [void][IO.Directory]::CreateDirectory($testOutput)
    $run = Start-Process -FilePath $portable -ArgumentList ('--self-test "' + $testOutput + '"') -WindowStyle Hidden -PassThru -Wait
    if ($run.ExitCode -ne 0) {
        $failureLog = Join-Path $testOutput 'ui-test-results.txt'
        if (Test-Path -LiteralPath $failureLog) { Get-Content -LiteralPath $failureLog | Write-Output }
        throw 'Standalone EXE test failed.'
    }
    Write-Output 'PASS: EXE runs by itself outside the source folder'
} finally {
    if ([IO.File]::Exists((Join-Path $testOutput '03-downloading-compact.png'))) { [IO.File]::Delete((Join-Path $testOutput '03-downloading-compact.png')) }
    # Remove only exact test files. Empty-only directory deletion preserves
    # anything unexpected; no recursive cleanup is used.
    foreach ($path in @($scriptPath, $planPath, $portable)) { [IO.File]::Delete($path) }
    foreach ($name in @('01-start.png','02-all-apps.png','03-downloading.png','04-ready.png','05-compact.png','06-installed-office.png','10-checking-files.png','ui-test-results.txt')) {
        [IO.File]::Delete((Join-Path $testOutput $name))
    }
    if ([IO.Directory]::Exists($testOutput)) { [IO.Directory]::Delete($testOutput, $false) }
    [IO.Directory]::Delete($scratch, $false)
}
