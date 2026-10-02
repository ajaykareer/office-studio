#requires -Version 5.1
<# KAREER Office Studio: read-only discovery and explicitly confirmed Office removal.
   -Library only defines functions, for offline tests. No uninstall runs on import. #>
[CmdletBinding()]
param(
    [ValidateSet('Inventory', 'Uninstall', 'ToolTerms')][string]$Action,
    [string]$RequestPath,
    [string]$WorkerPath = '',
    [switch]$Library
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
if (-not $WorkerPath) { $WorkerPath = Join-Path $PSScriptRoot 'DeploymentWorker.ps1' }
$script:MaintenanceWorkerPath = $WorkerPath
$script:MaintenanceLogPath = ''
$script:MaintenanceRebootRequired = $false

function Get-MaintenanceValue {
    param($Object, [string]$Name, $Default = '')
    if ($null -eq $Object) { return $Default }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
    } elseif ($null -ne $Object.PSObject.Properties[$Name]) { return $Object.$Name }
    return $Default
}

function Write-MaintenanceEvent {
    param([string]$Kind, [string]$Message, [int]$Percent = -2, $Inventory = $null, [string]$Text = '')
    $event = [ordered]@{ kind = $Kind; message = $Message }
    if ($Kind -in @('complete','error')) { $event.rebootRequired = $script:MaintenanceRebootRequired }
    if ($Percent -ne -2) { $event.percent = $Percent }
    if ($null -ne $Inventory) { $event.inventory = $Inventory }
    if ($Text) { $event.text = $Text }
    $json = ConvertTo-Json -InputObject $event -Compress -Depth 12
    [Console]::Out.WriteLine('@@' + $json)
    [Console]::Out.Flush()
    if ($script:MaintenanceLogPath) {
        [IO.File]::AppendAllText($script:MaintenanceLogPath, ([DateTime]::UtcNow.ToString('o') + ' ' + $json + [Environment]::NewLine), (New-Object Text.UTF8Encoding($false)))
    }
}

function New-MaintenanceItem {
    param([string]$Id, [string]$Name, [string]$Version, [string]$Architecture,
          [ValidateSet('ClickToRun','MSI','Store','Unsupported')][string]$Kind,
          [string[]]$ProductIds = @(), [bool]$Removable = $false, [string]$Details = '')
    [pscustomobject][ordered]@{
        Id = $Id; Name = $Name; Version = $Version; Architecture = $Architecture
        Kind = $Kind; ProductIds = @($ProductIds | Sort-Object -Unique)
        Removable = $Removable; Details = $Details
    }
}

function Get-MaintenanceFingerprint {
    param([object[]]$Items = @(), [string[]]$Warnings = @())
    $canonical = [ordered]@{ Items = @($Items | Sort-Object Id); Warnings = @($Warnings | Sort-Object -Unique) }
    $bytes = [Text.Encoding]::UTF8.GetBytes((ConvertTo-Json -InputObject $canonical -Compress -Depth 12))
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '').ToLowerInvariant() }
    finally { $sha.Dispose() }
}

function Test-OfficeRegistryName {
    param([string]$Name)
    # Match actual Office suites/apps. Exclude unrelated add-ins, runtimes and updates.
    if ($Name -match '(?i)(?:\bupdate\b|\bhotfix\b|\badd[ -]?in\b|\bruntime\b|\bredistributable\b|\bdatabase engine\b|\bproofing tools\b|\bfile validation\b|\bconnectivity\b|\bclick-to-run extensibility\b|\bweb components\b)') { return $false }
    return $Name -match '(?i)^Microsoft\s+(?:Office(?:\s|$)|365(?:\s|$)|Project(?:\s|$)|Visio(?:\s|$)|Word(?:\s|$)|Excel(?:\s|$)|PowerPoint(?:\s|$)|Outlook(?:\s|$)|OneNote(?:\s|$)|Access(?:\s|$)|Publisher(?:\s|$)|InfoPath(?:\s|$)|Lync(?:\s|$)|Skype for Business(?:\s|$))'
}

function ConvertTo-OfficeInventory {
    param([object[]]$RegistryEntries = @(), [object[]]$ClickToRunEntries = @(),
          [object[]]$StorePackages = @(), [string[]]$Warnings = @())
    $items = New-Object 'System.Collections.Generic.List[object]'
    $scanWarnings = New-Object 'System.Collections.Generic.List[string]'
    foreach ($warning in $Warnings) { if ($warning) { $scanWarnings.Add($warning) } }
    $c2rProductIds = @(); $c2rVersions = @(); $c2rArchitectures = @(); $hasC2R = $false; $hasVerifiedC2R = $false
    $authoritativeC2R = @($ClickToRunEntries | Where-Object { (Get-MaintenanceValue $_ 'OfficeProductReleaseIds') -and (Get-MaintenanceValue $_ 'OfficePackageVersion') })
    $c2rSources = $ClickToRunEntries
    if ($authoritativeC2R.Count -gt 0) { $c2rSources = $authoritativeC2R }
    foreach ($entry in $c2rSources) {
        $ids = [string](Get-MaintenanceValue $entry 'OfficeProductReleaseIds')
        $version = [string](Get-MaintenanceValue $entry 'OfficePackageVersion')
        if ($ids -and $version) { $hasVerifiedC2R = $true }
        if (-not $ids) { $ids = [string](Get-MaintenanceValue $entry 'ProductReleaseIds') }
        if (-not $version) { $version = [string](Get-MaintenanceValue $entry 'VersionToReport') }
        if (-not $version) { $version = [string](Get-MaintenanceValue $entry 'ClientVersionToReport') }
        if ($ids) {
            $hasC2R = $true
            $c2rProductIds += @($ids -split '[,;\s]+' | Where-Object { $_ })
            if ($version) { $c2rVersions += $version }
            $architecture = [string](Get-MaintenanceValue $entry 'Platform')
            if ($architecture) { $c2rArchitectures += $architecture }
        }
    }
    $c2rArp = @($RegistryEntries | Where-Object {
        (Test-OfficeRegistryName ([string](Get-MaintenanceValue $_ 'DisplayName'))) -and
        ([string](Get-MaintenanceValue $_ 'Publisher') -match '(?i)^Microsoft(?: Corporation)?$') -and
        ([string](Get-MaintenanceValue $_ 'UninstallString') -match '(?i)OfficeClickToRun\.exe')
    })
    if ($c2rArp.Count -gt 0) { $hasC2R = $true }
    if ($hasC2R) {
        $details = 'Removes every Click-to-Run Office product, including installed Project, Visio and languages, for this PC.'
        if (-not $hasVerifiedC2R) { $details += ' Legacy registration detected; build information is not verified by the Office inventory key.' }
        $name = 'Microsoft Office Click-to-Run'
        if ($c2rProductIds.Count -gt 0) { $name += ' (' + (($c2rProductIds | Sort-Object -Unique) -join ', ') + ')' }
        elseif ($c2rArp.Count -gt 0) { $name = (@($c2rArp | ForEach-Object { [string](Get-MaintenanceValue $_ 'DisplayName') } | Sort-Object -Unique) -join '; ') }
        if ($c2rVersions.Count -eq 0) { $c2rVersions = @($c2rArp | ForEach-Object { [string](Get-MaintenanceValue $_ 'DisplayVersion') } | Where-Object { $_ }) }
        $items.Add((New-MaintenanceItem -Id 'c2r' -Name $name -Version (($c2rVersions | Sort-Object -Unique) -join ', ') -Architecture (($c2rArchitectures | Sort-Object -Unique) -join ', ') -Kind ClickToRun -ProductIds @($c2rProductIds) -Removable $true -Details $details))
    } elseif ($ClickToRunEntries.Count -gt 0) {
        $items.Add((New-MaintenanceItem -Id 'c2r-remnants' -Name 'Office Click-to-Run registration remnants' -Version '' -Architecture '' -Kind Unsupported -Details 'Click-to-Run configuration exists without registered products. Use Microsoft deep cleanup or review Apps & features.'))
    }
    $seen = @{}
    # Prefer machine/64-bit registration for duplicate product codes and keep
    # the result stable regardless of registry enumeration order.
    $sortedRegistry = @($RegistryEntries | Sort-Object -Property { [string](Get-MaintenanceValue $_ 'KeyName') }, { [string](Get-MaintenanceValue $_ 'Hive') }, { [string](Get-MaintenanceValue $_ 'View') } -Descending)
    foreach ($entry in $sortedRegistry) {
        $name = [string](Get-MaintenanceValue $entry 'DisplayName')
        if (-not (Test-OfficeRegistryName $name)) { continue }
        if ([string](Get-MaintenanceValue $entry 'Publisher') -notmatch '(?i)^Microsoft(?: Corporation)?$') { continue }
        if ([string](Get-MaintenanceValue $entry 'UninstallString') -match '(?i)OfficeClickToRun\.exe') { continue }
        $key = [string](Get-MaintenanceValue $entry 'KeyName')
        $version = [string](Get-MaintenanceValue $entry 'DisplayVersion')
        $view = [string](Get-MaintenanceValue $entry 'View')
        $hive = [string](Get-MaintenanceValue $entry 'Hive')
        $architecture = ''; if ($view -eq 'Registry64') { $architecture = 'x64' } elseif ($view -eq 'Registry32') { $architecture = 'x86' }
        $isMsi = ([string](Get-MaintenanceValue $entry 'WindowsInstaller') -eq '1') -and ($key -match '^\{[A-Fa-f0-9]{8}-[A-Fa-f0-9]{4}-[A-Fa-f0-9]{4}-[A-Fa-f0-9]{4}-[A-Fa-f0-9]{12}\}$')
        if ($isMsi) {
            $id = 'msi:' + $key.ToUpperInvariant()
            if ($seen.ContainsKey($id)) { continue }; $seen[$id] = $true
            $items.Add((New-MaintenanceItem -Id $id -Name $name -Version $version -Architecture $architecture -Kind MSI -ProductIds @($key.ToUpperInvariant()) -Removable $true -Details "Registered Windows Installer Office product ($hive). Shared Office components may remain while other Office products use them."))
        } else {
            $id = 'unsupported:' + $hive + ':' + $view + ':' + $key
            if ($seen.ContainsKey($id)) { continue }; $seen[$id] = $true
            $items.Add((New-MaintenanceItem -Id $id -Name $name -Version $version -Architecture $architecture -Kind Unsupported -Details 'This Office registration cannot be safely removed with standard removal. Microsoft deep cleanup may support it.'))
        }
    }
    foreach ($provisioned in @($false, $true)) {
        $packages = @($StorePackages | Where-Object { ([string](Get-MaintenanceValue $_ 'Name') -ceq 'Microsoft.Office.Desktop') -and ([bool](Get-MaintenanceValue $_ 'IsProvisioned' $false) -eq $provisioned) })
        if ($packages.Count -eq 0) { continue }
        $ids = @($packages | ForEach-Object { [string](Get-MaintenanceValue $_ 'PackageFullName') } | Sort-Object -Unique)
        $valid = @($ids | Where-Object { $_ -notmatch '^Microsoft\.Office\.Desktop_[A-Za-z0-9._~\-]+$' }).Count -eq 0
        $id = 'store-installed'; $name = 'Microsoft Office Desktop (Microsoft Store)'; $details = 'Removes the exact Office Desktop package for all Windows users.'
        if ($provisioned) { $id = 'store-provisioned'; $name += ' - provisioned'; $details = 'Removes provisioning of Office Desktop so new Windows users do not automatically receive it.' }
        $kind = 'Store'; if (-not $valid) { $kind = 'Unsupported'; $details = 'The Store Office package identity is incomplete or invalid.' }
        $items.Add((New-MaintenanceItem -Id $id -Name $name -Version ((@($packages | ForEach-Object { [string](Get-MaintenanceValue $_ 'Version') } | Sort-Object -Unique)) -join ', ') -Architecture ((@($packages | ForEach-Object { [string](Get-MaintenanceValue $_ 'Architecture') } | Sort-Object -Unique)) -join ', ') -Kind $kind -ProductIds $ids -Removable $valid -Details $details))
    }
    $orderedItems = @($items.ToArray() | Sort-Object Id)
    $orderedWarnings = @($scanWarnings.ToArray() | Sort-Object -Unique)
    return [pscustomobject][ordered]@{ Items = $orderedItems; Warnings = $orderedWarnings; Fingerprint = Get-MaintenanceFingerprint -Items $orderedItems -Warnings $orderedWarnings }
}

function Get-MaintenanceProvisionedOfficePackages {
    try {
        return @(Get-AppxProvisionedPackage -Online -ErrorAction Stop | Where-Object { $_.DisplayName -ceq 'Microsoft.Office.Desktop' })
    } catch {
        $dismFailure = $_.Exception.Message.Trim()
        # The DISM provider can fail while opening its servicing registry keys
        # even for an elevated caller. Windows 10 2004+ also exposes an official
        # read-only package API which does not initialize the DISM provider.
        try {
            $null = [Windows.Management.Deployment.PackageManager,Windows.Management.Deployment,ContentType=WindowsRuntime]
            $manager = New-Object Windows.Management.Deployment.PackageManager
            foreach ($package in @($manager.FindProvisionedPackages())) {
                if ([string]$package.Id.Name -ceq 'Microsoft.Office.Desktop') {
                    $version = $package.Id.Version
                    [pscustomobject]@{
                        DisplayName = [string]$package.Id.Name; PackageName = [string]$package.Id.FullName
                        Version = ('{0}.{1}.{2}.{3}' -f $version.Major, $version.Minor, $version.Build, $version.Revision)
                        Architecture = [string]$package.Id.Architecture
                    }
                }
            }
        } catch { throw ('DISM inventory: ' + $dismFailure + ' Windows package inventory: ' + $_.Exception.Message.Trim()) }
    }
}

function Test-MaintenanceAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    try {
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } finally { $identity.Dispose() }
}

function Get-OfficeInventory {
    $entries = New-Object 'System.Collections.Generic.List[object]'
    $c2r = New-Object 'System.Collections.Generic.List[object]'
    $stores = New-Object 'System.Collections.Generic.List[object]'
    $warnings = New-Object 'System.Collections.Generic.List[string]'
    if (-not (Test-MaintenanceAdministrator)) { $warnings.Add('Office Studio is not running as administrator. Reopen it as administrator to complete the all-users Office scan.') }
    $views = @([Microsoft.Win32.RegistryView]::Registry32)
    if ([Environment]::Is64BitOperatingSystem) { $views = @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32) }
    foreach ($hive in @([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryHive]::CurrentUser)) {
        foreach ($view in $views) {
            $base = $null; $uninstall = $null
            try {
                $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hive, $view)
                $uninstall = $base.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall', $false)
                if ($null -ne $uninstall) {
                    foreach ($keyName in $uninstall.GetSubKeyNames()) {
                        $key = $null
                        try {
                            $key = $uninstall.OpenSubKey($keyName, $false)
                            if ($null -eq $key) { continue }
                            $entries.Add([pscustomobject]@{ Hive = $hive.ToString(); View = $view.ToString(); KeyName = $keyName; DisplayName = [string]$key.GetValue('DisplayName', ''); DisplayVersion = [string]$key.GetValue('DisplayVersion', ''); Publisher = [string]$key.GetValue('Publisher', ''); WindowsInstaller = $key.GetValue('WindowsInstaller', 0); UninstallString = [string]$key.GetValue('UninstallString', '') })
                        } catch { $warnings.Add("Could not read $hive/$view uninstall entry $keyName`: $($_.Exception.Message)") }
                        finally { if ($null -ne $key) { $key.Dispose() } }
                    }
                }
                $configuration = $null; $inventory = $null
                try {
                    $configuration = $base.OpenSubKey('SOFTWARE\Microsoft\Office\ClickToRun\Configuration', $false)
                    $inventory = $base.OpenSubKey('SOFTWARE\Microsoft\Office\ClickToRun\Inventory\Office\16.0', $false)
                    if ($null -ne $configuration -or $null -ne $inventory) {
                        $record = [ordered]@{ Hive = $hive.ToString(); View = $view.ToString(); OfficeProductReleaseIds = ''; OfficePackageVersion = ''; ProductReleaseIds = ''; VersionToReport = ''; ClientVersionToReport = ''; Platform = ''; InstallationPath = '' }
                        if ($null -ne $configuration) { foreach ($property in @('ProductReleaseIds','VersionToReport','ClientVersionToReport','Platform','InstallationPath')) { $record[$property] = [string]$configuration.GetValue($property, '') } }
                        if ($null -ne $inventory) { foreach ($property in @('OfficeProductReleaseIds','OfficePackageVersion')) { $record[$property] = [string]$inventory.GetValue($property, '') } }
                        $c2r.Add([pscustomobject]$record)
                    }
                } catch { $warnings.Add("Could not read $hive/$view Click-to-Run inventory: $($_.Exception.Message)") }
                finally { if ($null -ne $configuration) { $configuration.Dispose() }; if ($null -ne $inventory) { $inventory.Dispose() } }
            } catch { $warnings.Add("Could not scan $hive/$view Office registrations: $($_.Exception.Message)") }
            finally { if ($null -ne $uninstall) { $uninstall.Dispose() }; if ($null -ne $base) { $base.Dispose() } }
        }
    }
    try {
        $bundles = @(Get-AppxPackage -AllUsers -Name 'Microsoft.Office.Desktop' -PackageTypeFilter Bundle -ErrorAction Stop)
        $mainPackages = @(Get-AppxPackage -AllUsers -Name 'Microsoft.Office.Desktop' -PackageTypeFilter Main -ErrorAction Stop)
        $bundleVersions = @($bundles | ForEach-Object { [string]$_.Version })
        # Remove the parent bundle where one exists, retaining any independent
        # main package from another version for the explicit all-users removal.
        $packages = @($bundles) + @($mainPackages | Where-Object { [string]$_.Version -notin $bundleVersions })
        foreach ($package in $packages) { $stores.Add([pscustomobject]@{ Name = [string]$package.Name; PackageFullName = [string]$package.PackageFullName; Version = [string]$package.Version; Architecture = [string]$package.Architecture; IsProvisioned = $false }) }
    } catch { $warnings.Add('Could not scan Store Office for all users: ' + $_.Exception.Message) }
    try {
        foreach ($package in @(Get-MaintenanceProvisionedOfficePackages)) {
            $stores.Add([pscustomobject]@{ Name = [string]$package.DisplayName; PackageFullName = [string]$package.PackageName; Version = [string]$package.Version; Architecture = [string]$package.Architecture; IsProvisioned = $true })
        }
    } catch { $warnings.Add('Could not scan provisioned Store Office: ' + $_.Exception.Message) }
    return ConvertTo-OfficeInventory -RegistryEntries $entries.ToArray() -ClickToRunEntries $c2r.ToArray() -StorePackages $stores.ToArray() -Warnings $warnings.ToArray()
}

function Assert-MaintenanceSafePath {
    param([string]$Path)
    if ($Path -notmatch '^[A-Za-z]:\\' -or $Path -match '["<>|?*\x00-\x1f]' -or $Path.Substring(2).Contains(':')) { throw 'Choose a full local Windows folder path for removal logs.' }
    $full = [IO.Path]::GetFullPath($Path).TrimEnd('\')
    if ($full -eq [IO.Path]::GetPathRoot($full).TrimEnd('\')) { throw 'Removal needs a dedicated folder, not a drive root.' }
    foreach ($segment in $full.Split('\')) { if ($segment.EndsWith('.') -or $segment.EndsWith(' ')) { throw 'Removal folder names cannot end with a period or space.' } }
    $cursor = $full
    while ($cursor) {
        if (Test-Path -LiteralPath $cursor) {
            $item = Get-Item -LiteralPath $cursor -Force
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'Removal folders cannot contain symbolic links or junctions.' }
        }
        $parent = [IO.Path]::GetDirectoryName($cursor.TrimEnd('\'))
        if (-not $parent -or $parent -eq $cursor) { break }; $cursor = $parent
    }
    if ((Test-Path -LiteralPath $full) -and -not (Test-Path -LiteralPath $full -PathType Container)) { throw 'The removal log destination is a file.' }
    return $full
}

function Test-MaintenanceRequest {
    param([Parameter(Mandatory = $true)]$Request)
    $confirmed = Get-MaintenanceValue $Request 'Confirmed' $false
    if ($confirmed -isnot [bool] -or -not $confirmed) { throw 'Office removal requires explicit confirmation in Office Studio.' }
    $fingerprint = [string](Get-MaintenanceValue $Request 'Fingerprint')
    if ($fingerprint -notmatch '^[a-fA-F0-9]{64}$') { throw 'Refresh the Office scan before confirming removal.' }
    $mode = [string](Get-MaintenanceValue $Request 'Mode')
    if ($mode -notin @('Standard','Deep')) { throw 'Choose Standard removal or Microsoft deep cleanup.' }
    $workPath = Assert-MaintenanceSafePath -Path ([string](Get-MaintenanceValue $Request 'WorkPath'))
    return [pscustomobject][ordered]@{ Fingerprint = $fingerprint.ToLowerInvariant(); Confirmed = $true; Mode = $mode; WorkPath = $workPath }
}

function New-OfficeRemovalConfiguration {
    return '<?xml version="1.0" encoding="utf-8"?><Configuration><Remove All="TRUE" /><Display Level="Full" AcceptEULA="TRUE" /><Property Name="FORCEAPPSHUTDOWN" Value="FALSE" /></Configuration>'
}

function Assert-MaintenanceAdministrator {
    if (-not (Test-MaintenanceAdministrator)) { throw 'Office removal requires administrator access. Reopen Office Studio as administrator.' }
}

function Assert-NoRunningOfficeApplications {
    param([object[]]$Processes)
    if (-not $PSBoundParameters.ContainsKey('Processes')) { $Processes = @(Get-Process -ErrorAction Stop) }
    $names = @('WINWORD','EXCEL','POWERPNT','OUTLOOK','ONENOTE','MSACCESS','MSPUB','VISIO','WINPROJ','LYNC','INFOPATH','GROOVE','MSTORE','SETLANG','MSOUC','OIS','GRAPH','TEAMS','OfficeC2RClient','OfficeClickToRun')
    # OfficeClickToRun is a normal background service and must not block removal.
    $running = @($Processes | ForEach-Object { [string](Get-MaintenanceValue $_ 'ProcessName') } | Where-Object { $_ -in $names -and $_ -ne 'OfficeClickToRun' } | Sort-Object -Unique)
    if ($running.Count -gt 0) { throw ('Save your work and close these Office applications before removing Office: ' + ($running -join ', ') + '. No applications were force-closed.') }
}

function Assert-MaintenanceMicrosoftSignature {
    param([string]$Path)
    . $script:MaintenanceWorkerPath -Library
    Assert-MicrosoftSignature -Path $Path
}

function Save-MaintenanceDeploymentTool {
    param([string]$Path)
    . $script:MaintenanceWorkerPath -Library
    Save-MicrosoftDownload -Url 'https://officecdn.microsoft.com/pr/wsus/setup.exe' -Path $Path -Label 'Microsoft Office Deployment Tool'
}

function Save-MaintenanceScrubTool {
    param([string]$Directory)
    # The official command-line toolkit expires; download a fresh copy each run.
    $archive = Join-Path $Directory 'Microsoft-GetHelp.zip'
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $uri = [Uri]'https://aka.ms/SaRA_EnterpriseVersionFiles'; $response = $null
    try {
        for ($redirect = 0; $redirect -lt 10; $redirect++) {
            if ($uri.Scheme -ne 'https' -or $uri.UserInfo -or ($uri.Host -ne 'aka.ms' -and $uri.Host -ne 'download.microsoft.com')) { throw 'Microsoft cleanup download redirected to an unexpected address.' }
            $request = [Net.HttpWebRequest]::Create($uri); $request.AllowAutoRedirect = $false; $request.Timeout = 60000; $request.ReadWriteTimeout = 60000
            $response = $request.GetResponse()
            if ([int]$response.StatusCode -in @(301,302,303,307,308)) {
                $location = $response.Headers['Location']; $response.Dispose(); $response = $null
                if (-not $location) { throw 'Microsoft cleanup download returned an empty redirect.' }; $uri = New-Object Uri($uri, $location)
            } else { break }
        }
        if ($null -eq $response -or [int]$response.StatusCode -ne 200 -or $uri.Host -ne 'download.microsoft.com') { throw 'The official Microsoft cleanup toolkit could not be downloaded.' }
        Write-MaintenanceEvent -Kind progress -Message 'Downloading the current Microsoft Get Help cleanup toolkit...' -Percent -1
        $inputStream = $null; $outputStream = $null
        try {
            $inputStream = $response.GetResponseStream(); $outputStream = [IO.File]::Open($archive, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
            $buffer = New-Object byte[] 262144; [long]$received = 0
            while (($read = $inputStream.Read($buffer, 0, $buffer.Length)) -gt 0) { $outputStream.Write($buffer, 0, $read); $received += $read; if ($received -gt 1GB) { throw 'Microsoft cleanup toolkit download was unexpectedly large.' } }
            if ($received -eq 0 -or ($response.ContentLength -gt 0 -and $received -ne $response.ContentLength)) { throw 'Microsoft cleanup toolkit download was incomplete.' }
        } finally { if ($null -ne $outputStream) { $outputStream.Dispose() }; if ($null -ne $inputStream) { $inputStream.Dispose() } }
    } finally { if ($null -ne $response) { $response.Dispose() } }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $extractPath = Join-Path $Directory 'Microsoft-GetHelp'; [void][IO.Directory]::CreateDirectory($extractPath)
    $rootPrefix = [IO.Path]::GetFullPath($extractPath).TrimEnd('\') + '\'
    $zip = [IO.Compression.ZipFile]::OpenRead($archive)
    try {
        [long]$expanded = 0
        foreach ($entry in $zip.Entries) {
            $entryName = $entry.FullName.Replace('/', '\')
            if ($entryName -match '[:\x00-\x1f]' -or [IO.Path]::IsPathRooted($entryName)) { throw 'Microsoft cleanup archive contains an invalid path.' }
            $target = [IO.Path]::GetFullPath((Join-Path $extractPath $entryName))
            if (-not $target.StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase)) { throw 'Microsoft cleanup archive contains a path outside its folder.' }
            if ((($entry.ExternalAttributes -shr 16) -band 0xF000) -eq 0xA000) { throw 'Microsoft cleanup archive contains an unsupported symbolic link.' }
            $expanded += $entry.Length; if ($expanded -gt 2GB) { throw 'Microsoft cleanup archive was unexpectedly large.' }
            if ($entryName.EndsWith('\')) { [void][IO.Directory]::CreateDirectory($target); continue }
            [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($target))
            [IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $target, $false)
        }
    } finally { $zip.Dispose() }
    $executables = @(Get-ChildItem -LiteralPath $extractPath -Filter 'GetHelpCmd.exe' -File -Recurse)
    if ($executables.Count -ne 1) { throw 'The Microsoft cleanup archive does not contain exactly one GetHelpCmd.exe.' }
    # Third-party dependencies shipped by Microsoft are retained exactly as
    # received from its HTTPS archive. The entry point we execute must carry a
    # valid Microsoft Authenticode signature and is checked again before launch.
    Assert-MaintenanceMicrosoftSignature -Path $executables[0].FullName
    Write-MaintenanceEvent -Kind log -Message 'The current cleanup toolkit was extracted from the official Microsoft HTTPS download. GetHelpCmd.exe has a valid Microsoft signature; bundled dependencies retain their official archive contents.'
    return $executables[0].FullName
}

function Invoke-MaintenanceProcess {
    param([string]$FilePath, [string]$Arguments, [string]$Label, [string]$WorkingDirectory, [int[]]$SuccessExitCodes = @(0,3010))
    Assert-MaintenanceMicrosoftSignature -Path $FilePath
    $info = New-Object Diagnostics.ProcessStartInfo
    $info.FileName = $FilePath; $info.Arguments = $Arguments; $info.WorkingDirectory = $WorkingDirectory
    $info.UseShellExecute = $false; $info.CreateNoWindow = $true; $info.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
    $info.RedirectStandardOutput = $true; $info.RedirectStandardError = $true
    $process = New-Object Diagnostics.Process; $process.StartInfo = $info
    try {
        Write-MaintenanceEvent -Kind progress -Message $Label -Percent -1
        if (-not $process.Start()) { throw 'The Microsoft removal tool did not start.' }
        $stdout = $process.StandardOutput.ReadToEndAsync(); $stderr = $process.StandardError.ReadToEndAsync(); $timer = [Diagnostics.Stopwatch]::StartNew()
        while (-not $process.WaitForExit(1000)) { if ($timer.ElapsedMilliseconds -ge 10000) { Write-MaintenanceEvent -Kind progress -Message "$Label Keep Office Studio open. Microsoft may show its own progress window." -Percent -1; $timer.Restart() } }
        $process.WaitForExit(); $output = [string]$stdout.Result; $errors = [string]$stderr.Result
        foreach ($line in @($output -split '\r?\n') + @($errors -split '\r?\n')) { if ($line.Trim()) { Write-MaintenanceEvent -Kind log -Message $line } }
        if ($process.ExitCode -notin $SuccessExitCodes) { throw "$Label failed (exit code $($process.ExitCode)). Review the activity log and Microsoft message." }
        if ($process.ExitCode -eq 3010) { $script:MaintenanceRebootRequired = $true }
        return [pscustomobject]@{ ExitCode = $process.ExitCode; Output = $output; ErrorOutput = $errors }
    } finally { $process.Dispose() }
}

function Test-MaintenanceScrubResult {
    param($Result)
    $output = [string](Get-MaintenanceValue $Result 'Output') + "`n" + [string](Get-MaintenanceValue $Result 'ErrorOutput')
    if ([int](Get-MaintenanceValue $Result 'ExitCode' -1) -ne 0) { throw 'Microsoft deep cleanup did not exit successfully. Review the activity log.' }
    if ($output -match '(?im)^\s*(?:0?6|0?8|0?9|10|66|67|68)\s*[:\-]') { throw 'Microsoft deep cleanup reported an error or could not confirm an Office removal. Review the activity log.' }
    if ($output -notmatch '(?im)^\s*00:\s*Successfully completed this scenario\.(?:\s|$)') { throw 'Microsoft deep cleanup finished without a recognized success result. Removal is unverified; review the activity log.' }
    $script:MaintenanceRebootRequired = $true
}

function Invoke-MaintenanceToolTerms {
    param([Parameter(Mandatory = $true)]$Request)
    $workPath = Assert-MaintenanceSafePath -Path ([string](Get-MaintenanceValue $Request 'WorkPath'))
    $directory = Join-Path $workPath ('KAREER-Terms-' + [Guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($directory)
    $tool = Save-MaintenanceScrubTool -Directory $directory
    $licensePath = Join-Path $directory 'Microsoft-GetHelp-License.txt'
    # Microsoft's DisplayEULA mode only exports the terms. No scenario or
    # acceptance switch is passed and no Office removal is requested.
    $null = Invoke-MaintenanceProcess -FilePath $tool -Arguments ('-DisplayEULA "' + $licensePath + '"') -Label 'Reading Microsoft cleanup tool license terms...' -WorkingDirectory ([IO.Path]::GetDirectoryName($tool)) -SuccessExitCodes @(0)
    if (-not (Test-Path -LiteralPath $licensePath -PathType Leaf)) { throw 'The Microsoft tool did not produce its license terms.' }
    $licenseFile = Get-Item -LiteralPath $licensePath -Force
    if (($licenseFile.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0 -or $licenseFile.Length -le 0 -or $licenseFile.Length -gt 1MB) { throw 'The Microsoft license terms file is invalid.' }
    $terms = [IO.File]::ReadAllText($licensePath)
    Write-MaintenanceEvent -Kind terms -Message 'Microsoft cleanup tool license terms' -Text $terms
    Write-MaintenanceEvent -Kind complete -Message 'Microsoft cleanup tool license terms are ready to review.' -Percent 100
}

function Remove-MaintenanceStorePackages {
    param([object[]]$Items)
    foreach ($item in @($Items | Where-Object { $_.Kind -eq 'Store' } | Sort-Object Id -Descending)) {
        foreach ($packageId in $item.ProductIds) {
            if ($packageId -notmatch '^Microsoft\.Office\.Desktop_[A-Za-z0-9._~\-]+$') { throw 'The Office Store package identity is invalid.' }
            Write-MaintenanceEvent -Kind progress -Message ('Removing ' + $item.Name + '...') -Percent -1
            if ($item.Id -eq 'store-provisioned') {
                $present = @(Get-MaintenanceProvisionedOfficePackages | Where-Object { $_.PackageName -ceq $packageId })
                if ($present.Count -gt 0) { $null = Remove-AppxProvisionedPackage -Online -PackageName $packageId -ErrorAction Stop }
            } elseif ($item.Id -eq 'store-installed') {
                $present = @(Get-AppxPackage -AllUsers -Name 'Microsoft.Office.Desktop' -PackageTypeFilter @('Main','Bundle') -ErrorAction Stop | Where-Object { $_.Name -ceq 'Microsoft.Office.Desktop' -and $_.PackageFullName -ceq $packageId })
                if ($present.Count -gt 0) { Remove-AppxPackage -AllUsers -Package $packageId -ErrorAction Stop }
            }
            else { throw 'The Office Store registration is unsupported.' }
        }
    }
}

function Invoke-OfficeRemoval {
    param([Parameter(Mandatory = $true)]$Request)
    $Request = Test-MaintenanceRequest -Request $Request
    Assert-MaintenanceAdministrator
    $script:MaintenanceRebootRequired = $false
    $mutex = $null; $ownsMutex = $false; $operationLock = $null; $removalStarted = $false; $finalInventoryEmitted = $false
    try {
        $mutex = New-Object Threading.Mutex($false, 'Global\KAREER.OfficeStudio.OfficeOperation')
        try { $ownsMutex = $mutex.WaitOne(0) } catch [Threading.AbandonedMutexException] { $ownsMutex = $true }
        if (-not $ownsMutex) { throw 'Another Office Studio installation or removal is running. Wait for it to finish.' }
        Write-MaintenanceEvent -Kind status -Message 'Checking the confirmed Office inventory again...'
        $before = Get-OfficeInventory
        Write-MaintenanceEvent -Kind inventory -Message 'Office inventory before removal.' -Inventory $before
        if ($before.Fingerprint -cne $Request.Fingerprint) { throw 'The installed Office inventory changed after confirmation. Refresh and review the removal list again.' }
        if (@($before.Warnings).Count -gt 0) { throw 'The Office scan is incomplete. Resolve the scan warnings before removing Office.' }
        if ($Request.Mode -eq 'Standard' -and @($before.Items | Where-Object { -not $_.Removable }).Count -gt 0) { throw 'Some detected Office registrations need Microsoft deep cleanup or manual removal. Standard removal has not started.' }
        if ($Request.Mode -eq 'Standard' -and @($before.Items).Count -eq 0) { throw 'No registered Office products were found to remove.' }
        Assert-NoRunningOfficeApplications
        [void][IO.Directory]::CreateDirectory($Request.WorkPath)
        $lockPath = Join-Path $Request.WorkPath '.kareer-maintenance.lock'
        $null = Assert-MaintenanceSafePath -Path $Request.WorkPath
        if ((Test-Path -LiteralPath $lockPath) -and ((Get-Item -LiteralPath $lockPath -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'The removal folder lock is a symbolic link.' }
        try { $operationLock = [IO.File]::Open($lockPath, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None) }
        catch [IO.IOException] { throw 'The removal log folder is unavailable or is already in use.' }
        $runPath = Join-Path $Request.WorkPath ('KAREER-Removal-' + [DateTime]::Now.ToString('yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($runPath)
        $script:MaintenanceLogPath = Join-Path $runPath 'removal-events.log'
        Write-MaintenanceEvent -Kind log -Message "Removal logs: $runPath"
        Write-MaintenanceEvent -Kind log -Message 'Personal documents are retained. Standalone OneDrive, new Teams and new Outlook are managed separately. Classic Office-bundled components may be removed. Office Studio will not restart Windows automatically.'
        if ($Request.Mode -eq 'Deep') {
            $tool = Save-MaintenanceScrubTool -Directory $runPath
            Assert-NoRunningOfficeApplications
            $removalStarted = $true
            $result = Invoke-MaintenanceProcess -FilePath $tool -Arguments '-S OfficeScrubScenario -AcceptEula -OfficeVersion All' -Label 'Microsoft is deeply removing all supported Office installations...' -WorkingDirectory ([IO.Path]::GetDirectoryName($tool)) -SuccessExitCodes @(0)
            Test-MaintenanceScrubResult -Result $result
        } else {
            if (@($before.Items | Where-Object { $_.Kind -eq 'ClickToRun' }).Count -gt 0) {
                $tool = Join-Path $runPath 'OfficeSetup.exe'
                Save-MaintenanceDeploymentTool -Path $tool
                $configuration = Join-Path $runPath 'Remove-Office.xml'
                [IO.File]::WriteAllText($configuration, (New-OfficeRemovalConfiguration), (New-Object Text.UTF8Encoding($false)))
                Assert-NoRunningOfficeApplications
                $removalStarted = $true
                $null = Invoke-MaintenanceProcess -FilePath $tool -Arguments ('/configure "' + $configuration + '"') -Label 'Removing all Click-to-Run Office products, Project, Visio and languages...' -WorkingDirectory $runPath
            }
            foreach ($item in @($before.Items | Where-Object { $_.Kind -eq 'MSI' })) {
                Assert-NoRunningOfficeApplications
                $productId = [string]$item.ProductIds[0]
                if ($productId -notmatch '^\{[A-Fa-f0-9]{8}-[A-Fa-f0-9]{4}-[A-Fa-f0-9]{4}-[A-Fa-f0-9]{4}-[A-Fa-f0-9]{12}\}$') { throw 'The registered Office MSI product code is invalid.' }
                $msiexec = Join-Path ([Environment]::GetFolderPath('Windows')) 'System32\msiexec.exe'
                if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) { $msiexec = Join-Path ([Environment]::GetFolderPath('Windows')) 'Sysnative\msiexec.exe' }
                $msiLog = Join-Path $runPath ('MSI-' + $productId.Trim('{}') + '.log')
                $removalStarted = $true
                $null = Invoke-MaintenanceProcess -FilePath $msiexec -Arguments ('/x ' + $productId + ' /passive /norestart /L*v "' + $msiLog + '"') -Label ('Removing ' + $item.Name + '...') -WorkingDirectory $runPath -SuccessExitCodes @(0,3010)
            }
        }
        # Store registration/provisioning is separate from C2R/MSI and can remain after scrub.
        if (@($before.Items | Where-Object { $_.Kind -eq 'Store' }).Count -gt 0) {
            Assert-NoRunningOfficeApplications
            $removalStarted = $true
            Remove-MaintenanceStorePackages -Items @($before.Items)
        }
        Write-MaintenanceEvent -Kind status -Message 'Checking which Office registrations remain...'
        $after = Get-OfficeInventory
        Write-MaintenanceEvent -Kind inventory -Message 'Office inventory after removal.' -Inventory $after
        $finalInventoryEmitted = $true
        if (@($after.Warnings).Count -gt 0) { throw 'The removal tools finished, but the final scan is incomplete. Removal cannot be verified. Review the scan warnings and logs.' }
        if (@($after.Items).Count -gt 0) {
            $message = 'Office registrations remain after removal. Review the remaining list and removal logs.'
            if ($script:MaintenanceRebootRequired) { $message = 'Microsoft requested a restart and Office registrations still remain. Restart Windows, then refresh the scan; removal is not yet verified.' }
            throw $message
        }
        $message = 'Office removal completed. No Office suite registrations were found in the final scan. Personal documents are retained. Standalone OneDrive, new Teams and new Outlook are managed separately.'
        if ($script:MaintenanceRebootRequired) { $message += ' Restart Windows to finish cleanup; Office Studio will not restart automatically.' }
        Write-MaintenanceEvent -Kind complete -Message $message -Percent 100
    } catch {
        $failure = $_
        if ($removalStarted -and -not $finalInventoryEmitted) {
            try { $remaining = Get-OfficeInventory; Write-MaintenanceEvent -Kind inventory -Message 'Current Office inventory after the removal attempt.' -Inventory $remaining }
            catch { Write-MaintenanceEvent -Kind log -Message ('Could not complete the final inventory: ' + $_.Exception.Message) }
        }
        Write-MaintenanceEvent -Kind log -Message ('Removal stopped: ' + $failure.Exception.Message)
        throw $failure
    } finally {
        if ($null -ne $operationLock) { $operationLock.Dispose() }
        if ($ownsMutex -and $null -ne $mutex) { $mutex.ReleaseMutex() }
        if ($null -ne $mutex) { $mutex.Dispose() }
        $script:MaintenanceLogPath = ''
    }
}

if (-not $Library) {
    try {
        [Console]::OutputEncoding = New-Object Text.UTF8Encoding($false)
        if ($Action -eq 'Inventory') {
            Write-MaintenanceEvent -Kind status -Message 'Checking installed Office suites, applications and Store packages...'
            $inventory = Get-OfficeInventory
            $message = 'Office scan complete.'
            if (@($inventory.Warnings).Count -gt 0) { $message = 'Office scan finished with warnings. Some installations could not be checked.' }
            elseif (@($inventory.Items).Count -eq 0) { $message = 'No Office suite registrations found.' }
            Write-MaintenanceEvent -Kind inventory -Message $message -Inventory $inventory
            Write-MaintenanceEvent -Kind complete -Message $message -Percent 100
        } elseif ($Action -eq 'ToolTerms') {
            if (-not $RequestPath -or -not (Test-Path -LiteralPath $RequestPath -PathType Leaf)) { throw 'A tool license request is required.' }
            $request = Get-Content -LiteralPath $RequestPath -Raw -Encoding UTF8 | ConvertFrom-Json
            Invoke-MaintenanceToolTerms -Request $request
        } elseif ($Action -eq 'Uninstall') {
            if (-not $RequestPath -or -not (Test-Path -LiteralPath $RequestPath -PathType Leaf)) { throw 'A confirmed removal request is required.' }
            $request = Get-Content -LiteralPath $RequestPath -Raw -Encoding UTF8 | ConvertFrom-Json
            Invoke-OfficeRemoval -Request $request
        } else { throw 'Choose -Action Inventory, -Action Uninstall or -Action ToolTerms.' }
        exit 0
    } catch { Write-MaintenanceEvent -Kind error -Message $_.Exception.Message; exit 1 }
}
