#Requires -Version 5.1
<#
.SYNOPSIS
    Offline regression tests for KAREER Office Studio maintenance.
.DESCRIPTION
    Uses synthetic installation records and test doubles. Never downloads,
    uninstalls applications, changes Windows settings, or removes user files.
#>
[CmdletBinding()]
param([string]$WorkerPath)

$ErrorActionPreference = 'Stop'
if (-not $WorkerPath) { $WorkerPath = Join-Path $PSScriptRoot 'OfficeMaintenance.ps1' }
$script:MaintenanceSourcePath = [IO.Path]::GetFullPath($WorkerPath)
$script:Passed = 0
$script:Failures = New-Object 'System.Collections.Generic.List[string]'
$script:TestRoot = Join-Path ([IO.Path]::GetTempPath()) ('KareerMaintenance-Tests-' + [Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($script:TestRoot)

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}
function Assert-Equal {
    param($Actual, $Expected, [string]$Message)
    if ($Actual -cne $Expected) { throw ('{0}. Expected [{1}], got [{2}].' -f $Message, $Expected, $Actual) }
}
function Assert-Throws {
    param([scriptblock]$Body, [string]$Message)
    $didThrow = $false
    try { & $Body | Out-Null } catch { $didThrow = $true }
    if (-not $didThrow) { throw $Message }
}
function Invoke-Check {
    param([string]$Name, [scriptblock]$Body)
    try {
        & $Body
        $script:Passed++
        Write-Host ('PASS  ' + $Name) -ForegroundColor Green
    } catch {
        $detail = $Name + ': ' + $_.Exception.Message
        $script:Failures.Add($detail)
        Write-Host ('FAIL  ' + $detail) -ForegroundColor Red
    }
}

function New-RegistryFixture {
    param(
        [string]$Name = 'Microsoft Office Professional Plus 2016',
        [string]$KeyName = '{90160000-0011-0000-1000-0000000FF1CE}',
        [string]$Publisher = 'Microsoft Corporation',
        [string]$View = 'Registry64',
        [string]$Hive = 'LocalMachine',
        [int]$WindowsInstaller = 1
    )
    [pscustomobject]@{
        Hive = $Hive; View = $View; KeyName = $KeyName; DisplayName = $Name
        DisplayVersion = '16.0.4266.1001'; Publisher = $Publisher
        WindowsInstaller = $WindowsInstaller; UninstallString = ('MsiExec.exe /I' + $KeyName)
    }
}

function New-ClickToRunFixture {
    param([string]$Products = 'O365ProPlusRetail,VisioProRetail', [string]$View = 'Registry64')
    [pscustomobject]@{
        Hive = 'LocalMachine'; View = $View; ProductReleaseIds = $Products
        VersionToReport = '16.0.19127.20154'; ClientVersionToReport = '16.0.19127.20154'
        Platform = 'x64'; InstallationPath = 'C:\Program Files\Microsoft Office'
    }
}

function New-StoreFixture {
    param([string]$Name = 'Microsoft.Office.Desktop', [bool]$Provisioned = $false)
    [pscustomobject]@{
        Name = $Name; PackageFullName = ($Name + '_16051.14326.20962.0_x86__8wekyb3d8bbwe')
        Version = '16051.14326.20962.0'; Architecture = 'x86'; IsProvisioned = $Provisioned
    }
}

function New-RemovalRequest {
    param([string]$Fingerprint = ('a' * 64), [string]$Mode = 'Standard')
    [pscustomobject]@{ Confirmed = $true; Fingerprint = $Fingerprint; Mode = $Mode; WorkPath = (Join-Path $script:TestRoot 'RemovalWork') }
}

try {
    if (-not (Test-Path -LiteralPath $WorkerPath -PathType Leaf)) { throw ('Worker was not found: ' + $WorkerPath) }
    $tokens = $null; $parseErrors = $null
    $null = [Management.Automation.Language.Parser]::ParseFile($WorkerPath, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count -gt 0) { throw ($parseErrors | Out-String) }
    . $WorkerPath -Library

    # Aliases have precedence over imported commands. These fail fast if any
    # test accidentally reaches a real Windows mutator or external operation.
    function Deny-TestExternalOperation { throw 'Offline maintenance tests forbid Windows application operations, settings changes, downloads and real process launches.' }
    foreach ($command in @('Remove-AppxPackage', 'Remove-AppxProvisionedPackage', 'Remove-ProvisionedAppxPackage', 'Remove-WindowsPackage',
        'Get-AppxPackage', 'Get-AppxProvisionedPackage', 'Start-Process', 'Stop-Process', 'Stop-Service', 'Set-Service',
        'Restart-Computer', 'Stop-Computer', 'Set-ExecutionPolicy', 'Set-WinHomeLocation', 'Set-ItemProperty',
        'New-ItemProperty', 'Remove-ItemProperty', 'Invoke-WebRequest', 'Invoke-RestMethod', 'Get-AuthenticodeSignature')) {
        Set-Alias -Name $command -Value Deny-TestExternalOperation -Scope Global -Force
    }

    $script:RealAppCheck = ${function:Assert-NoRunningOfficeApplications}
    $script:RealProcessRunner = ${function:Invoke-MaintenanceProcess}
    $script:Downloads = New-Object 'System.Collections.Generic.List[string]'
    $script:Processes = New-Object 'System.Collections.Generic.List[object]'
    $script:StoreCalls = New-Object 'System.Collections.Generic.List[object]'
    $script:Events = New-Object 'System.Collections.Generic.List[object]'
    $script:InventoryQueue = New-Object 'System.Collections.Generic.Queue[object]'
    $script:LastInventory = $null
    $script:InventoryReads = 0
    $script:SimulatedProcesses = @()
    $script:MockExitCode = 0
    $script:MockToolOutput = '00: Successfully completed this scenario.'
    $script:MockTermsFile = $true
    function Assert-MaintenanceAdministrator { }
    function Assert-MaintenanceMicrosoftSignature { param([string]$Path) }
    function Get-OfficeInventory {
        $script:InventoryReads++
        if ($script:InventoryQueue.Count -gt 0) { $script:LastInventory = $script:InventoryQueue.Dequeue() }
        if ($null -eq $script:LastInventory) { throw 'Offline test has no inventory fixture.' }
        return $script:LastInventory
    }
    function Assert-NoRunningOfficeApplications {
        param([object[]]$Processes)
        if ($PSBoundParameters.ContainsKey('Processes')) { & $script:RealAppCheck -Processes $Processes }
        else { & $script:RealAppCheck -Processes $script:SimulatedProcesses }
    }
    function Write-MaintenanceEvent {
        param([string]$Kind, [string]$Message, [int]$Percent = -2, $Inventory = $null, [bool]$RebootRequired = $false, [string]$Text = '')
        $script:Events.Add([pscustomobject]@{ Kind = $Kind; Message = $Message; Inventory = $Inventory; RebootRequired = $RebootRequired; Text = $Text })
    }
    function Save-MaintenanceDeploymentTool {
        param([string]$Path)
        $script:Downloads.Add('ODT')
        [IO.File]::WriteAllText($Path, 'FAKE-ODT-NEVER-EXECUTE')
    }
    function Save-MaintenanceScrubTool {
        param([string]$Directory)
        $script:Downloads.Add('Scrub')
        $path = Join-Path $Directory 'GetHelpCmd.exe'
        [IO.File]::WriteAllText($path, 'FAKE-SCRUB-NEVER-EXECUTE')
        return $path
    }
    function Invoke-MaintenanceProcess {
        param([string]$FilePath, [string]$Arguments, [string]$Label, [string]$WorkingDirectory, [int[]]$SuccessExitCodes = @(0, 3010))
        $script:Processes.Add([pscustomobject]@{ FilePath = $FilePath; Arguments = $Arguments; WorkingDirectory = $WorkingDirectory })
        if ($script:MockExitCode -notin $SuccessExitCodes) { throw ('Simulated Microsoft tool failure, exit code ' + $script:MockExitCode) }
        if ($Arguments -match '^-DisplayEULA "(.+)"$' -and $script:MockTermsFile) { [IO.File]::WriteAllText($Matches[1], 'TEST LICENSE TERMS - offline fixture only') }
        if ($script:MockExitCode -eq 3010) { $script:MaintenanceRebootRequired = $true }
        [pscustomobject]@{ ExitCode = $script:MockExitCode; Output = $script:MockToolOutput; ErrorOutput = '' }
    }
    function Remove-MaintenanceStorePackages {
        param([object[]]$Items)
        # Replace the entire AppX adapter. Never invoke any Windows AppX command.
        foreach ($item in @($Items | Where-Object { $_.Kind -eq 'Store' })) {
            foreach ($package in @($item.ProductIds)) {
                $script:StoreCalls.Add([pscustomobject]@{ Scope = $item.Id; Package = $package })
            }
        }
    }
    function Reset-Scenario {
        param($Before, $After)
        $script:Downloads.Clear(); $script:Processes.Clear(); $script:StoreCalls.Clear(); $script:Events.Clear(); $script:InventoryQueue.Clear()
        $script:InventoryQueue.Enqueue($Before)
        if ($null -ne $After) { $script:InventoryQueue.Enqueue($After) }
        $script:LastInventory = $null; $script:InventoryReads = 0
        $script:MockExitCode = 0; $script:MockToolOutput = '00: Successfully completed this scenario.'
        $script:MockTermsFile = $true
        $script:SimulatedProcesses = @()
    }
    function Assert-NoRemovalActivity {
        Assert-Equal $script:Downloads.Count 0 'A blocked removal still downloaded a removal tool'
        Assert-Equal $script:Processes.Count 0 'A blocked removal still launched a removal tool'
        Assert-Equal $script:StoreCalls.Count 0 'A blocked removal still removed a Store package'
    }

    Invoke-Check 'Click-to-Run metadata preserves product IDs, version and architecture' {
        $inventory = ConvertTo-OfficeInventory -RegistryEntries @() -ClickToRunEntries @((New-ClickToRunFixture)) -StorePackages @() -Warnings @()
        Assert-Equal @($inventory.Items).Count 1 'One Click-to-Run installation should produce one inventory entry'
        $item = @($inventory.Items)[0]
        Assert-True (@($item.ProductIds) -contains 'O365ProPlusRetail') 'Microsoft 365 product ID was lost'
        Assert-True (@($item.ProductIds) -contains 'VisioProRetail') 'Shared Visio product ID was lost'
        Assert-Equal $item.Version '16.0.19127.20154' 'Click-to-Run version changed'
        Assert-Equal $item.Architecture 'x64' 'Click-to-Run architecture changed'
        Assert-True $item.Removable 'Known Click-to-Run installation was not eligible for removal'
        Assert-True ($inventory.Fingerprint -match '^[a-f0-9]{64}$') 'Inventory fingerprint was not SHA-256'
    }

    Invoke-Check 'Authoritative Office inventory overrides stale legacy Click-to-Run metadata' {
        $current = New-ClickToRunFixture -Products 'ProPlus2024Volume'
        $current | Add-Member -NotePropertyName OfficeProductReleaseIds -NotePropertyValue 'ProPlus2024Volume'
        $current | Add-Member -NotePropertyName OfficePackageVersion -NotePropertyValue '16.0.17932.20540'
        $stale = New-ClickToRunFixture -Products 'ProPlus2019Retail' -View 'Registry32'
        $stale.Platform = 'x86'
        $inventory = ConvertTo-OfficeInventory -ClickToRunEntries @($current, $stale)
        Assert-Equal @($inventory.Items).Count 1 'One Office installation became multiple targets'
        Assert-Equal @($inventory.Items)[0].Version '16.0.17932.20540' 'Stale configuration version overrode the authoritative inventory'
        Assert-True (@($inventory.Items)[0].ProductIds -notcontains 'ProPlus2019Retail') 'Stale product IDs polluted the current installation'
        Assert-Equal @($inventory.Items)[0].Architecture 'x64' 'Stale architecture polluted the current installation'
    }

    Invoke-Check 'MSI inventory keeps Office and excludes separate applications and shared runtimes' {
        $entries = @(
            (New-RegistryFixture),
            (New-RegistryFixture -Name 'Microsoft OneDrive' -KeyName '{11111111-1111-1111-1111-111111111111}'),
            (New-RegistryFixture -Name 'Microsoft Teams' -KeyName '{22222222-2222-2222-2222-222222222222}'),
            (New-RegistryFixture -Name 'Microsoft Visual Studio 2010 Tools for Office Runtime (x64)' -KeyName '{33333333-3333-3333-3333-333333333333}'),
            (New-RegistryFixture -Name 'Microsoft Access Database Engine 2016 (English)' -KeyName '{44444444-4444-4444-4444-444444444444}'),
            (New-RegistryFixture -Name 'Contoso Office Suite' -Publisher 'Contoso' -KeyName '{55555555-5555-5555-5555-555555555555}')
        )
        $inventory = ConvertTo-OfficeInventory -RegistryEntries $entries -ClickToRunEntries @() -StorePackages @() -Warnings @()
        Assert-Equal @($inventory.Items | Where-Object { $_.Removable }).Count 1 'Unrelated applications or shared runtimes were marked for removal'
        Assert-Equal @($inventory.Items | Where-Object { $_.Removable })[0].Name 'Microsoft Office Professional Plus 2016' 'The Office MSI product was lost'
    }

    Invoke-Check 'Duplicate registry views do not create duplicate removals' {
        $inventory = ConvertTo-OfficeInventory -RegistryEntries @((New-RegistryFixture), (New-RegistryFixture -View 'Registry32')) -ClickToRunEntries @((New-ClickToRunFixture), (New-ClickToRunFixture -View 'Registry32')) -StorePackages @() -Warnings @()
        Assert-Equal @($inventory.Items | Where-Object { $_.Removable }).Count 2 'Duplicate records became duplicate removal targets'
    }

    Invoke-Check 'Malformed MSI product GUID cannot become a removal target' {
        $inventory = ConvertTo-OfficeInventory -RegistryEntries @((New-RegistryFixture -KeyName 'OfficeSetup /quiet')) -ClickToRunEntries @() -StorePackages @() -Warnings @()
        Assert-Equal @($inventory.Items | Where-Object { $_.Removable }).Count 0 'An invalid MSI product code became removable'
        Assert-True (@($inventory.Items | Where-Object { $_.Kind -eq 'Unsupported' -and $_.Details }).Count -gt 0) 'Unrecognized Office registration was silently dropped'
    }

    Invoke-Check 'Office Store detection does not target OneDrive, Teams, OneNote or the Office hub' {
        $packages = @((New-StoreFixture), (New-StoreFixture -Name 'Microsoft.OneDriveSync'), (New-StoreFixture -Name 'MSTeams'), (New-StoreFixture -Name 'Microsoft.Office.OneNote'), (New-StoreFixture -Name 'Microsoft.MicrosoftOfficeHub'))
        $inventory = ConvertTo-OfficeInventory -RegistryEntries @() -ClickToRunEntries @() -StorePackages $packages -Warnings @()
        Assert-Equal @($inventory.Items).Count 1 'An unrelated Store application was included in Office removal'
    }

    Invoke-Check 'Mismatched Office Store identity is shown as unsupported, never removable' {
        $package = New-StoreFixture
        $package.PackageFullName = 'MSTeams_999.0.0.0_x64__8wekyb3d8bbwe'
        $inventory = ConvertTo-OfficeInventory -StorePackages @($package)
        Assert-Equal @($inventory.Items | Where-Object { $_.Removable }).Count 0 'A mismatched Store package became removable'
        Assert-Equal @($inventory.Items | Where-Object { $_.Kind -eq 'Unsupported' }).Count 1 'Unsupported Office Store identity was silently hidden'
    }

    Invoke-Check 'Partial detection warnings survive inventory normalization' {
        $inventory = ConvertTo-OfficeInventory -RegistryEntries @() -ClickToRunEntries @() -StorePackages @() -Warnings @('Registry32 could not be read; inventory may be incomplete.')
        Assert-True (@($inventory.Warnings) -contains 'Registry32 could not be read; inventory may be incomplete.') 'Partial detection warning was lost'
    }

    Invoke-Check 'Inventory fingerprint is stable across record ordering and changes with installed products' {
        $office = New-RegistryFixture
        $visio = New-RegistryFixture -Name 'Microsoft Visio Professional 2016' -KeyName '{90160000-0051-0000-1000-0000000FF1CE}'
        $first = ConvertTo-OfficeInventory -RegistryEntries @($office, $visio) -ClickToRunEntries @() -StorePackages @() -Warnings @()
        $reordered = ConvertTo-OfficeInventory -RegistryEntries @($visio, $office) -ClickToRunEntries @() -StorePackages @() -Warnings @()
        $changed = ConvertTo-OfficeInventory -RegistryEntries @($office) -ClickToRunEntries @() -StorePackages @() -Warnings @()
        Assert-Equal $first.Fingerprint $reordered.Fingerprint 'Record order changed the inventory fingerprint'
        Assert-True ($first.Fingerprint -cne $changed.Fingerprint) 'Changed installed products retained the same fingerprint'
    }

    Invoke-Check 'Removal confirmation requires a real true boolean' {
        foreach ($confirmation in @($false, 'true', 1)) {
            $request = New-RemovalRequest
            $request.Confirmed = $confirmation
            Assert-Throws { Test-MaintenanceRequest -Request $request } 'An unconfirmed or malformed request was accepted'
        }
        $request = New-RemovalRequest
        $request.PSObject.Properties.Remove('Confirmed')
        Assert-Throws { Test-MaintenanceRequest -Request $request } 'A request without confirmation was accepted'
    }

    Invoke-Check 'Valid confirmed removal modes are accepted and missing mode is rejected' {
        Test-MaintenanceRequest -Request (New-RemovalRequest -Mode 'Standard') | Out-Null
        Test-MaintenanceRequest -Request (New-RemovalRequest -Mode 'Deep') | Out-Null
        $request = New-RemovalRequest
        $request.PSObject.Properties.Remove('Mode')
        Assert-Throws { Test-MaintenanceRequest -Request $request } 'A request without an explicit removal mode was accepted'
    }

    Invoke-Check 'Removal request rejects malformed fingerprints and unsafe work paths' {
        $request = New-RemovalRequest -Fingerprint 'not-a-fingerprint'
        Assert-Throws { Test-MaintenanceRequest -Request $request } 'Malformed fingerprint was accepted'
        foreach ($path in @('..\Elsewhere', 'C:\')) {
            $request = New-RemovalRequest
            $request.WorkPath = $path
            Assert-Throws { Test-MaintenanceRequest -Request $request } 'An unsafe removal working directory was accepted'
        }
    }

    Invoke-Check 'Removal XML targets Office without adding products or forcing apps closed' {
        [xml]$xml = New-OfficeRemovalConfiguration
        Assert-Equal @($xml.SelectNodes('/Configuration/Remove')).Count 1 'There should be one Office remove instruction'
        Assert-Equal $xml.Configuration.Remove.All 'TRUE' 'Remove All was not explicitly selected'
        Assert-True ($null -eq $xml.SelectSingleNode('/Configuration/Add')) 'Removal XML unexpectedly contains installation instructions'
        Assert-True ($null -eq $xml.SelectSingleNode('/Configuration/Property[@Name="FORCEAPPSHUTDOWN" and @Value="TRUE"]')) 'Removal XML would force-close running Office'
    }

    Invoke-Check 'Running Office apps block removal while unrelated apps do not' {
        Assert-Throws { Assert-NoRunningOfficeApplications -Processes @([pscustomobject]@{ ProcessName = 'WINWORD' }) } 'Running Word did not block removal'
        Assert-Throws { Assert-NoRunningOfficeApplications -Processes @([pscustomobject]@{ ProcessName = 'EXCEL' }) } 'Running Excel did not block removal'
        Assert-NoRunningOfficeApplications -Processes @([pscustomobject]@{ ProcessName = 'notepad' }, [pscustomobject]@{ ProcessName = 'OneDrive' })
    }

    Invoke-Check 'Missing confirmation and inventory mismatch stop before removal activity' {
        $inventory = ConvertTo-OfficeInventory -ClickToRunEntries @((New-ClickToRunFixture))
        Reset-Scenario -Before $inventory
        $request = New-RemovalRequest -Fingerprint $inventory.Fingerprint
        $request.PSObject.Properties.Remove('Confirmed')
        Assert-Throws { Invoke-OfficeRemoval -Request $request } 'Unconfirmed removal started'
        Assert-NoRemovalActivity
        Reset-Scenario -Before $inventory
        Assert-Throws { Invoke-OfficeRemoval -Request (New-RemovalRequest -Fingerprint ('b' * 64)) } 'Removal accepted a stale inventory fingerprint'
        Assert-NoRemovalActivity
    }

    Invoke-Check 'Partial inventory and unsupported standard removal stop before changes' {
        foreach ($mode in @('Standard', 'Deep')) {
            $inventory = ConvertTo-OfficeInventory -ClickToRunEntries @((New-ClickToRunFixture)) -Warnings @('The Store scan could not be completed.')
            Reset-Scenario -Before $inventory
            Assert-Throws { Invoke-OfficeRemoval -Request (New-RemovalRequest -Fingerprint $inventory.Fingerprint -Mode $mode) } 'Removal started with an incomplete scan'
            Assert-NoRemovalActivity
        }
        $inventory = ConvertTo-OfficeInventory -RegistryEntries @((New-RegistryFixture -KeyName 'UnknownOfficeRegistration'))
        Reset-Scenario -Before $inventory
        Assert-Throws { Invoke-OfficeRemoval -Request (New-RemovalRequest -Fingerprint $inventory.Fingerprint) } 'Standard removal accepted an unknown Office registration'
        Assert-NoRemovalActivity
    }

    Invoke-Check 'Running Office applications block confirmed removal before downloads' {
        $inventory = ConvertTo-OfficeInventory -ClickToRunEntries @((New-ClickToRunFixture))
        Reset-Scenario -Before $inventory
        $script:SimulatedProcesses = @([pscustomobject]@{ ProcessName = 'OUTLOOK' })
        Assert-Throws { Invoke-OfficeRemoval -Request (New-RemovalRequest -Fingerprint $inventory.Fingerprint) } 'Removal started while Outlook was running'
        Assert-NoRemovalActivity
        $script:SimulatedProcesses = @()
    }

    Invoke-Check 'Standard removal uses exact ODT and MSI targets, then requires a clean rescan' {
        $msi = New-RegistryFixture
        $msi.UninstallString = 'untrusted-program.exe /unexpected'
        $inventory = ConvertTo-OfficeInventory -RegistryEntries @($msi) -ClickToRunEntries @((New-ClickToRunFixture))
        Reset-Scenario -Before $inventory -After (ConvertTo-OfficeInventory)
        Invoke-OfficeRemoval -Request (New-RemovalRequest -Fingerprint $inventory.Fingerprint)
        Assert-Equal $script:Downloads.Count 1 'Standard removal downloaded an unexpected toolkit'
        Assert-Equal $script:Downloads[0] 'ODT' 'Standard removal selected the wrong toolkit'
        Assert-Equal $script:Processes.Count 2 'Expected exactly one C2R and one MSI removal'
        Assert-True ($script:Processes[0].Arguments -match '^/configure "(.+)"$') 'C2R removal did not use the generated configuration'
        [xml]$config = [IO.File]::ReadAllText($Matches[1])
        Assert-Equal $config.Configuration.Remove.All 'TRUE' 'Generated configuration did not remove the approved Office scope'
        Assert-True ($script:Processes[1].Arguments.StartsWith('/x {90160000-0011-0000-1000-0000000FF1CE} /passive /norestart ')) 'MSI removal did not use the validated product GUID and suppress restart'
        Assert-True ($script:Processes[1].FilePath.EndsWith('\msiexec.exe')) 'MSI removal executed an arbitrary registered uninstall string'
        Assert-Equal $script:InventoryReads 2 'Successful removal did not rescan the installation'
        Assert-Equal @($script:Events | Where-Object { $_.Kind -eq 'complete' }).Count 1 'Successful removal did not emit a verified completion'
    }

    Invoke-Check 'Store removal targets only exact Office Desktop packages and both scopes' {
        $inventory = ConvertTo-OfficeInventory -StorePackages @((New-StoreFixture), (New-StoreFixture -Provisioned $true), (New-StoreFixture -Name 'MSTeams'))
        Reset-Scenario -Before $inventory -After (ConvertTo-OfficeInventory)
        Invoke-OfficeRemoval -Request (New-RemovalRequest -Fingerprint $inventory.Fingerprint)
        Assert-Equal $script:StoreCalls.Count 2 'Expected installed and provisioned Office Store removal'
        Assert-True (@($script:StoreCalls | Where-Object { $_.Package -notlike 'Microsoft.Office.Desktop_*' }).Count -eq 0) 'A separate Store app was removed'
        Assert-True @($script:StoreCalls | Where-Object { $_.Scope -eq 'store-installed' }).Count 'Installed package scope was not passed to the Store adapter'
        Assert-True @($script:StoreCalls | Where-Object { $_.Scope -eq 'store-provisioned' }).Count 'Provisioned package scope was not passed to the Store adapter'
        Assert-Equal $script:Processes.Count 0 'Store-only removal launched an unnecessary installer'
        # Inspect scope switches without calling destructive Windows AppX APIs.
        $ast = [Management.Automation.Language.Parser]::ParseFile($script:MaintenanceSourcePath, [ref]$tokens, [ref]$parseErrors)
        $adapter = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Remove-MaintenanceStorePackages' }, $true)
        $commands = @($adapter.FindAll({ param($node) $node -is [Management.Automation.Language.CommandAst] -and $node.GetCommandName() -in @('Remove-AppxPackage', 'Remove-AppxProvisionedPackage') }, $true))
        Assert-Equal $commands.Count 2 'Store adapter contains an unexpected removal command set'
        Assert-True (@($commands | Where-Object { $_.Extent.Text -match 'Remove-AppxPackage\s+-AllUsers\s+-Package\s+\$packageId' }).Count -eq 1) 'Installed Store removal lacks its explicit package and all-users scope'
        Assert-True (@($commands | Where-Object { $_.Extent.Text -match 'Remove-AppxProvisionedPackage\s+-Online\s+-PackageName\s+\$packageId' }).Count -eq 1) 'Provisioned Store removal lacks its explicit package scope'
    }

    Invoke-Check 'Installer failure never reports complete and refreshes remaining inventory' {
        $inventory = ConvertTo-OfficeInventory -ClickToRunEntries @((New-ClickToRunFixture))
        Reset-Scenario -Before $inventory -After $inventory
        $script:MockExitCode = 1603
        Assert-Throws { Invoke-OfficeRemoval -Request (New-RemovalRequest -Fingerprint $inventory.Fingerprint) } 'Installer failure was swallowed'
        Assert-Equal @($script:Events | Where-Object { $_.Kind -eq 'complete' }).Count 0 'Failed removal reported completion'
        Assert-True ($script:InventoryReads -ge 2) 'Failure did not refresh remaining Office registrations'
    }

    Invoke-Check 'Remaining registrations and failed final scan prevent success' {
        $inventory = ConvertTo-OfficeInventory -ClickToRunEntries @((New-ClickToRunFixture))
        foreach ($after in @($inventory, (ConvertTo-OfficeInventory -Warnings @('Final scan unavailable.')))) {
            Reset-Scenario -Before $inventory -After $after
            Assert-Throws { Invoke-OfficeRemoval -Request (New-RemovalRequest -Fingerprint $inventory.Fingerprint) } 'Unverified removal reported success'
            Assert-Equal @($script:Events | Where-Object { $_.Kind -eq 'complete' }).Count 0 'Unverified removal emitted completion'
        }
    }

    Invoke-Check 'Restart-required result is retained and still requires a clean final scan' {
        $inventory = ConvertTo-OfficeInventory -RegistryEntries @((New-RegistryFixture))
        Reset-Scenario -Before $inventory -After (ConvertTo-OfficeInventory)
        $script:MockExitCode = 3010
        Invoke-OfficeRemoval -Request (New-RemovalRequest -Fingerprint $inventory.Fingerprint)
        Assert-True $script:MaintenanceRebootRequired 'Restart-required status was lost'
        Assert-True (@($script:Events | Where-Object { $_.Kind -eq 'complete' -and $_.Message -match 'Restart Windows' }).Count -eq 1) 'Restart instruction was omitted'
        Reset-Scenario -Before $inventory -After $inventory
        $script:MockExitCode = 3010
        Assert-Throws { Invoke-OfficeRemoval -Request (New-RemovalRequest -Fingerprint $inventory.Fingerprint) } 'Restart-required with remaining Office registrations was reported as complete'
        Assert-Equal @($script:Events | Where-Object { $_.Kind -eq 'complete' }).Count 0 'Remaining Office registrations emitted completion'
    }

    Invoke-Check 'Deep cleanup accepts explicit empty-inventory consent and invokes only Office scenario' {
        $inventory = ConvertTo-OfficeInventory
        Reset-Scenario -Before $inventory -After $inventory
        Invoke-OfficeRemoval -Request (New-RemovalRequest -Fingerprint $inventory.Fingerprint -Mode 'Deep')
        Assert-Equal $script:Downloads[0] 'Scrub' 'Deep cleanup did not fetch the fresh Microsoft toolkit'
        Assert-Equal $script:Processes.Count 1 'Deep cleanup launched extra programs'
        Assert-Equal $script:Processes[0].Arguments '-S OfficeScrubScenario -AcceptEula -OfficeVersion All' 'Deep cleanup selected an unexpected scenario'
        Assert-True $script:MaintenanceRebootRequired 'Deep cleanup did not retain its restart requirement'
    }

    Invoke-Check 'Deep cleanup requires recognized success output, even with exit code zero' {
        Test-MaintenanceScrubResult -Result ([pscustomobject]@{ ExitCode = 0; Output = '00: Successfully completed this scenario.'; ErrorOutput = '' })
        foreach ($output in @('', '00: Unrecognized result.', '06: Office applications are running.', '08: Multiple versions detected.')) {
            Assert-Throws { Test-MaintenanceScrubResult -Result ([pscustomobject]@{ ExitCode = 0; Output = $output; ErrorOutput = '' }) } 'Unrecognized or failed deep cleanup output was accepted'
        }
    }

    Invoke-Check 'ToolTerms only displays terms without accepting a license or starting removal' {
        Reset-Scenario -Before (ConvertTo-OfficeInventory)
        Invoke-MaintenanceToolTerms -Request ([pscustomobject]@{ WorkPath = (Join-Path $script:TestRoot 'TermsWork') })
        Assert-Equal $script:Processes.Count 1 'Viewing terms launched unexpected processes'
        Assert-True ($script:Processes[0].Arguments -match '^-DisplayEULA "[^"]+"$') 'Viewing terms did not use only the display-EULA switch'
        Assert-True ($script:Processes[0].Arguments -notmatch 'AcceptEula|OfficeScrubScenario|OfficeVersion') 'Viewing terms accepted a license or requested removal'
        Assert-Equal $script:InventoryReads 0 'Viewing terms unnecessarily inspected installed Office'
        Assert-Equal $script:StoreCalls.Count 0 'Viewing terms requested Store package removal'
        Assert-Equal @($script:Events | Where-Object { $_.Kind -eq 'terms' -and $_.Text -eq 'TEST LICENSE TERMS - offline fixture only' }).Count 1 'Terms text was not returned for user review'
        Reset-Scenario -Before (ConvertTo-OfficeInventory)
        $script:MockTermsFile = $false
        Assert-Throws { Invoke-MaintenanceToolTerms -Request ([pscustomobject]@{ WorkPath = (Join-Path $script:TestRoot 'NoTermsWork') }) } 'Missing tool terms file was accepted'
    }
} finally {
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    $resolvedTestRoot = [IO.Path]::GetFullPath($script:TestRoot)
    if ($resolvedTestRoot.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -and
        [IO.Path]::GetFileName($resolvedTestRoot).StartsWith('KareerMaintenance-Tests-')) {
        if (Test-Path -LiteralPath $resolvedTestRoot) { Remove-Item -LiteralPath $resolvedTestRoot -Recurse -Force }
    } else { throw 'Refusing to clean up outside the test temporary folder.' }
}

Write-Host ('{0} passed; {1} failed.' -f $script:Passed, $script:Failures.Count)
if ($script:Failures.Count -gt 0) { throw ($script:Failures -join [Environment]::NewLine) }
