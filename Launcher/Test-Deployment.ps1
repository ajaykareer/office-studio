#Requires -Version 5.1
<#
.SYNOPSIS
    Offline regression checks for the KAREER Office Studio deployment worker.
.DESCRIPTION
    Runs only in a temporary test directory. Network requests and external
    installers are replaced with test doubles. Never downloads or installs Office.
#>
[CmdletBinding()]
param(
    [string]$WorkerPath
)

$ErrorActionPreference = 'Stop'
if (-not $WorkerPath) { $WorkerPath = Join-Path $PSScriptRoot 'DeploymentWorker.ps1' }
$script:Passed = 0
$script:Failures = New-Object 'System.Collections.Generic.List[string]'
$script:TestRoot = Join-Path ([IO.Path]::GetTempPath()) ('KareerOfficeStudio-Tests-' + [Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($script:TestRoot)

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

function Assert-Equal {
    param($Actual, $Expected, [string]$Message)
    if ($Actual -cne $Expected) {
        throw ('{0}. Expected [{1}], got [{2}].' -f $Message, $Expected, $Actual)
    }
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

function New-TestPlan {
    param(
        [string]$Branch = 'ProPlus2024Volume',
        [string]$Channel = 'PerpetualVL2024',
        [string[]]$Components = @('Word', 'Excel'),
        [string]$Architecture = '64',
        [string]$Language = 'en-us',
        [string]$Folder = 'Office files & settings'
    )
    [pscustomobject]@{
        PlanId = ('a' * 64)
        Branch = $Branch
        Channel = $Channel
        Components = $Components
        Architecture = $Architecture
        Language = $Language
        DownloadPath = (Join-Path $script:TestRoot $Folder)
    }
}

function New-TestPayload {
    param($Plan, [string]$Version = '16.0.19127.20154')
    $dataPath = Join-Path $Plan.DownloadPath 'Office\Data'
    $versionPath = Join-Path $dataPath $Version
    [void][IO.Directory]::CreateDirectory($versionPath)
    $arch = 'x64'
    if ($Plan.Architecture -eq '32') { $arch = 'x86' }
    [IO.File]::WriteAllText((Join-Path $dataPath ('v' + $Plan.Architecture + '.cab')), 'test-catalog-data')
    [IO.File]::WriteAllText((Join-Path $versionPath ('stream.' + $arch + '.x-none.dat')), 'test-office-data')
    [IO.File]::WriteAllText((Join-Path $versionPath ('stream.' + $arch + '.' + $Plan.Language + '.dat')), 'test-language-data')
}

function New-ReadyFixture {
    param([string[]]$Components = @('Word', 'Excel'))
    $plan = Test-DeploymentPlan -Plan (New-TestPlan -Components $Components -Folder ([Guid]::NewGuid().ToString('N')))
    New-TestPayload -Plan $plan
    [IO.File]::WriteAllText((Join-Path $plan.DownloadPath 'setup.exe'), 'fake-microsoft-tool')
    [IO.File]::WriteAllText((Join-Path $plan.DownloadPath 'Config.xml'), (New-OfficeConfiguration -Plan $plan -OfficeVersion '16.0.19127.20154'))
    foreach ($addon in @(Get-AddonDownloads -Plan $plan)) {
        $path = Join-Path $plan.DownloadPath $addon.RelativePath
        [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($path))
        [IO.File]::WriteAllText($path, 'fake-addon-data')
    }
    $null = Write-ReadyState -Plan $plan -Payload (Get-OfficePayload -Plan $plan)
    return $plan
}

try {
    if (-not (Test-Path -LiteralPath $WorkerPath -PathType Leaf)) {
        throw ('Worker was not found: ' + $WorkerPath)
    }
    $tokens = $null
    $parseErrors = $null
    $null = [Management.Automation.Language.Parser]::ParseFile($WorkerPath, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count -gt 0) { throw ($parseErrors | Out-String) }
    # Offline tests must not contend with a user's real Office operation.
    $isolatedWorker = Join-Path $script:TestRoot 'test-worker.ps1'
    [IO.File]::WriteAllText($isolatedWorker, ([IO.File]::ReadAllText($WorkerPath).Replace('Global\KAREER.OfficeStudio.OfficeOperation', ('Local\KAREER.OfficeStudio.Test.' + [Guid]::NewGuid().ToString('N')))))
    . $isolatedWorker -Library

    # Test doubles prevent all network and installer activity. Real payload checks,
    # plan validation, configuration generation and file hashing remain in use.
    $script:RealDeploymentProcess = ${function:Invoke-DeploymentProcess}
    $script:Downloads = New-Object 'System.Collections.Generic.List[object]'
    $script:Processes = New-Object 'System.Collections.Generic.List[object]'
    $script:Events = New-Object 'System.Collections.Generic.List[object]'
    $script:FailDownloadProcess = $false
    function Assert-MicrosoftSignature { param([string]$Path) }
    function Write-DeploymentEvent {
        param([string]$Kind, [string]$Message, [int]$Percent = -2)
        $script:Events.Add([pscustomobject]@{ Kind = $Kind; Message = $Message; Percent = $Percent })
    }
    function Save-MicrosoftDownload {
        param([string]$Url, [string]$Path, [string]$Label)
        $script:Downloads.Add([pscustomobject]@{ Url = $Url; Path = $Path; Label = $Label })
        [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($Path))
        [IO.File]::WriteAllText($Path, 'fake-downloaded-installer')
    }
    function Invoke-DeploymentProcess {
        param([string]$FilePath, [string]$Arguments, [string]$Label, $Plan, [int[]]$SuccessExitCodes)
        $script:Processes.Add([pscustomobject]@{ FilePath = $FilePath; Arguments = $Arguments; SuccessExitCodes = $SuccessExitCodes })
        if ($Arguments -notmatch '^/download ') { throw ('Offline test refused an installer action: ' + $Arguments) }
        if ($script:FailDownloadProcess) { throw 'Simulated Office download failure (exit code 123).' }
        if ($Arguments -notmatch '^/download "(.+)"$') { throw 'Download configuration was not correctly quoted.' }
        [xml]$downloadConfig = [IO.File]::ReadAllText($Matches[1])
        $downloadPlan = $Plan | ConvertTo-Json -Depth 6 | ConvertFrom-Json
        $downloadPlan.DownloadPath = [string]$downloadConfig.Configuration.Add.SourcePath
        New-TestPayload -Plan $downloadPlan
        return [pscustomobject]@{ ExitCode = 0; Output = ''; ErrorOutput = '' }
    }

    foreach ($profile in @(
        @('ProPlus2019Retail', 'Current'),
        @('ProPlus2021Volume', 'PerpetualVL2021'),
        @('ProPlus2024Volume', 'PerpetualVL2024'),
        @('O365ProPlusRetail', 'Current')
    )) {
        $testProfile = $profile
        Invoke-Check ('Word and Excel selection for ' + $testProfile[0]) {
            $plan = Test-DeploymentPlan -Plan (New-TestPlan -Branch $testProfile[0] -Channel $testProfile[1])
            [xml]$config = New-OfficeConfiguration -Plan $plan
            $product = $config.SelectSingleNode('/Configuration/Add/Product[@ID="' + $testProfile[0] + '"]')
            Assert-True ($null -ne $product) 'The chosen Office product was not configured'
            Assert-Equal $config.Configuration.Add.Channel $testProfile[1] 'Office channel changed'
            Assert-True ($null -eq $product.SelectSingleNode('ExcludeApp[@ID="Word"]')) 'Word was excluded'
            Assert-True ($null -eq $product.SelectSingleNode('ExcludeApp[@ID="Excel"]')) 'Excel was excluded'
            foreach ($app in @('Access', 'Outlook', 'OneNote', 'PowerPoint')) {
                Assert-True ($null -ne $product.SelectSingleNode('ExcludeApp[@ID="' + $app + '"]')) ('Unselected app was not excluded: ' + $app)
            }
            Assert-Equal $config.Configuration.Add.OfficeClientEdition '64' 'Architecture changed'
            Assert-Equal $config.Configuration.Add.SourcePath $plan.DownloadPath 'Download folder did not survive XML escaping'
            Assert-Equal $product.Language.ID 'en-us' 'Language changed'
        }
    }

    Invoke-Check '32-bit configuration and version pinning' {
        $plan = Test-DeploymentPlan -Plan (New-TestPlan -Architecture '32')
        [xml]$config = New-OfficeConfiguration -Plan $plan -OfficeVersion '16.0.19127.20154'
        Assert-Equal $config.Configuration.Add.OfficeClientEdition '32' '32-bit selection was not retained'
        Assert-Equal $config.Configuration.Add.Version '16.0.19127.20154' 'Install did not pin the downloaded Office version'
    }

    Invoke-Check 'Configuration does not force-close Office or uninstall existing products' {
        [xml]$config = New-OfficeConfiguration -Plan (Test-DeploymentPlan -Plan (New-TestPlan))
        Assert-True ($null -eq $config.SelectSingleNode('/Configuration/Property[@Name="FORCEAPPSHUTDOWN" and @Value="TRUE"]')) 'Office apps would be forcibly closed'
        Assert-True ($null -eq $config.SelectSingleNode('/Configuration/Remove | /Configuration/RemoveMSI')) 'Configuration would uninstall existing products'
    }

    Invoke-Check 'Wrong profile-channel combinations are rejected' {
        Assert-Throws { Test-DeploymentPlan -Plan (New-TestPlan -Channel 'Current') } 'Office 2024 accepted the Current channel'
        Assert-Throws { Test-DeploymentPlan -Plan (New-TestPlan -Branch 'ProPlus2021Volume' -Channel 'PerpetualVL2024') } 'Office 2021 accepted the Office 2024 channel'
    }

    Invoke-Check 'Unknown product and component inputs are rejected' {
        Assert-Throws { Test-DeploymentPlan -Plan (New-TestPlan -Branch 'NotAnOfficeProduct') } 'Unknown product was accepted'
        Assert-Throws { Test-DeploymentPlan -Plan (New-TestPlan -Components @('Word', 'NotAnOfficeApp')) } 'Unknown app was accepted'
    }

    Invoke-Check 'Empty app selection is rejected' {
        Assert-Throws { Test-DeploymentPlan -Plan (New-TestPlan -Components @()) } 'Empty selection was accepted'
    }

    Invoke-Check 'Invalid architecture is rejected' {
        Assert-Throws { Test-DeploymentPlan -Plan (New-TestPlan -Architecture '128') } 'Unknown architecture was accepted'
    }

    Invoke-Check 'Relative download folder is rejected' {
        $plan = New-TestPlan
        $plan.DownloadPath = '..\Elsewhere'
        Assert-Throws { Test-DeploymentPlan -Plan $plan } 'Relative download folder was accepted'
    }

    Invoke-Check 'Plan JSON round trip and malformed JSON rejection' {
        $path = Join-Path $script:TestRoot 'plan.json'
        New-TestPlan | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $path -Encoding UTF8
        $plan = Read-DeploymentPlan -Path $path
        Assert-Equal $plan.Branch 'ProPlus2024Volume' 'Plan read changed the selected product'
        Set-Content -LiteralPath $path -Value '{this is not valid JSON' -Encoding UTF8
        Assert-Throws { Read-DeploymentPlan -Path $path } 'Invalid JSON was accepted'
    }

    Invoke-Check 'Missing Office payload blocks install readiness' {
        $plan = Test-DeploymentPlan -Plan (New-TestPlan -Folder 'NoPayload')
        Assert-Throws { Get-OfficePayload -Plan $plan } 'A missing Office payload was accepted'
    }

    Invoke-Check 'Publisher is rejected for Office 2024 and Microsoft 365' {
        Assert-Throws { Test-DeploymentPlan -Plan (New-TestPlan -Components @('Word', 'Publisher')) } 'Office 2024 accepted Publisher'
        Assert-Throws { Test-DeploymentPlan -Plan (New-TestPlan -Branch 'O365ProPlusRetail' -Channel 'Current' -Components @('Word', 'Publisher')) } 'Microsoft 365 accepted Publisher'
    }

    Invoke-Check 'Addons alone and duplicate components are rejected' {
        Assert-Throws { Test-DeploymentPlan -Plan (New-TestPlan -Components @('OneDrive', 'Teams')) } 'An addon-only selection was accepted'
        Assert-Throws { Test-DeploymentPlan -Plan (New-TestPlan -Components @('Word', 'Word')) } 'Duplicate apps were accepted'
    }

    Invoke-Check 'Office payload requires its language stream and version catalog' {
        $plan = Test-DeploymentPlan -Plan (New-TestPlan -Folder 'IncompletePayload')
        New-TestPayload -Plan $plan
        $languagePath = Join-Path $plan.DownloadPath 'Office\Data\16.0.19127.20154\stream.x64.en-us.dat'
        Remove-Item -LiteralPath $languagePath
        Assert-Throws { Get-OfficePayload -Plan $plan } 'Missing language stream was accepted'
        New-TestPayload -Plan $plan
        Remove-Item -LiteralPath (Join-Path $plan.DownloadPath 'Office\Data\v64.cab')
        Assert-Throws { Get-OfficePayload -Plan $plan } 'Missing version catalog was accepted'
    }

    Invoke-Check 'Ready state accepts a complete unchanged download' {
        $plan = New-ReadyFixture
        $ready = Test-ReadyState -Plan $plan
        Assert-Equal $ready.PlanId $plan.PlanId 'Ready state returned a different plan'
        Assert-Equal $ready.OfficeVersion '16.0.19127.20154' 'Ready state returned a different Office version'
    }

    Invoke-Check 'Ready state rejects changed selection even when PlanId is reused' {
        $plan = New-ReadyFixture
        $plan.Components = @('Word', 'PowerPoint')
        Assert-Throws { Test-ReadyState -Plan $plan } 'Readiness survived a change to the selected apps'
    }

    Invoke-Check 'Ready state rejects changed PlanId' {
        $plan = New-ReadyFixture
        $plan.PlanId = ('b' * 64)
        Assert-Throws { Test-ReadyState -Plan $plan } 'Readiness survived a change to PlanId'
    }

    Invoke-Check 'Ready state rejects modified Office configuration' {
        $plan = New-ReadyFixture
        [IO.File]::AppendAllText((Join-Path $plan.DownloadPath 'Config.xml'), ' ')
        Assert-Throws { Test-ReadyState -Plan $plan } 'Modified configuration was accepted'
    }

    Invoke-Check 'Ready state rejects same-size payload corruption' {
        $plan = New-ReadyFixture
        [IO.File]::WriteAllText((Join-Path $plan.DownloadPath 'Office\Data\16.0.19127.20154\stream.x64.x-none.dat'), 'TEST-office-data')
        Assert-Throws { Test-ReadyState -Plan $plan } 'A same-size corrupted Office stream was accepted'
    }

    Invoke-Check 'Ready state rejects a missing selected addon' {
        $plan = New-ReadyFixture -Components @('Word', 'OneDrive', 'Teams')
        Remove-Item -LiteralPath (Join-Path $plan.DownloadPath 'Addons\MSTeams.msix')
        Assert-Throws { Test-ReadyState -Plan $plan } 'A missing selected Teams package was accepted'
    }

    Invoke-Check 'Download only downloads, then validates without running an installer' {
        $script:Downloads.Clear(); $script:Processes.Clear(); $script:Events.Clear()
        $plan = Test-DeploymentPlan -Plan (New-TestPlan -Components @('Word', 'Excel', 'OneDrive', 'Teams') -Folder 'MockDownload')
        Invoke-Deployment -Action Download -Plan $plan
        Assert-Equal $script:Processes.Count 1 'Download launched an unexpected number of processes'
        Assert-True ($script:Processes[0].Arguments.StartsWith('/download ')) 'Download invoked installation'
        Assert-Equal $script:Downloads.Count 4 'Download did not fetch the Office tool and all selected addons'
        Assert-True (Test-Path -LiteralPath (Join-Path $plan.DownloadPath 'ready.json')) 'A successful download did not create its verification record'
        Assert-Equal @($script:Events | Where-Object { $_.Kind -eq 'complete' }).Count 1 'Download did not produce one completion event'
        Invoke-Deployment -Action Validate -Plan $plan
        Assert-Equal $script:Processes.Count 1 'Validating files launched an installer'
    }

    Invoke-Check 'Failed repeat download removes stale readiness and never reports success' {
        $plan = New-ReadyFixture
        [IO.File]::AppendAllText((Join-Path $plan.DownloadPath 'Config.xml'), 'corrupt')
        $script:Events.Clear()
        $script:FailDownloadProcess = $true
        try {
            Assert-Throws { Invoke-Deployment -Action Download -Plan $plan } 'Failed download did not report an error'
            Assert-True (-not (Test-Path -LiteralPath (Join-Path $plan.DownloadPath 'ready.json'))) 'An earlier ready marker survived a failed new download'
            Assert-Equal @($script:Events | Where-Object { $_.Kind -eq 'complete' }).Count 0 'Failed download reported success'
        } finally { $script:FailDownloadProcess = $false }
    }

    Invoke-Check 'A newer cached Office version cannot replace the freshly downloaded version' {
        $plan = Test-DeploymentPlan -Plan (New-TestPlan -Folder 'OldHigherVersion')
        New-TestPayload -Plan $plan -Version '16.0.99999.99999'
        Invoke-Deployment -Action Download -Plan $plan
        [xml]$config = [IO.File]::ReadAllText((Join-Path $plan.DownloadPath 'Config.xml'))
        Assert-Equal $config.Configuration.Add.Version '16.0.19127.20154' 'Download selected an older cache from another channel'
        $ready = Test-ReadyState -Plan $plan
        Assert-Equal $ready.OfficeVersion '16.0.19127.20154' 'Ready state selected an older cache from another channel'
        Assert-Equal @(Get-ChildItem -LiteralPath $plan.DownloadPath -Directory -Force | Where-Object { $_.Name -like '.kareer-download-*' }).Count 0 'A completed download left staging folders behind'
    }

    Invoke-Check 'A verified repeat download uses no network or deployment process' {
        $plan = New-ReadyFixture
        $script:Downloads.Clear(); $script:Processes.Clear(); $script:Events.Clear()
        Invoke-Deployment -Action Download -Plan $plan
        Assert-Equal $script:Downloads.Count 0 'A verified package was downloaded again'
        Assert-Equal $script:Processes.Count 0 'A verified package launched setup'
        Assert-Equal @($script:Events | Where-Object { $_.Kind -eq 'complete' }).Count 1 'Cached reuse did not complete'
    }

    Invoke-Check 'Changing suite apps reuses verified files and updates exclusions' {
        $plan = New-ReadyFixture
        $plan.Components = @('Word'); $plan.PlanId = ('b' * 64)
        $script:Downloads.Clear(); $script:Processes.Clear()
        Invoke-Deployment -Action Download -Plan $plan
        Assert-Equal $script:Downloads.Count 0 'Changing suite apps downloaded files again'
        Assert-Equal $script:Processes.Count 0 'Changing suite apps launched setup'
        [xml]$config = [IO.File]::ReadAllText((Join-Path $plan.DownloadPath 'Config.xml'))
        Assert-True ($null -ne $config.SelectSingleNode('/Configuration/Add/Product/ExcludeApp[@ID="Excel"]')) 'Removed Excel was not excluded'
        $null = Test-ReadyState -Plan $plan
    }

    Invoke-Check 'Cached reuse downloads only newly selected addons' {
        $plan = New-ReadyFixture -Components @('Word', 'Teams')
        $plan.Components = @('Word', 'Teams', 'OneDrive'); $plan.PlanId = ('c' * 64)
        $script:Downloads.Clear(); $script:Processes.Clear()
        Invoke-Deployment -Action Download -Plan $plan
        Assert-Equal $script:Downloads.Count 1 'Previously verified addons were downloaded again'
        Assert-Equal $script:Downloads[0].Label 'OneDrive' 'New addon was not downloaded'
        Assert-Equal $script:Processes.Count 0 'Addon download ran an installer'
        $null = Test-ReadyState -Plan $plan
    }

    Invoke-Check 'Interrupted legacy staging is reused and its version is pinned' {
        $plan = Test-DeploymentPlan -Plan (New-TestPlan -Folder 'InterruptedDownload')
        $partial = $plan | ConvertTo-Json -Depth 6 | ConvertFrom-Json
        $partial.DownloadPath = Join-Path $plan.DownloadPath ('.kareer-download-' + [Guid]::NewGuid().ToString('N'))
        New-TestPayload -Plan $partial
        [IO.File]::WriteAllText((Join-Path $partial.DownloadPath 'Config.xml'), (New-OfficeConfiguration -Plan $partial))
        $resume = Get-ResumeStage -Plan $plan
        Assert-Equal $resume.Path $partial.DownloadPath 'Partial download was discarded'
        Assert-Equal $resume.Version '16.0.19127.20154' 'Partial version was not pinned'
        $script:Processes.Clear()
        Invoke-Deployment -Action Download -Plan $plan
        Assert-True ($script:Processes[0].Arguments.Contains($partial.DownloadPath)) 'Retry did not pass the same folder to Microsoft'
        $null = Test-ReadyState -Plan $plan
    }

    Invoke-Check 'Other channel or language staging is never reused' {
        $plan = Test-DeploymentPlan -Plan (New-TestPlan -Folder 'DifferentProfile')
        $partial = $plan | ConvertTo-Json -Depth 6 | ConvertFrom-Json
        $partial.DownloadPath = Join-Path $plan.DownloadPath ('.kareer-download-' + [Guid]::NewGuid().ToString('N'))
        New-TestPayload -Plan $partial
        $partial.Language = 'fr-fr'
        [IO.File]::WriteAllText((Join-Path $partial.DownloadPath 'Config.xml'), (New-OfficeConfiguration -Plan $partial))
        Assert-True ($null -eq (Get-ResumeStage -Plan $plan)) 'Other-language stage was reused'
        $partial.Language = 'en-us'; $partial.Branch = 'ProPlus2021Volume'; $partial.Channel = 'PerpetualVL2021'
        [IO.File]::WriteAllText((Join-Path $partial.DownloadPath 'Config.xml'), (New-OfficeConfiguration -Plan $partial))
        Assert-True ($null -eq (Get-ResumeStage -Plan $plan)) 'Other-channel stage was reused'
    }

    Invoke-Check 'Process runner propagates errors and records restart-required exit codes' {
        # The real runner starts only a harmless Windows command shell here; it
        # immediately exits with a chosen code and never invokes an installer.
        $plan = New-ReadyFixture
        $cmd = Join-Path $env:WINDIR 'System32\cmd.exe'
        Assert-Throws { & $script:RealDeploymentProcess -FilePath $cmd -Arguments '/d /c exit 7' -Label 'Test process' -Plan $plan -SuccessExitCodes @(0) } 'A failed process was reported as successful'
        $script:RebootRequired = $false
        $result = & $script:RealDeploymentProcess -FilePath $cmd -Arguments '/d /c exit 3010' -Label 'Test process' -Plan $plan -SuccessExitCodes @(0, 3010)
        Assert-Equal $result.ExitCode 3010 'Restart-required process result was lost'
        Assert-True $script:RebootRequired 'Restart-required exit code was not recorded'
    }
} finally {
    # Guard recursive cleanup to the unique directory created by this test run.
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    $resolvedTestRoot = [IO.Path]::GetFullPath($script:TestRoot)
    if ($resolvedTestRoot.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -and
        [IO.Path]::GetFileName($resolvedTestRoot).StartsWith('KareerOfficeStudio-Tests-')) {
        if (Test-Path -LiteralPath $resolvedTestRoot) { Remove-Item -LiteralPath $resolvedTestRoot -Recurse -Force }
    } else {
        throw 'Refusing to clean up a test directory outside the temporary folder.'
    }
}

Write-Host ('{0} passed; {1} failed.' -f $script:Passed, $script:Failures.Count)
if ($script:Failures.Count -gt 0) {
    throw ($script:Failures -join [Environment]::NewLine)
}
