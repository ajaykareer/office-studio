#requires -Version 5.1
<#
  KAREER Office Studio deployment worker. Uses only Microsoft's deployment tools.
  -Library loads the functions for offline validation and tests without taking action.
#>
[CmdletBinding()]
param(
    [ValidateSet('Download', 'Install', 'Validate')][string]$Action,
    [string]$PlanPath,
    [string]$StartSignal,
    [switch]$Library
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$script:RebootRequired = $false

function Write-DeploymentEvent {
    param([string]$Kind, [string]$Message, [int]$Percent = -2, [string]$Stage = '', [long]$BytesOnDisk = -1)
    $event = [ordered]@{ kind = $Kind; message = $Message }
    if ($Percent -ne -2) { $event.percent = $Percent }
    if ($Stage) { $event.stage = $Stage }
    if ($BytesOnDisk -ge 0) { $event.bytesOnDisk = $BytesOnDisk }
    [Console]::Out.WriteLine('@@' + (ConvertTo-Json -InputObject $event -Compress -Depth 8))
    [Console]::Out.Flush()
}

function Assert-NoReparsePoint {
    param([string]$Path)
    $cursor = [System.IO.Path]::GetFullPath($Path)
    while ($cursor) {
        if (Test-Path -LiteralPath $cursor) {
            $item = Get-Item -LiteralPath $cursor -Force
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw "For predictable file handling, choose a folder without symbolic links or junctions: $cursor"
            }
        }
        $parent = [IO.Path]::GetDirectoryName($cursor.TrimEnd('\'))
        if (-not $parent -or $parent -eq $cursor) { break }
        $cursor = $parent
    }
}

function Test-DeploymentPlan {
    param([Parameter(Mandatory = $true)]$Plan)
    foreach ($name in @('PlanId', 'Branch', 'Channel', 'Components', 'Architecture', 'Language', 'DownloadPath')) {
        if ($null -eq $Plan.PSObject.Properties[$name] -or $null -eq $Plan.$name) { throw "The deployment plan is missing $name." }
    }
    if ([string]$Plan.PlanId -notmatch '^[A-Fa-f0-9]{64}$') { throw 'The deployment plan identifier is invalid.' }
    $channels = @{
        ProPlus2019Retail = 'Current'; ProPlus2021Volume = 'PerpetualVL2021'
        ProPlus2024Volume = 'PerpetualVL2024'; O365ProPlusRetail = 'Current'
    }
    $branch = [string]$Plan.Branch
    if (-not $channels.ContainsKey($branch)) { throw 'Choose a supported Office edition.' }
    if ([string]$Plan.Channel -cne $channels[$branch]) { throw 'The Office update channel does not match the selected edition.' }
    if ([string]$Plan.Architecture -notin @('32', '64')) { throw 'Office architecture must be 32 or 64.' }
    if ([string]$Plan.Architecture -eq '64' -and -not [Environment]::Is64BitOperatingSystem) { throw '64-bit Office requires 64-bit Windows.' }
    $language = ([string]$Plan.Language).ToLowerInvariant()
    if ($language -notin @('en-us', 'en-gb', 'fr-fr', 'es-es', 'de-de', 'it-it', 'pt-br', 'hi-in')) { throw 'Choose a supported Office language.' }
    $allowed = @('Word', 'Excel', 'PowerPoint', 'Outlook', 'OneNote', 'Access', 'Publisher', 'OneDrive', 'Teams')
    $components = @()
    foreach ($component in @($Plan.Components)) {
        $match = @($allowed | Where-Object { $_ -ceq [string]$component })
        if ($match.Count -ne 1) { throw "Unsupported Office component: $component" }
        if ($components -contains [string]$component) { throw 'The deployment plan contains a duplicate component.' }
        $components += [string]$component
    }
    if (@($components | Where-Object { $_ -notin @('OneDrive', 'Teams') }).Count -eq 0) { throw 'Select at least one Office application, such as Word or Excel.' }
    if ($components -contains 'Publisher' -and $branch -in @('ProPlus2024Volume', 'O365ProPlusRetail')) {
        throw 'Publisher is not available for Office 2024 or Microsoft 365 in this installer.'
    }
    $destination = [string]$Plan.DownloadPath
    if ($destination -notmatch '^(?:[A-Za-z]:\\|\\\\[^\\]+\\[^\\]+\\)' -or $destination -match '["<>|?*\x00-\x1f]') {
        throw 'The download folder must be a full Windows folder path.'
    }
    if ($destination.Substring(2).Contains(':')) { throw 'The download folder cannot contain an alternate data stream.' }
    $destination = [IO.Path]::GetFullPath($destination).TrimEnd('\')
    if ($destination -eq [IO.Path]::GetPathRoot($destination).TrimEnd('\')) { throw 'Choose a dedicated folder, rather than a drive root.' }
    foreach ($segment in $destination.Split('\')) {
        if ($segment.EndsWith('.') -or $segment.EndsWith(' ')) { throw 'Folder names cannot end with a period or space.' }
    }
    Assert-NoReparsePoint -Path $destination
    if ((Test-Path -LiteralPath $destination) -and -not (Test-Path -LiteralPath $destination -PathType Container)) { throw 'The download destination is a file, not a folder.' }
    return [pscustomobject][ordered]@{
        PlanId = ([string]$Plan.PlanId).ToLowerInvariant(); Branch = $branch; Channel = $channels[$branch]
        Components = @($components | Sort-Object); Architecture = [string]$Plan.Architecture
        Language = $language; DownloadPath = $destination
    }
}

function Read-DeploymentPlan {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw 'The deployment plan file could not be found.' }
    $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    return Test-DeploymentPlan -Plan $raw
}

function New-OfficeConfiguration {
    param([Parameter(Mandatory = $true)]$Plan, [string]$OfficeVersion = '')
    $Plan = Test-DeploymentPlan -Plan $Plan
    if ($OfficeVersion -and $OfficeVersion -notmatch '^16\.0\.\d+\.\d+$') { throw 'Invalid downloaded Office version.' }
    $xml = New-Object System.Xml.XmlDocument
    [void]$xml.AppendChild($xml.CreateXmlDeclaration('1.0', 'utf-8', $null))
    $root = $xml.CreateElement('Configuration'); [void]$xml.AppendChild($root)
    $add = $xml.CreateElement('Add'); [void]$root.AppendChild($add)
    $add.SetAttribute('OfficeClientEdition', $Plan.Architecture)
    $add.SetAttribute('Channel', $Plan.Channel)
    $add.SetAttribute('SourcePath', $Plan.DownloadPath)
    $add.SetAttribute('AllowCdnFallback', 'FALSE')
    if ($OfficeVersion) { $add.SetAttribute('Version', $OfficeVersion) }
    $product = $xml.CreateElement('Product'); $product.SetAttribute('ID', $Plan.Branch); [void]$add.AppendChild($product)
    $language = $xml.CreateElement('Language'); $language.SetAttribute('ID', $Plan.Language); [void]$product.AppendChild($language)
    $suiteComponents = @('Word', 'Excel', 'PowerPoint', 'Outlook', 'OneNote', 'Access')
    if ($Plan.Branch -notin @('ProPlus2024Volume', 'O365ProPlusRetail')) { $suiteComponents += 'Publisher' }
    $excluded = $suiteComponents | Where-Object { $_ -notin $Plan.Components }
    foreach ($id in @($excluded) + @('Groove', 'Lync', 'OneDrive', 'Teams', 'OutlookForWindows')) {
        $exclude = $xml.CreateElement('ExcludeApp'); $exclude.SetAttribute('ID', $id); [void]$product.AppendChild($exclude)
    }
    $display = $xml.CreateElement('Display'); $display.SetAttribute('Level', 'Full'); $display.SetAttribute('AcceptEULA', 'TRUE'); [void]$root.AppendChild($display)
    $property = $xml.CreateElement('Property'); $property.SetAttribute('Name', 'FORCEAPPSHUTDOWN'); $property.SetAttribute('Value', 'FALSE'); [void]$root.AppendChild($property)
    return $xml.OuterXml
}

function Get-NativeArchitecture {
    $architecture = $env:PROCESSOR_ARCHITEW6432
    if (-not $architecture) { $architecture = $env:PROCESSOR_ARCHITECTURE }
    if ($architecture -eq 'ARM64') { return 'arm64' }
    if ([Environment]::Is64BitOperatingSystem) { return 'x64' }
    return 'x86'
}

function Get-AddonDownloads {
    param($Plan)
    $native = Get-NativeArchitecture
    if ($Plan.Components -contains 'OneDrive') {
        $linkId = @{ x64 = '844652'; x86 = '2181213'; arm64 = '2282608' }[$native]
        [pscustomobject]@{ Name = 'OneDrive'; RelativePath = 'Addons\OneDriveSetup.exe'; Url = "https://go.microsoft.com/fwlink/?linkid=$linkId" }
    }
    if ($Plan.Components -contains 'Teams') {
        [pscustomobject]@{ Name = 'Teams installer'; RelativePath = 'Addons\teamsbootstrapper.exe'; Url = 'https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409' }
        $linkId = @{ x64 = '2196106'; x86 = '2196060'; arm64 = '2196207' }[$native]
        [pscustomobject]@{ Name = 'Teams'; RelativePath = 'Addons\MSTeams.msix'; Url = "https://go.microsoft.com/fwlink/?linkid=$linkId&clcid=0x409" }
    }
}

function Assert-MicrosoftSignature {
    param([Parameter(Mandatory = $true)][string]$Path)
    Assert-NoReparsePoint -Path $Path
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "A required Microsoft installer is missing: $Path" }
    $signature = Get-AuthenticodeSignature -LiteralPath $Path
    if ($signature.Status -ne 'Valid' -or $null -eq $signature.SignerCertificate -or $signature.SignerCertificate.Subject -notmatch '(?:^|,\s*)O=Microsoft Corporation(?:,|$)') {
        throw "Microsoft signature verification failed for $([IO.Path]::GetFileName($Path)) ($($signature.Status)). Download again; do not run this file manually."
    }
}

function Save-MicrosoftDownload {
    param([string]$Url, [string]$Path, [string]$Label)
    Assert-NoReparsePoint -Path $Path
    $parent = [IO.Path]::GetDirectoryName($Path)
    [void][IO.Directory]::CreateDirectory($parent)
    $temporary = Join-Path $parent (([IO.Path]::GetFileNameWithoutExtension($Path)) + '.' + [Guid]::NewGuid().ToString('N') + '.download' + [IO.Path]::GetExtension($Path))
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $uri = [Uri]$Url
        $response = $null
        for ($redirect = 0; $redirect -lt 10; $redirect++) {
            if ($uri.Scheme -ne 'https' -or $uri.UserInfo) { throw 'Microsoft download redirected to an unsupported address.' }
            $request = [Net.HttpWebRequest]::Create($uri)
            $request.AllowAutoRedirect = $false; $request.Timeout = 60000; $request.ReadWriteTimeout = 60000
            $request.UserAgent = 'KAREER-Office-Studio/1.0'
            $response = $request.GetResponse()
            if ([int]$response.StatusCode -in @(301, 302, 303, 307, 308)) {
                $location = $response.Headers['Location']; $response.Dispose(); $response = $null
                if (-not $location) { throw 'Microsoft download returned an empty redirect.' }
                $uri = New-Object Uri($uri, $location)
            } else { break }
        }
        if ($null -eq $response -or [int]$response.StatusCode -ne 200) { throw 'Microsoft download could not be completed.' }
        $inputStream = $null; $outputStream = $null
        try {
            $inputStream = $response.GetResponseStream()
            $outputStream = [IO.File]::Open($temporary, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
            $buffer = New-Object byte[] 262144; [long]$received = 0; [long]$expected = $response.ContentLength
            $timer = [Diagnostics.Stopwatch]::StartNew()
            Write-DeploymentEvent -Kind progress -Message "Downloading $Label..." -Percent -1
            while (($count = $inputStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                $outputStream.Write($buffer, 0, $count); $received += $count
                if ($timer.ElapsedMilliseconds -ge 1000) {
                    $percentage = -1
                    if ($expected -gt 0) { $percentage = [int][Math]::Min(99, [Math]::Floor(100.0 * $received / $expected)) }
                    Write-DeploymentEvent -Kind progress -Message ("Downloading {0}: {1:N1} MB" -f $Label, ($received / 1MB)) -Percent $percentage
                    $timer.Restart()
                }
            }
            if ($received -le 0 -or ($expected -gt 0 -and $received -ne $expected)) { throw "The $Label download was incomplete." }
        } finally {
            if ($null -ne $outputStream) { $outputStream.Dispose() }
            if ($null -ne $inputStream) { $inputStream.Dispose() }
            $response.Dispose()
        }
        Write-DeploymentEvent -Kind status -Message "Verifying Microsoft's signature for $Label..."
        Assert-MicrosoftSignature -Path $temporary
        Move-Item -LiteralPath $temporary -Destination $Path -Force
        Write-DeploymentEvent -Kind log -Message "$Label downloaded and Microsoft signature verified."
    } finally {
        if (Test-Path -LiteralPath $temporary -PathType Leaf) { Remove-Item -LiteralPath $temporary -Force }
    }
}

function Get-OfficePayload {
    param($Plan, [string]$OfficeVersion = '')
    $data = Join-Path $Plan.DownloadPath 'Office\Data'
    Assert-NoReparsePoint -Path $data
    if (-not (Test-Path -LiteralPath $data -PathType Container)) { throw 'The downloaded Office data folder is missing. Download the selected edition first.' }
    $architecture = 'x86'; if ($Plan.Architecture -eq '64') { $architecture = 'x64' }
    $candidates = @(Get-ChildItem -LiteralPath $data -Directory | Where-Object { $_.Name -match '^16\.0\.\d+\.\d+$' } | Sort-Object { [version]$_.Name } -Descending)
    if ($OfficeVersion) { $candidates = @($candidates | Where-Object { $_.Name -ceq $OfficeVersion }) }
    $versionDirectory = $null
    foreach ($candidate in $candidates) {
        Assert-NoReparsePoint -Path $candidate.FullName
        $neutral = Join-Path $candidate.FullName "stream.$architecture.x-none.dat"
        $localized = Join-Path $candidate.FullName "stream.$architecture.$($Plan.Language).dat"
        if ((Test-Path -LiteralPath $neutral -PathType Leaf) -and (Test-Path -LiteralPath $localized -PathType Leaf)) {
            if ((Get-Item -LiteralPath $neutral).Length -gt 0 -and (Get-Item -LiteralPath $localized).Length -gt 0) { $versionDirectory = $candidate; break }
        }
    }
    if ($null -eq $versionDirectory) { throw "Required $architecture Office and $($Plan.Language) language files are missing or incomplete. Download again." }
    $versionCab = Join-Path $data ('v' + $Plan.Architecture + '.cab')
    if (-not (Test-Path -LiteralPath $versionCab -PathType Leaf) -or (Get-Item -LiteralPath $versionCab).Length -eq 0) { throw 'The Office version catalog is missing or incomplete. Download again.' }
    # Reject linked subdirectories before traversing the payload tree.
    $directories = New-Object 'System.Collections.Generic.Queue[string]'
    $directories.Enqueue($data)
    $files = New-Object 'System.Collections.Generic.List[object]'
    while ($directories.Count -gt 0) {
        $directory = $directories.Dequeue()
        foreach ($entry in Get-ChildItem -LiteralPath $directory -Force) {
            if (($entry.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw "The Office cache contains a symbolic link or junction: $($entry.FullName)" }
            if ($entry.PSIsContainer) { $directories.Enqueue($entry.FullName) } else { $files.Add($entry) }
        }
    }
    return [pscustomobject]@{ Version = $versionDirectory.Name; Files = @($files.ToArray() | Sort-Object FullName) }
}

function Get-RequiredFilePaths {
    param($Plan, $Payload)
    $paths = @((Join-Path $Plan.DownloadPath 'Config.xml'), (Join-Path $Plan.DownloadPath 'setup.exe'))
    foreach ($file in $Payload.Files) { $paths += $file.FullName }
    foreach ($addon in @(Get-AddonDownloads -Plan $Plan)) { $paths += Join-Path $Plan.DownloadPath $addon.RelativePath }
    return @($paths | Sort-Object -Unique)
}

function Get-FileRecord {
    param($Plan, [string]$Path)
    Assert-NoReparsePoint -Path $Path
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "A downloaded file is missing: $Path" }
    $file = Get-Item -LiteralPath $Path
    if ($file.Length -le 0) { throw "A downloaded file is empty: $Path" }
    return [pscustomobject][ordered]@{
        RelativePath = $file.FullName.Substring($Plan.DownloadPath.Length + 1)
        Length = $file.Length
        Sha256 = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
    }
}

function Write-ReadyState {
    param($Plan, $Payload)
    $records = @(); $paths = @(Get-RequiredFilePaths -Plan $Plan -Payload $Payload); $index = 0
    foreach ($path in $paths) {
        $index++
        Write-DeploymentEvent -Kind progress -Message "Checking downloaded file $index of $($paths.Count)..." -Percent -1
        $records += Get-FileRecord -Plan $Plan -Path $path
    }
    $ready = [pscustomobject][ordered]@{
        SchemaVersion = 1; PlanId = $Plan.PlanId; Plan = $Plan; CreatedUtc = [DateTime]::UtcNow.ToString('o')
        OfficeVersion = $Payload.Version; NativeArchitecture = (Get-NativeArchitecture); Files = $records
    }
    $readyPath = Join-Path $Plan.DownloadPath 'ready.json'
    $temporary = Join-Path $Plan.DownloadPath ('ready.' + [Guid]::NewGuid().ToString('N') + '.tmp')
    try {
        [IO.File]::WriteAllText($temporary, ($ready | ConvertTo-Json -Depth 10), (New-Object Text.UTF8Encoding($false)))
        Move-Item -LiteralPath $temporary -Destination $readyPath -Force
    } finally { if (Test-Path -LiteralPath $temporary -PathType Leaf) { Remove-Item -LiteralPath $temporary -Force } }
    return $ready
}

function Test-ReadyState {
    param([Parameter(Mandatory = $true)]$Plan)
    $Plan = Test-DeploymentPlan -Plan $Plan
    $readyPath = Join-Path $Plan.DownloadPath 'ready.json'
    Assert-NoReparsePoint -Path $readyPath
    if (-not (Test-Path -LiteralPath $readyPath -PathType Leaf)) { throw 'This selection has not completed downloading. Download it first.' }
    $ready = Get-Content -LiteralPath $readyPath -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($property in @('SchemaVersion', 'PlanId', 'Plan', 'OfficeVersion', 'NativeArchitecture', 'Files')) {
        if ($null -eq $ready.PSObject.Properties[$property]) { throw 'The download verification record is incomplete. Download again.' }
    }
    if ($ready.SchemaVersion -ne 1 -or $ready.PlanId -cne $Plan.PlanId) { throw 'The downloaded files belong to a different selection. Download the current selection first.' }
    $recordedPlan = Test-DeploymentPlan -Plan $ready.Plan
    if (($recordedPlan | ConvertTo-Json -Compress -Depth 6) -cne ($Plan | ConvertTo-Json -Compress -Depth 6)) { throw 'The download settings have changed. Download the current selection first.' }
    if ($ready.NativeArchitecture -cne (Get-NativeArchitecture)) { throw 'These files were prepared for a different Windows architecture. Download again on this PC.' }
    if ([string]$ready.OfficeVersion -notmatch '^16\.0\.\d+\.\d+$') { throw 'The downloaded Office version record is invalid.' }
    $configPath = Join-Path $Plan.DownloadPath 'Config.xml'
    Assert-NoReparsePoint -Path $configPath
    if (-not (Test-Path -LiteralPath $configPath -PathType Leaf)) { throw 'The Office configuration is missing. Download again.' }
    $expectedConfig = New-OfficeConfiguration -Plan $Plan -OfficeVersion $ready.OfficeVersion
    if ([IO.File]::ReadAllText($configPath) -cne $expectedConfig) { throw 'The Office configuration has been changed. Download again before installing.' }
    $payload = Get-OfficePayload -Plan $Plan -OfficeVersion $ready.OfficeVersion
    $paths = @(Get-RequiredFilePaths -Plan $Plan -Payload $payload)
    if (@($ready.Files).Count -ne $paths.Count) { throw 'The downloaded file list has changed. Download again.' }
    $recordMap = @{}
    foreach ($record in @($ready.Files)) {
        foreach ($field in @('RelativePath', 'Length', 'Sha256')) {
            if ($null -eq $record.PSObject.Properties[$field]) { throw 'A downloaded file verification record is incomplete.' }
        }
        if ($recordMap.ContainsKey([string]$record.RelativePath)) { throw 'The downloaded file verification record contains duplicates.' }
        $recordMap.Add([string]$record.RelativePath, $record)
    }
    $index = 0
    foreach ($path in $paths) {
        $index++
        Write-DeploymentEvent -Kind progress -Message "Verifying file $index of $($paths.Count) before installation..." -Percent -1
        $actual = Get-FileRecord -Plan $Plan -Path $path
        if (-not $recordMap.ContainsKey($actual.RelativePath)) { throw 'A downloaded file is absent from the verification record. Download again.' }
        $record = $recordMap[$actual.RelativePath]
        if ($actual.Length -ne $record.Length -or $actual.Sha256 -cne [string]$record.Sha256) { throw "A downloaded file changed or is damaged: $($actual.RelativePath). Download again." }
    }
    Assert-MicrosoftSignature -Path (Join-Path $Plan.DownloadPath 'setup.exe')
    foreach ($addon in @(Get-AddonDownloads -Plan $Plan)) { Assert-MicrosoftSignature -Path (Join-Path $Plan.DownloadPath $addon.RelativePath) }
    return $ready
}

function Invoke-DeploymentProcess {
    param([string]$FilePath, [string]$Arguments, [string]$Label, $Plan, [int[]]$SuccessExitCodes = @(0, 3010), [string]$MonitorPath = '')
    Assert-MicrosoftSignature -Path $FilePath
    $info = New-Object Diagnostics.ProcessStartInfo
    $info.FileName = $FilePath; $info.Arguments = $Arguments; $info.WorkingDirectory = $Plan.DownloadPath
    $info.UseShellExecute = $false; $info.CreateNoWindow = $true; $info.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
    $info.RedirectStandardOutput = $true; $info.RedirectStandardError = $true
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $info
    try {
        Write-DeploymentEvent -Kind progress -Message $Label -Percent -1
        if (-not $process.Start()) { throw 'The Microsoft installer did not start.' }
        $stdout = $process.StandardOutput.ReadToEndAsync(); $stderr = $process.StandardError.ReadToEndAsync()
        $timer = [Diagnostics.Stopwatch]::StartNew()
        $elapsed = [Diagnostics.Stopwatch]::StartNew()
        while (-not $process.WaitForExit(1000)) {
            if ($timer.ElapsedMilliseconds -ge 10000) {
                if ($MonitorPath) {
                    # Office does not expose a reliable total or percentage. Count
                    # bytes present in this operation's isolated download folder.
                    [long]$receivedBytes = 0
                    $folders = New-Object 'System.Collections.Generic.Queue[string]'
                    $folders.Enqueue($MonitorPath)
                    while ($folders.Count -gt 0) {
                        $folder = $folders.Dequeue()
                        foreach ($entry in @(Get-ChildItem -LiteralPath $folder -Force -ErrorAction SilentlyContinue)) {
                            if (($entry.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
                            if ($entry.PSIsContainer) { $folders.Enqueue($entry.FullName) }
                            else {
                                # Windows directory enumeration can report length zero
                                # until the downloader closes an open stream. Query the
                                # file itself instead of relying on cached directory data.
                                try { $entry.Refresh(); $receivedBytes += $entry.Length } catch { }
                            }
                        }
                    }
                    $message = 'Microsoft setup is running. Office files on disk: {0:N1} MB. File sizes may include reserved space.' -f ($receivedBytes / 1MB)
                    Write-DeploymentEvent -Kind progress -Message $message -Percent -1 -BytesOnDisk $receivedBytes
                } else {
                    Write-DeploymentEvent -Kind progress -Message "$Label Please keep Office Studio open." -Percent -1
                }
                $timer.Restart()
            }
        }
        $process.WaitForExit()
        $output = [string]$stdout.Result; $errors = [string]$stderr.Result
        foreach ($line in @($output -split '\r?\n') + @($errors -split '\r?\n')) {
            if ($line.Trim()) { Write-DeploymentEvent -Kind log -Message $line }
        }
        if ($process.ExitCode -notin $SuccessExitCodes) { throw "$Label failed (exit code $($process.ExitCode)). Check the activity log and Microsoft's installer message." }
        if ($process.ExitCode -eq 3010) { $script:RebootRequired = $true }
        return [pscustomobject]@{ ExitCode = $process.ExitCode; Output = $output; ErrorOutput = $errors }
    } finally { $process.Dispose() }
}

function Get-ReusableDownload {
    param($Plan)
    $marker = Join-Path $Plan.DownloadPath 'ready.json'
    if (-not (Test-Path -LiteralPath $marker -PathType Leaf)) { return $null }
    try {
        Assert-NoReparsePoint -Path $marker
        $saved = Get-Content -LiteralPath $marker -Raw -Encoding UTF8 | ConvertFrom-Json
        $previous = Test-DeploymentPlan -Plan $saved.Plan
        foreach ($name in @('Branch', 'Channel', 'Architecture', 'Language', 'DownloadPath')) {
            if ($previous.$name -cne $Plan.$name) { return $null }
        }
        Write-DeploymentEvent -Kind stage -Stage 'checking-saved' -Message 'Checking your saved files. A complete matching package can be reused without another Office download.'
        return (Test-ReadyState -Plan $previous)
    } catch {
        Write-DeploymentEvent -Kind log -Message ('Saved package could not be reused: ' + $_.Exception.Message)
        return $null
    }
}

function Get-ResumeStage {
    param($Plan)
    # Includes partial downloads created by older Studio versions. Only reuse a
    # stage whose own configuration identifies this exact Office download profile.
    foreach ($folder in @(Get-ChildItem -LiteralPath $Plan.DownloadPath -Directory -Force | Where-Object { $_.Name -match '^\.kareer-download-[a-f0-9]{32}$' } | Sort-Object LastWriteTimeUtc -Descending)) {
        try {
            Assert-NoReparsePoint -Path $folder.FullName
            $configuration = Join-Path $folder.FullName 'Config.xml'
            Assert-NoReparsePoint -Path $configuration
            $xml = New-Object Xml.XmlDocument
            $xml.XmlResolver = $null
            $xml.Load($configuration)
            $add = $xml.SelectSingleNode('/Configuration/Add')
            $products = @($xml.SelectNodes('/Configuration/Add/Product'))
            $languages = @($xml.SelectNodes('/Configuration/Add/Product/Language'))
            if ($null -eq $add -or $products.Count -ne 1 -or $languages.Count -ne 1) { continue }
            if ($add.GetAttribute('SourcePath') -ine $folder.FullName -or $add.GetAttribute('OfficeClientEdition') -cne $Plan.Architecture -or $add.GetAttribute('Channel') -cne $Plan.Channel -or $products[0].GetAttribute('ID') -cne $Plan.Branch -or $languages[0].GetAttribute('ID') -cne $Plan.Language) { continue }
            $version = $add.GetAttribute('Version')
            if ($version -and $version -notmatch '^16\.0\.\d+\.\d+$') { continue }
            $data = Join-Path $folder.FullName 'Office\Data'
            Assert-NoReparsePoint -Path $data
            if (-not $version -and (Test-Path -LiteralPath $data -PathType Container)) {
                $versions = @(Get-ChildItem -LiteralPath $data -Directory | Where-Object { $_.Name -match '^16\.0\.\d+\.\d+$' } | Sort-Object { [version]$_.Name } -Descending)
                if ($versions.Count -gt 0) { Assert-NoReparsePoint -Path $versions[0].FullName; $version = $versions[0].Name }
            }
            return [pscustomobject]@{ Path = $folder.FullName; Version = $version }
        } catch { Write-DeploymentEvent -Kind log -Message ('Skipping an unusable partial download: ' + $_.Exception.Message) }
    }
    return $null
}

function Invoke-DeploymentCore {
    param([ValidateSet('Download', 'Install', 'Validate')][string]$Action, $Plan)
    $Plan = Test-DeploymentPlan -Plan $Plan
    $script:RebootRequired = $false
    $configPath = Join-Path $Plan.DownloadPath 'Config.xml'
    $setupPath = Join-Path $Plan.DownloadPath 'setup.exe'
    if ($Action -eq 'Validate') {
        Write-DeploymentEvent -Kind status -Message 'Checking the saved download against this selection...'
        $null = Test-ReadyState -Plan $Plan
        Write-DeploymentEvent -Kind complete -Message 'The saved download is verified and ready to install.' -Percent 100
        return
    }
    if ($Action -eq 'Download') {
        Write-DeploymentEvent -Kind stage -Stage 'preparing' -Message 'Preparing Microsoft setup and looking for matching saved files.'
        [void][IO.Directory]::CreateDirectory($Plan.DownloadPath)
        $readyPath = Join-Path $Plan.DownloadPath 'ready.json'
        Assert-NoReparsePoint -Path $readyPath
        $cached = Get-ReusableDownload -Plan $Plan
        if ($null -ne $cached) {
            if ($cached.PlanId -ceq $Plan.PlanId -and ($cached.Plan | ConvertTo-Json -Compress -Depth 6) -ceq ($Plan | ConvertTo-Json -Compress -Depth 6)) {
                Write-DeploymentEvent -Kind complete -Message ('Saved Office ' + $cached.OfficeVersion + ' verified. No download was needed. Your apps are ready to install.') -Percent 100
                return
            }
            Remove-Item -LiteralPath $readyPath -Force
            foreach ($addon in @(Get-AddonDownloads -Plan $Plan)) {
                if (@(Get-AddonDownloads -Plan $cached.Plan | ForEach-Object { $_.RelativePath }) -notcontains $addon.RelativePath) {
                    Write-DeploymentEvent -Kind stage -Stage 'downloading' -Message ('Reusing your Office package and downloading the selected ' + $addon.Name + ' addon.')
                    Save-MicrosoftDownload -Url $addon.Url -Path (Join-Path $Plan.DownloadPath $addon.RelativePath) -Label $addon.Name
                }
            }
            Assert-NoReparsePoint -Path $configPath
            [IO.File]::WriteAllText($configPath, (New-OfficeConfiguration -Plan $Plan -OfficeVersion $cached.OfficeVersion), (New-Object Text.UTF8Encoding($false)))
            Write-DeploymentEvent -Kind stage -Stage 'checking' -Message 'Checking the saved package against your new app selection.'
            $null = Write-ReadyState -Plan $Plan -Payload (Get-OfficePayload -Plan $Plan -OfficeVersion $cached.OfficeVersion)
            Write-DeploymentEvent -Kind complete -Message ('Reused Office ' + $cached.OfficeVersion + ' for your selected apps. Ready to install.') -Percent 100
            return
        }
        if (Test-Path -LiteralPath $readyPath -PathType Leaf) { Remove-Item -LiteralPath $readyPath -Force }
        Write-DeploymentEvent -Kind stage -Stage 'preparing' -Message 'Getting Microsoft setup ready for your download.'
        Assert-NoReparsePoint -Path $configPath
        $toolReady = $false
        if (Test-Path -LiteralPath $setupPath -PathType Leaf) {
            try { Assert-NoReparsePoint -Path $setupPath; Assert-MicrosoftSignature -Path $setupPath; $toolReady = $true } catch { }
        }
        if (-not $toolReady) { Save-MicrosoftDownload -Url 'https://officecdn.microsoft.com/pr/wsus/setup.exe' -Path $setupPath -Label 'Office Deployment Tool' }
        $resume = Get-ResumeStage -Plan $Plan
        $stagePath = Join-Path $Plan.DownloadPath ('.kareer-download-' + [Guid]::NewGuid().ToString('N'))
        $resumeVersion = ''
        if ($null -ne $resume) {
            $stagePath = $resume.Path; $resumeVersion = $resume.Version
            Write-DeploymentEvent -Kind status -Message ('Resuming previous download. Microsoft will check existing files in ' + $stagePath)
        }
        [void][IO.Directory]::CreateDirectory($stagePath)
        $stagePlan = $Plan | ConvertTo-Json -Depth 6 | ConvertFrom-Json
        $stagePlan.DownloadPath = $stagePath
        $stageConfig = Join-Path $stagePath 'Config.xml'
        [IO.File]::WriteAllText($stageConfig, (New-OfficeConfiguration -Plan $stagePlan -OfficeVersion $resumeVersion), (New-Object Text.UTF8Encoding($false)))
        if ($null -ne $resume) {
            Write-DeploymentEvent -Kind stage -Stage 'reusing' -Message 'Checking and reusing your saved files. Microsoft will download any files it still needs.'
        } else {
            Write-DeploymentEvent -Kind stage -Stage 'downloading' -Message 'Downloading shared Office files for your selected apps.'
        }
        $null = Invoke-DeploymentProcess -FilePath $setupPath -Arguments ('/download "' + $stageConfig + '"') -Label 'Downloading Office files...' -Plan $Plan -SuccessExitCodes @(0) -MonitorPath (Join-Path $stagePath 'Office')
        Write-DeploymentEvent -Kind status -Message 'Microsoft finished the Office download. Saving the files to your package folder.'
        $stagedPayload = Get-OfficePayload -Plan $stagePlan -OfficeVersion $resumeVersion
        $expectedRoot = $Plan.DownloadPath.TrimEnd('\') + '\'
        foreach ($file in $stagedPayload.Files) {
            $relativePath = $file.FullName.Substring($stagePath.Length + 1)
            $destination = [IO.Path]::GetFullPath((Join-Path $Plan.DownloadPath $relativePath))
            if (-not $destination.StartsWith($expectedRoot, [StringComparison]::OrdinalIgnoreCase)) { throw 'A downloaded file has an invalid destination.' }
            Assert-NoReparsePoint -Path $destination
            [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($destination))
            Move-Item -LiteralPath $file.FullName -Destination $destination -Force
        }
        # Only this invocation's uniquely named staging directory is removed.
        $resolvedStage = [IO.Path]::GetFullPath($stagePath)
        if (-not $resolvedStage.StartsWith($expectedRoot, [StringComparison]::OrdinalIgnoreCase) -or [IO.Path]::GetFileName($resolvedStage) -notmatch '^\.kareer-download-[a-f0-9]{32}$') {
            throw 'The staging folder could not be safely cleaned up.'
        }
        Assert-NoReparsePoint -Path $resolvedStage
        Remove-Item -LiteralPath $resolvedStage -Recurse -Force
        $payload = Get-OfficePayload -Plan $Plan -OfficeVersion $stagedPayload.Version
        foreach ($addon in @(Get-AddonDownloads -Plan $Plan)) {
            Write-DeploymentEvent -Kind stage -Stage 'downloading' -Message ('Office files are saved. Downloading the selected ' + $addon.Name + ' addon.')
            Save-MicrosoftDownload -Url $addon.Url -Path (Join-Path $Plan.DownloadPath $addon.RelativePath) -Label $addon.Name
        }
        # Pin /configure to the exact version that was downloaded successfully.
        [IO.File]::WriteAllText($configPath, (New-OfficeConfiguration -Plan $Plan -OfficeVersion $payload.Version), (New-Object Text.UTF8Encoding($false)))
        Write-DeploymentEvent -Kind stage -Stage 'checking' -Message 'Checking the downloaded files before making your package ready to install.'
        $null = Write-ReadyState -Plan $Plan -Payload $payload
        Write-DeploymentEvent -Kind complete -Message 'Download complete. Your selected applications are ready. Click Install when you are ready.' -Percent 100
        return
    }
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw 'Installation requires administrator access. Reopen KAREER Office Studio as administrator.' }
    Write-DeploymentEvent -Kind status -Message 'Verifying the download before installing...'
    $null = Test-ReadyState -Plan $Plan
    $null = Invoke-DeploymentProcess -FilePath $setupPath -Arguments ('/configure "' + $configPath + '"') -Label 'Installing your selected Office applications...' -Plan $Plan
    if ($Plan.Components -contains 'OneDrive') {
        $null = Invoke-DeploymentProcess -FilePath (Join-Path $Plan.DownloadPath 'Addons\OneDriveSetup.exe') -Arguments '/allusers /silent' -Label 'Installing OneDrive for this PC...' -Plan $Plan -SuccessExitCodes @(0, 3010)
    }
    if ($Plan.Components -contains 'Teams') {
        $result = Invoke-DeploymentProcess -FilePath (Join-Path $Plan.DownloadPath 'Addons\teamsbootstrapper.exe') -Arguments ('-p -o "' + (Join-Path $Plan.DownloadPath 'Addons\MSTeams.msix') + '"') -Label 'Installing Teams for this PC...' -Plan $Plan -SuccessExitCodes @(0)
        # Teams emits structured status; reject an explicit failure even if exit code is zero.
        if ($result.Output -match '"success"\s*:\s*false') { throw 'The Microsoft Teams installer reported failure. See the activity log for its error code.' }
    }
    $message = 'Installation completed successfully. Open your applications from the Start menu.'
    if ($script:RebootRequired) { $message = 'Installation completed. Restart Windows to finish setup.' }
    Write-DeploymentEvent -Kind complete -Message $message -Percent 100
}

function Invoke-Deployment {
    param([ValidateSet('Download', 'Install', 'Validate')][string]$Action, $Plan)
    $Plan = Test-DeploymentPlan -Plan $Plan
    if ($Action -eq 'Validate') {
        Invoke-DeploymentCore -Action $Action -Plan $Plan
        return
    }
    # A second copy of the launcher must not replace files while installation or
    # another download is using this folder. Validation itself remains read-only.
    $lockPath = Join-Path $Plan.DownloadPath '.kareer-operation.lock'
    Assert-NoReparsePoint -Path $lockPath
    $operationLock = $null
    $officeMutex = $null; $ownsOfficeMutex = $false
    try {
        $officeMutex = New-Object Threading.Mutex($false, 'Global\KAREER.OfficeStudio.OfficeOperation')
        try { $ownsOfficeMutex = $officeMutex.WaitOne(0) }
        catch [Threading.AbandonedMutexException] { $ownsOfficeMutex = $true }
        if (-not $ownsOfficeMutex) { throw 'Another Office Studio installation or removal is running. Wait for it to finish.' }
        if ($Action -eq 'Download') { [void][IO.Directory]::CreateDirectory($Plan.DownloadPath) }
        try { $operationLock = [IO.File]::Open($lockPath, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None) }
        catch [IO.IOException] { throw 'Another operation is using this download folder, or the folder is unavailable. Close the other installer and try again.' }
        Invoke-DeploymentCore -Action $Action -Plan $Plan
    } finally {
        if ($null -ne $operationLock) { $operationLock.Dispose() }
        if ($ownsOfficeMutex -and $null -ne $officeMutex) { $officeMutex.ReleaseMutex() }
        if ($null -ne $officeMutex) { $officeMutex.Dispose() }
    }
}

if (-not $Library) {
    try {
        if ($StartSignal) {
            # The launcher assigns this process to its private download job before
            # releasing this gate. A failed assignment must never start Office setup.
            $deadline = [DateTime]::UtcNow.AddSeconds(30)
            while (-not [IO.File]::Exists($StartSignal)) {
                if ([DateTime]::UtcNow -gt $deadline) { throw 'The download could not be attached to its cancellation controller.' }
                Start-Sleep -Milliseconds 100
            }
        }
        [Console]::OutputEncoding = New-Object Text.UTF8Encoding($false)
        if (-not $Action -or -not $PlanPath) { throw 'Both -Action and -PlanPath are required.' }
        $plan = Read-DeploymentPlan -Path $PlanPath
        Invoke-Deployment -Action $Action -Plan $plan
        exit 0
    } catch {
        Write-DeploymentEvent -Kind error -Message $_.Exception.Message
        exit 1
    }
}
