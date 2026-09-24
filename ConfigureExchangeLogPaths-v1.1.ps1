<#
.SYNOPSIS
Configures supported Exchange Server transport log paths, compares them with a reference server, or exports the required Set-* commands.

.DESCRIPTION
ConfigureExchangeLogPaths-v1.1.ps1 has two operating modes.

1. New root mode
   Run without -SourceServer. The script detects each target server's Exchange install path and relocates the higher-growth transport logs to a new root while preserving the Exchange directory structure below V15.
   If -NewRoot is omitted, the script displays the detected Exchange install path(s) and prompts for the new root.

   The following paths are relocated in this mode:
   - Message tracking
   - Receive protocol logs
   - Send protocol logs
   - Connectivity logs

   Lower-volume Agent and Routing paths are intentionally left unchanged in new root mode.
   DNS, pipeline tracing, mailbox delivery throttling, and other Reference-only additions are also not moved by NewRoot mode.

2. Reference server mode
   Use -SourceServer with -TargetServers. The script compares supported log path settings with the source server and builds the target configuration using install-path-aware translation.

   - If a source path is the Exchange default, the target receives its own equivalent default path based on the target Exchange install path.
   - If a source custom path is still under the source Exchange install path, the relative path below V15 is preserved under the target Exchange install path.
   - If a source custom path is outside the Exchange install path, the literal custom path is proposed for the target.

   Reference mode also includes supported Agent, Routing, DNS, mailbox delivery throttling, and pipeline tracing paths so an existing server's custom layout can be assessed more completely.
   PipelineTracingPath is Review Only because pipeline tracing can capture complete message contents; the script never relocates it automatically.

By default, the script shows a Current -> Proposed preview, asks for confirmation, applies only required Apply-policy changes, and verifies the target server(s) afterward.
Review Only paths remain visible in preview/comparison reports but are never changed automatically. If target drive availability is Missing or cannot be verified (Unknown), automatic apply is blocked for that setting.
When -OutputFile is specified without -CompareOnly, no Exchange changes are applied. The required Set-* commands are written to a TXT file for manual review.
When -CompareOnly is specified in reference mode, all supported source/target path values are displayed and no changes are applied. If -OutputFile is also specified, the same comparison is written to a TXT report; Set-* commands are not generated.
Before a confirmed Apply, the script saves a fresh TXT configuration snapshot under the script folder's ConfigSnapshots directory. Existing snapshot files are never overwritten.

The script reads the Exchange installation path from the msExchInstallPath attribute on the Exchange server object in Active Directory. This avoids assuming that Exchange is installed on C: or in the default Program Files location.

.PARAMETER SourceServer
Reference Exchange server. Valid only in Reference mode.

.PARAMETER TargetServers
One or more Exchange servers to compare/configure. Required in Reference mode. In New root mode, if omitted, the local computer is used.

.PARAMETER NewRoot
New root directory that replaces the Exchange installation root through V15 in New root mode. Example: D:\EXCLOG.
If omitted, the script prompts for the value.

.PARAMETER CompareOnly
Displays all supported source and target log path values without applying changes. Valid only in Reference mode. If -OutputFile is also specified, the comparison is written to a TXT report.

.PARAMETER OutputFile
Optional TXT output path. In normal mode, exports the required Exchange PowerShell commands instead of applying changes. With -CompareOnly, exports a comparison report instead of Set-* commands.

.PARAMETER Help
Displays a short usage guide and exits without initializing Exchange Management Shell or making changes.

.EXAMPLE
.\ConfigureExchangeLogPaths-v1.1.ps1
Detects the local Exchange install path, prompts for a new log root, previews changes for higher-growth logs, asks for confirmation, applies changes, and verifies the result.

.EXAMPLE
.\ConfigureExchangeLogPaths-v1.1.ps1 -TargetServers EX01,EX02 -NewRoot E:\EXCLOG
Moves the higher-growth supported transport logs on EX01 and EX02 under E:\EXCLOG while preserving the Exchange folder structure below V15.

.EXAMPLE
.\ConfigureExchangeLogPaths-v1.1.ps1 -SourceServer EX01 -TargetServers EX02,EX03
Uses EX01 as the reference. Default paths are translated to each target's own Exchange install path and custom paths are reproduced using the rules described above.

.EXAMPLE
.\ConfigureExchangeLogPaths-v1.1.ps1 -SourceServer EX01 -TargetServers EX02,EX03 -CompareOnly
Displays all supported log path settings for EX01 versus EX02 and EX03. No changes are applied.

.EXAMPLE
.\ConfigureExchangeLogPaths-v1.1.ps1 -SourceServer EX01 -TargetServers EX02,EX03 -OutputFile C:\Temp\ExchangeLogPath-Commands.txt
Does not change Exchange. Exports only the required Set-* commands to a TXT file for manual review.

.EXAMPLE
.\ConfigureExchangeLogPaths-v1.1.ps1 -SourceServer EX01 -TargetServers EX02,EX03 -CompareOnly -OutputFile C:\Temp\ExchangeLogPath-Compare.txt
Displays the comparison on screen and writes the same comparison to a TXT report. No Set-* commands are generated and no Exchange changes are applied.

.NOTES
Author  : Ceyhun Kirmizitas
Version : 1.1
Date    : 17/09/2026
Scope   : Exchange Server 2016, Exchange Server 2019, Exchange Server SE

Change Log
----------
1.1 - Expanded Reference mode coverage for DNS, pipeline tracing, and mailbox
      delivery throttling log paths. Added Apply/Review Only policies, null-default
      awareness, blocked automatic apply when target drive validation is Unknown,
      and automatic pre-change configuration snapshots before Apply.
1.0 - Initial release

This script changes supported Exchange transport log path properties only.
It does not relocate mailbox databases, mailbox transaction logs, queue database files, queue transaction logs, IIS logs, HttpProxy logs, or diagnostic log folders that do not have a supported Exchange Set-* path property in this scope.
#>

[CmdletBinding(DefaultParameterSetName = 'NewRoot')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Reference')]
    [ValidateNotNullOrEmpty()]
    [string]$SourceServer,

    [Parameter(Mandatory = $true, ParameterSetName = 'Reference')]
    [Parameter(Mandatory = $false, ParameterSetName = 'NewRoot')]
    [ValidateNotNullOrEmpty()]
    [string[]]$TargetServers,

    [Parameter(Mandatory = $false, ParameterSetName = 'NewRoot')]
    [string]$NewRoot,

    [Parameter(Mandatory = $false, ParameterSetName = 'Reference')]
    [switch]$CompareOnly,

    [Parameter(Mandatory = $false, ParameterSetName = 'Reference')]
    [Parameter(Mandatory = $false, ParameterSetName = 'NewRoot')]
    [string]$OutputFile,

    [switch]$Help
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$script:ScriptBaseName = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)

if ($Help) {
    @"
ConfigureExchangeLogPaths-v1.1.ps1
Compare or configure supported Exchange transport log paths.

COMMON USAGE
  New-root mode on the local server:
    .\ConfigureExchangeLogPaths-v1.1.ps1

  New-root mode on selected servers:
    .\ConfigureExchangeLogPaths-v1.1.ps1 -TargetServers EX01,EX02 -NewRoot E:\EXCLOG

  Compare source and targets only:
    .\ConfigureExchangeLogPaths-v1.1.ps1 -SourceServer EX01 -TargetServers EX02,EX03 -CompareOnly

  Export required commands without applying:
    .\ConfigureExchangeLogPaths-v1.1.ps1 -SourceServer EX01 -TargetServers EX02,EX03 -OutputFile C:\Temp\LogPath-Commands.txt

NOTES
  - Normal mode: preview -> confirmation -> apply -> verify.
  - -CompareOnly: no Exchange changes.
  - -OutputFile: exports commands/report and does not apply Exchange changes.
  - Review Only paths are shown but are not changed automatically.
  - For full help:
      Get-Help .\ConfigureExchangeLogPaths-v1.1.ps1 -Full
"@ | Write-Host
    return
}


# ---------------------------------------------------------------------------
# Supported log path map
# ---------------------------------------------------------------------------
# MoveInNewRootMode controls the intentionally small default relocation scope.
# Reference mode uses every entry so custom Agent/Routing layouts on an existing
# server can also be reproduced on replacement/new Exchange servers.
$LogPathMap = @(
    [PSCustomObject]@{ ServiceKey = 'Transport'; ServiceName = 'Transport Service'; GetCmd = 'Get-TransportService'; SetCmd = 'Set-TransportService'; Property = 'MessageTrackingLogPath';      DefaultKind = 'ExchangeRelative'; DefaultRelative = 'TransportRoles\Logs\MessageTracking';                    Policy = 'Apply';      MoveInNewRootMode = $true  },
    [PSCustomObject]@{ ServiceKey = 'Transport'; ServiceName = 'Transport Service'; GetCmd = 'Get-TransportService'; SetCmd = 'Set-TransportService'; Property = 'ConnectivityLogPath';         DefaultKind = 'ExchangeRelative'; DefaultRelative = 'TransportRoles\Logs\Hub\Connectivity';                  Policy = 'Apply';      MoveInNewRootMode = $true  },
    [PSCustomObject]@{ ServiceKey = 'Transport'; ServiceName = 'Transport Service'; GetCmd = 'Get-TransportService'; SetCmd = 'Set-TransportService'; Property = 'ReceiveProtocolLogPath';      DefaultKind = 'ExchangeRelative'; DefaultRelative = 'TransportRoles\Logs\Hub\ProtocolLog\SmtpReceive';       Policy = 'Apply';      MoveInNewRootMode = $true  },
    [PSCustomObject]@{ ServiceKey = 'Transport'; ServiceName = 'Transport Service'; GetCmd = 'Get-TransportService'; SetCmd = 'Set-TransportService'; Property = 'SendProtocolLogPath';         DefaultKind = 'ExchangeRelative'; DefaultRelative = 'TransportRoles\Logs\Hub\ProtocolLog\SmtpSend';          Policy = 'Apply';      MoveInNewRootMode = $true  },
    [PSCustomObject]@{ ServiceKey = 'Transport'; ServiceName = 'Transport Service'; GetCmd = 'Get-TransportService'; SetCmd = 'Set-TransportService'; Property = 'AgentLogPath';                DefaultKind = 'ExchangeRelative'; DefaultRelative = 'TransportRoles\Logs\Hub\AgentLog';                       Policy = 'Apply';      MoveInNewRootMode = $false },
    [PSCustomObject]@{ ServiceKey = 'Transport'; ServiceName = 'Transport Service'; GetCmd = 'Get-TransportService'; SetCmd = 'Set-TransportService'; Property = 'RoutingTableLogPath';         DefaultKind = 'ExchangeRelative'; DefaultRelative = 'TransportRoles\Logs\Hub\Routing';                        Policy = 'Apply';      MoveInNewRootMode = $false },
    [PSCustomObject]@{ ServiceKey = 'Transport'; ServiceName = 'Transport Service'; GetCmd = 'Get-TransportService'; SetCmd = 'Set-TransportService'; Property = 'DnsLogPath';                  DefaultKind = 'Null';             DefaultRelative = $null;                                                     Policy = 'Apply';      MoveInNewRootMode = $false },
    [PSCustomObject]@{ ServiceKey = 'Transport'; ServiceName = 'Transport Service'; GetCmd = 'Get-TransportService'; SetCmd = 'Set-TransportService'; Property = 'PipelineTracingPath';         DefaultKind = 'ExchangeRelative'; DefaultRelative = 'TransportRoles\Logs\Hub\PipelineTracing';                Policy = 'ReviewOnly'; MoveInNewRootMode = $false },

    [PSCustomObject]@{ ServiceKey = 'FrontEnd';  ServiceName = 'Front End Transport'; GetCmd = 'Get-FrontEndTransportService'; SetCmd = 'Set-FrontEndTransportService'; Property = 'ConnectivityLogPath';    DefaultKind = 'ExchangeRelative'; DefaultRelative = 'TransportRoles\Logs\FrontEnd\Connectivity';               Policy = 'Apply';      MoveInNewRootMode = $true  },
    [PSCustomObject]@{ ServiceKey = 'FrontEnd';  ServiceName = 'Front End Transport'; GetCmd = 'Get-FrontEndTransportService'; SetCmd = 'Set-FrontEndTransportService'; Property = 'ReceiveProtocolLogPath'; DefaultKind = 'ExchangeRelative'; DefaultRelative = 'TransportRoles\Logs\FrontEnd\ProtocolLog\SmtpReceive';    Policy = 'Apply';      MoveInNewRootMode = $true  },
    [PSCustomObject]@{ ServiceKey = 'FrontEnd';  ServiceName = 'Front End Transport'; GetCmd = 'Get-FrontEndTransportService'; SetCmd = 'Set-FrontEndTransportService'; Property = 'SendProtocolLogPath';    DefaultKind = 'ExchangeRelative'; DefaultRelative = 'TransportRoles\Logs\FrontEnd\ProtocolLog\SmtpSend';       Policy = 'Apply';      MoveInNewRootMode = $true  },
    [PSCustomObject]@{ ServiceKey = 'FrontEnd';  ServiceName = 'Front End Transport'; GetCmd = 'Get-FrontEndTransportService'; SetCmd = 'Set-FrontEndTransportService'; Property = 'AgentLogPath';           DefaultKind = 'ExchangeRelative'; DefaultRelative = 'TransportRoles\Logs\FrontEnd\AgentLog';                    Policy = 'Apply';      MoveInNewRootMode = $false },
    [PSCustomObject]@{ ServiceKey = 'FrontEnd';  ServiceName = 'Front End Transport'; GetCmd = 'Get-FrontEndTransportService'; SetCmd = 'Set-FrontEndTransportService'; Property = 'RoutingTableLogPath';    DefaultKind = 'ExchangeRelative'; DefaultRelative = 'TransportRoles\Logs\FrontEnd\Routing';                     Policy = 'Apply';      MoveInNewRootMode = $false },
    [PSCustomObject]@{ ServiceKey = 'FrontEnd';  ServiceName = 'Front End Transport'; GetCmd = 'Get-FrontEndTransportService'; SetCmd = 'Set-FrontEndTransportService'; Property = 'DnsLogPath';             DefaultKind = 'Null';             DefaultRelative = $null;                                                     Policy = 'Apply';      MoveInNewRootMode = $false },

    [PSCustomObject]@{ ServiceKey = 'Mailbox';   ServiceName = 'Mailbox Transport'; GetCmd = 'Get-MailboxTransportService'; SetCmd = 'Set-MailboxTransportService'; Property = 'ConnectivityLogPath';             DefaultKind = 'ExchangeRelative'; DefaultRelative = 'TransportRoles\Logs\Mailbox\Connectivity';                Policy = 'Apply';      MoveInNewRootMode = $true  },
    [PSCustomObject]@{ ServiceKey = 'Mailbox';   ServiceName = 'Mailbox Transport'; GetCmd = 'Get-MailboxTransportService'; SetCmd = 'Set-MailboxTransportService'; Property = 'ReceiveProtocolLogPath';          DefaultKind = 'ExchangeRelative'; DefaultRelative = 'TransportRoles\Logs\Mailbox\ProtocolLog\SmtpReceive';     Policy = 'Apply';      MoveInNewRootMode = $true  },
    [PSCustomObject]@{ ServiceKey = 'Mailbox';   ServiceName = 'Mailbox Transport'; GetCmd = 'Get-MailboxTransportService'; SetCmd = 'Set-MailboxTransportService'; Property = 'SendProtocolLogPath';             DefaultKind = 'ExchangeRelative'; DefaultRelative = 'TransportRoles\Logs\Mailbox\ProtocolLog\SmtpSend';        Policy = 'Apply';      MoveInNewRootMode = $true  },
    [PSCustomObject]@{ ServiceKey = 'Mailbox';   ServiceName = 'Mailbox Transport'; GetCmd = 'Get-MailboxTransportService'; SetCmd = 'Set-MailboxTransportService'; Property = 'MailboxDeliveryAgentLogPath';     DefaultKind = 'ExchangeRelative'; DefaultRelative = 'TransportRoles\Logs\Mailbox\AgentLog\Delivery';           Policy = 'Apply';      MoveInNewRootMode = $false },
    [PSCustomObject]@{ ServiceKey = 'Mailbox';   ServiceName = 'Mailbox Transport'; GetCmd = 'Get-MailboxTransportService'; SetCmd = 'Set-MailboxTransportService'; Property = 'MailboxSubmissionAgentLogPath';   DefaultKind = 'ExchangeRelative'; DefaultRelative = 'TransportRoles\Logs\Mailbox\AgentLog\Submission';         Policy = 'Apply';      MoveInNewRootMode = $false },
    [PSCustomObject]@{ ServiceKey = 'Mailbox';   ServiceName = 'Mailbox Transport'; GetCmd = 'Get-MailboxTransportService'; SetCmd = 'Set-MailboxTransportService'; Property = 'RoutingTableLogPath';             DefaultKind = 'ExchangeRelative'; DefaultRelative = 'TransportRoles\Logs\Mailbox\Routing';                      Policy = 'Apply';      MoveInNewRootMode = $false },
    [PSCustomObject]@{ ServiceKey = 'Mailbox';   ServiceName = 'Mailbox Transport'; GetCmd = 'Get-MailboxTransportService'; SetCmd = 'Set-MailboxTransportService'; Property = 'MailboxDeliveryThrottlingLogPath';DefaultKind = 'ExchangeRelative'; DefaultRelative = 'TransportRoles\Logs\Throttling\Delivery';                   Policy = 'Apply';      MoveInNewRootMode = $false },
    [PSCustomObject]@{ ServiceKey = 'Mailbox';   ServiceName = 'Mailbox Transport'; GetCmd = 'Get-MailboxTransportService'; SetCmd = 'Set-MailboxTransportService'; Property = 'PipelineTracingPath';             DefaultKind = 'ExchangeRelative'; DefaultRelative = 'TransportRoles\Logs\Mailbox\PipelineTracing';               Policy = 'ReviewOnly'; MoveInNewRootMode = $false }
)

# ---------------------------------------------------------------------------
# Exchange Management Shell
# ---------------------------------------------------------------------------
# The script can be started from a normal Windows PowerShell 5.1 session.
# If Exchange cmdlets are not already loaded, initialize the local Exchange
# Management Shell using the server's installed RemoteExchange.ps1.
function Initialize-ExchangeShell {
    if (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue) { return }
    if (-not $env:ExchangeInstallPath) { throw 'Exchange Management Shell commands were not found.' }

    $remoteExchange = Join-Path $env:ExchangeInstallPath 'bin\RemoteExchange.ps1'
    if (-not (Test-Path $remoteExchange)) { throw 'Exchange Management Shell commands were not found.' }

    . $remoteExchange
    Connect-ExchangeServer -Auto -AllowClobber | Out-Null
}

# ---------------------------------------------------------------------------
# Path helpers
# ---------------------------------------------------------------------------
# Path comparisons are normalized and case-insensitive because Windows paths
# may be returned with different casing or a trailing backslash even when they
# identify the same directory. These helpers deliberately do not require the
# path to exist; Exchange may create a configured log directory later.
function Normalize-PathText {
    param([AllowNull()][string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    $value = $Path.Trim().Replace('/', '\')
    while ($value.Length -gt 3 -and $value.EndsWith('\')) {
        $value = $value.Substring(0, $value.Length - 1)
    }
    return $value
}

function Join-PathText {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$Relative
    )

    $left = (Normalize-PathText -Path $Root).TrimEnd('\')
    $right = $Relative.Trim().TrimStart('\').Replace('/', '\')
    return "$left\$right"
}

function Test-PathUnderRoot {
    param(
        [AllowNull()][string]$Path,
        [Parameter(Mandatory = $true)][string]$Root
    )

    $pathValue = Normalize-PathText -Path $Path
    $rootValue = Normalize-PathText -Path $Root
    if (-not $pathValue -or -not $rootValue) { return $false }

    if ($pathValue.Equals($rootValue, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    return $pathValue.StartsWith($rootValue + '\', [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-RelativePathUnderRoot {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Root
    )

    $pathValue = Normalize-PathText -Path $Path
    $rootValue = Normalize-PathText -Path $Root
    if (-not (Test-PathUnderRoot -Path $pathValue -Root $rootValue)) {
        throw "Path '$Path' is not under root '$Root'."
    }

    if ($pathValue.Length -eq $rootValue.Length) { return '' }
    return $pathValue.Substring($rootValue.Length).TrimStart('\')
}

function Test-EquivalentPath {
    param(
        [AllowNull()][string]$Left,
        [AllowNull()][string]$Right
    )

    $leftValue = Normalize-PathText -Path $Left
    $rightValue = Normalize-PathText -Path $Right
    if ($null -eq $leftValue -and $null -eq $rightValue) { return $true }
    if ($null -eq $leftValue -or $null -eq $rightValue) { return $false }
    return $leftValue.Equals($rightValue, [System.StringComparison]::OrdinalIgnoreCase)
}

# Classify the CURRENT configured value against the default path calculated
# from that server's own Exchange install root. This is the key distinction
# that prevents C:-installed and D:-installed Exchange servers from being
# reported as different when both are actually using their defaults.
function Get-PathClassification {
    param(
        [AllowNull()][string]$CurrentPath,
        [AllowNull()][string]$DefaultPath,
        [Parameter(Mandatory = $true)][ValidateSet('ExchangeRelative','Null')][string]$DefaultKind
    )

    if ($DefaultKind -eq 'Null') {
        if ([string]::IsNullOrWhiteSpace($CurrentPath)) { return 'Default/Null' }
        return 'Custom'
    }

    if ([string]::IsNullOrWhiteSpace($CurrentPath)) { return 'Disabled/Null' }
    if (Test-EquivalentPath -Left $CurrentPath -Right $DefaultPath) { return 'Default' }
    return 'Custom'
}

function ConvertTo-DisplayPath {
    param([AllowNull()][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '<null>' }
    return $Path
}

# ---------------------------------------------------------------------------
# Server discovery
# ---------------------------------------------------------------------------
function Get-NormalizedTargets {
    param([Parameter(Mandatory = $true)][string[]]$Servers)

    $result = @($Servers | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique)
    if ($result.Count -eq 0) { throw 'At least one target server is required.' }
    return $result
}

# Never assume C:\Program Files\Microsoft\Exchange Server\V15.
# Exchange can be installed on another drive/path, so read msExchInstallPath
# from the Exchange server object in Active Directory for every server.
function Get-ExchangeInstallPath {
    param([Parameter(Mandatory = $true)][string]$Server)

    # msExchInstallPath is stored on the Exchange server object in AD and works
    # even when different Exchange servers were installed on different drives.
    $serverObject = Get-ExchangeServer -Identity $Server -ErrorAction Stop
    $distinguishedName = [string]$serverObject.DistinguishedName
    if ([string]::IsNullOrWhiteSpace($distinguishedName)) {
        throw "Unable to determine the Active Directory object for Exchange server '$Server'."
    }

    try {
        $adsi = [ADSI]("LDAP://{0}" -f $distinguishedName)
        $installPath = [string]$adsi.Properties['msExchInstallPath'].Value
    }
    catch {
        throw "Unable to read msExchInstallPath for '$Server'. $($_.Exception.Message)"
    }

    if ([string]::IsNullOrWhiteSpace($installPath)) {
        # Local fallback is useful if AD replication has not yet populated the attribute.
        if ($Server -ieq $env:COMPUTERNAME -and $env:ExchangeInstallPath) {
            $installPath = $env:ExchangeInstallPath
        }
        else {
            throw "msExchInstallPath is empty for Exchange server '$Server'."
        }
    }

    return Normalize-PathText -Path $installPath
}

function Test-TargetDriveAvailability {
    param(
        [Parameter(Mandatory = $true)][string]$Server,
        [AllowNull()][string]$Path
    )

    $pathValue = Normalize-PathText -Path $Path
    if (-not $pathValue) { return 'NotApplicable' }
    if ($pathValue -notmatch '^([A-Za-z]):\\') { return 'NotApplicable' }

    $drive = $Matches[1].ToUpper() + ':'
    try {
        $disk = Get-WmiObject -Class Win32_LogicalDisk -ComputerName $Server -Filter "DeviceID='$drive'" -ErrorAction Stop
        if ($disk) { return 'Available' }
        return 'Missing'
    }
    catch {
        # A blocked WMI/DCOM path is not proof that the drive is missing.
        return 'Unknown'
    }
}

# Build one normalized snapshot per server. Besides the current path, retain
# the calculated default and Default/Custom classification so later stages
# can reason about intent instead of blindly copying path strings.
function Get-ServerSnapshot {
    param([Parameter(Mandatory = $true)][string]$Server)

    $installPath = Get-ExchangeInstallPath -Server $Server
    $serviceCache = @{}
    $items = New-Object System.Collections.ArrayList

    foreach ($entry in $LogPathMap) {
        if (-not $serviceCache.ContainsKey($entry.ServiceKey)) {
            $serviceCache[$entry.ServiceKey] = & $entry.GetCmd -Identity $Server -ErrorAction Stop
        }

        $serviceObject = $serviceCache[$entry.ServiceKey]
        $propertyInfo = $serviceObject.PSObject.Properties[$entry.Property]
        if (-not $propertyInfo) {
            throw "Property '$($entry.Property)' was not returned by $($entry.GetCmd) for '$Server'."
        }

        $currentPath = if ($null -eq $propertyInfo.Value) { $null } else { [string]$propertyInfo.Value }
        $defaultPath = if ($entry.DefaultKind -eq 'Null') { $null } else { Join-PathText -Root $installPath -Relative $entry.DefaultRelative }
        $classification = Get-PathClassification -CurrentPath $currentPath -DefaultPath $defaultPath -DefaultKind $entry.DefaultKind

        [void]$items.Add([PSCustomObject]@{
            SettingId          = "$($entry.ServiceKey).$($entry.Property)"
            ServiceKey         = $entry.ServiceKey
            ServiceName        = $entry.ServiceName
            GetCmd             = $entry.GetCmd
            SetCmd             = $entry.SetCmd
            Property           = $entry.Property
            DefaultKind        = $entry.DefaultKind
            DefaultRelative    = $entry.DefaultRelative
            Policy             = $entry.Policy
            MoveInNewRootMode  = $entry.MoveInNewRootMode
            CurrentPath        = $currentPath
            DefaultPath        = $defaultPath
            Classification     = $classification
        })
    }

    return [PSCustomObject]@{
        Server      = $Server
        InstallPath = $installPath
        Items       = @($items)
    }
}

function Get-SnapshotItem {
    param(
        [Parameter(Mandatory = $true)]$Snapshot,
        [Parameter(Mandatory = $true)][string]$SettingId
    )

    $item = @($Snapshot.Items | Where-Object { $_.SettingId -eq $SettingId })
    if ($item.Count -ne 1) { throw "Unable to resolve setting '$SettingId' on '$($Snapshot.Server)'." }
    return $item[0]
}

# ---------------------------------------------------------------------------
# Pre-change configuration snapshot
# ---------------------------------------------------------------------------
function Get-PreChangeSnapshotPath {
    param([Parameter(Mandatory = $true)][string]$ScriptBaseName)

    $snapshotRoot = Join-Path -Path $PSScriptRoot -ChildPath 'ConfigSnapshots'
    if (-not (Test-Path -LiteralPath $snapshotRoot -PathType Container)) {
        New-Item -Path $snapshotRoot -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $candidate = Join-Path -Path $snapshotRoot -ChildPath ("{0}_{1}.txt" -f $ScriptBaseName, $timestamp)
    $suffix = 2
    while (Test-Path -LiteralPath $candidate) {
        $candidate = Join-Path -Path $snapshotRoot -ChildPath ("{0}_{1}_{2}.txt" -f $ScriptBaseName, $timestamp, $suffix)
        $suffix++
    }

    return [System.IO.Path]::GetFullPath($candidate)
}

function Add-LogPathSnapshotLines {
    param(
        [Parameter(Mandatory = $true)][System.Collections.ArrayList]$Lines,
        [Parameter(Mandatory = $true)]$Snapshot,
        [Parameter(Mandatory = $true)][string]$Role
    )

    [void]$Lines.Add(("=== {0}: {1} ===" -f $Role, $Snapshot.Server))
    [void]$Lines.Add(("ExchangeInstallPath = {0}" -f (ConvertTo-DisplayPath -Path $Snapshot.InstallPath)))
    [void]$Lines.Add('')

    foreach ($item in @($Snapshot.Items | Sort-Object ServiceName,Property)) {
        [void]$Lines.Add(("[{0} - {1}]" -f $item.ServiceName, $item.Property))
        [void]$Lines.Add(("Current        = {0}" -f (ConvertTo-DisplayPath -Path $item.CurrentPath)))
        [void]$Lines.Add(("Classification = {0}" -f $item.Classification))
        [void]$Lines.Add(("Default        = {0}" -f (ConvertTo-DisplayPath -Path $item.DefaultPath)))
        [void]$Lines.Add(("Policy         = {0}" -f $item.Policy))
        [void]$Lines.Add('')
    }
}

function Export-PreChangeSnapshot {
    param(
        [Parameter(Mandatory = $true)][string[]]$Targets,
        [AllowNull()][string]$SourceServer
    )

    $path = Get-PreChangeSnapshotPath -ScriptBaseName $script:ScriptBaseName
    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add('Exchange log path pre-change snapshot')
    [void]$lines.Add(("Script        : {0}" -f $script:ScriptBaseName))
    [void]$lines.Add(("Snapshot Time : {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')))
    [void]$lines.Add('Purpose       : Current configuration captured immediately before Apply')
    [void]$lines.Add('')

    if (-not [string]::IsNullOrWhiteSpace($SourceServer)) {
        $freshSource = Get-ServerSnapshot -Server $SourceServer
        Add-LogPathSnapshotLines -Lines $lines -Snapshot $freshSource -Role 'SOURCE'
    }

    foreach ($target in $Targets) {
        $freshTarget = Get-ServerSnapshot -Server $target
        Add-LogPathSnapshotLines -Lines $lines -Snapshot $freshTarget -Role 'TARGET'
    }

    Set-Content -LiteralPath $path -Value @($lines) -Encoding UTF8 -ErrorAction Stop
    return $path
}

# ---------------------------------------------------------------------------
# Desired configuration
# ---------------------------------------------------------------------------
# Reference mode is semantic, not a literal clone:
#   1. Source Default -> calculate the target's own default path.
#   2. Source Custom inside Exchange root -> preserve the relative path below
#      the source Exchange root and rebuild it below the target Exchange root.
#   3. Source Custom outside Exchange root -> preserve the literal custom path.
# This keeps the layout consistent even when source and target Exchange binaries
# were installed on different drives.
function Get-ReferenceDesiredPath {
    param(
        [AllowNull()][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$SourceClassification,
        [Parameter(Mandatory = $true)][string]$SourceInstallPath,
        [Parameter(Mandatory = $true)][string]$TargetInstallPath,
        [Parameter(Mandatory = $true)][ValidateSet('ExchangeRelative','Null')][string]$DefaultKind,
        [AllowNull()][string]$DefaultRelative
    )

    # A null source path is recorded and compared but is not pushed automatically.
    # This is especially important for settings such as DnsLogPath whose product
    # default is $null and whose enabled state is managed by a separate property.
    if ([string]::IsNullOrWhiteSpace($SourcePath)) {
        return [PSCustomObject]@{
            DesiredPath = $null
            Mapping     = 'Skipped - source path is null/default-disabled'
            Manage      = $false
        }
    }

    if ($SourceClassification -eq 'Default' -and $DefaultKind -eq 'ExchangeRelative') {
        return [PSCustomObject]@{
            DesiredPath = Join-PathText -Root $TargetInstallPath -Relative $DefaultRelative
            Mapping     = 'Default translated to target Exchange install path'
            Manage      = $true
        }
    }

    if (Test-PathUnderRoot -Path $SourcePath -Root $SourceInstallPath) {
        $relative = Get-RelativePathUnderRoot -Path $SourcePath -Root $SourceInstallPath
        return [PSCustomObject]@{
            DesiredPath = Join-PathText -Root $TargetInstallPath -Relative $relative
            Mapping     = 'Custom path under Exchange root translated to target Exchange install path'
            Manage      = $true
        }
    }

    return [PSCustomObject]@{
        DesiredPath = Normalize-PathText -Path $SourcePath
        Mapping     = 'Literal custom path copied from source'
        Manage      = $true
    }
}

# Build the full Source -> Target proposal. Literal custom paths receive an
# additional drive-presence check because E:\EXCLOG on the source is unsafe to
# apply automatically if the target server has no E: drive.
function New-ReferenceDesiredConfiguration {
    param(
        [Parameter(Mandatory = $true)]$SourceSnapshot,
        [Parameter(Mandatory = $true)]$TargetSnapshot
    )

    $desiredItems = New-Object System.Collections.ArrayList

    foreach ($sourceItem in $SourceSnapshot.Items) {
        $targetItem = Get-SnapshotItem -Snapshot $TargetSnapshot -SettingId $sourceItem.SettingId
        $mapping = Get-ReferenceDesiredPath `
            -SourcePath $sourceItem.CurrentPath `
            -SourceClassification $sourceItem.Classification `
            -SourceInstallPath $SourceSnapshot.InstallPath `
            -TargetInstallPath $TargetSnapshot.InstallPath `
            -DefaultKind $sourceItem.DefaultKind `
            -DefaultRelative $sourceItem.DefaultRelative

        $driveStatus = if ($mapping.Manage -and $mapping.Mapping -eq 'Literal custom path copied from source') {
            Test-TargetDriveAvailability -Server $TargetSnapshot.Server -Path $mapping.DesiredPath
        }
        else {
            'NotChecked'
        }

        $manage = $mapping.Manage -and ($sourceItem.Policy -eq 'Apply')
        $reviewNote = $null
        if ($sourceItem.Policy -eq 'ReviewOnly') {
            $manage = $false
            $reviewNote = 'Review Only. This setting is reported but is not applied automatically.'
        }
        if ($driveStatus -eq 'Missing') {
            $manage = $false
            $reviewNote = 'Target drive is missing. This setting will not be applied automatically.'
        }
        elseif ($driveStatus -eq 'Unknown') {
            $manage = $false
            $reviewNote = 'Target drive could not be verified remotely. Automatic apply is blocked; review or export the command for manual use.'
        }

        [void]$desiredItems.Add([PSCustomObject]@{
            SettingId             = $sourceItem.SettingId
            ServiceKey            = $sourceItem.ServiceKey
            ServiceName           = $sourceItem.ServiceName
            SetCmd                = $sourceItem.SetCmd
            Property              = $sourceItem.Property
            Policy                = $sourceItem.Policy
            SourcePath            = $sourceItem.CurrentPath
            SourceClassification  = $sourceItem.Classification
            TargetCurrentPath     = $targetItem.CurrentPath
            TargetClassification  = $targetItem.Classification
            DesiredPath           = $mapping.DesiredPath
            Mapping               = $mapping.Mapping
            DriveStatus           = $driveStatus
            Manage                = $manage
            ReviewNote            = $reviewNote
        })
    }

    return [PSCustomObject]@{
        Target = $TargetSnapshot.Server
        Items  = @($desiredItems)
    }
}

# NewRoot mode intentionally relocates only entries marked MoveInNewRootMode.
# The user supplies only a root such as D:\EXCLOG; the Exchange-relative
# directory names remain unchanged, for example:
#   <ExchangeRoot>\TransportRoles\Logs\Hub\ProtocolLog\SmtpReceive
# becomes
#   D:\EXCLOG\TransportRoles\Logs\Hub\ProtocolLog\SmtpReceive
function New-NewRootDesiredConfiguration {
    param(
        [Parameter(Mandatory = $true)]$TargetSnapshot,
        [Parameter(Mandatory = $true)][string]$Root
    )

    $desiredItems = New-Object System.Collections.ArrayList
    foreach ($targetItem in @($TargetSnapshot.Items | Where-Object { $_.MoveInNewRootMode })) {
        $desiredPath = Join-PathText -Root $Root -Relative $targetItem.DefaultRelative
        $driveStatus = Test-TargetDriveAvailability -Server $TargetSnapshot.Server -Path $desiredPath
        $manage = $true
        $reviewNote = $null

        if ($driveStatus -eq 'Missing') {
            $manage = $false
            $reviewNote = 'Target drive is missing. This setting will not be applied automatically.'
        }
        elseif ($driveStatus -eq 'Unknown') {
            $manage = $false
            $reviewNote = 'Target drive could not be verified remotely. Automatic apply is blocked; review or export the command for manual use.'
        }

        [void]$desiredItems.Add([PSCustomObject]@{
            SettingId             = $targetItem.SettingId
            ServiceKey            = $targetItem.ServiceKey
            ServiceName           = $targetItem.ServiceName
            SetCmd                = $targetItem.SetCmd
            Property              = $targetItem.Property
            Policy                = $targetItem.Policy
            SourcePath            = $null
            SourceClassification  = $null
            TargetCurrentPath     = $targetItem.CurrentPath
            TargetClassification  = $targetItem.Classification
            DesiredPath           = $desiredPath
            Mapping               = 'New root replaces Exchange install root through V15; relative structure preserved'
            DriveStatus           = $driveStatus
            Manage                = $manage
            ReviewNote            = $reviewNote
        })
    }

    return [PSCustomObject]@{
        Target = $TargetSnapshot.Server
        Items  = @($desiredItems)
    }
}

# ---------------------------------------------------------------------------
# Compare and change plans
# ---------------------------------------------------------------------------
# Keep comparison separate from execution. CompareOnly shows every supported
# setting, while the normal change plan contains only values that actually
# differ. This makes repeated runs idempotent from the operator's perspective.
function New-ComparisonPlan {
    param([Parameter(Mandatory = $true)]$DesiredConfiguration)

    $result = New-Object System.Collections.ArrayList
    foreach ($item in $DesiredConfiguration.Items) {
        $isSame = Test-EquivalentPath -Left $item.TargetCurrentPath -Right $item.DesiredPath
        $status = if ($isSame) {
            'Same'
        }
        elseif ($item.Policy -eq 'ReviewOnly') {
            'Review Only'
        }
        elseif (-not $item.Manage) {
            'Skipped'
        }
        else {
            'Different'
        }

        [void]$result.Add([PSCustomObject]@{
            Target                = $DesiredConfiguration.Target
            ServiceKey            = $item.ServiceKey
            ServiceName           = $item.ServiceName
            Property              = $item.Property
            Policy                = $item.Policy
            SourcePath            = $item.SourcePath
            SourceClassification  = $item.SourceClassification
            CurrentPath           = $item.TargetCurrentPath
            CurrentClassification = $item.TargetClassification
            ProposedPath          = $item.DesiredPath
            Mapping               = $item.Mapping
            DriveStatus           = $item.DriveStatus
            ReviewNote            = $item.ReviewNote
            Status                = $status
        })
    }
    return @($result)
}

function New-ChangePlan {
    param([Parameter(Mandatory = $true)]$DesiredConfiguration)

    $result = New-Object System.Collections.ArrayList
    foreach ($item in $DesiredConfiguration.Items) {
        if (-not $item.Manage) { continue }
        if (Test-EquivalentPath -Left $item.TargetCurrentPath -Right $item.DesiredPath) { continue }

        [void]$result.Add([PSCustomObject]@{
            Target                = $DesiredConfiguration.Target
            ServiceKey            = $item.ServiceKey
            ServiceName           = $item.ServiceName
            SetCmd                = $item.SetCmd
            Property              = $item.Property
            Policy                = $item.Policy
            SourcePath            = $item.SourcePath
            SourceClassification  = $item.SourceClassification
            CurrentPath           = $item.TargetCurrentPath
            CurrentClassification = $item.TargetClassification
            NewValue              = $item.DesiredPath
            Mapping               = $item.Mapping
            DriveStatus           = $item.DriveStatus
            ReviewNote            = $item.ReviewNote
        })
    }
    return @($result)
}

function Show-ComparisonPlan {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Comparisons)

    foreach ($item in $Comparisons) {
        Write-Host ''
        Write-Host "[$($item.Target)] $($item.ServiceName) - $($item.Property)" -ForegroundColor Cyan
        Write-Host "  Policy          : $($item.Policy)" -ForegroundColor $(if ($item.Policy -eq 'ReviewOnly') { 'Yellow' } else { 'DarkGray' })
        if ($null -ne $item.SourceClassification) {
            Write-Host "  Source          : $(ConvertTo-DisplayPath -Path $item.SourcePath)"
            Write-Host "  Source Type     : $($item.SourceClassification)"
        }
        Write-Host "  Target Current  : $(ConvertTo-DisplayPath -Path $item.CurrentPath)"
        Write-Host "  Target Type     : $($item.CurrentClassification)"
        Write-Host "  Proposed        : $(ConvertTo-DisplayPath -Path $item.ProposedPath)"
        Write-Host "  Mapping         : $($item.Mapping)"
        if ($item.DriveStatus -ne 'NotChecked') { Write-Host "  Drive Status    : $($item.DriveStatus)" }
        if ($item.ReviewNote) { Write-Host "  Review          : $($item.ReviewNote)" -ForegroundColor Yellow }

        switch ($item.Status) {
            'Same'      { Write-Host '  Status          : Same' -ForegroundColor Green }
            'Different' { Write-Host '  Status          : Different' -ForegroundColor Red }
            default     { Write-Host "  Status          : $($item.Status)" -ForegroundColor Yellow }
        }
    }

    Write-Host ''
    foreach ($target in @($Comparisons.Target | Select-Object -Unique)) {
        $targetItems = @($Comparisons | Where-Object { $_.Target -eq $target })
        $different = @($targetItems | Where-Object { $_.Status -eq 'Different' }).Count
        $reviewOnly = @($targetItems | Where-Object { $_.Status -eq 'Review Only' }).Count
        $skipped = @($targetItems | Where-Object { $_.Status -eq 'Skipped' }).Count
        Write-Host "[$target] Compared: $($targetItems.Count)  Different: $different  Review Only: $reviewOnly  Skipped: $skipped" -ForegroundColor Cyan
    }
}

# CompareOnly + OutputFile writes human-readable comparison data only. Keeping
# this separate from Export-CommandPlan prevents a read-only comparison from
# accidentally producing executable Set-* commands.
function Export-ComparisonReport {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Comparisons,
        [Parameter(Mandatory = $true)][string[]]$Targets,
        [Parameter(Mandatory = $true)][string]$SourceServer
    )

    $lines = New-Object System.Collections.ArrayList
    foreach ($line in @(
        '# ConfigureExchangeLogPaths-v1.1.ps1 comparison report'
        '# Tool Author : Ceyhun Kirmizitas'
        "# Generated   : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        "# Operator    : $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
        "# Source      : $SourceServer"
        "# Targets     : $($Targets -join ', ')"
        '# Mode        : Compare only (read-only)'
        '# No Set-* commands were generated and no Exchange changes were applied.'
        ''
    )) { [void]$lines.Add($line) }

    foreach ($target in $Targets) {
        $targetItems = @($Comparisons | Where-Object { $_.Target -eq $target })
        [void]$lines.Add('=====================================================================')
        [void]$lines.Add("Target: $target")
        [void]$lines.Add('=====================================================================')
        [void]$lines.Add('')

        foreach ($item in $targetItems) {
            [void]$lines.Add("$($item.ServiceName) - $($item.Property)")
            [void]$lines.Add("  Policy          : $($item.Policy)")
            if ($null -ne $item.SourceClassification) {
                [void]$lines.Add("  Source          : $(ConvertTo-DisplayPath -Path $item.SourcePath)")
                [void]$lines.Add("  Source Type     : $($item.SourceClassification)")
            }
            [void]$lines.Add("  Target Current  : $(ConvertTo-DisplayPath -Path $item.CurrentPath)")
            [void]$lines.Add("  Target Type     : $($item.CurrentClassification)")
            [void]$lines.Add("  Proposed        : $(ConvertTo-DisplayPath -Path $item.ProposedPath)")
            [void]$lines.Add("  Mapping         : $($item.Mapping)")
            if ($item.DriveStatus -ne 'NotChecked') { [void]$lines.Add("  Drive Status    : $($item.DriveStatus)") }
            if ($item.ReviewNote) { [void]$lines.Add("  Review          : $($item.ReviewNote)") }
            [void]$lines.Add("  Status          : $($item.Status)")
            [void]$lines.Add('')
        }

        $different = @($targetItems | Where-Object { $_.Status -eq 'Different' }).Count
        $reviewOnly = @($targetItems | Where-Object { $_.Status -eq 'Review Only' }).Count
        $skipped = @($targetItems | Where-Object { $_.Status -eq 'Skipped' }).Count
        [void]$lines.Add("Summary: Compared=$($targetItems.Count) Different=$different ReviewOnly=$reviewOnly Skipped=$skipped")
        [void]$lines.Add('')
    }

    Set-Content -Path $Path -Value $lines -Encoding UTF8
}

function Show-ChangePlan {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Changes,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$DesiredConfigurations
    )

    if ($Changes.Count -eq 0) {
        Write-Host ''
        Write-Host 'No log path changes are required.' -ForegroundColor Green
    }
    else {
        foreach ($item in $Changes) {
            Write-Host ''
            Write-Host "[$($item.Target)] $($item.ServiceName) - $($item.Property)" -ForegroundColor Cyan
            if ($null -ne $item.SourceClassification) {
                Write-Host "  Source          : $(ConvertTo-DisplayPath -Path $item.SourcePath)"
                Write-Host "  Source Type     : $($item.SourceClassification)"
            }
            Write-Host "  Current         : $(ConvertTo-DisplayPath -Path $item.CurrentPath)"
            Write-Host "  Current Type    : $($item.CurrentClassification)"
            Write-Host "  New             : $(ConvertTo-DisplayPath -Path $item.NewValue)"
            Write-Host "  Mapping         : $($item.Mapping)"
            if ($item.DriveStatus -ne 'NotChecked') { Write-Host "  Drive Status    : $($item.DriveStatus)" }
            if ($item.ReviewNote) { Write-Host "  Review          : $($item.ReviewNote)" -ForegroundColor Yellow }
        }
    }

    $skipped = New-Object System.Collections.ArrayList
    foreach ($configuration in $DesiredConfigurations) {
        foreach ($item in @($configuration.Items | Where-Object {
            -not $_.Manage -and -not (Test-EquivalentPath -Left $_.TargetCurrentPath -Right $_.DesiredPath)
        })) {
            [void]$skipped.Add([PSCustomObject]@{ Target = $configuration.Target; Item = $item })
        }
    }

    foreach ($entry in $skipped) {
        Write-Host ''
        Write-Host "[$($entry.Target)] $($entry.Item.ServiceName) - $($entry.Item.Property)" -ForegroundColor Yellow
        Write-Host "  Current         : $(ConvertTo-DisplayPath -Path $entry.Item.TargetCurrentPath)"
        Write-Host "  Proposed        : $(ConvertTo-DisplayPath -Path $entry.Item.DesiredPath)"
        Write-Host "  Policy          : $($entry.Item.Policy)"
        Write-Host "  Status          : $(if ($entry.Item.Policy -eq 'ReviewOnly') { 'Review Only' } else { 'Skipped' })"
        Write-Host "  Reason          : $($entry.Item.ReviewNote)"
    }
}

# ---------------------------------------------------------------------------
# Apply / export
# ---------------------------------------------------------------------------
# Group path changes by Exchange service so one Set-* cmdlet can update all
# changed properties for that service in a single operation.
function New-TargetOperations {
    param(
        [Parameter(Mandatory = $true)][string]$Target,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Changes
    )

    $operations = New-Object System.Collections.ArrayList
    foreach ($serviceKey in @($Changes.ServiceKey | Select-Object -Unique)) {
        $serviceChanges = @($Changes | Where-Object { $_.ServiceKey -eq $serviceKey })
        if ($serviceChanges.Count -eq 0) { continue }

        $params = [ordered]@{ Identity = $Target }
        foreach ($change in $serviceChanges) {
            $params[$change.Property] = $change.NewValue
        }

        [void]$operations.Add([PSCustomObject]@{
            Target     = $Target
            ServiceKey = $serviceKey
            ServiceName = $serviceChanges[0].ServiceName
            SetCmd     = $serviceChanges[0].SetCmd
            Parameters = $params
            Changes    = $serviceChanges
        })
    }
    return @($operations)
}

function Invoke-ConfigurationPlan {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Changes,
        [Parameter(Mandatory = $true)][string[]]$Targets
    )

    foreach ($target in $Targets) {
        $targetChanges = @($Changes | Where-Object { $_.Target -eq $target })
        foreach ($operation in @(New-TargetOperations -Target $target -Changes $targetChanges)) {
            Write-Host "Applying $($operation.ServiceName) changes on $target..." -ForegroundColor Cyan
            $operationParameters = $operation.Parameters
            & $operation.SetCmd @operationParameters
        }
    }
}

function ConvertTo-CommandLiteral {
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return '$null' }
    $text = [string]$Value
    return "'$($text.Replace("'", "''"))'"
}

function Resolve-OutputFilePath {
    param([Parameter(Mandatory = $true)][string]$Path)

    $resolved = $Path
    if ([System.IO.Path]::GetExtension($resolved) -ne '.txt') {
        $resolved = $resolved + '.txt'
    }

    $parent = Split-Path -Parent $resolved
    if ($parent -and -not (Test-Path $parent)) {
        New-Item -Path $parent -ItemType Directory -Force | Out-Null
    }
    return $resolved
}

# OutputFile is a review mode: generate the same Set-* operations that would
# otherwise be applied, but write them to TXT and do not modify Exchange.
function Export-CommandPlan {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Changes,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$DesiredConfigurations,
        [Parameter(Mandatory = $true)][string[]]$Targets,
        [string]$SourceServer,
        [string]$NewRoot
    )

    $lines = New-Object System.Collections.ArrayList
    foreach ($line in @(
        '# ConfigureExchangeLogPaths-v1.1.ps1 generated command file'
        '# Tool Author : Ceyhun Kirmizitas'
        "# Generated   : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        "# Operator    : $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
        "# Source      : $(if ($SourceServer) { $SourceServer } else { '<New root mode>' })"
        "# New Root    : $(if ($NewRoot) { $NewRoot } else { '<Reference mode>' })"
        "# Targets     : $($Targets -join ', ')"
        '# WARNING: Review every command before running it in Exchange Management Shell.'
        '# ConfigureExchangeLogPaths-v1.1.ps1 did not apply any Exchange changes.'
        ''
    )) { [void]$lines.Add($line) }

    foreach ($target in $Targets) {
        $targetChanges = @($Changes | Where-Object { $_.Target -eq $target })
        $targetDesired = @($DesiredConfigurations | Where-Object { $_.Target -eq $target } | Select-Object -First 1)
        $reviewItems = @()
        if ($targetDesired.Count -gt 0) {
            $reviewItems = @($targetDesired[0].Items | Where-Object {
                -not $_.Manage -and -not (Test-EquivalentPath -Left $_.TargetCurrentPath -Right $_.DesiredPath)
            })
        }
        if ($targetChanges.Count -eq 0 -and $reviewItems.Count -eq 0) { continue }

        [void]$lines.Add('# =====================================================================')
        [void]$lines.Add("# Target: $target")
        [void]$lines.Add('# =====================================================================')
        [void]$lines.Add('')

        foreach ($reviewItem in $reviewItems) {
            [void]$lines.Add("# REVIEW ONLY / SKIPPED: $($reviewItem.ServiceName) - $($reviewItem.Property)")
            [void]$lines.Add("# Policy : $($reviewItem.Policy)")
            [void]$lines.Add("# Current: $(ConvertTo-DisplayPath -Path $reviewItem.TargetCurrentPath)")
            [void]$lines.Add("# Source : $(ConvertTo-DisplayPath -Path $reviewItem.SourcePath)")
            [void]$lines.Add("# Proposed: $(ConvertTo-DisplayPath -Path $reviewItem.DesiredPath)")
            if ($reviewItem.ReviewNote) { [void]$lines.Add("# Reason : $($reviewItem.ReviewNote)") }
            [void]$lines.Add('# No command generated for this setting.')
            [void]$lines.Add('')
        }

        foreach ($operation in @(New-TargetOperations -Target $target -Changes $targetChanges)) {
            [void]$lines.Add("# $($operation.ServiceName)")
            foreach ($change in $operation.Changes) {
                [void]$lines.Add("# $($change.Property): $(ConvertTo-DisplayPath -Path $change.CurrentPath) -> $(ConvertTo-DisplayPath -Path $change.NewValue)")
                if ($change.ReviewNote) { [void]$lines.Add("# REVIEW: $($change.ReviewNote)") }
            }

            $parts = New-Object System.Collections.ArrayList
            [void]$parts.Add($operation.SetCmd)
            foreach ($key in $operation.Parameters.Keys) {
                [void]$parts.Add("-$key $(ConvertTo-CommandLiteral -Value $operation.Parameters[$key])")
            }
            [void]$lines.Add(($parts -join ' '))
            [void]$lines.Add('')
        }
    }

    if ($Changes.Count -eq 0) {
        [void]$lines.Add('# No changes are required.')
    }

    Set-Content -Path $Path -Value $lines -Encoding UTF8
}

# ---------------------------------------------------------------------------
# Verification
# ---------------------------------------------------------------------------
# Re-read Exchange after Apply and compare actual values with the desired plan.
# Verification is deliberately based on fresh snapshots rather than assuming
# successful Set-* cmdlet execution means the final state is correct.
function Test-ConfigurationPlan {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$DesiredConfigurations
    )

    $results = New-Object System.Collections.ArrayList
    foreach ($configuration in $DesiredConfigurations) {
        $snapshot = Get-ServerSnapshot -Server $configuration.Target
        foreach ($desiredItem in @($configuration.Items | Where-Object { $_.Manage })) {
            $actualItem = Get-SnapshotItem -Snapshot $snapshot -SettingId $desiredItem.SettingId
            $status = if (Test-EquivalentPath -Left $actualItem.CurrentPath -Right $desiredItem.DesiredPath) { 'OK' } else { 'Mismatch' }
            [void]$results.Add([PSCustomObject]@{
                Target   = $configuration.Target
                Service  = $desiredItem.ServiceName
                Property = $desiredItem.Property
                Expected = $desiredItem.DesiredPath
                Actual   = $actualItem.CurrentPath
                Status   = $status
            })
        }
    }
    return @($results)
}

function Show-Verification {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Results)

    Write-Host ''
    Write-Host 'Verification summary' -ForegroundColor Cyan
    foreach ($result in $Results) {
        $color = if ($result.Status -eq 'OK') { 'Green' } else { 'Red' }
        Write-Host "[$($result.Target)] $($result.Service) - $($result.Property): $($result.Status)" -ForegroundColor $color
        if ($result.Status -ne 'OK') {
            Write-Host "  Expected : $(ConvertTo-DisplayPath -Path $result.Expected)"
            Write-Host "  Actual   : $(ConvertTo-DisplayPath -Path $result.Actual)"
        }
    }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
# Execution flow:
#   discover -> snapshot -> build desired state -> compare/change plan ->
#   CompareOnly (optionally with TXT report) OR OutputFile command export OR
#   confirmed Apply -> fresh verification.
Initialize-ExchangeShell

if ($PSCmdlet.ParameterSetName -eq 'Reference') {
    $targets = Get-NormalizedTargets -Servers $TargetServers
    if (@($targets | Where-Object { $_ -ieq $SourceServer }).Count -gt 0) {
        throw 'SourceServer cannot also be listed in TargetServers.'
    }

    Write-Host "Source server : $SourceServer" -ForegroundColor Cyan
    Write-Host "Target servers: $($targets -join ', ')" -ForegroundColor Cyan
    if ($CompareOnly) { Write-Host 'Mode          : Compare only (read-only)' -ForegroundColor Cyan }

    $sourceSnapshot = Get-ServerSnapshot -Server $SourceServer
    Write-Host "Source Exchange install path: $($sourceSnapshot.InstallPath)" -ForegroundColor DarkCyan
}
else {
    $targets = if ($TargetServers) { Get-NormalizedTargets -Servers $TargetServers } else { @($env:COMPUTERNAME) }
}

$targetSnapshots = @{}
foreach ($target in $targets) {
    $snapshot = Get-ServerSnapshot -Server $target
    $targetSnapshots[$target] = $snapshot
    Write-Host "[$target] Exchange install path: $($snapshot.InstallPath)" -ForegroundColor DarkCyan
}

if ($PSCmdlet.ParameterSetName -eq 'NewRoot') {
    if ([string]::IsNullOrWhiteSpace($NewRoot)) {
        Write-Host ''
        Write-Host 'The new root replaces the Exchange installation path through V15.' -ForegroundColor Yellow
        Write-Host 'Example: C:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\... -> D:\EXCLOG\TransportRoles\Logs\...' -ForegroundColor DarkGray
        $NewRoot = Read-Host 'New log root (example: D:\EXCLOG)'
    }

    $NewRoot = Normalize-PathText -Path $NewRoot
    if ([string]::IsNullOrWhiteSpace($NewRoot) -or $NewRoot -notmatch '^(?:[A-Za-z]:\\|\\\\)') {
        throw 'NewRoot must be an absolute local or UNC path, for example D:\EXCLOG.'
    }
    Write-Host "New log root : $NewRoot" -ForegroundColor Cyan
}

$desiredConfigurations = New-Object System.Collections.ArrayList
$allChanges = New-Object System.Collections.ArrayList
$allComparisons = New-Object System.Collections.ArrayList

foreach ($target in $targets) {
    $desired = if ($PSCmdlet.ParameterSetName -eq 'Reference') {
        New-ReferenceDesiredConfiguration -SourceSnapshot $sourceSnapshot -TargetSnapshot $targetSnapshots[$target]
    }
    else {
        New-NewRootDesiredConfiguration -TargetSnapshot $targetSnapshots[$target] -Root $NewRoot
    }

    [void]$desiredConfigurations.Add($desired)

    if ($CompareOnly) {
        foreach ($comparison in @(New-ComparisonPlan -DesiredConfiguration $desired)) {
            [void]$allComparisons.Add($comparison)
        }
    }
    else {
        foreach ($change in @(New-ChangePlan -DesiredConfiguration $desired)) {
            [void]$allChanges.Add($change)
        }
    }
}

if ($CompareOnly) {
    Show-ComparisonPlan -Comparisons @($allComparisons)

    # OutputFile changes only the destination of the read-only comparison. It
    # does not turn CompareOnly into command-export mode.
    if (-not [string]::IsNullOrWhiteSpace($OutputFile)) {
        $outputPath = Resolve-OutputFilePath -Path $OutputFile
        Export-ComparisonReport -Path $outputPath -Comparisons @($allComparisons) -Targets $targets -SourceServer $SourceServer

        Write-Host ''
        Write-Host "Comparison report created: $outputPath" -ForegroundColor Green
        Write-Host 'No Set-* commands were generated and no Exchange configuration changes were applied.' -ForegroundColor Yellow
    }
    return
}

Show-ChangePlan -Changes @($allChanges) -DesiredConfigurations @($desiredConfigurations)

# -OutputFile is explicit export-only mode. No Set-* cmdlet is executed here.
if (-not [string]::IsNullOrWhiteSpace($OutputFile)) {
    $outputPath = Resolve-OutputFilePath -Path $OutputFile
    Export-CommandPlan -Path $outputPath -Changes @($allChanges) -DesiredConfigurations @($desiredConfigurations) -Targets $targets -SourceServer $SourceServer -NewRoot $NewRoot

    Write-Host ''
    Write-Host "Command file created: $outputPath" -ForegroundColor Green
    Write-Host 'No Exchange configuration changes were applied. Review the TXT file and run the required commands manually.' -ForegroundColor Yellow
    return
}

if ($allChanges.Count -eq 0) {
    Write-Host 'No Exchange configuration changes were applied.' -ForegroundColor Green
    return
}

$confirmation = Read-Host 'Continue and apply these changes? [Y/N]'
if ($confirmation -notmatch '^[Yy]$') {
    Write-Host 'Operation cancelled. No Exchange configuration changes were applied.' -ForegroundColor Yellow
    return
}

$sourceForSnapshot = if ($PSCmdlet.ParameterSetName -eq 'Reference') { $SourceServer } else { $null }
$snapshotPath = Export-PreChangeSnapshot -Targets $targets -SourceServer $sourceForSnapshot
Write-Host ''
Write-Host "Configuration snapshot created: $snapshotPath" -ForegroundColor Green

Invoke-ConfigurationPlan -Changes @($allChanges) -Targets $targets

$verification = @(Test-ConfigurationPlan -DesiredConfigurations @($desiredConfigurations))
Show-Verification -Results $verification

if (@($verification | Where-Object { $_.Status -ne 'OK' }).Count -gt 0) {
    throw 'One or more log path settings did not match the expected value after configuration.'
}
