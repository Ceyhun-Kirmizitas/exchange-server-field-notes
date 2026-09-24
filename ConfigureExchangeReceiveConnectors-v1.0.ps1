<#
.SYNOPSIS
Compares and migrates Exchange Server Receive Connector configuration from a reference server to one or more target servers.

.DESCRIPTION
ConfigureExchangeReceiveConnectors-v1.1.ps1 is intended for side-by-side Exchange server deployments, upgrades, and replacement-server work.
It uses an existing Exchange server as the reference and compares its Receive Connectors with one or more target servers.

The script inventories both built-in and custom Receive Connectors. Built-in connectors are mapped by connector family (for example, Default Frontend <ServerName> -> Default Frontend <TargetServer>) and are never recreated or deleted. Custom connectors are matched by name and can be created when missing.

The following connector settings are compared and can be aligned when supported:
- Enabled
- Bindings (local IP address and TCP port)
- RemoteIPRanges
- Fqdn
- AuthMechanism
- PermissionGroups
- RequireTLS
- TlsCertificateName
- TlsDomainCapabilities
- ProtocolLoggingLevel
- MaxMessageSize
- Banner

TransportRole is displayed and used when a missing custom connector is created, but it is Review Only for an existing connector.

Anonymous relay is detected independently from PermissionGroups. The script checks for the explicit ms-Exch-SMTP-Accept-Any-Recipient right assigned to NT AUTHORITY\ANONYMOUS LOGON. Anonymous relay state is always displayed. By default, the extended relay permission is Review Only. Use -IncludeAnonymousRelayPermissions to include adding or removing that explicit permission in the change plan.

Microsoft recommends a dedicated Front End Transport Receive Connector with tightly scoped RemoteIPRanges for anonymous relay. If anonymous relay is detected on a built-in connector, the script reports a warning but does not hide the configuration.

By default, the script shows a Current -> Proposed preview, asks for confirmation, applies only required supported changes, and verifies the target server(s) afterward.
When -OutputFile is specified without -CompareOnly, no Exchange changes are applied and the required New-ReceiveConnector, Set-ReceiveConnector, Add-ADPermission, and Remove-ADPermission commands are written to a TXT file for manual review.
When -CompareOnly and -OutputFile are used together, the comparison is displayed and also written to a TXT report. No executable commands are generated and no Exchange changes are applied.
Before a confirmed Apply, the script saves a fresh TXT configuration snapshot under the script folder's ConfigSnapshots directory. Existing snapshot files are never overwritten.

.PARAMETER SourceServer
Reference Exchange server whose Receive Connector configuration is used as the source.

.PARAMETER TargetServers
One or more Exchange servers to compare/configure.

.PARAMETER ConnectorName
Optional source connector name(s) to process. Matching is case-insensitive. For built-in connectors, the family name (for example, Default Frontend) can also be used. If omitted, all Receive Connectors on SourceServer are processed.

.PARAMETER IncludeAnonymousRelayPermissions
Includes the explicit ms-Exch-SMTP-Accept-Any-Recipient permission for NT AUTHORITY\ANONYMOUS LOGON in the apply/export plan. Without this switch, anonymous relay is detected and displayed but the extended permission is Review Only.

.PARAMETER CompareOnly
Displays all selected source/target connector values without applying changes. If -OutputFile is also specified, the same comparison is written to a TXT report.

.PARAMETER OutputFile
Optional TXT output path. In normal mode, exports the required Exchange PowerShell commands instead of applying changes. With -CompareOnly, exports a comparison report instead of executable commands.

.PARAMETER Help
Displays a short usage guide and exits without initializing Exchange Management Shell or making changes.

.EXAMPLE
.\ConfigureExchangeReceiveConnectors-v1.1.ps1 -SourceServer EX16-01 -TargetServers EXSE-01,EXSE-02
Compares all Receive Connectors on EX16-01 with both target servers, previews supported changes, asks for confirmation, applies the changes, and verifies the result.

.EXAMPLE
.\ConfigureExchangeReceiveConnectors-v1.1.ps1 -SourceServer EX16-01 -TargetServers EXSE-01,EXSE-02 -ConnectorName "App Relay"
Processes only the source connector named App Relay.

.EXAMPLE
.\ConfigureExchangeReceiveConnectors-v1.1.ps1 -SourceServer EX16-01 -TargetServers EXSE-01,EXSE-02 -CompareOnly
Displays all selected Receive Connector properties and anonymous relay state without applying changes.

.EXAMPLE
.\ConfigureExchangeReceiveConnectors-v1.1.ps1 -SourceServer EX16-01 -TargetServers EXSE-01,EXSE-02 -CompareOnly -OutputFile C:\Temp\ReceiveConnector-Compare.txt
Displays the read-only comparison and writes the same data to a TXT report. No executable commands are generated.

.EXAMPLE
.\ConfigureExchangeReceiveConnectors-v1.1.ps1 -SourceServer EX16-01 -TargetServers EXSE-01,EXSE-02 -OutputFile C:\Temp\ReceiveConnector-Commands.txt
Exports the required Receive Connector commands to TXT without changing Exchange.

.EXAMPLE
.\ConfigureExchangeReceiveConnectors-v1.1.ps1 -SourceServer EX16-01 -TargetServers EXSE-01,EXSE-02 -IncludeAnonymousRelayPermissions
Includes explicit anonymous relay permission alignment in the preview/apply plan.

.NOTES
Author  : Ceyhun Kirmizitas
Version : 1.1
Date    : 23/09/2026
Scope   : Exchange Server 2016, Exchange Server 2019, Exchange Server SE

Change Log
----------
1.1 - 23/09/2026
- Added built-in -Help quick usage guidance.
- Standardized the release filename and in-script examples as ConfigureExchangeReceiveConnectors-v1.1.ps1.

1.0 - Initial release, including automatic pre-change configuration snapshots before Apply.

Migration safety model:
- Apply: supported connector settings that can be aligned after preview/confirmation.
- Review Only: TransportRole changes on existing connectors and anonymous relay permissions unless explicitly enabled.
- Skip: unsafe automatic operations such as recreating/deleting built-in connectors or applying a specific binding when the target IP address cannot be verified.

Built-in Receive connectors are never recreated or deleted.
This script does not manage Send Connectors, certificates, Extended Protection, IIS bindings, firewall/NAT rules, or SMTP functional testing.
#>

#requires -version 5.1

[CmdletBinding(DefaultParameterSetName = 'Run')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Run')]
    [ValidateNotNullOrEmpty()]
    [string]$SourceServer,

    [Parameter(Mandatory = $true, ParameterSetName = 'Run')]
    [ValidateNotNullOrEmpty()]
    [string[]]$TargetServers,

    [Parameter(Mandatory = $false, ParameterSetName = 'Run')]
    [string[]]$ConnectorName,

    [Parameter(Mandatory = $false, ParameterSetName = 'Run')]
    [switch]$IncludeAnonymousRelayPermissions,

    [Parameter(Mandatory = $false, ParameterSetName = 'Run')]
    [switch]$CompareOnly,

    [Parameter(Mandatory = $false, ParameterSetName = 'Run')]
    [string]$OutputFile,

    [Parameter(Mandatory = $true, ParameterSetName = 'Help')]
    [switch]$Help
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:ScriptBaseName = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)

if ($Help) {
    @"
ConfigureExchangeReceiveConnectors-v1.1.ps1
Compare or migrate Exchange Receive Connector configuration.

COMMON USAGE
  Compare and migrate all connectors:
    .\ConfigureExchangeReceiveConnectors-v1.1.ps1 -SourceServer EX16-01 -TargetServers EXSE-01,EXSE-02

  Process one connector only:
    .\ConfigureExchangeReceiveConnectors-v1.1.ps1 -SourceServer EX16-01 -TargetServers EXSE-01,EXSE-02 -ConnectorName "App Relay"

  Compare only:
    .\ConfigureExchangeReceiveConnectors-v1.1.ps1 -SourceServer EX16-01 -TargetServers EXSE-01,EXSE-02 -CompareOnly

  Export required commands without applying:
    .\ConfigureExchangeReceiveConnectors-v1.1.ps1 -SourceServer EX16-01 -TargetServers EXSE-01,EXSE-02 -OutputFile C:\Temp\ReceiveConnector-Commands.txt

  Include explicit anonymous relay permission alignment:
    .\ConfigureExchangeReceiveConnectors-v1.1.ps1 -SourceServer EX16-01 -TargetServers EXSE-01,EXSE-02 -IncludeAnonymousRelayPermissions

NOTES
  - Normal mode: preview -> confirmation -> apply -> verify.
  - -CompareOnly: no Exchange changes.
  - -OutputFile: exports commands/report and does not apply Exchange changes.
  - Anonymous relay permissions are Review Only unless -IncludeAnonymousRelayPermissions is used.
  - Built-in Receive Connectors are never recreated or deleted.
  - For full help:
      Get-Help .\ConfigureExchangeReceiveConnectors-v1.1.ps1 -Full
"@ | Write-Host
    return
}

$AnonymousLogonPrincipal = 'NT AUTHORITY\ANONYMOUS LOGON'
$AnonymousRelayRight = 'ms-Exch-SMTP-Accept-Any-Recipient'

# Built-in connector family names created by Exchange Setup on Mailbox servers.
# The source server suffix is replaced with the target server name when mapping.
$BuiltInConnectorFamilies = @(
    'Default Frontend',
    'Client Frontend',
    'Outbound Proxy Frontend',
    'Default',
    'Client Proxy'
)

# Keep the intentionally small migration scope in one place. Deep throttling and
# transport-tuning settings are outside v1.1 to avoid copying environment-specific
# tuning simply because it exists on an older server.
$ConnectorPropertyMap = @(
    [PSCustomObject]@{ Name = 'Enabled';              Policy = 'Apply'      },
    [PSCustomObject]@{ Name = 'Bindings';             Policy = 'Apply'      },
    [PSCustomObject]@{ Name = 'RemoteIPRanges';       Policy = 'Apply'      },
    [PSCustomObject]@{ Name = 'Fqdn';                 Policy = 'Apply'      },
    [PSCustomObject]@{ Name = 'AuthMechanism';        Policy = 'Apply'      },
    [PSCustomObject]@{ Name = 'PermissionGroups';     Policy = 'Apply'      },
    [PSCustomObject]@{ Name = 'RequireTLS';           Policy = 'Apply'      },
    [PSCustomObject]@{ Name = 'TlsCertificateName';   Policy = 'Apply'      },
    [PSCustomObject]@{ Name = 'TlsDomainCapabilities';Policy = 'Apply'      },
    [PSCustomObject]@{ Name = 'ProtocolLoggingLevel'; Policy = 'Apply'      },
    [PSCustomObject]@{ Name = 'MaxMessageSize';       Policy = 'Apply'      },
    [PSCustomObject]@{ Name = 'Banner';               Policy = 'Apply'      },
    [PSCustomObject]@{ Name = 'TransportRole';        Policy = 'ReviewOnly' }
)

# ---------------------------------------------------------------------------
# Exchange Management Shell
# ---------------------------------------------------------------------------
function Initialize-ExchangeShell {
    if (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue) { return }
    if (-not $env:ExchangeInstallPath) { throw 'Exchange Management Shell commands were not found.' }

    $remoteExchange = Join-Path $env:ExchangeInstallPath 'bin\RemoteExchange.ps1'
    if (-not (Test-Path $remoteExchange)) { throw 'Exchange Management Shell commands were not found.' }

    . $remoteExchange
    Connect-ExchangeServer -Auto -AllowClobber | Out-Null
    if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
        throw 'Exchange Management Shell could not be initialized.'
    }
}

# ---------------------------------------------------------------------------
# Generic helpers
# ---------------------------------------------------------------------------
function Get-NormalizedTargets {
    param([Parameter(Mandatory = $true)][string[]]$Servers)

    $result = @($Servers | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique)
    if ($result.Count -eq 0) { throw 'At least one target server is required.' }
    return $result
}

function Resolve-ExchangeServerName {
    param([Parameter(Mandatory = $true)][string]$Identity)

    $server = Get-ExchangeServer -Identity $Identity -ErrorAction Stop
    if ([string]::IsNullOrWhiteSpace([string]$server.Name)) {
        throw "Unable to resolve Exchange server name for '$Identity'."
    }
    return [string]$server.Name
}

function Normalize-ByteSizeValue {
    param($Value)

    if ($null -eq $Value) { return $null }
    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    if ($text -ieq 'Unlimited') { return 'Unlimited' }

    # Exchange often renders ByteQuantifiedSize as "36 MB (37,748,736 bytes)".
    # Keep a cmdlet-friendly size token instead of serializing the display suffix.
    if ($text -match '^(?<number>\d+(?:[\.,]\d+)?)\s*(?<unit>B|KB|MB|GB|TB)\b') {
        return (($Matches['number'] -replace ',', '.') + $Matches['unit'])
    }

    if ($text -match '\((?<bytes>[\d,\. ]+)\s+bytes\)') {
        $bytes = $Matches['bytes'] -replace '[^0-9]', ''
        if ($bytes) { return "${bytes}B" }
    }

    return $text
}

function Normalize-ConnectorPropertyValue {
    param(
        [Parameter(Mandatory = $true)][string]$PropertyName,
        $Value
    )

    if ($null -eq $Value) { return $null }

    switch ($PropertyName) {
        { $_ -in @('Bindings','RemoteIPRanges','TlsDomainCapabilities') } {
            return @($Value | ForEach-Object { [string]$_ })
        }
        { $_ -in @('AuthMechanism','PermissionGroups') } {
            return @(([string]$Value).Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        }
        { $_ -in @('Enabled','RequireTLS') } {
            return [bool]$Value
        }
        'MaxMessageSize' {
            return Normalize-ByteSizeValue -Value $Value
        }
        { $_ -in @('Fqdn','TlsCertificateName','ProtocolLoggingLevel','Banner','TransportRole') } {
            $text = [string]$Value
            if ([string]::IsNullOrWhiteSpace($text)) { return $null }
            return $text
        }
        default {
            return $Value
        }
    }
}

function ConvertTo-DisplayValue {
    param($Value)

    if ($null -eq $Value) { return '<null>' }
    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $items = @($Value | ForEach-Object { [string]$_ })
        if ($items.Count -eq 0) { return '<empty>' }
        return ($items -join ', ')
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return '<empty>' }
    return $text
}

function ConvertTo-ComparableValue {
    param($Value)

    if ($null -eq $Value) { return '<null>' }
    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        return (@($Value | ForEach-Object { ([string]$_).Trim().ToLowerInvariant() } | Sort-Object) -join '|')
    }
    return ([string]$Value).Trim().ToLowerInvariant()
}

function Test-EquivalentValue {
    param($Left, $Right)
    return (ConvertTo-ComparableValue -Value $Left) -eq (ConvertTo-ComparableValue -Value $Right)
}

function ConvertTo-PowerShellLiteral {
    param($Value)

    if ($null -eq $Value) { return '$null' }
    if ($Value -is [bool]) { return $(if ($Value) { '$true' } else { '$false' }) }

    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $items = @($Value | ForEach-Object { ConvertTo-PowerShellLiteral -Value ([string]$_) })
        return '@(' + ($items -join ', ') + ')'
    }

    $text = [string]$Value
    return "'" + $text.Replace("'", "''") + "'"
}

function Resolve-OutputFilePath {
    param([Parameter(Mandatory = $true)][string]$Path)

    $resolved = $Path.Trim()
    if ([string]::IsNullOrWhiteSpace($resolved)) { throw 'OutputFile cannot be empty.' }

    $extension = [System.IO.Path]::GetExtension($resolved)
    if ([string]::IsNullOrWhiteSpace($extension)) {
        $resolved = "$resolved.txt"
    }
    elseif ($extension -ne '.txt') {
        throw 'OutputFile must use the .txt extension.'
    }

    if (-not [System.IO.Path]::IsPathRooted($resolved)) {
        $resolved = Join-Path -Path (Get-Location).Path -ChildPath $resolved
    }

    $parent = Split-Path -Parent $resolved
    if ($parent -and -not (Test-Path -LiteralPath $parent -PathType Container)) {
        New-Item -Path $parent -ItemType Directory -Force | Out-Null
    }

    return [System.IO.Path]::GetFullPath($resolved)
}

# ---------------------------------------------------------------------------
# Connector identity and discovery
# ---------------------------------------------------------------------------
function Get-ConnectorDescriptor {
    param(
        [Parameter(Mandatory = $true)]$Connector,
        [Parameter(Mandatory = $true)][string]$Server
    )

    $name = [string]$Connector.Name
    foreach ($family in $BuiltInConnectorFamilies) {
        $expectedName = "$family $Server"
        if ($name.Equals($expectedName, [System.StringComparison]::OrdinalIgnoreCase)) {
            return [PSCustomObject]@{
                Type        = 'BuiltIn'
                Family      = $family
                CanonicalId = "BuiltIn:$family"
            }
        }
    }

    return [PSCustomObject]@{
        Type        = 'Custom'
        Family      = $null
        CanonicalId = "Custom:$name"
    }
}

function Get-AnonymousRelayState {
    param([Parameter(Mandatory = $true)]$Connector)

    try {
        $permissions = @(Get-ADPermission -Identity $Connector.Identity -User $AnonymousLogonPrincipal -ErrorAction Stop |
            Where-Object { $_.Deny -eq $false -and $_.IsInherited -eq $false })

        $rights = New-Object System.Collections.ArrayList
        foreach ($permission in $permissions) {
            foreach ($right in @($permission.ExtendedRights)) {
                $rightText = [string]$right
                if (-not [string]::IsNullOrWhiteSpace($rightText) -and -not ($rights -contains $rightText)) {
                    [void]$rights.Add($rightText)
                }
            }
        }

        $enabled = @($rights | Where-Object { $_ -ieq $AnonymousRelayRight }).Count -gt 0
        return [PSCustomObject]@{
            Enabled     = $enabled
            QueryStatus = 'Available'
            Rights      = @($rights)
            Error       = $null
        }
    }
    catch {
        return [PSCustomObject]@{
            Enabled     = $null
            QueryStatus = 'Unknown'
            Rights      = @()
            Error       = $_.Exception.Message
        }
    }
}

function Get-ReceiveConnectorSnapshot {
    param([Parameter(Mandatory = $true)][string]$Server)

    $serverObject = Get-ExchangeServer -Identity $Server -ErrorAction Stop
    $serverName = [string]$serverObject.Name
    $serverFqdn = [string]$serverObject.Fqdn
    if ([string]::IsNullOrWhiteSpace($serverName)) { throw "Unable to resolve Exchange server name for '$Server'." }

    $items = New-Object System.Collections.ArrayList

    foreach ($connector in @(Get-ReceiveConnector -Server $serverName -ErrorAction Stop)) {
        $descriptor = Get-ConnectorDescriptor -Connector $connector -Server $serverName
        $relay = Get-AnonymousRelayState -Connector $connector
        $values = [ordered]@{}

        foreach ($property in $ConnectorPropertyMap) {
            $propertyInfo = $connector.PSObject.Properties[$property.Name]
            if (-not $propertyInfo) {
                throw "Property '$($property.Name)' was not returned for Receive Connector '$($connector.Identity)'."
            }
            $values[$property.Name] = Normalize-ConnectorPropertyValue -PropertyName $property.Name -Value $propertyInfo.Value
        }

        [void]$items.Add([PSCustomObject]@{
            Server          = $serverName
            ServerFqdn      = $serverFqdn
            Name            = [string]$connector.Name
            Identity        = [string]$connector.Identity
            ConnectorType   = $descriptor.Type
            BuiltInFamily   = $descriptor.Family
            CanonicalId     = $descriptor.CanonicalId
            Values          = $values
            AnonymousRelay  = $relay
        })
    }

    return [PSCustomObject]@{
        Server     = $serverName
        ServerFqdn = $serverFqdn
        Items      = @($items)
    }
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

function Add-ReceiveConnectorSnapshotLines {
    param(
        [Parameter(Mandatory = $true)][System.Collections.ArrayList]$Lines,
        [Parameter(Mandatory = $true)]$Snapshot,
        [Parameter(Mandatory = $true)][string]$Role
    )

    [void]$Lines.Add(("=== {0}: {1} ===" -f $Role, $Snapshot.Server))
    [void]$Lines.Add(("ServerFqdn = {0}" -f (ConvertTo-DisplayValue $Snapshot.ServerFqdn)))
    [void]$Lines.Add('')

    foreach ($connector in @($Snapshot.Items | Sort-Object Name)) {
        [void]$Lines.Add(("[{0}]" -f $connector.Name))
        [void]$Lines.Add(("Identity      = {0}" -f $connector.Identity))
        [void]$Lines.Add(("ConnectorType = {0}" -f $connector.ConnectorType))
        if ($connector.BuiltInFamily) { [void]$Lines.Add(("BuiltInFamily = {0}" -f $connector.BuiltInFamily)) }

        foreach ($property in $ConnectorPropertyMap) {
            [void]$Lines.Add(("{0} = {1}" -f $property.Name, (ConvertTo-DisplayValue $connector.Values[$property.Name])))
        }

        [void]$Lines.Add(("AnonymousRelay.Enabled     = {0}" -f (ConvertTo-DisplayValue $connector.AnonymousRelay.Enabled)))
        [void]$Lines.Add(("AnonymousRelay.QueryStatus = {0}" -f (ConvertTo-DisplayValue $connector.AnonymousRelay.QueryStatus)))
        [void]$Lines.Add(("AnonymousRelay.Rights      = {0}" -f (ConvertTo-DisplayValue $connector.AnonymousRelay.Rights)))
        if ($connector.AnonymousRelay.Error) {
            [void]$Lines.Add(("AnonymousRelay.Error       = {0}" -f $connector.AnonymousRelay.Error))
        }
        [void]$Lines.Add('')
    }
}

function Export-PreChangeSnapshot {
    param(
        [Parameter(Mandatory = $true)][string[]]$Targets,
        [Parameter(Mandatory = $true)][string]$SourceServer
    )

    $path = Get-PreChangeSnapshotPath -ScriptBaseName $script:ScriptBaseName
    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add('Exchange Receive Connector pre-change snapshot')
    [void]$lines.Add(("Script        : {0}" -f $script:ScriptBaseName))
    [void]$lines.Add(("Snapshot Time : {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')))
    [void]$lines.Add('Purpose       : Current configuration captured immediately before Apply')
    [void]$lines.Add('')

    $freshSource = Get-ReceiveConnectorSnapshot -Server $SourceServer
    Add-ReceiveConnectorSnapshotLines -Lines $lines -Snapshot $freshSource -Role 'SOURCE'

    foreach ($target in $Targets) {
        $freshTarget = Get-ReceiveConnectorSnapshot -Server $target
        Add-ReceiveConnectorSnapshotLines -Lines $lines -Snapshot $freshTarget -Role 'TARGET'
    }

    Set-Content -LiteralPath $path -Value @($lines) -Encoding UTF8 -ErrorAction Stop
    return $path
}

function Select-SourceConnector {
    param(
        [Parameter(Mandatory = $true)]$SourceSnapshot,
        [string[]]$Names
    )

    if (-not $Names -or @($Names | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count -eq 0) {
        return @($SourceSnapshot.Items)
    }

    $selected = New-Object System.Collections.ArrayList
    foreach ($requestedName in $Names) {
        $needle = $requestedName.Trim()
        $matches = @($SourceSnapshot.Items | Where-Object {
            $_.Name -ieq $needle -or ($_.ConnectorType -eq 'BuiltIn' -and $_.BuiltInFamily -ieq $needle)
        })

        if ($matches.Count -eq 0) {
            throw "Receive Connector '$needle' was not found on source server '$($SourceSnapshot.Server)'."
        }

        foreach ($match in $matches) {
            if (@($selected | Where-Object { $_.CanonicalId -eq $match.CanonicalId }).Count -eq 0) {
                [void]$selected.Add($match)
            }
        }
    }

    return @($selected)
}

function Get-TargetConnectorName {
    param(
        [Parameter(Mandatory = $true)]$SourceConnector,
        [Parameter(Mandatory = $true)][string]$TargetServer
    )

    if ($SourceConnector.ConnectorType -eq 'BuiltIn') {
        return "$($SourceConnector.BuiltInFamily) $TargetServer"
    }
    return $SourceConnector.Name
}

function Get-TargetConnector {
    param(
        [Parameter(Mandatory = $true)]$TargetSnapshot,
        [Parameter(Mandatory = $true)]$SourceConnector
    )

    $targetName = Get-TargetConnectorName -SourceConnector $SourceConnector -TargetServer $TargetSnapshot.Server
    $matches = @($TargetSnapshot.Items | Where-Object { $_.Name -ieq $targetName })
    if ($matches.Count -gt 1) { throw "Multiple target Receive Connectors matched '$targetName' on '$($TargetSnapshot.Server)'." }
    if ($matches.Count -eq 0) { return $null }
    return $matches[0]
}

# ---------------------------------------------------------------------------
# Binding safety
# ---------------------------------------------------------------------------
function Get-BindingIPAddress {
    param([Parameter(Mandatory = $true)]$Binding)

    $text = ([string]$Binding).Trim()
    if ($text -match '^\[(?<ip>.+)\]:(?<port>\d+)$') { return $Matches['ip'] }
    if ($text -match '^(?<ip>[^:]+):(?<port>\d+)$') { return $Matches['ip'] }
    return $null
}

function Test-TargetBindingAvailability {
    param(
        [Parameter(Mandatory = $true)][string]$Server,
        [Parameter(Mandatory = $true)]$Bindings
    )

    $specificIPs = New-Object System.Collections.ArrayList
    foreach ($binding in @($Bindings)) {
        $ip = Get-BindingIPAddress -Binding $binding
        if ([string]::IsNullOrWhiteSpace($ip)) { return 'Unknown' }
        if ($ip -in @('0.0.0.0','::')) { continue }
        if (-not ($specificIPs -contains $ip)) { [void]$specificIPs.Add($ip) }
    }

    if ($specificIPs.Count -eq 0) { return 'Wildcard' }

    try {
        $targetIPs = New-Object System.Collections.ArrayList
        $adapters = @(Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $Server -Filter 'IPEnabled=True' -ErrorAction Stop)
        foreach ($adapter in $adapters) {
            foreach ($ip in @($adapter.IPAddress)) {
                if ($ip -and -not ($targetIPs -contains [string]$ip)) { [void]$targetIPs.Add([string]$ip) }
            }
        }

        foreach ($requiredIP in $specificIPs) {
            if (@($targetIPs | Where-Object { $_ -ieq $requiredIP }).Count -eq 0) { return 'Missing' }
        }
        return 'Available'
    }
    catch {
        # WMI/DCOM failure is not proof that the address is absent. Treat it as
        # unknown and do not automatically alter a specific local binding.
        return 'Unknown'
    }
}

function Get-DesiredConnectorPropertyValue {
    param(
        [Parameter(Mandatory = $true)][string]$PropertyName,
        [Parameter(Mandatory = $true)]$SourceConnector,
        [Parameter(Mandatory = $true)]$TargetSnapshot
    )

    $sourceValue = $SourceConnector.Values[$PropertyName]
    $mapping = 'Literal source value'
    $desiredValue = $sourceValue

    # The default Receive Connector FQDN is commonly the server FQDN. Copying the
    # literal old server FQDN to a replacement server would be wrong, so translate
    # that server-relative value to the target server's own FQDN. A custom/shared
    # FQDN remains a literal migration value.
    if ($PropertyName -eq 'Fqdn' -and
        -not [string]::IsNullOrWhiteSpace([string]$sourceValue) -and
        -not [string]::IsNullOrWhiteSpace([string]$SourceConnector.ServerFqdn) -and
        ([string]$sourceValue).Equals([string]$SourceConnector.ServerFqdn, [System.StringComparison]::OrdinalIgnoreCase)) {
        $desiredValue = $TargetSnapshot.ServerFqdn
        $mapping = 'Source server FQDN translated to target server FQDN'
    }

    return [PSCustomObject]@{
        SourceValue  = $sourceValue
        DesiredValue = $desiredValue
        Mapping      = $mapping
    }
}

function Test-TargetTlsCertificateAvailability {
    param(
        [Parameter(Mandatory = $true)][string]$Server,
        [AllowNull()][string]$TlsCertificateName
    )

    if ([string]::IsNullOrWhiteSpace($TlsCertificateName)) { return 'NotApplicable' }

    try {
        $certificates = @(Get-ExchangeCertificate -Server $Server -ErrorAction Stop)
        foreach ($certificate in $certificates) {
            $identifier = "<I>$($certificate.Issuer)<S>$($certificate.Subject)"
            if (-not $identifier.Equals($TlsCertificateName, [System.StringComparison]::OrdinalIgnoreCase)) { continue }

            $statusOK = $true
            if ($certificate.PSObject.Properties['Status']) {
                $statusOK = ([string]$certificate.Status) -ieq 'Valid'
            }

            $smtpOK = $true
            if ($certificate.PSObject.Properties['Services']) {
                $smtpOK = ([string]$certificate.Services) -match '(?i)SMTP'
            }

            if ($statusOK -and $smtpOK) { return 'Available' }
        }
        return 'Missing'
    }
    catch {
        return 'Unknown'
    }
}

# ---------------------------------------------------------------------------
# Plan construction
# ---------------------------------------------------------------------------
function New-ConnectorPlan {
    param(
        [Parameter(Mandatory = $true)]$SourceConnector,
        [Parameter(Mandatory = $true)]$TargetSnapshot,
        [switch]$ManageAnonymousRelay
    )

    $targetConnector = Get-TargetConnector -TargetSnapshot $TargetSnapshot -SourceConnector $SourceConnector
    $targetName = Get-TargetConnectorName -SourceConnector $SourceConnector -TargetServer $TargetSnapshot.Server
    $targetExists = $null -ne $targetConnector
    $canCreate = (-not $targetExists) -and $SourceConnector.ConnectorType -eq 'Custom'
    $bindingStatus = Test-TargetBindingAvailability -Server $TargetSnapshot.Server -Bindings $SourceConnector.Values['Bindings']

    $reviewNotes = New-Object System.Collections.ArrayList
    if (-not $targetExists -and $SourceConnector.ConnectorType -eq 'BuiltIn') {
        [void]$reviewNotes.Add('Built-in connector is missing on the target. The script will not recreate Exchange built-in connectors.')
    }
    if ($bindingStatus -eq 'Missing') {
        [void]$reviewNotes.Add('One or more source binding IP addresses are not configured on the target server. Binding changes are skipped.')
    }
    elseif ($bindingStatus -eq 'Unknown') {
        [void]$reviewNotes.Add('Target binding IP addresses could not be verified remotely. Specific binding changes are skipped.')
    }
    if ($SourceConnector.ConnectorType -eq 'BuiltIn' -and $SourceConnector.AnonymousRelay.Enabled -eq $true) {
        [void]$reviewNotes.Add('Anonymous relay is enabled on a built-in connector. Microsoft recommends a dedicated Front End Transport connector with tightly scoped RemoteIPRanges.')
    }

    $comparisons = New-Object System.Collections.ArrayList
    $tlsCertificateStatus = 'NotChecked'
    foreach ($property in $ConnectorPropertyMap) {
        $desired = Get-DesiredConnectorPropertyValue -PropertyName $property.Name -SourceConnector $SourceConnector -TargetSnapshot $TargetSnapshot
        $sourceValue = $desired.SourceValue
        $desiredValue = $desired.DesiredValue
        $targetValue = if ($targetExists) { $targetConnector.Values[$property.Name] } else { $null }
        $status = if (-not $targetExists) { 'Missing' } elseif (Test-EquivalentValue -Left $desiredValue -Right $targetValue) { 'Same' } else { 'Different' }
        $manage = $false
        $note = $null

        if ($property.Policy -eq 'ReviewOnly') {
            $note = 'Review Only on existing connectors.'
        }
        elseif ($SourceConnector.ConnectorType -eq 'BuiltIn' -and -not $targetExists) {
            $note = 'Built-in connector will not be created automatically.'
        }
        elseif ($property.Name -eq 'Bindings' -and $status -ne 'Same' -and $bindingStatus -in @('Missing','Unknown')) {
            $note = "Binding safety check: $bindingStatus."
        }
        elseif ($property.Name -eq 'TlsCertificateName' -and $status -ne 'Same' -and -not [string]::IsNullOrWhiteSpace([string]$desiredValue)) {
            $tlsCertificateStatus = Test-TargetTlsCertificateAvailability -Server $TargetSnapshot.Server -TlsCertificateName ([string]$desiredValue)
            if ($tlsCertificateStatus -eq 'Available') {
                $manage = $true
            }
            else {
                $note = "Target TLS certificate check: $tlsCertificateStatus. TlsCertificateName is not changed automatically."
            }
        }
        elseif ($property.Policy -eq 'Apply' -and $status -ne 'Same') {
            $manage = $true
        }

        # TransportRole is required when creating a missing custom connector, but
        # it remains Review Only for an existing connector.
        if ($property.Name -eq 'TransportRole' -and $canCreate) {
            $manage = $true
            $note = 'Used only when the missing custom connector is created.'
        }

        [void]$comparisons.Add([PSCustomObject]@{
            Property     = $property.Name
            Policy       = $property.Policy
            SourceValue  = $sourceValue
            DesiredValue = $desiredValue
            TargetValue  = $targetValue
            Mapping      = $desired.Mapping
            Status       = $status
            Manage       = $manage
            ReviewNote   = $note
        })
    }

    # A custom connector cannot be safely created if a specific source binding is
    # missing or cannot be verified on the target.
    if ($canCreate -and $bindingStatus -in @('Missing','Unknown')) {
        $canCreate = $false
        foreach ($comparison in $comparisons) { $comparison.Manage = $false }
        [void]$reviewNotes.Add('Missing custom connector was not marked for creation because its binding could not be safely validated.')
    }

    $sourceRelay = $SourceConnector.AnonymousRelay
    $targetRelay = if ($targetExists) {
        $targetConnector.AnonymousRelay
    }
    else {
        [PSCustomObject]@{ Enabled = $false; QueryStatus = 'Missing'; Rights = @(); Error = $null }
    }

    $relayStatus = if ($sourceRelay.QueryStatus -ne 'Available') {
        'Unknown'
    }
    elseif ($targetExists -and $targetRelay.QueryStatus -ne 'Available') {
        'Unknown'
    }
    elseif ($sourceRelay.Enabled -eq $targetRelay.Enabled) {
        'Same'
    }
    else {
        'Different'
    }

    $relayPolicy = if ($ManageAnonymousRelay) { 'Apply' } else { 'ReviewOnly' }
    $relayManage = $false
    $relayNote = $null

    if ($sourceRelay.QueryStatus -ne 'Available') {
        $relayNote = "Source anonymous relay permission could not be queried: $($sourceRelay.Error)"
    }
    elseif ($targetExists -and $targetRelay.QueryStatus -ne 'Available') {
        $relayNote = "Target anonymous relay permission could not be queried: $($targetRelay.Error)"
    }
    elseif (-not $targetExists -and $SourceConnector.ConnectorType -eq 'BuiltIn') {
        $relayNote = 'Built-in connector is missing and will not be created.'
    }
    elseif ($relayStatus -eq 'Different' -and $ManageAnonymousRelay) {
        if ($targetExists -or $canCreate) { $relayManage = $true }
    }
    elseif ($relayStatus -eq 'Different') {
        $relayNote = 'Review Only. Use -IncludeAnonymousRelayPermissions to align the explicit anonymous relay permission.'
    }

    return [PSCustomObject]@{
        SourceServer      = $SourceConnector.Server
        TargetServer      = $TargetSnapshot.Server
        SourceName        = $SourceConnector.Name
        TargetName        = $targetName
        ConnectorType     = $SourceConnector.ConnectorType
        BuiltInFamily     = $SourceConnector.BuiltInFamily
        TargetExists      = $targetExists
        CanCreate         = $canCreate
        BindingStatus     = $bindingStatus
        TlsCertificateStatus = $tlsCertificateStatus
        SourceConnector   = $SourceConnector
        TargetConnector   = $targetConnector
        Comparisons       = @($comparisons)
        AnonymousRelay    = [PSCustomObject]@{
            SourceEnabled = $sourceRelay.Enabled
            TargetEnabled = $targetRelay.Enabled
            SourceStatus  = $sourceRelay.QueryStatus
            TargetStatus  = $targetRelay.QueryStatus
            Status        = $relayStatus
            Policy        = $relayPolicy
            Manage        = $relayManage
            ReviewNote    = $relayNote
        }
        ReviewNotes       = @($reviewNotes)
    }
}

function Test-PlanHasManagedChanges {
    param([Parameter(Mandatory = $true)]$Plan)

    if ($Plan.CanCreate) { return $true }
    if (@($Plan.Comparisons | Where-Object { $_.Manage -and $_.Status -ne 'Same' }).Count -gt 0) { return $true }
    if ($Plan.AnonymousRelay.Manage -and $Plan.AnonymousRelay.Status -eq 'Different') { return $true }
    return $false
}

# ---------------------------------------------------------------------------
# Display
# ---------------------------------------------------------------------------
function Show-ComparisonPlan {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Plans)

    Write-Host ''
    Write-Host 'Receive Connector comparison' -ForegroundColor Cyan
    Write-Host '----------------------------' -ForegroundColor Cyan

    foreach ($plan in ($Plans | Sort-Object TargetServer,TargetName)) {
        Write-Host ''
        Write-Host "[$($plan.TargetServer)] $($plan.SourceName) -> $($plan.TargetName)" -ForegroundColor Yellow
        Write-Host "  Type            : $($plan.ConnectorType)"
        Write-Host "  Target Exists   : $($plan.TargetExists)"
        Write-Host "  Binding Check   : $($plan.BindingStatus)"
        if ($plan.TlsCertificateStatus -ne 'NotChecked') { Write-Host "  TLS Cert Check  : $($plan.TlsCertificateStatus)" }

        foreach ($item in $plan.Comparisons) {
            $statusColor = if ($item.Status -eq 'Same') { 'Green' } elseif ($item.Status -eq 'Different' -or $item.Status -eq 'Missing') { 'Red' } else { 'Yellow' }
            Write-Host "  $($item.Property)"
            Write-Host "    Source        : $(ConvertTo-DisplayValue -Value $item.SourceValue)"
            if (-not (Test-EquivalentValue -Left $item.SourceValue -Right $item.DesiredValue)) { Write-Host "    Proposed      : $(ConvertTo-DisplayValue -Value $item.DesiredValue)" -ForegroundColor Green }
            Write-Host "    Target        : $(if ($plan.TargetExists) { ConvertTo-DisplayValue -Value $item.TargetValue } else { '<missing connector>' })"
            Write-Host "    Policy        : $($item.Policy)"
            Write-Host "    Status        : $($item.Status)" -ForegroundColor $statusColor
            if ($item.ReviewNote) { Write-Host "    Review        : $($item.ReviewNote)" -ForegroundColor Yellow }
        }

        Write-Host '  Anonymous Relay'
        Write-Host "    Source        : $(if ($null -eq $plan.AnonymousRelay.SourceEnabled) { 'Unknown' } else { [string]$plan.AnonymousRelay.SourceEnabled })"
        Write-Host "    Target        : $(if ($null -eq $plan.AnonymousRelay.TargetEnabled) { 'Unknown' } else { [string]$plan.AnonymousRelay.TargetEnabled })"
        Write-Host "    Permission    : $AnonymousRelayRight"
        Write-Host "    Policy        : $($plan.AnonymousRelay.Policy)"
        Write-Host "    Status        : $($plan.AnonymousRelay.Status)" -ForegroundColor $(if ($plan.AnonymousRelay.Status -eq 'Same') { 'Green' } elseif ($plan.AnonymousRelay.Status -eq 'Different') { 'Red' } else { 'Yellow' })
        if ($plan.AnonymousRelay.ReviewNote) { Write-Host "    Review        : $($plan.AnonymousRelay.ReviewNote)" -ForegroundColor Yellow }

        foreach ($note in $plan.ReviewNotes) { Write-Host "  REVIEW          : $note" -ForegroundColor Yellow }
    }

    Write-Host ''
    Write-Host 'Summary' -ForegroundColor Cyan
    Write-Host '-------' -ForegroundColor Cyan
    foreach ($target in @($Plans.TargetServer | Select-Object -Unique)) {
        $targetPlans = @($Plans | Where-Object { $_.TargetServer -eq $target })
        $different = 0
        $reviewOnly = 0
        foreach ($plan in $targetPlans) {
            $different += @($plan.Comparisons | Where-Object { $_.Status -ne 'Same' }).Count
            if ($plan.AnonymousRelay.Status -ne 'Same') { $different++ }
            $reviewOnly += @($plan.Comparisons | Where-Object { $_.Policy -eq 'ReviewOnly' -and $_.Status -ne 'Same' }).Count
            if ($plan.AnonymousRelay.Policy -eq 'ReviewOnly' -and $plan.AnonymousRelay.Status -ne 'Same') { $reviewOnly++ }
        }
        Write-Host "${target}: Connectors=$($targetPlans.Count) Differences=$different ReviewOnly=$reviewOnly" -ForegroundColor Cyan
    }
}

function Show-ChangePlan {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Plans)

    Write-Host ''
    Write-Host 'Preview' -ForegroundColor Cyan
    Write-Host '-------' -ForegroundColor Cyan

    $managedCount = 0
    foreach ($plan in ($Plans | Sort-Object TargetServer,TargetName)) {
        $hasManaged = Test-PlanHasManagedChanges -Plan $plan
        $hasReview = @($plan.Comparisons | Where-Object { $_.Status -ne 'Same' -and -not $_.Manage }).Count -gt 0 -or
            ($plan.AnonymousRelay.Status -ne 'Same' -and -not $plan.AnonymousRelay.Manage) -or $plan.ReviewNotes.Count -gt 0

        if (-not $hasManaged -and -not $hasReview) { continue }

        Write-Host ''
        Write-Host "[$($plan.TargetServer)] $($plan.SourceName) -> $($plan.TargetName)" -ForegroundColor Yellow
        Write-Host "  Type            : $($plan.ConnectorType)"

        if ($plan.CanCreate) {
            Write-Host '  Action          : CREATE custom connector' -ForegroundColor Green
            $managedCount++
        }
        elseif (-not $plan.TargetExists -and $plan.ConnectorType -eq 'BuiltIn') {
            Write-Host '  Action          : REVIEW ONLY - built-in connector missing' -ForegroundColor Yellow
        }

        foreach ($item in @($plan.Comparisons | Where-Object { $_.Status -ne 'Same' })) {
            if ($item.Manage -and $item.Property -ne 'TransportRole') {
                Write-Host "  SET $($item.Property)"
                Write-Host "    Current       : $(if ($plan.TargetExists) { ConvertTo-DisplayValue -Value $item.TargetValue } else { '<missing connector>' })"
                Write-Host "    New           : $(ConvertTo-DisplayValue -Value $item.DesiredValue)" -ForegroundColor Green
                $managedCount++
            }
            elseif (-not $item.Manage) {
                Write-Host "  REVIEW $($item.Property): $(ConvertTo-DisplayValue -Value $item.TargetValue) -> $(ConvertTo-DisplayValue -Value $item.DesiredValue)" -ForegroundColor Yellow
                if ($item.ReviewNote) { Write-Host "    Reason        : $($item.ReviewNote)" -ForegroundColor Yellow }
            }
        }

        if ($plan.AnonymousRelay.Status -eq 'Different') {
            if ($plan.AnonymousRelay.Manage) {
                $relayAction = if ($plan.AnonymousRelay.SourceEnabled) { 'ENABLE' } else { 'DISABLE' }
                Write-Host "  Anonymous Relay: $relayAction ($AnonymousRelayRight)" -ForegroundColor Green
                $managedCount++
            }
            else {
                Write-Host "  Anonymous Relay: REVIEW ONLY ($AnonymousRelayRight)" -ForegroundColor Yellow
                if ($plan.AnonymousRelay.ReviewNote) { Write-Host "    Reason        : $($plan.AnonymousRelay.ReviewNote)" -ForegroundColor Yellow }
            }
        }

        foreach ($note in $plan.ReviewNotes) { Write-Host "  REVIEW          : $note" -ForegroundColor Yellow }
    }

    if ($managedCount -eq 0) {
        Write-Host ''
        Write-Host 'No automatic Receive Connector changes are required. Review-only differences may still be listed above.' -ForegroundColor Green
    }
}

# ---------------------------------------------------------------------------
# Operations
# ---------------------------------------------------------------------------
function New-SetCommandLine {
    param(
        [Parameter(Mandatory = $true)][string]$CommandName,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Parameters,
        [string[]]$Switches
    )

    $parts = New-Object System.Collections.ArrayList
    [void]$parts.Add($CommandName)
    foreach ($key in $Parameters.Keys) {
        [void]$parts.Add("-$key")
        [void]$parts.Add((ConvertTo-PowerShellLiteral -Value $Parameters[$key]))
    }
    foreach ($switchName in @($Switches)) { if ($switchName) { [void]$parts.Add("-$switchName") } }
    [void]$parts.Add('-ErrorAction')
    [void]$parts.Add('Stop')
    return ($parts -join ' ')
}

function Get-TargetIdentityText {
    param([Parameter(Mandatory = $true)]$Plan)
    if ($Plan.TargetExists) { return [string]$Plan.TargetConnector.Identity }
    return "$($Plan.TargetServer)\$($Plan.TargetName)"
}

function New-TargetOperations {
    param([Parameter(Mandatory = $true)]$Plan)

    $operations = New-Object System.Collections.ArrayList
    $identityText = Get-TargetIdentityText -Plan $Plan

    if ($Plan.CanCreate) {
        $newParams = [ordered]@{
            Name           = $Plan.TargetName
            Server         = $Plan.TargetServer
            TransportRole  = $Plan.SourceConnector.Values['TransportRole']
            Bindings       = $Plan.SourceConnector.Values['Bindings']
            RemoteIPRanges = $Plan.SourceConnector.Values['RemoteIPRanges']
        }
        [void]$operations.Add([PSCustomObject]@{
            Type       = 'ExchangeCmdlet'
            Command    = 'New-ReceiveConnector'
            Parameters = $newParams
            Switches   = @('Custom')
            Note       = 'Create missing custom Receive Connector.'
        })
    }

    if ($Plan.TargetExists -or $Plan.CanCreate) {
        $setParams = [ordered]@{ Identity = $identityText }
        foreach ($item in @($Plan.Comparisons | Where-Object { $_.Manage -and $_.Status -ne 'Same' })) {
            if ($item.Property -eq 'TransportRole') { continue }
            if ($Plan.CanCreate -and $item.Property -in @('Bindings','RemoteIPRanges')) { continue }
            $setParams[$item.Property] = $item.DesiredValue
        }

        if ($setParams.Count -gt 1) {
            [void]$operations.Add([PSCustomObject]@{
                Type       = 'ExchangeCmdlet'
                Command    = 'Set-ReceiveConnector'
                Parameters = $setParams
                Switches   = @()
                Note       = 'Align supported Receive Connector properties.'
            })
        }
    }

    if ($Plan.AnonymousRelay.Manage -and $Plan.AnonymousRelay.Status -eq 'Different' -and ($Plan.TargetExists -or $Plan.CanCreate)) {
        [void]$operations.Add([PSCustomObject]@{
            Type       = $(if ($Plan.AnonymousRelay.SourceEnabled) { 'AddAnonymousRelay' } else { 'RemoveAnonymousRelay' })
            Command    = $(if ($Plan.AnonymousRelay.SourceEnabled) { 'Add-ADPermission' } else { 'Remove-ADPermission' })
            Parameters = [ordered]@{ Identity = $identityText }
            Switches   = @()
            Note       = 'Align explicit anonymous relay permission.'
        })
    }

    return @($operations)
}

function Invoke-TargetOperation {
    param([Parameter(Mandatory = $true)]$Operation)

    if ($Operation.Type -eq 'ExchangeCmdlet') {
        $parameters = $Operation.Parameters
        if ($Operation.Switches -contains 'Custom') {
            & $Operation.Command @parameters -Custom -ErrorAction Stop | Out-Null
        }
        else {
            & $Operation.Command @parameters -ErrorAction Stop | Out-Null
        }
        return
    }

    $identity = $Operation.Parameters['Identity']
    $connector = Get-ReceiveConnector -Identity $identity -ErrorAction Stop

    if ($Operation.Type -eq 'AddAnonymousRelay') {
        $connector | Add-ADPermission -User $AnonymousLogonPrincipal -ExtendedRights $AnonymousRelayRight -ErrorAction Stop | Out-Null
        return
    }

    if ($Operation.Type -eq 'RemoveAnonymousRelay') {
        $connector | Remove-ADPermission -User $AnonymousLogonPrincipal -ExtendedRights $AnonymousRelayRight -Confirm:$false -ErrorAction Stop | Out-Null
        return
    }

    throw "Unknown operation type '$($Operation.Type)'."
}

function Invoke-ConfigurationPlan {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Plans)

    foreach ($plan in ($Plans | Sort-Object TargetServer,TargetName)) {
        $operations = @(New-TargetOperations -Plan $plan)
        if ($operations.Count -eq 0) { continue }

        Write-Host "Configuring [$($plan.TargetServer)] $($plan.TargetName)..." -ForegroundColor Cyan
        foreach ($operation in $operations) {
            Invoke-TargetOperation -Operation $operation
        }
    }
}

# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------
function ConvertTo-RelayCommandLine {
    param([Parameter(Mandatory = $true)]$Operation)

    $identity = ConvertTo-PowerShellLiteral -Value $Operation.Parameters['Identity']
    $user = ConvertTo-PowerShellLiteral -Value $AnonymousLogonPrincipal
    $right = ConvertTo-PowerShellLiteral -Value $AnonymousRelayRight

    if ($Operation.Type -eq 'AddAnonymousRelay') {
        return "Get-ReceiveConnector -Identity $identity -ErrorAction Stop | Add-ADPermission -User $user -ExtendedRights $right -ErrorAction Stop"
    }
    if ($Operation.Type -eq 'RemoveAnonymousRelay') {
        return "Get-ReceiveConnector -Identity $identity -ErrorAction Stop | Remove-ADPermission -User $user -ExtendedRights $right -Confirm:`$false -ErrorAction Stop"
    }
    throw "Unsupported relay operation type '$($Operation.Type)'."
}

function Export-CommandPlan {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Plans,
        [Parameter(Mandatory = $true)][string[]]$Targets,
        [Parameter(Mandatory = $true)][string]$Source,
        [switch]$IncludeRelay
    )

    $lines = New-Object System.Collections.ArrayList
    foreach ($line in @(
        '# ConfigureExchangeReceiveConnectors-v1.1.ps1 generated command file'
        '# Tool Author : Ceyhun Kirmizitas'
        "# Generated   : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        "# Operator    : $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
        "# Source      : $Source"
        "# Targets     : $($Targets -join ', ')"
        "# Anonymous relay permissions: $(if ($IncludeRelay) { 'Included' } else { 'Review Only' })"
        '# WARNING: Review every command before running it in Exchange Management Shell.'
        '# No Exchange configuration changes were applied by ConfigureExchangeReceiveConnectors-v1.1.ps1.'
        ''
    )) { [void]$lines.Add($line) }

    $operationCount = 0
    foreach ($plan in ($Plans | Sort-Object TargetServer,TargetName)) {
        $operations = @(New-TargetOperations -Plan $plan)
        $hasReview = @($plan.Comparisons | Where-Object { $_.Status -ne 'Same' -and -not $_.Manage }).Count -gt 0 -or
            ($plan.AnonymousRelay.Status -ne 'Same' -and -not $plan.AnonymousRelay.Manage) -or $plan.ReviewNotes.Count -gt 0

        if ($operations.Count -eq 0 -and -not $hasReview) { continue }

        [void]$lines.Add('# =====================================================================')
        [void]$lines.Add("# Target: $($plan.TargetServer) | Connector: $($plan.TargetName) | Type: $($plan.ConnectorType)")
        [void]$lines.Add('# =====================================================================')

        foreach ($note in $plan.ReviewNotes) { [void]$lines.Add("# REVIEW: $note") }
        foreach ($item in @($plan.Comparisons | Where-Object { $_.Status -ne 'Same' -and -not $_.Manage })) {
            [void]$lines.Add("# REVIEW: $($item.Property): $(ConvertTo-DisplayValue -Value $item.TargetValue) -> $(ConvertTo-DisplayValue -Value $item.DesiredValue). $($item.ReviewNote)")
        }
        if ($plan.AnonymousRelay.Status -ne 'Same' -and -not $plan.AnonymousRelay.Manage) {
            [void]$lines.Add("# REVIEW: Anonymous Relay: target=$(ConvertTo-DisplayValue -Value $plan.AnonymousRelay.TargetEnabled) source=$(ConvertTo-DisplayValue -Value $plan.AnonymousRelay.SourceEnabled). $($plan.AnonymousRelay.ReviewNote)")
        }

        foreach ($operation in $operations) {
            [void]$lines.Add("# $($operation.Note)")
            if ($operation.Type -eq 'ExchangeCmdlet') {
                [void]$lines.Add((New-SetCommandLine -CommandName $operation.Command -Parameters $operation.Parameters -Switches $operation.Switches))
            }
            else {
                [void]$lines.Add((ConvertTo-RelayCommandLine -Operation $operation))
            }
            [void]$lines.Add('')
            $operationCount++
        }

        if ($operations.Count -eq 0) { [void]$lines.Add('') }
    }

    if ($operationCount -eq 0) { [void]$lines.Add('# No executable changes are required.') }
    Set-Content -LiteralPath $Path -Value @($lines) -Encoding UTF8
}

function Export-ComparisonReport {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Plans,
        [Parameter(Mandatory = $true)][string[]]$Targets,
        [Parameter(Mandatory = $true)][string]$Source,
        [switch]$IncludeRelay
    )

    $lines = New-Object System.Collections.ArrayList
    foreach ($line in @(
        '# ConfigureExchangeReceiveConnectors-v1.1.ps1 comparison report'
        '# Tool Author : Ceyhun Kirmizitas'
        "# Generated   : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        "# Operator    : $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
        "# Source      : $Source"
        "# Targets     : $($Targets -join ', ')"
        "# Anonymous relay permissions: $(if ($IncludeRelay) { 'Apply policy selected' } else { 'Review Only' })"
        '# Mode        : Compare only (read-only)'
        '# No Set-* commands were generated and no Exchange changes were applied.'
        ''
    )) { [void]$lines.Add($line) }

    foreach ($plan in ($Plans | Sort-Object TargetServer,TargetName)) {
        [void]$lines.Add('=====================================================================')
        [void]$lines.Add("Target: $($plan.TargetServer)")
        [void]$lines.Add("Source Connector: $($plan.SourceName)")
        [void]$lines.Add("Target Connector: $($plan.TargetName)")
        [void]$lines.Add("Type: $($plan.ConnectorType)")
        [void]$lines.Add("Target Exists: $($plan.TargetExists)")
        [void]$lines.Add("Binding Check: $($plan.BindingStatus)")
        if ($plan.TlsCertificateStatus -ne 'NotChecked') { [void]$lines.Add("TLS Certificate Check: $($plan.TlsCertificateStatus)") }
        [void]$lines.Add('=====================================================================')
        [void]$lines.Add('')

        foreach ($item in $plan.Comparisons) {
            [void]$lines.Add($item.Property)
            [void]$lines.Add("  Source   : $(ConvertTo-DisplayValue -Value $item.SourceValue)")
            [void]$lines.Add("  Proposed : $(ConvertTo-DisplayValue -Value $item.DesiredValue)")
            [void]$lines.Add("  Target   : $(if ($plan.TargetExists) { ConvertTo-DisplayValue -Value $item.TargetValue } else { '<missing connector>' })")
            [void]$lines.Add("  Mapping  : $($item.Mapping)")
            [void]$lines.Add("  Policy : $($item.Policy)")
            [void]$lines.Add("  Status : $($item.Status)")
            if ($item.ReviewNote) { [void]$lines.Add("  Review : $($item.ReviewNote)") }
            [void]$lines.Add('')
        }

        [void]$lines.Add('Anonymous Relay')
        [void]$lines.Add("  Principal  : $AnonymousLogonPrincipal")
        [void]$lines.Add("  Permission : $AnonymousRelayRight")
        [void]$lines.Add("  Source     : $(if ($null -eq $plan.AnonymousRelay.SourceEnabled) { 'Unknown' } else { [string]$plan.AnonymousRelay.SourceEnabled })")
        [void]$lines.Add("  Target     : $(if ($null -eq $plan.AnonymousRelay.TargetEnabled) { 'Unknown' } else { [string]$plan.AnonymousRelay.TargetEnabled })")
        [void]$lines.Add("  Policy     : $($plan.AnonymousRelay.Policy)")
        [void]$lines.Add("  Status     : $($plan.AnonymousRelay.Status)")
        if ($plan.AnonymousRelay.ReviewNote) { [void]$lines.Add("  Review     : $($plan.AnonymousRelay.ReviewNote)") }
        [void]$lines.Add('')

        foreach ($note in $plan.ReviewNotes) { [void]$lines.Add("REVIEW: $note") }
        if ($plan.ReviewNotes.Count -gt 0) { [void]$lines.Add('') }
    }

    Set-Content -LiteralPath $Path -Value @($lines) -Encoding UTF8
}

# ---------------------------------------------------------------------------
# Verification
# ---------------------------------------------------------------------------
function Test-ConfigurationPlan {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Plans)

    $results = New-Object System.Collections.ArrayList
    $freshByTarget = @{}

    foreach ($target in @($Plans.TargetServer | Select-Object -Unique)) {
        $freshByTarget[$target] = Get-ReceiveConnectorSnapshot -Server $target
    }

    foreach ($plan in $Plans) {
        $freshSnapshot = $freshByTarget[$plan.TargetServer]
        $freshConnector = @($freshSnapshot.Items | Where-Object { $_.Name -ieq $plan.TargetName } | Select-Object -First 1)

        if ($freshConnector.Count -eq 0) {
            $wasExpectedToExist = $plan.TargetExists -or $plan.CanCreate
            [void]$results.Add([PSCustomObject]@{
                Target    = $plan.TargetServer
                Connector = $plan.TargetName
                Item      = 'Connector'
                Status    = $(if ($wasExpectedToExist) { 'Mismatch' } else { 'Skipped' })
                Expected  = $(if ($wasExpectedToExist) { 'Present' } else { 'Not managed' })
                Actual    = 'Missing'
            })
            continue
        }

        $fresh = $freshConnector[0]
        foreach ($comparison in $plan.Comparisons) {
            if (-not $comparison.Manage) { continue }
            $actual = $fresh.Values[$comparison.Property]
            $expected = $comparison.DesiredValue
            $status = if (Test-EquivalentValue -Left $actual -Right $expected) { 'Verified' } else { 'Mismatch' }
            [void]$results.Add([PSCustomObject]@{
                Target    = $plan.TargetServer
                Connector = $plan.TargetName
                Item      = $comparison.Property
                Status    = $status
                Expected  = ConvertTo-DisplayValue -Value $expected
                Actual    = ConvertTo-DisplayValue -Value $actual
            })
        }

        if ($plan.AnonymousRelay.Manage) {
            $actualRelay = $fresh.AnonymousRelay
            $status = if ($actualRelay.QueryStatus -eq 'Available' -and $actualRelay.Enabled -eq $plan.AnonymousRelay.SourceEnabled) { 'Verified' } else { 'Mismatch' }
            [void]$results.Add([PSCustomObject]@{
                Target    = $plan.TargetServer
                Connector = $plan.TargetName
                Item      = 'Anonymous Relay'
                Status    = $status
                Expected  = [string]$plan.AnonymousRelay.SourceEnabled
                Actual    = $(if ($actualRelay.QueryStatus -eq 'Available') { [string]$actualRelay.Enabled } else { 'Unknown' })
            })
        }
    }

    return @($results)
}

function Show-Verification {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Results)

    Write-Host ''
    Write-Host 'Verification summary' -ForegroundColor Cyan
    Write-Host '--------------------' -ForegroundColor Cyan

    if ($Results.Count -eq 0) {
        Write-Host 'No managed changes required verification.' -ForegroundColor Green
        return
    }

    foreach ($result in $Results) {
        $color = if ($result.Status -eq 'Verified') { 'Green' } elseif ($result.Status -eq 'Mismatch') { 'Red' } else { 'Yellow' }
        Write-Host "[$($result.Target)] $($result.Connector) - $($result.Item): $($result.Status)" -ForegroundColor $color
        if ($result.Status -eq 'Mismatch') {
            Write-Host "  Expected : $($result.Expected)"
            Write-Host "  Actual   : $($result.Actual)"
        }
    }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
Initialize-ExchangeShell

$sourceServerName = Resolve-ExchangeServerName -Identity $SourceServer
$resolvedTargets = @($TargetServers | ForEach-Object { Resolve-ExchangeServerName -Identity $_ })
$targets = Get-NormalizedTargets -Servers $resolvedTargets
if (@($targets | Where-Object { $_ -ieq $sourceServerName }).Count -gt 0) {
    throw 'SourceServer cannot also be listed in TargetServers.'
}

Write-Host "Source server : $sourceServerName" -ForegroundColor Cyan
Write-Host "Target servers: $($targets -join ', ')" -ForegroundColor Cyan
Write-Host "Anonymous relay permissions: $(if ($IncludeAnonymousRelayPermissions) { 'Included' } else { 'Review Only' })" -ForegroundColor Cyan
if ($CompareOnly) { Write-Host 'Mode          : Compare only (read-only)' -ForegroundColor Cyan }

$sourceSnapshot = Get-ReceiveConnectorSnapshot -Server $sourceServerName
$sourceConnectors = @(Select-SourceConnector -SourceSnapshot $sourceSnapshot -Names $ConnectorName)
if ($sourceConnectors.Count -eq 0) { throw 'No source Receive Connectors were selected.' }

$targetSnapshots = @{}
foreach ($target in $targets) {
    $targetSnapshots[$target] = Get-ReceiveConnectorSnapshot -Server $target
}

$plans = New-Object System.Collections.ArrayList
foreach ($target in $targets) {
    foreach ($sourceConnector in $sourceConnectors) {
        [void]$plans.Add((New-ConnectorPlan -SourceConnector $sourceConnector -TargetSnapshot $targetSnapshots[$target] -ManageAnonymousRelay:$IncludeAnonymousRelayPermissions))
    }
}

if ($CompareOnly) {
    Show-ComparisonPlan -Plans @($plans)

    if (-not [string]::IsNullOrWhiteSpace($OutputFile)) {
        $outputPath = Resolve-OutputFilePath -Path $OutputFile
        Export-ComparisonReport -Path $outputPath -Plans @($plans) -Targets $targets -Source $sourceServerName -IncludeRelay:$IncludeAnonymousRelayPermissions
        Write-Host ''
        Write-Host "Comparison report created: $outputPath" -ForegroundColor Green
        Write-Host 'No Set-* commands were generated and no Exchange changes were applied.' -ForegroundColor Yellow
    }
    return
}

Show-ChangePlan -Plans @($plans)
$managedPlans = @($plans | Where-Object { Test-PlanHasManagedChanges -Plan $_ })

if (-not [string]::IsNullOrWhiteSpace($OutputFile)) {
    $outputPath = Resolve-OutputFilePath -Path $OutputFile
    Export-CommandPlan -Path $outputPath -Plans @($plans) -Targets $targets -Source $sourceServerName -IncludeRelay:$IncludeAnonymousRelayPermissions
    Write-Host ''
    Write-Host "Command file created: $outputPath" -ForegroundColor Green
    Write-Host 'No Exchange configuration changes were applied. Review the TXT file and run the required commands manually.' -ForegroundColor Yellow
    return
}

if ($managedPlans.Count -eq 0) {
    Write-Host ''
    Write-Host 'No automatic Receive Connector changes are required.' -ForegroundColor Green
    return
}

$confirmation = Read-Host 'Continue and apply these changes? [Y/N]'
if ($confirmation -notmatch '^[Yy]$') {
    Write-Host 'Operation cancelled. No Exchange configuration changes were applied.' -ForegroundColor Yellow
    return
}

$snapshotPath = Export-PreChangeSnapshot -Targets $targets -SourceServer $sourceServerName
Write-Host ''
Write-Host "Configuration snapshot created: $snapshotPath" -ForegroundColor Green

Invoke-ConfigurationPlan -Plans @($plans)
$verification = @(Test-ConfigurationPlan -Plans @($plans))
Show-Verification -Results $verification

if (@($verification | Where-Object { $_.Status -eq 'Mismatch' }).Count -gt 0) {
    throw 'Differences remain after configuration. Review the verification summary.'
}

Write-Host ''
Write-Host 'Receive Connector configuration completed and verified.' -ForegroundColor Green
