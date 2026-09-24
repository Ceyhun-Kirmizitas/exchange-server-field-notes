<#
.SYNOPSIS
Configures Exchange Server Client Access settings or exports the required Set-* commands for manual execution.

.DESCRIPTION
ConfigureExchangeURLs-v2.1.ps1 has two operating modes.

1. Interactive URL mode
   Run without -SourceServer. Internal, external, and Autodiscover SCP namespaces can be supplied as parameters or entered interactively.
   If -TargetServers is omitted, the local computer is used. Authentication settings are not changed.

2. Reference server mode
   Use -SourceServer with -TargetServers. URL/hostname settings are compared with the source server and applied to the target server(s).
   Add -IncludeAuthentication to include supported authentication settings from the source server, including OWA LogonFormat and DefaultDomain.
   Add -CompareOnly to display a read-only source/target comparison without applying changes.
   Migration-sensitive settings such as the PowerShell virtual directory, ASA/Kerberos state, EWS MRS Proxy, and Outlook Anywhere SSL offloading are shown as Review Only and are not changed automatically.

By default, the script shows a Current -> New preview, asks for confirmation, applies only the required changes, and verifies the target server(s) afterward.
Before a confirmed Apply, the script saves a fresh TXT configuration snapshot under the script folder's ConfigSnapshots directory. Existing snapshot files are never overwritten.
When -OutputFile is specified without -CompareOnly, the script does not apply changes and writes the required Exchange PowerShell commands to a TXT file for manual review.
When -CompareOnly and -OutputFile are used together, the comparison is displayed and also written to a TXT report. No Set-* commands are generated and no Exchange changes are applied.

.PARAMETER SourceServer
Reference Exchange server. Valid only in Reference mode.

.PARAMETER TargetServers
One or more Exchange servers to compare/configure. Required in Reference mode and optional in Interactive mode.

.PARAMETER IncludeAuthentication
Includes supported authentication settings from SourceServer. OWA LogonFormat and DefaultDomain are included. PowerShell virtual directory settings remain Review Only. Valid only in Reference mode.

.PARAMETER CompareOnly
Displays all compared SourceServer and TargetServers values without applying changes.
Different values are highlighted in red. Valid only in Reference mode. If -OutputFile is also specified, the comparison is written to a TXT report.

.PARAMETER InternalNamespace
Internal Client Access namespace for Interactive mode. If omitted, the script prompts for it.

.PARAMETER ExternalNamespace
External Client Access namespace for Interactive mode. If omitted, the script prompts and defaults to InternalNamespace.

.PARAMETER AutodiscoverSCPNamespace
Autodiscover SCP namespace for Interactive mode. If omitted, the script prompts with a suggested value.

.PARAMETER OutputFile
Optional TXT output path. In normal mode, exports the required PowerShell Set-* commands instead of applying changes. With -CompareOnly, exports a comparison report instead of Set-* commands.

.PARAMETER Help
Displays a short usage guide and exits without initializing Exchange Management Shell or making changes.

.EXAMPLE
.\ConfigureExchangeURLs-v2.1.ps1
Prompts for namespaces, previews the required changes for the local Exchange server, asks for confirmation, applies them, and verifies the result.

.EXAMPLE
.\ConfigureExchangeURLs-v2.1.ps1 -TargetServers EX01,EX02 -InternalNamespace mail.contoso.com -ExternalNamespace mail.contoso.com -AutodiscoverSCPNamespace autodiscover.contoso.com
Applies URL/hostname changes to EX01 and EX02 without changing authentication.

.EXAMPLE
.\ConfigureExchangeURLs-v2.1.ps1 -SourceServer EX01 -TargetServers EX02,EX03 -IncludeAuthentication
Aligns URL/hostname and supported authentication settings on EX02 and EX03 with EX01.

.EXAMPLE
.\ConfigureExchangeURLs-v2.1.ps1 -SourceServer EX01 -TargetServers EX02,EX03 -CompareOnly -IncludeAuthentication
Displays all URL/hostname and supported authentication values for EX01 versus EX02 and EX03. Different values are highlighted in red. No changes are applied.

.EXAMPLE
.\ConfigureExchangeURLs-v2.1.ps1 -SourceServer EX01 -TargetServers EX02,EX03 -IncludeAuthentication -OutputFile C:\Temp\ExchangeURL-Commands.txt
Does not change Exchange. Exports the required URL/hostname and authentication commands to a TXT file for manual review.

.EXAMPLE
.\ConfigureExchangeURLs-v2.1.ps1 -SourceServer EX01 -TargetServers EX02,EX03 -CompareOnly -OutputFile C:\Temp\ExchangeURL-Compare.txt
Displays the source/target comparison and writes the same read-only comparison to a TXT report. No Set-* commands are generated and no Exchange changes are applied.

.NOTES
Author  : Ceyhun Kirmizitas
Version : 2.1
Date    : 17/09/2026
Scope   : Exchange Server 2016, Exchange Server 2019, Exchange Server SE

Change Log
----------
2.1 - Added migration-safe Review Only handling for the PowerShell virtual directory,
      ASA/Kerberos visibility, EWS MRS Proxy review, Outlook Anywhere SSL offloading
      review, OWA LogonFormat/DefaultDomain migration, strict post-change verification,
      and automatic pre-change configuration snapshots before Apply.
2.0 - Initial release

Migration safety model:
- Apply: settings that the tool can safely align after preview/confirmation.
- Review Only: settings that are important during migration but are not changed automatically.
- Skip: settings outside this script's scope.

The PowerShell virtual directory is Review Only and is never changed automatically. Microsoft recommends modifying it only at the request of Microsoft Customer Service and Support.
The script intentionally does not copy Extended Protection settings, certificates, IIS bindings, or other server-specific configuration.
#>

[CmdletBinding(DefaultParameterSetName = 'Interactive')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Reference')]
    [ValidateNotNullOrEmpty()]
    [string]$SourceServer,

    [Parameter(Mandatory = $true, ParameterSetName = 'Reference')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')]
    [ValidateNotNullOrEmpty()]
    [string[]]$TargetServers,

    [Parameter(Mandatory = $false, ParameterSetName = 'Reference')]
    [switch]$IncludeAuthentication,

    [Parameter(Mandatory = $false, ParameterSetName = 'Reference')]
    [switch]$CompareOnly,

    [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')]
    [string]$InternalNamespace,

    [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')]
    [string]$ExternalNamespace,

    [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')]
    [string]$AutodiscoverSCPNamespace,

    [Parameter(Mandatory = $false, ParameterSetName = 'Reference')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')]
    [string]$OutputFile,

    [switch]$Help
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$script:ScriptBaseName = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)

if ($Help) {
    @"
ConfigureExchangeURLs-v2.1.ps1
Compare or configure Exchange client-access URLs, namespaces, and supported authentication settings.

COMMON USAGE
  Interactive namespace mode:
    .\ConfigureExchangeURLs-v2.1.ps1

  Set namespaces on selected servers:
    .\ConfigureExchangeURLs-v2.1.ps1 -TargetServers EX01,EX02 -InternalNamespace mail.contoso.com -ExternalNamespace mail.contoso.com -AutodiscoverSCPNamespace autodiscover.contoso.com

  Compare source and targets, including supported authentication:
    .\ConfigureExchangeURLs-v2.1.ps1 -SourceServer EX01 -TargetServers EX02,EX03 -CompareOnly -IncludeAuthentication

  Export required commands without applying:
    .\ConfigureExchangeURLs-v2.1.ps1 -SourceServer EX01 -TargetServers EX02,EX03 -IncludeAuthentication -OutputFile C:\Temp\ExchangeURL-Commands.txt

NOTES
  - Normal mode: preview -> confirmation -> apply -> verify.
  - -CompareOnly: no Exchange changes.
  - -OutputFile: exports commands/report and does not apply Exchange changes.
  - PowerShell vDir, ASA/Kerberos, MRS Proxy and SSL offload items remain Review Only where designed.
  - For full help:
      Get-Help .\ConfigureExchangeURLs-v2.1.ps1 -Full
"@ | Write-Host
    return
}

# ---------------------------------------------------------------------------
# Configuration Map
# ---------------------------------------------------------------------------
# Keep the service metadata in one place so discovery, comparison, export,
# and apply logic use the same cmdlet and authentication-property definitions.
$VirtualDirectoryMap = @(
    [PSCustomObject]@{ Name = 'OWA';          Key = 'Owa';          GetCmd = 'Get-OwaVirtualDirectory';          SetCmd = 'Set-OwaVirtualDirectory';          Path = '/owa';                         Policy = 'Apply';      AuthProperties = @('BasicAuthentication','DigestAuthentication','WindowsAuthentication','FormsAuthentication','AdfsAuthentication','OAuthAuthentication','LogonFormat','DefaultDomain'); ReviewProperties = @() },
    [PSCustomObject]@{ Name = 'ECP';          Key = 'Ecp';          GetCmd = 'Get-EcpVirtualDirectory';          SetCmd = 'Set-EcpVirtualDirectory';          Path = '/ecp';                         Policy = 'Apply';      AuthProperties = @('BasicAuthentication','DigestAuthentication','WindowsAuthentication','FormsAuthentication','AdfsAuthentication','OAuthAuthentication'); ReviewProperties = @() },
    [PSCustomObject]@{ Name = 'EWS';          Key = 'Ews';          GetCmd = 'Get-WebServicesVirtualDirectory';  SetCmd = 'Set-WebServicesVirtualDirectory';  Path = '/EWS/Exchange.asmx';           Policy = 'Apply';      AuthProperties = @('BasicAuthentication','DigestAuthentication','WindowsAuthentication','WSSecurityAuthentication','OAuthAuthentication','CertificateAuthentication'); ReviewProperties = @('MRSProxyEnabled') },
    [PSCustomObject]@{ Name = 'MAPI';         Key = 'Mapi';         GetCmd = 'Get-MapiVirtualDirectory';         SetCmd = 'Set-MapiVirtualDirectory';         Path = '/mapi';                        Policy = 'Apply';      AuthProperties = @('IISAuthenticationMethods'); ReviewProperties = @() },
    [PSCustomObject]@{ Name = 'ActiveSync';   Key = 'ActiveSync';   GetCmd = 'Get-ActiveSyncVirtualDirectory';   SetCmd = 'Set-ActiveSyncVirtualDirectory';   Path = '/Microsoft-Server-ActiveSync'; Policy = 'Apply';      AuthProperties = @('BasicAuthEnabled','WindowsAuthEnabled','ClientCertAuth'); ReviewProperties = @() },
    [PSCustomObject]@{ Name = 'OAB';          Key = 'Oab';          GetCmd = 'Get-OabVirtualDirectory';          SetCmd = 'Set-OabVirtualDirectory';          Path = '/OAB';                         Policy = 'Apply';      AuthProperties = @('BasicAuthentication','WindowsAuthentication','OAuthAuthentication'); ReviewProperties = @() },
    [PSCustomObject]@{ Name = 'PowerShell';   Key = 'PowerShell';   GetCmd = 'Get-PowerShellVirtualDirectory';   SetCmd = 'Set-PowerShellVirtualDirectory';   Path = '/powershell';                  Policy = 'ReviewOnly'; AuthProperties = @('BasicAuthentication','WindowsAuthentication','CertificateAuthentication'); ReviewProperties = @('RequireSSL') },
    [PSCustomObject]@{ Name = 'Autodiscover'; Key = 'Autodiscover'; GetCmd = 'Get-AutodiscoverVirtualDirectory'; SetCmd = 'Set-AutodiscoverVirtualDirectory'; Path = $null;                          Policy = 'Apply';      AuthProperties = @('BasicAuthentication','DigestAuthentication','WindowsAuthentication','WSSecurityAuthentication','OAuthAuthentication'); ReviewProperties = @() }
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
    if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) { throw 'Exchange Management Shell could not be initialized.' }
}

function Normalize-Namespace {
    param([Parameter(Mandatory = $true)][string]$Value)

    $valueToParse = $Value.Trim()
    if ([string]::IsNullOrWhiteSpace($valueToParse)) { throw 'Namespace cannot be empty.' }

    if ($valueToParse -match '^[a-zA-Z][a-zA-Z0-9+.-]*://') {
        try { return ([Uri]$valueToParse).DnsSafeHost }
        catch { throw "Invalid URL or namespace: $Value" }
    }

    $valueToParse = $valueToParse.Trim('/')
    if ($valueToParse.Contains('/')) { $valueToParse = $valueToParse.Split('/')[0] }
    if ([string]::IsNullOrWhiteSpace($valueToParse)) { throw "Invalid namespace: $Value" }
    return $valueToParse
}

function Get-SuggestedAutodiscoverNamespace {
    param([Parameter(Mandatory = $true)][string]$ClientAccessNamespace)

    $labels = @($ClientAccessNamespace.Split('.') | Where-Object { $_ })
    if ($labels.Count -ge 3) { return "autodiscover.$($labels[1..($labels.Count - 1)] -join '.')" }
    return "autodiscover.$ClientAccessNamespace"
}

# ---------------------------------------------------------------------------
# Build Configuration Snapshot
# ---------------------------------------------------------------------------
# Discovery is read-only. A snapshot is collected before any change plan is built.
function Get-DefaultVirtualDirectory {
    param(
        [Parameter(Mandatory = $true)][string]$CommandName,
        [Parameter(Mandatory = $true)][string]$Server
    )

    $items = @(& $CommandName -Server $Server -ErrorAction Stop)
    if ($items.Count -eq 0) { throw "No virtual directory was returned by $CommandName for server $Server." }

    $defaultSite = @($items | Where-Object {
        ([string]$_.Identity) -like '*Default Web Site*' -or ([string]$_.Name) -like '*Default Web Site*'
    })
    if ($defaultSite.Count -gt 0) { return $defaultSite[0] }
    return $items[0]
}

function Get-AlternateServiceAccountInfo {
    param([Parameter(Mandatory = $true)]$ClientAccessService)

    $result = [ordered]@{
        Account = 'Not Configured'
        Status  = 'Not Configured'
    }

    $property = $ClientAccessService.PSObject.Properties['AlternateServiceAccountConfiguration']
    if ($null -eq $property -or $null -eq $property.Value) { return [PSCustomObject]$result }

    $configuration = ([string]$property.Value).Trim()
    if ([string]::IsNullOrWhiteSpace($configuration) -or
        $configuration -match '(?i)^<?Not\s*Set>?$' -or
        $configuration -match '(?i)Latest:\s*<?Not\s*Set>?') {
        return [PSCustomObject]$result
    }

    $result.Account = 'Configured'
    $result.Status = 'Configured'
    if ($configuration -match '(?is)Latest:\s*(?<Latest>.*?)(?:\s+Previous:|$)' -and
        $Matches['Latest'] -match '(?<Account>[^\s,]+\\[^\s,]+)\s*$') {
        $result.Account = $Matches['Account']
    }

    return [PSCustomObject]$result
}

function Get-ExchangeClientAccessSnapshot {
    param([Parameter(Mandatory = $true)][string]$Server)

    $null = Get-ExchangeServer -Identity $Server -ErrorAction Stop
    $clientAccess = Get-ClientAccessService -Identity $Server -IncludeAlternateServiceAccountCredentialStatus -ErrorAction Stop
    $data = [ordered]@{
        Server       = $Server
        ClientAccess = $clientAccess
        ASA          = Get-AlternateServiceAccountInfo -ClientAccessService $clientAccess
        OutlookAny   = Get-DefaultVirtualDirectory -CommandName 'Get-OutlookAnywhere' -Server $Server
    }

    foreach ($entry in $VirtualDirectoryMap) {
        $data[$entry.Key] = Get-DefaultVirtualDirectory -CommandName $entry.GetCmd -Server $Server
    }
    return [PSCustomObject]$data
}

function ConvertTo-DisplayValue {
    param($Value)

    if ($null -eq $Value) { return '<null>' }
    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        return (@($Value) | ForEach-Object { [string]$_ }) -join ','
    }
    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return '<empty>' }
    return $text
}

function ConvertTo-ComparableValue {
    param($Value)

    if ($null -eq $Value) { return '<null>' }
    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        return (@($Value) | ForEach-Object { [string]$_ } | Sort-Object) -join '|'
    }
    return [string]$Value
}

function Get-SnapshotComponent {
    param(
        [Parameter(Mandatory = $true)]$Snapshot,
        [Parameter(Mandatory = $true)][string]$Key
    )

    return $Snapshot.PSObject.Properties[$Key].Value
}

# ---------------------------------------------------------------------------
# Pre-change configuration snapshot
# ---------------------------------------------------------------------------
# A real Apply always writes a fresh, human-readable snapshot first. CompareOnly
# and OutputFile modes remain read-only/export-only and do not create snapshots.
# Snapshot failure is intentionally fatal so configuration is never changed
# without a recoverable record of the pre-change state.
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

function Add-ClientAccessSnapshotLines {
    param(
        [Parameter(Mandatory = $true)][System.Collections.ArrayList]$Lines,
        [Parameter(Mandatory = $true)]$Snapshot,
        [Parameter(Mandatory = $true)][string]$Role
    )

    [void]$Lines.Add(("=== {0}: {1} ===" -f $Role, $Snapshot.Server))
    [void]$Lines.Add('')

    foreach ($entry in $VirtualDirectoryMap) {
        $component = Get-SnapshotComponent -Snapshot $Snapshot -Key $entry.Key
        [void]$Lines.Add(("[{0}]" -f $entry.Name))

        $propertyNames = New-Object System.Collections.ArrayList
        if ($entry.Path) {
            [void]$propertyNames.Add('InternalUrl')
            [void]$propertyNames.Add('ExternalUrl')
        }
        foreach ($name in $entry.AuthProperties) { [void]$propertyNames.Add($name) }
        foreach ($name in $entry.ReviewProperties) { [void]$propertyNames.Add($name) }

        foreach ($propertyName in @($propertyNames | Select-Object -Unique)) {
            $property = $component.PSObject.Properties[$propertyName]
            if ($property) {
                [void]$Lines.Add(("{0} = {1}" -f $propertyName, (ConvertTo-DisplayValue $property.Value)))
            }
        }
        [void]$Lines.Add('')
    }

    [void]$Lines.Add('[Autodiscover SCP]')
    [void]$Lines.Add(("AutoDiscoverServiceInternalUri = {0}" -f (ConvertTo-DisplayValue $Snapshot.ClientAccess.AutoDiscoverServiceInternalUri)))
    [void]$Lines.Add('')

    [void]$Lines.Add('[Alternate Service Account / Kerberos]')
    [void]$Lines.Add(("Account = {0}" -f (ConvertTo-DisplayValue $Snapshot.ASA.Account)))
    [void]$Lines.Add(("Status  = {0}" -f (ConvertTo-DisplayValue $Snapshot.ASA.Status)))
    [void]$Lines.Add('')

    [void]$Lines.Add('[Outlook Anywhere]')
    foreach ($propertyName in @(
        'InternalHostname','ExternalHostname','InternalClientsRequireSsl','ExternalClientsRequireSsl',
        'InternalClientAuthenticationMethod','ExternalClientAuthenticationMethod','IISAuthenticationMethods','SSLOffloading'
    )) {
        $property = $Snapshot.OutlookAny.PSObject.Properties[$propertyName]
        if ($property) {
            [void]$Lines.Add(("{0} = {1}" -f $propertyName, (ConvertTo-DisplayValue $property.Value)))
        }
    }
    [void]$Lines.Add('')
}

function Export-PreChangeSnapshot {
    param(
        [Parameter(Mandatory = $true)][string[]]$Targets,
        [AllowNull()][string]$SourceServer
    )

    $path = Get-PreChangeSnapshotPath -ScriptBaseName $script:ScriptBaseName
    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add('Exchange configuration pre-change snapshot')
    [void]$lines.Add(("Script        : {0}" -f $script:ScriptBaseName))
    [void]$lines.Add(("Snapshot Time : {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')))
    [void]$lines.Add('Purpose       : Current configuration captured immediately before Apply')
    [void]$lines.Add('')
    [void]$lines.Add('PowerShell virtual directory: Review Only - not changed automatically by this script.')
    [void]$lines.Add('')

    if (-not [string]::IsNullOrWhiteSpace($SourceServer)) {
        $freshSource = Get-ExchangeClientAccessSnapshot -Server $SourceServer
        Add-ClientAccessSnapshotLines -Lines $lines -Snapshot $freshSource -Role 'SOURCE'
    }

    foreach ($target in $Targets) {
        $freshTarget = Get-ExchangeClientAccessSnapshot -Server $target
        Add-ClientAccessSnapshotLines -Lines $lines -Snapshot $freshTarget -Role 'TARGET'
    }

    Set-Content -LiteralPath $path -Value @($lines) -Encoding UTF8 -ErrorAction Stop
    return $path
}

# ---------------------------------------------------------------------------
# Build Desired Configuration
# ---------------------------------------------------------------------------
function New-DesiredConfiguration {
    param(
        [Parameter(Mandatory = $true)]$TargetSnapshot,
        $SourceSnapshot,
        [string]$InternalNamespace,
        [string]$ExternalNamespace,
        [string]$AutodiscoverNamespace,
        [switch]$CopyAuthentication
    )

    $components = New-Object System.Collections.ArrayList

    foreach ($entry in $VirtualDirectoryMap) {
        $values = [ordered]@{}
        $sourceObject = if ($SourceSnapshot) { Get-SnapshotComponent -Snapshot $SourceSnapshot -Key $entry.Key } else { $null }

        if ($entry.Path) {
            if ($SourceSnapshot) {
                $values['InternalUrl'] = $sourceObject.InternalUrl
                $values['ExternalUrl'] = $sourceObject.ExternalUrl
            }
            else {
                $values['InternalUrl'] = "https://$InternalNamespace$($entry.Path)"
                $values['ExternalUrl'] = "https://$ExternalNamespace$($entry.Path)"
            }
        }

        if ($SourceSnapshot -and $CopyAuthentication) {
            foreach ($propertyName in $entry.AuthProperties) { $values[$propertyName] = $sourceObject.$propertyName }
        }

        if ($values.Count -gt 0) {
            $reviewNote = if ($entry.Policy -eq 'ReviewOnly') {
                'Review Only. This virtual directory is not changed automatically.'
            }
            else { $null }

            [void]$components.Add([PSCustomObject]@{
                Key = $entry.Key; Name = $entry.Name; Values = $values; Policy = $entry.Policy; ReviewNote = $reviewNote
            })
        }

        # These settings are migration-relevant but environment dependent. Show
        # source/target differences without cloning them automatically.
        if ($SourceSnapshot -and $entry.ReviewProperties.Count -gt 0) {
            $reviewValues = [ordered]@{}
            foreach ($propertyName in $entry.ReviewProperties) { $reviewValues[$propertyName] = $sourceObject.$propertyName }
            [void]$components.Add([PSCustomObject]@{
                Key = $entry.Key
                Name = "$($entry.Name) (Review Only)"
                Values = $reviewValues
                Policy = 'ReviewOnly'
                ReviewNote = 'Migration-sensitive setting. Review the target design before changing it.'
            })
        }
    }

    $scpValue = if ($SourceSnapshot) {
        $SourceSnapshot.ClientAccess.AutoDiscoverServiceInternalUri
    }
    else {
        "https://$AutodiscoverNamespace/Autodiscover/Autodiscover.xml"
    }
    [void]$components.Add([PSCustomObject]@{
        Key = 'ClientAccess'; Name = 'Autodiscover SCP'; Values = [ordered]@{ AutoDiscoverServiceInternalUri = $scpValue }; Policy = 'Apply'; ReviewNote = $null
    })

    if ($SourceSnapshot) {
        [void]$components.Add([PSCustomObject]@{
            Key = 'ASA'
            Name = 'Alternate Service Account (Kerberos)'
            Values = [ordered]@{ Account = $SourceSnapshot.ASA.Account; Status = $SourceSnapshot.ASA.Status }
            Policy = 'ReviewOnly'
            ReviewNote = 'ASA/Kerberos credentials are not deployed by this script. Review and roll the ASA credential separately when required.'
        })
    }

    $oaValues = [ordered]@{}
    if ($SourceSnapshot) {
        foreach ($propertyName in @('InternalHostname','ExternalHostname','InternalClientsRequireSsl','ExternalClientsRequireSsl')) {
            $oaValues[$propertyName] = $SourceSnapshot.OutlookAny.$propertyName
        }
        if ($CopyAuthentication) {
            foreach ($propertyName in @('InternalClientAuthenticationMethod','ExternalClientAuthenticationMethod','IISAuthenticationMethods')) {
                $oaValues[$propertyName] = $SourceSnapshot.OutlookAny.$propertyName
            }
        }
    }
    else {
        $oaValues['InternalHostname'] = $InternalNamespace
        $oaValues['ExternalHostname'] = $ExternalNamespace
    }
    [void]$components.Add([PSCustomObject]@{ Key = 'OutlookAny'; Name = 'Outlook Anywhere'; Values = $oaValues; Policy = 'Apply'; ReviewNote = $null })

    if ($SourceSnapshot) {
        [void]$components.Add([PSCustomObject]@{
            Key = 'OutlookAny'
            Name = 'Outlook Anywhere (Review Only)'
            Values = [ordered]@{ SSLOffloading = $SourceSnapshot.OutlookAny.SSLOffloading }
            Policy = 'ReviewOnly'
            ReviewNote = 'SSL offloading depends on the load balancer/TLS topology and is not changed automatically.'
        })
    }

    return [PSCustomObject]@{ Target = $TargetSnapshot.Server; Components = @($components) }
}

# ---------------------------------------------------------------------------
# Build Change Plan
# ---------------------------------------------------------------------------
# Only properties that differ between Current and Desired are added to the plan.
function New-ChangePlan {
    param(
        [Parameter(Mandatory = $true)]$TargetSnapshot,
        [Parameter(Mandatory = $true)]$DesiredConfiguration,
        [switch]$ApplyOnly
    )

    $plan = New-Object System.Collections.ArrayList
    foreach ($component in $DesiredConfiguration.Components) {
        $policy = if ($component.PSObject.Properties['Policy']) { [string]$component.Policy } else { 'Apply' }
        if ($ApplyOnly -and $policy -ne 'Apply') { continue }

        $currentObject = Get-SnapshotComponent -Snapshot $TargetSnapshot -Key $component.Key
        foreach ($propertyName in $component.Values.Keys) {
            $currentValue = $currentObject.$propertyName
            $newValue = $component.Values[$propertyName]
            if ((ConvertTo-ComparableValue $currentValue) -ne (ConvertTo-ComparableValue $newValue)) {
                [void]$plan.Add([PSCustomObject]@{
                    Target         = $TargetSnapshot.Server
                    ComponentKey   = $component.Key
                    Component      = $component.Name
                    Setting        = $propertyName
                    Current        = ConvertTo-DisplayValue $currentValue
                    New            = ConvertTo-DisplayValue $newValue
                    CurrentValue   = $currentValue
                    NewValue       = $newValue
                    Policy         = $policy
                    ReviewNote     = $component.ReviewNote
                })
            }
        }
    }
    return @($plan)
}

# ---------------------------------------------------------------------------
# Compare Reference and Target Values
# ---------------------------------------------------------------------------
# CompareOnly deliberately builds a complete comparison list, including
# matching values. This is separate from New-ChangePlan, which contains only
# differences needed for export/apply operations.
function New-ComparisonPlan {
    param(
        [Parameter(Mandatory = $true)]$TargetSnapshot,
        [Parameter(Mandatory = $true)]$DesiredConfiguration
    )

    $items = New-Object System.Collections.ArrayList
    foreach ($component in $DesiredConfiguration.Components) {
        $policy = if ($component.PSObject.Properties['Policy']) { [string]$component.Policy } else { 'Apply' }
        $targetObject = Get-SnapshotComponent -Snapshot $TargetSnapshot -Key $component.Key
        foreach ($propertyName in $component.Values.Keys) {
            $sourceValue = $component.Values[$propertyName]
            $targetValue = $targetObject.$propertyName
            $status = if ((ConvertTo-ComparableValue $sourceValue) -eq (ConvertTo-ComparableValue $targetValue)) { 'Same' } else { 'Different' }

            [void]$items.Add([PSCustomObject]@{
                Target       = $TargetSnapshot.Server
                ComponentKey = $component.Key
                Component    = $component.Name
                Setting      = $propertyName
                Source       = ConvertTo-DisplayValue $sourceValue
                TargetValue  = ConvertTo-DisplayValue $targetValue
                Policy       = $policy
                ReviewNote   = $component.ReviewNote
                Status       = $status
            })
        }
    }
    return @($items)
}

function Show-ComparisonPlan {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Comparisons)

    Write-Host ''
    Write-Host 'Comparison' -ForegroundColor Cyan
    Write-Host '----------' -ForegroundColor Cyan

    $currentTarget = $null
    foreach ($item in ($Comparisons | Sort-Object Target,Component,Setting)) {
        if ($currentTarget -ne $item.Target) {
            $currentTarget = $item.Target
            Write-Host ''
            Write-Host "[$currentTarget]" -ForegroundColor Yellow
        }

        Write-Host ("{0} - {1}" -f $item.Component, $item.Setting)
        if ($item.Policy -eq 'ReviewOnly') { Write-Host '  Policy : Review Only' -ForegroundColor Yellow }
        if ($item.ReviewNote) { Write-Host ("  Review : {0}" -f $item.ReviewNote) -ForegroundColor Yellow }
        Write-Host ("  Source : {0}" -f $item.Source)
        if ($item.Status -eq 'Different') {
            Write-Host ("  Target : {0}" -f $item.TargetValue) -ForegroundColor Red
            Write-Host '  Status : Different' -ForegroundColor Red
        }
        else {
            Write-Host ("  Target : {0}" -f $item.TargetValue)
            Write-Host '  Status : Same' -ForegroundColor Green
        }
    }

    Write-Host ''
    Write-Host 'Summary' -ForegroundColor Cyan
    Write-Host '-------' -ForegroundColor Cyan
    foreach ($target in @($Comparisons.Target | Select-Object -Unique)) {
        $targetItems = @($Comparisons | Where-Object { $_.Target -eq $target })
        $differentCount = @($targetItems | Where-Object { $_.Status -eq 'Different' }).Count
        Write-Host ("{0}: {1} compared, {2} different" -f $target, $targetItems.Count, $differentCount) -ForegroundColor $(if ($differentCount -gt 0) { 'Red' } else { 'Green' })
    }
}


# CompareOnly + OutputFile is a reporting combination, not a command-export
# combination. Keeping this function separate from Export-CommandPlan makes it
# impossible for a read-only comparison to accidentally emit executable Set-* commands.
function Export-ComparisonReport {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Comparisons,
        [Parameter(Mandatory = $true)][string[]]$Targets,
        [Parameter(Mandatory = $true)][string]$SourceServer,
        [switch]$IncludeAuthentication
    )

    $lines = New-Object System.Collections.ArrayList
    foreach ($line in @(
        '# ConfigureExchangeURLs-v2.1.ps1 comparison report'
        '# Tool Author : Ceyhun Kirmizitas'
        "# Generated   : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        "# Operator    : $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
        "# Source      : $SourceServer"
        "# Targets     : $($Targets -join ', ')"
        "# Authentication: $(if ($IncludeAuthentication) { 'Included from source' } else { 'Not compared' })"
        '# Mode        : Compare only (read-only)'
        '# No Set-* commands were generated and no Exchange changes were applied.'
        ''
    )) { [void]$lines.Add($line) }

    foreach ($target in $Targets) {
        $targetItems = @($Comparisons | Where-Object { $_.Target -eq $target } | Sort-Object Component,Setting)
        [void]$lines.Add('=====================================================================')
        [void]$lines.Add("Target: $target")
        [void]$lines.Add('=====================================================================')
        [void]$lines.Add('')

        foreach ($item in $targetItems) {
            [void]$lines.Add("$($item.Component) - $($item.Setting)")
            [void]$lines.Add("  Policy : $($item.Policy)")
            if ($item.ReviewNote) { [void]$lines.Add("  Review : $($item.ReviewNote)") }
            [void]$lines.Add("  Source : $($item.Source)")
            [void]$lines.Add("  Target : $($item.TargetValue)")
            [void]$lines.Add("  Status : $($item.Status)")
            [void]$lines.Add('')
        }

        $differentCount = @($targetItems | Where-Object { $_.Status -eq 'Different' }).Count
        [void]$lines.Add("Summary: Compared=$($targetItems.Count) Different=$differentCount")
        [void]$lines.Add('')
    }

    Set-Content -Path $Path -Value $lines -Encoding UTF8
}

function Show-ChangePlan {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Changes)

    if ($Changes.Count -eq 0) {
        Write-Host 'No changes are required.' -ForegroundColor Green
        return
    }

    Write-Host ''
    Write-Host 'Preview' -ForegroundColor Cyan
    Write-Host '-------' -ForegroundColor Cyan
    $currentTarget = $null
    foreach ($change in ($Changes | Sort-Object Target,Component,Setting)) {
        if ($currentTarget -ne $change.Target) {
            $currentTarget = $change.Target
            Write-Host ''
            Write-Host "[$currentTarget]" -ForegroundColor Yellow
        }
        Write-Host ("{0} - {1}" -f $change.Component, $change.Setting)
        if ($change.Policy -eq 'ReviewOnly') {
            Write-Host '  Policy  : Review Only - not applied automatically' -ForegroundColor Yellow
            if ($change.ReviewNote) { Write-Host ("  Review  : {0}" -f $change.ReviewNote) -ForegroundColor Yellow }
        }
        Write-Host ("  Current : {0}" -f $change.Current)
        Write-Host ("  New     : {0}" -f $change.New) -ForegroundColor $(if ($change.Policy -eq 'ReviewOnly') { 'Yellow' } else { 'Green' })
    }
}

function Get-PlanForComponent {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Plan,
        [Parameter(Mandatory = $true)][string]$ComponentKey
    )
    return @($Plan | Where-Object { $_.ComponentKey -eq $ComponentKey })
}

function Get-DesiredComponent {
    param(
        [Parameter(Mandatory = $true)]$DesiredConfiguration,
        [Parameter(Mandatory = $true)][string]$ComponentKey
    )
    return $DesiredConfiguration.Components | Where-Object { $_.Key -eq $ComponentKey } | Select-Object -First 1
}

# ---------------------------------------------------------------------------
# Export Commands
# ---------------------------------------------------------------------------
# The export path serializes values as PowerShell literals but never executes
# the generated command text. Engineers can review the TXT file before use.
function ConvertTo-PowerShellLiteral {
    param($Value)

    if ($null -eq $Value) { return '$null' }
    if ($Value -is [bool]) { return $(if ($Value) { '$true' } else { '$false' }) }
    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $items = @($Value | ForEach-Object { ConvertTo-PowerShellLiteral -Value $_ })
        return '@(' + ($items -join ', ') + ')'
    }

    $text = [string]$Value
    return "'" + $text.Replace("'", "''") + "'"
}

function Resolve-OutputFilePath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        $candidate = "ConfigureExchangeURLs-Commands-{0}.txt" -f (Get-Date -Format 'yyyyMMdd-HHmmss')
    }
    else {
        $candidate = $Path.Trim()
        $extension = [System.IO.Path]::GetExtension($candidate)
        if ([string]::IsNullOrWhiteSpace($extension)) {
            $candidate = "$candidate.txt"
        }
        elseif ($extension -ne '.txt') {
            throw 'OutputFile must use the .txt extension.'
        }
    }

    if (-not [System.IO.Path]::IsPathRooted($candidate)) {
        $candidate = Join-Path -Path (Get-Location).Path -ChildPath $candidate
    }

    $parent = Split-Path -Parent $candidate
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
        throw "OutputFile directory does not exist: $parent"
    }

    return [System.IO.Path]::GetFullPath($candidate)
}

function New-SetCommandLine {
    param(
        [Parameter(Mandatory = $true)][string]$CommandName,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Parameters
    )

    $parts = New-Object System.Collections.ArrayList
    [void]$parts.Add($CommandName)
    foreach ($name in $Parameters.Keys) {
        [void]$parts.Add("-$name")
        [void]$parts.Add((ConvertTo-PowerShellLiteral -Value $Parameters[$name]))
    }
    [void]$parts.Add('-ErrorAction')
    [void]$parts.Add('Stop')
    return ($parts -join ' ')
}

function Get-VirtualDirectoryCommandParameters {
    param(
        [Parameter(Mandatory = $true)]$MapEntry,
        [Parameter(Mandatory = $true)]$TargetSnapshot,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Plan
    )

    $changes = Get-PlanForComponent -Plan $Plan -ComponentKey $MapEntry.Key
    if ($changes.Count -eq 0) { return $null }

    $targetObject = Get-SnapshotComponent -Snapshot $TargetSnapshot -Key $MapEntry.Key
    $params = [ordered]@{ Identity = $targetObject.Identity }
    foreach ($change in $changes) { $params[$change.Setting] = $change.NewValue }

    # MAPI requires IISAuthenticationMethods when URL values are changed.
    # Preserve the target value unless authentication is explicitly part of the change plan.
    if ($MapEntry.Key -eq 'Mapi' -and ($changes.Setting -contains 'InternalUrl' -or $changes.Setting -contains 'ExternalUrl')) {
        if (-not $params.Contains('IISAuthenticationMethods')) {
            $params['IISAuthenticationMethods'] = $targetObject.IISAuthenticationMethods
        }
    }

    return $params
}

function Get-OutlookAnywhereCommandParameters {
    param(
        [Parameter(Mandatory = $true)]$TargetSnapshot,
        [Parameter(Mandatory = $true)]$DesiredConfiguration,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Plan
    )

    $changes = Get-PlanForComponent -Plan $Plan -ComponentKey 'OutlookAny'
    if ($changes.Count -eq 0) { return $null }

    $desired = Get-DesiredComponent -DesiredConfiguration $DesiredConfiguration -ComponentKey 'OutlookAny'
    $params = [ordered]@{ Identity = $TargetSnapshot.OutlookAny.Identity }
    foreach ($change in $changes) { $params[$change.Setting] = $change.NewValue }

    # Outlook Anywhere requires companion parameters when hostname values are changed.
    # Re-send the existing SSL/auth values unless the reference plan explicitly changes them.
    if ($changes.Setting -contains 'InternalHostname' -or $changes.Setting -contains 'ExternalHostname') {
        if (-not $params.Contains('InternalClientsRequireSsl')) {
            $params['InternalClientsRequireSsl'] = if ($desired.Values.Contains('InternalClientsRequireSsl')) { $desired.Values['InternalClientsRequireSsl'] } else { $TargetSnapshot.OutlookAny.InternalClientsRequireSsl }
        }
        if (-not $params.Contains('ExternalClientsRequireSsl')) {
            $params['ExternalClientsRequireSsl'] = if ($desired.Values.Contains('ExternalClientsRequireSsl')) { $desired.Values['ExternalClientsRequireSsl'] } else { $TargetSnapshot.OutlookAny.ExternalClientsRequireSsl }
        }
        if (-not $params.Contains('ExternalClientAuthenticationMethod')) {
            $params['ExternalClientAuthenticationMethod'] = if ($desired.Values.Contains('ExternalClientAuthenticationMethod')) { $desired.Values['ExternalClientAuthenticationMethod'] } else { $TargetSnapshot.OutlookAny.ExternalClientAuthenticationMethod }
        }
        if ($null -eq $params['ExternalClientAuthenticationMethod']) {
            throw "Cannot determine Outlook Anywhere ExternalClientAuthenticationMethod for $($TargetSnapshot.Server)."
        }
    }

    return $params
}

function New-TargetOperations {
    param(
        [Parameter(Mandatory = $true)]$TargetSnapshot,
        [Parameter(Mandatory = $true)]$DesiredConfiguration,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Plan
    )

    $operations = New-Object System.Collections.ArrayList
    # Review Only items remain visible in preview/reporting but must never reach a Set-* cmdlet.
    $applyPlan = @($Plan | Where-Object { $_.Policy -eq 'Apply' })
    $desiredOwa = Get-DesiredComponent -DesiredConfiguration $DesiredConfiguration -ComponentKey 'Owa'
    $desiredEcp = Get-DesiredComponent -DesiredConfiguration $DesiredConfiguration -ComponentKey 'Ecp'
    $adfsEnabled = $false
    if ($desiredOwa -and $desiredOwa.Values.Contains('AdfsAuthentication')) { $adfsEnabled = $adfsEnabled -or [bool]$desiredOwa.Values['AdfsAuthentication'] }
    if ($desiredEcp -and $desiredEcp.Values.Contains('AdfsAuthentication')) { $adfsEnabled = $adfsEnabled -or [bool]$desiredEcp.Values['AdfsAuthentication'] }

    # OWA/ECP are paired. Emit ECP first when ADFS is enabled so both sides stay aligned.
    $orderedKeys = if ($adfsEnabled) { @('Ecp','Owa') } else { @('Owa','Ecp') }
    $entries = @(
        foreach ($key in $orderedKeys) { $VirtualDirectoryMap | Where-Object { $_.Key -eq $key } | Select-Object -First 1 }
        $VirtualDirectoryMap | Where-Object { $_.Key -notin @('Owa','Ecp') }
    )

    foreach ($entry in $entries) {
        $changes = Get-PlanForComponent -Plan $applyPlan -ComponentKey $entry.Key
        if ($changes.Count -eq 0) { continue }
        $params = Get-VirtualDirectoryCommandParameters -MapEntry $entry -TargetSnapshot $TargetSnapshot -Plan $applyPlan
        [void]$operations.Add([PSCustomObject]@{ CommandName = $entry.SetCmd; Parameters = $params; Changes = $changes; ReviewNote = $null })
    }

    $scpChanges = Get-PlanForComponent -Plan $applyPlan -ComponentKey 'ClientAccess'
    if ($scpChanges.Count -gt 0) {
        [void]$operations.Add([PSCustomObject]@{
            CommandName = 'Set-ClientAccessService'
            Parameters  = [ordered]@{ Identity = $TargetSnapshot.Server; AutoDiscoverServiceInternalUri = $scpChanges[0].NewValue }
            Changes     = $scpChanges
            ReviewNote  = $null
        })
    }

    $oaChanges = Get-PlanForComponent -Plan $applyPlan -ComponentKey 'OutlookAny'
    if ($oaChanges.Count -gt 0) {
        [void]$operations.Add([PSCustomObject]@{
            CommandName = 'Set-OutlookAnywhere'
            Parameters  = Get-OutlookAnywhereCommandParameters -TargetSnapshot $TargetSnapshot -DesiredConfiguration $DesiredConfiguration -Plan $applyPlan
            Changes     = $oaChanges
            ReviewNote  = $null
        })
    }
    return @($operations)
}

function Export-CommandPlan {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Changes,
        [Parameter(Mandatory = $true)][hashtable]$TargetSnapshots,
        [Parameter(Mandatory = $true)][hashtable]$DesiredByTarget,
        [Parameter(Mandatory = $true)][string[]]$Targets,
        [string]$SourceServer,
        [switch]$IncludeAuthentication
    )

    $lines = New-Object System.Collections.ArrayList
    foreach ($line in @(
        '# ConfigureExchangeURLs-v2.1.ps1 generated command file'
        '# Tool Author : Ceyhun Kirmizitas'
        "# Generated   : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        "# Operator    : $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
        "# Source      : $(if ($SourceServer) { $SourceServer } else { '<Interactive namespace mode>' })"
        "# Targets     : $($Targets -join ', ')"
        "# Authentication: $(if ($IncludeAuthentication) { 'Included from source' } else { 'Not changed' })"
        '# WARNING: Review every command before running it in Exchange Management Shell.'
        '# ConfigureExchangeURLs-v2.1.ps1 did not apply any Exchange changes.'
        ''
    )) { [void]$lines.Add($line) }

    foreach ($targetServer in $Targets) {
        $targetPlan = @($Changes | Where-Object { $_.Target -eq $targetServer })
        if ($targetPlan.Count -eq 0) { continue }
        [void]$lines.Add('# =====================================================================')
        [void]$lines.Add("# Target: $targetServer")
        [void]$lines.Add('# =====================================================================')
        [void]$lines.Add('')

        $reviewItems = @($targetPlan | Where-Object { $_.Policy -eq 'ReviewOnly' })
        foreach ($reviewItem in $reviewItems) {
            [void]$lines.Add("# REVIEW ONLY: $($reviewItem.Component) - $($reviewItem.Setting)")
            [void]$lines.Add("# Current: $($reviewItem.Current)")
            [void]$lines.Add("# Source/Reference: $($reviewItem.New)")
            if ($reviewItem.ReviewNote) { [void]$lines.Add("# Note   : $($reviewItem.ReviewNote)") }
            [void]$lines.Add('# No command generated for this setting.')
            [void]$lines.Add('')
        }

        $operations = @(New-TargetOperations -TargetSnapshot $TargetSnapshots[$targetServer] -DesiredConfiguration $DesiredByTarget[$targetServer] -Plan $targetPlan)
        foreach ($operation in $operations) {
            if ($operation.ReviewNote) { [void]$lines.Add("# REVIEW: $($operation.ReviewNote)") }
            foreach ($change in $operation.Changes) {
                [void]$lines.Add("# $($change.Component) - $($change.Setting)")
                [void]$lines.Add("# Current: $($change.Current)")
                [void]$lines.Add("# New    : $($change.New)")
            }
            [void]$lines.Add((New-SetCommandLine -CommandName $operation.CommandName -Parameters $operation.Parameters))
            [void]$lines.Add('')
        }
    }
    Set-Content -LiteralPath $Path -Value @($lines) -Encoding UTF8
}

# ---------------------------------------------------------------------------
# Apply Configuration
# ---------------------------------------------------------------------------
# Apply uses the same parameter-building helpers as command export. This keeps
# the preview/export/apply paths aligned and avoids executing generated text.
function Invoke-SetCommand {
    param(
        [Parameter(Mandatory = $true)][string]$CommandName,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Parameters
    )

    & $CommandName @Parameters -ErrorAction Stop
}

function Invoke-ConfigurationPlan {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Changes,
        [Parameter(Mandatory = $true)][hashtable]$TargetSnapshots,
        [Parameter(Mandatory = $true)][hashtable]$DesiredByTarget,
        [Parameter(Mandatory = $true)][string[]]$Targets
    )

    foreach ($targetServer in $Targets) {
        $targetPlan = @($Changes | Where-Object { $_.Target -eq $targetServer })
        if ($targetPlan.Count -eq 0) { continue }
        Write-Host "Configuring $targetServer..." -ForegroundColor Cyan
        $reviewItems = @($targetPlan | Where-Object { $_.Policy -eq 'ReviewOnly' })
        foreach ($reviewItem in $reviewItems) {
            [void]$lines.Add("# REVIEW ONLY: $($reviewItem.Component) - $($reviewItem.Setting)")
            [void]$lines.Add("# Current: $($reviewItem.Current)")
            [void]$lines.Add("# Source/Reference: $($reviewItem.New)")
            if ($reviewItem.ReviewNote) { [void]$lines.Add("# Note   : $($reviewItem.ReviewNote)") }
            [void]$lines.Add('# No command generated for this setting.')
            [void]$lines.Add('')
        }

        $operations = @(New-TargetOperations -TargetSnapshot $TargetSnapshots[$targetServer] -DesiredConfiguration $DesiredByTarget[$targetServer] -Plan $targetPlan)
        foreach ($operation in $operations) {
            Invoke-SetCommand -CommandName $operation.CommandName -Parameters $operation.Parameters
        }
    }
}

# ---------------------------------------------------------------------------
# Verify Configuration
# ---------------------------------------------------------------------------
# Re-query each target after apply and compare it with the same Desired object.
# Verification reports remaining differences instead of assuming Set-* succeeded.
function Test-ConfigurationPlan {
    param(
        [Parameter(Mandatory = $true)][hashtable]$DesiredByTarget,
        [Parameter(Mandatory = $true)][string[]]$Targets
    )

    $results = New-Object System.Collections.ArrayList
    foreach ($targetServer in $Targets) {
        try {
            $snapshot = Get-ExchangeClientAccessSnapshot -Server $targetServer
            $remaining = @(New-ChangePlan -TargetSnapshot $snapshot -DesiredConfiguration $DesiredByTarget[$targetServer] -ApplyOnly)
            [void]$results.Add([PSCustomObject]@{
                Server           = $targetServer
                Status           = $(if ($remaining.Count -eq 0) { 'Verified' } else { 'Differences remain' })
                RemainingChanges = $remaining.Count
            })
        }
        catch {
            [void]$results.Add([PSCustomObject]@{
                Server           = $targetServer
                Status           = 'Verification failed'
                RemainingChanges = $null
            })
        }
    }
    return @($results)
}

function Get-NormalizedTargets {
    param([string[]]$Servers)

    $result = @($Servers | Where-Object { $_ -and $_.Trim() } | ForEach-Object { $_.Trim() } | Select-Object -Unique)
    if ($result.Count -eq 0) { throw 'At least one target server is required.' }
    return $result
}

# Main
Initialize-ExchangeShell


if ($PSCmdlet.ParameterSetName -eq 'Reference') {
    $targets = Get-NormalizedTargets -Servers $TargetServers
    if ($targets -contains $SourceServer) { throw 'SourceServer cannot also be listed in TargetServers.' }

    Write-Host "Source server : $SourceServer" -ForegroundColor Cyan
    Write-Host "Target servers: $($targets -join ', ')" -ForegroundColor Cyan
    Write-Host "Authentication: $(if ($IncludeAuthentication) { 'Included' } else { 'Not changed' })" -ForegroundColor Cyan
    Write-Host 'PowerShell vDir : Not changed (Review Only)' -ForegroundColor Yellow
    if ($CompareOnly) { Write-Host 'Mode          : Compare only (read-only)' -ForegroundColor Cyan }
    $sourceSnapshot = Get-ExchangeClientAccessSnapshot -Server $SourceServer
}
else {
    $targets = if ($TargetServers) { Get-NormalizedTargets -Servers $TargetServers } else { @($env:COMPUTERNAME) }

    if ([string]::IsNullOrWhiteSpace($InternalNamespace)) {
        $InternalNamespace = Normalize-Namespace -Value (Read-Host 'Internal namespace (example: mail.contoso.com)')
    }
    else { $InternalNamespace = Normalize-Namespace -Value $InternalNamespace }

    if ([string]::IsNullOrWhiteSpace($ExternalNamespace)) {
        $externalInput = Read-Host "External namespace (Enter = $InternalNamespace)"
        $ExternalNamespace = if ([string]::IsNullOrWhiteSpace($externalInput)) { $InternalNamespace } else { Normalize-Namespace -Value $externalInput }
    }
    else { $ExternalNamespace = Normalize-Namespace -Value $ExternalNamespace }

    $suggestedAutodiscover = Get-SuggestedAutodiscoverNamespace -ClientAccessNamespace $InternalNamespace
    if ([string]::IsNullOrWhiteSpace($AutodiscoverSCPNamespace)) {
        $autodiscoverInput = Read-Host "Autodiscover SCP namespace (Enter = $suggestedAutodiscover)"
        $AutodiscoverSCPNamespace = if ([string]::IsNullOrWhiteSpace($autodiscoverInput)) { $suggestedAutodiscover } else { Normalize-Namespace -Value $autodiscoverInput }
    }
    else { $AutodiscoverSCPNamespace = Normalize-Namespace -Value $AutodiscoverSCPNamespace }

    Write-Host "Internal namespace        : $InternalNamespace" -ForegroundColor Cyan
    Write-Host "External namespace        : $ExternalNamespace" -ForegroundColor Cyan
    Write-Host "Autodiscover SCP namespace: $AutodiscoverSCPNamespace" -ForegroundColor Cyan
    Write-Host "Target servers            : $($targets -join ', ')" -ForegroundColor Cyan
    Write-Host 'Authentication            : Not changed' -ForegroundColor Cyan
    Write-Host 'PowerShell vDir            : Not changed (Review Only)' -ForegroundColor Yellow
}

$targetSnapshots = @{}
$desiredByTarget = @{}
$allChanges = New-Object System.Collections.ArrayList
$allComparisons = New-Object System.Collections.ArrayList
foreach ($targetServer in $targets) {
    $snapshot = Get-ExchangeClientAccessSnapshot -Server $targetServer
    $targetSnapshots[$targetServer] = $snapshot
    $desired = if ($PSCmdlet.ParameterSetName -eq 'Reference') {
        New-DesiredConfiguration -TargetSnapshot $snapshot -SourceSnapshot $sourceSnapshot -CopyAuthentication:$IncludeAuthentication
    }
    else {
        New-DesiredConfiguration -TargetSnapshot $snapshot -InternalNamespace $InternalNamespace -ExternalNamespace $ExternalNamespace -AutodiscoverNamespace $AutodiscoverSCPNamespace
    }
    $desiredByTarget[$targetServer] = $desired
    if ($CompareOnly) {
        foreach ($comparison in @(New-ComparisonPlan -TargetSnapshot $snapshot -DesiredConfiguration $desired)) { [void]$allComparisons.Add($comparison) }
    }
    else {
        foreach ($change in @(New-ChangePlan -TargetSnapshot $snapshot -DesiredConfiguration $desired)) { [void]$allChanges.Add($change) }
    }
}

if ($CompareOnly) {
    Show-ComparisonPlan -Comparisons @($allComparisons)

    # OutputFile only adds a persistent copy of the read-only comparison. It
    # does not change CompareOnly into command-export or apply mode.
    if (-not [string]::IsNullOrWhiteSpace($OutputFile)) {
        $outputPath = Resolve-OutputFilePath -Path $OutputFile
        Export-ComparisonReport -Path $outputPath -Comparisons @($allComparisons) -Targets $targets -SourceServer $SourceServer -IncludeAuthentication:$IncludeAuthentication

        Write-Host ''
        Write-Host "Comparison report created: $outputPath" -ForegroundColor Green
        Write-Host 'No Set-* commands were generated and no Exchange configuration changes were applied.' -ForegroundColor Yellow
    }
    return
}

Show-ChangePlan -Changes @($allChanges)
if ($allChanges.Count -eq 0) {
    Write-Host 'No changes are required.' -ForegroundColor Green
    return
}

$actionableChanges = @($allChanges | Where-Object { $_.Policy -eq 'Apply' })

# -OutputFile is an explicit export-only mode. When it is present, no Set-*
# cmdlet is executed by this script.
if (-not [string]::IsNullOrWhiteSpace($OutputFile)) {
    $outputPath = Resolve-OutputFilePath -Path $OutputFile
    Export-CommandPlan -Path $outputPath -Changes @($allChanges) -TargetSnapshots $targetSnapshots -DesiredByTarget $desiredByTarget -Targets $targets -SourceServer $SourceServer -IncludeAuthentication:$IncludeAuthentication

    Write-Host ''
    Write-Host "Command file created: $outputPath" -ForegroundColor Green
    Write-Host 'No Exchange configuration changes were applied. Review the TXT file and run the required commands manually.' -ForegroundColor Yellow
    return
}

if ($actionableChanges.Count -eq 0) {
    Write-Host 'No automatic changes are required. Review Only differences were shown above.' -ForegroundColor Yellow
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
Write-Host 'PowerShell vDir was not changed. It remains Review Only.' -ForegroundColor Yellow

Invoke-ConfigurationPlan -Changes @($allChanges) -TargetSnapshots $targetSnapshots -DesiredByTarget $desiredByTarget -Targets $targets

$verification = @(Test-ConfigurationPlan -DesiredByTarget $desiredByTarget -Targets $targets)
Write-Host ''
Write-Host 'Verification summary' -ForegroundColor Cyan
Write-Host '--------------------' -ForegroundColor Cyan
$verification | Format-Table Server,Status,RemainingChanges -AutoSize

if (@($verification | Where-Object { $_.Status -ne 'Verified' }).Count -gt 0) {
    throw 'One or more targets could not be fully verified after configuration. Review the summary above.'
}
else {
    Write-Host 'Configuration completed and verified.' -ForegroundColor Green
}
