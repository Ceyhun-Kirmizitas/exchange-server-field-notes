<#
.SYNOPSIS
Reviews, configures, clones, backs up, or restores supported Exchange Server Client Access URL/namespace configuration.

.DESCRIPTION
ExchangeURLManager.ps1 provides five Manager operations plus an explicit guided Interactive Configure mode.

1. Review mode
   Use -Review with -Server to display the current supported Client Access configuration.
   Review is read-only. A server-specific discovery failure does not stop the remaining servers where practical.

2. Configure mode
   Run without -SourceServer or another mode switch. Internal, external, and Autodiscover SCP namespaces can be supplied as parameters or entered interactively.
   If -Server is omitted, the local computer is used. Authentication settings are not changed unless explicitly requested.
   Use -ClearExternalUrls to remove supported external URLs/hostnames.
   Use -IncludePowerShellUrls to include PowerShell InternalUrl/ExternalUrl only.
   Outlook Anywhere (RPC over HTTP) SSL requirements and DefaultAuthenticationMethod can be changed only when their explicit parameters are supplied.

3. Clone mode
   Use -SourceServer with -TargetServer. The default behavior is a read-only source/target comparison; no Exchange configuration changes are made.
   Add -CompareOnly when you want to state the read-only intent explicitly.
   Add -ApplyChanges explicitly to enter Clone / Apply mode.
   -CompareOnly and -ApplyChanges cannot be used together.
   Add -IncludeAuthentication to include supported authentication settings from the source server, including OWA LogonFormat and DefaultDomain.
   Add -IncludeOutlookAnywhereSslRequirements to copy Outlook Anywhere (RPC over HTTP) InternalClientsRequireSsl and ExternalClientsRequireSsl explicitly during Apply.
   Add -IncludePowerShellUrls to include only PowerShell InternalUrl/ExternalUrl during Apply.
   At an interactive console, long comparison output is paged according to the current console height without splitting a setting/result block. Press ENTER to continue or Q to exit.

4. Backup mode
   Use -Backup with -Server to create one versioned JSON backup per server.
   Backups are read-only and contain the supported Apply-capable values plus Review Only visibility data.
   Secrets, passwords, SecureStrings, certificate private keys, and ASA credential material are not exported.
   The default location is ConfigBackups under the script folder. -BackupPath can specify a directory or, for a single server, a .json file.

5. Restore mode
   Use -Restore with -BackupFile to build a same-server restore plan from an ExchangeURLManager JSON backup.
   Cross-server restore is intentionally blocked; use Clone for server-to-server migration.
   URL/hostname settings are restored by default.
   Authentication, PowerShell URLs, and Outlook Anywhere SSL requirements require their explicit opt-in switches.
   Review Only values from the backup remain visible but are never applied automatically.

Interactive Configure
   Use -Interactive for guided field configuration.
   Interactive mode collects server/component/value selections only; it uses the same Compare -> Preview -> Confirm -> JSON backup -> Apply -> Verify engine as normal Configure.
   Common namespace, component-selection, and advanced per-component styles are available.
   Advanced mode uses 1=Internal only, 2=External only, 3=Both, 4=Skip for every dual-sided URL/hostname component.

Outlook Anywhere settings apply only to the legacy RPC over HTTP /rpc endpoint.
MAPI over HTTP is managed separately through the MAPI virtual directory.

Configure, Clone / Apply, and Restore use:
Discover -> Compare -> Preview -> Confirm -> JSON pre-change backup -> Apply -> Verify.

Clone is read-only by default. -CompareOnly can be used as an explicit read-only intent switch.
The Apply path is available only when -ApplyChanges is explicitly supplied.
-CompareOnly and -ApplyChanges are mutually exclusive.

A real Apply is blocked if the fresh pre-change JSON backup cannot be created.
The manual Backup mode and automatic pre-change backup use the same versioned JSON schema.

When -OutputFile is specified in Configure, Clone / Apply, or Restore, no Exchange changes are applied and only required Set-* commands are exported.
With the default read-only Clone comparison, -OutputFile writes the comparison report.
With Review, -OutputFile writes the read-only review report.

.PARAMETER Interactive
Runs guided Interactive Configure mode. The script collects server/component/value choices and then uses the normal Configure planning/apply engine.

.PARAMETER Review
Runs read-only Review mode. Requires Server.

.PARAMETER Backup
Runs read-only Backup mode. Requires Server.

.PARAMETER Restore
Runs Restore mode. Requires BackupFile.

.PARAMETER SourceServer
Reference Exchange server. Valid only in Clone mode.

.PARAMETER Server
One or more Exchange servers for Review, Backup, Configure, or Interactive Configure.
Required in Review and Backup modes. Optional in Configure and Interactive modes; if omitted, the local computer is used.

.PARAMETER TargetServer
One or more target Exchange servers. Required only in Clone mode.

.PARAMETER BackupFile
ExchangeURLManager JSON backup used by Restore mode.
The target server is read from the backup and must match the live Exchange server identity.

.PARAMETER BackupPath
Optional Backup-mode output path.
For multiple servers, specify a directory.
For one server, specify either a directory or a .json filename.
Existing files are never overwritten.

.PARAMETER IncludeAuthentication
Includes supported authentication settings in Clone or Restore.
In Review mode, authentication settings are hidden by default and are displayed only when this switch is used.
PowerShell authentication remains Review Only.

.PARAMETER CompareOnly
Explicitly requests the read-only Clone comparison.
This switch is optional because SourceServer + TargetServer already defaults to read-only comparison.
It cannot be used together with -ApplyChanges.

.PARAMETER ApplyChanges
Explicitly enables the Clone Apply path.
Without this switch, SourceServer + TargetServer always runs a read-only comparison and no Exchange configuration changes are made.
It cannot be used together with -CompareOnly.
When -OutputFile is used together with -ApplyChanges, required Set-* commands are exported and no Exchange configuration changes are applied.

.PARAMETER InternalNamespace
Internal Client Access namespace for Configure mode. If omitted, the script prompts for it.

.PARAMETER ExternalNamespace
External Client Access namespace for Configure mode. If omitted, the script prompts and defaults to InternalNamespace.

.PARAMETER ClearExternalUrls
Explicitly clears supported ExternalUrl values and Outlook Anywhere (RPC over HTTP) ExternalHostname in Configure mode.
Cannot be used together with -ExternalNamespace.

.PARAMETER IncludePowerShellUrls
Explicitly includes PowerShell InternalUrl/ExternalUrl in Configure, Clone, or Restore.
PowerShell authentication, RequireSSL, and Extended Protection remain Review Only.

.PARAMETER IncludeOutlookAnywhereSslRequirements
Explicitly includes Outlook Anywhere (RPC over HTTP) InternalClientsRequireSsl and ExternalClientsRequireSsl in Clone or Restore.
If omitted, the target server's current SSL requirement values are preserved.

.PARAMETER OutlookAnywhereInternalClientsRequireSsl
Optional Configure-mode Boolean value for Outlook Anywhere (RPC over HTTP) InternalClientsRequireSsl.
If omitted, the current server value is preserved.

.PARAMETER OutlookAnywhereExternalClientsRequireSsl
Optional Configure-mode Boolean value for Outlook Anywhere (RPC over HTTP) ExternalClientsRequireSsl.
If omitted, the current server value is preserved.

.PARAMETER OutlookAnywhereDefaultAuthenticationMethod
Optional Configure-mode Outlook Anywhere (RPC over HTTP) DefaultAuthenticationMethod.
Valid values: Basic, Ntlm, Negotiate.
If omitted, Outlook Anywhere (RPC over HTTP) authentication is not changed.

.PARAMETER AutodiscoverSCPNamespace
Autodiscover SCP namespace for Configure mode. If omitted, the script prompts with a suggested value.

.PARAMETER OutputFile
Optional TXT output path.
Review writes a read-only report.
Configure/Interactive/Clone/Restore export required Set-* commands without applying changes.
Default Clone comparison writes the read-only comparison report.
Clone / Apply with -OutputFile exports required Set-* commands without applying them.

.PARAMETER Help
Displays a short usage guide and exits without initializing Exchange Management Shell or making changes.

.EXAMPLE
.\ExchangeURLManager.ps1 -Review -Server EX01
Reviews the current supported Client Access configuration on EX01. No changes are applied.

.EXAMPLE
.\ExchangeURLManager.ps1 -Review -Server EX01,EX02 -OutputFile C:\Temp\ExchangeURL-Review.txt
Reviews EX01 and EX02 and writes the same read-only report to a TXT file.

.EXAMPLE
.\ExchangeURLManager.ps1 -Interactive
Starts guided Interactive Configure. Choose Common namespace, Select components, or Advanced per-component configuration.

.EXAMPLE
.\ExchangeURLManager.ps1 -Interactive -Server EX01,EX02
Starts guided Interactive Configure for EX01 and EX02. Advanced mode shows current values for both servers before each component choice.

.EXAMPLE
.\ExchangeURLManager.ps1
Prompts for namespaces, previews the required changes for the local Exchange server, asks for confirmation, creates a JSON pre-change backup, applies the changes, and verifies the result.

.EXAMPLE
.\ExchangeURLManager.ps1 -Server EX01,EX02 -InternalNamespace mail.contoso.com -ExternalNamespace mail.contoso.com -AutodiscoverSCPNamespace autodiscover.contoso.com
Applies URL/hostname changes to EX01 and EX02 without changing authentication.

.EXAMPLE
.\ExchangeURLManager.ps1 -SourceServer EX01 -TargetServer EX02,EX03 -IncludeAuthentication
Displays all URL/hostname and supported authentication values for EX01 versus EX02 and EX03. No changes are applied.

.EXAMPLE
.\ExchangeURLManager.ps1 -SourceServer EX01 -TargetServer EX02,EX03 -CompareOnly -IncludeAuthentication
Explicitly requests the same read-only source/target comparison. No changes are applied.

.EXAMPLE
.\ExchangeURLManager.ps1 -SourceServer EX01 -TargetServer EX02,EX03 -ApplyChanges -IncludeAuthentication
Aligns URL/hostname and supported authentication settings on EX02 and EX03 with EX01 after Preview, confirmation, and successful pre-change backup creation.

.EXAMPLE
.\ExchangeURLManager.ps1 -Backup -Server EX01,EX02
Creates one versioned JSON backup per server under ConfigBackups.

.EXAMPLE
.\ExchangeURLManager.ps1 -Backup -Server EX01 -BackupPath C:\Temp\EX01-ExchangeURL.json
Creates a single-server JSON backup at the specified path.

.EXAMPLE
.\ExchangeURLManager.ps1 -Restore -BackupFile C:\Temp\EX01-ExchangeURL.json
Restores supported URL/hostname settings to the same server identified by the backup after preview and confirmation.

.EXAMPLE
.\ExchangeURLManager.ps1 -Restore -BackupFile C:\Temp\EX01-ExchangeURL.json -IncludeAuthentication -IncludePowerShellUrls -IncludeOutlookAnywhereSslRequirements
Restores supported URL/hostname settings and explicitly includes supported authentication, PowerShell URLs, and Outlook Anywhere SSL requirements.

.EXAMPLE
.\ExchangeURLManager.ps1 -Server EX01 -InternalNamespace mail.contoso.com -ClearExternalUrls
Configures the internal namespace and clears supported external URLs/hostnames.

.EXAMPLE
.\ExchangeURLManager.ps1 -Server EX01 -InternalNamespace mail.contoso.com -ExternalNamespace mail.contoso.com -OutlookAnywhereInternalClientsRequireSsl $true -OutlookAnywhereExternalClientsRequireSsl $true -OutlookAnywhereDefaultAuthenticationMethod Ntlm
Configures namespaces and explicitly aligns Outlook Anywhere (RPC over HTTP) SSL requirements and default authentication.

.NOTES
Author        : Ceyhun Kirmizitas
Version       : 1.0
Date          : 28/09/2026
Applies to    : Exchange Server Mailbox servers
Compatibility : Intended for Exchange Server 2016, Exchange Server 2019, and Exchange Server Subscription Edition.
Validation    : No hard Exchange-version gate is enforced.
Shell         : Windows PowerShell 5.1
Website    : https://ceyhunkirmizitas.net
GitHub     : https://github.com/Ceyhun-Kirmizitas
LinkedIn   : https://www.linkedin.com/in/ceyhun-kirmizitas/

Change Log
----------
1.0 - Initial ExchangeURLManager release.
      Provides Review, Configure, Interactive Configure, Clone, Backup, and Restore modes.
      Uses versioned JSON backups for manual Backup and automatic pre-change recovery.
      Keeps authentication, PowerShell URL, and Outlook Anywhere SSL migration behind explicit opt-in controls where required.
      Keeps ASA/Kerberos, EWS MRS Proxy, Outlook Anywhere SSL offloading, PowerShell authentication/RequireSSL, and other migration-sensitive settings Review Only.
      ASA status discovery is isolated from normal Client Access discovery so an optional ASA read failure does not stop URL/namespace operations.
      Outlook Anywhere is treated separately from MAPI over HTTP; its SSL/authentication settings apply only to the legacy /rpc endpoint.

Migration safety model:
- Apply: settings that the tool can safely align after preview/confirmation.
- Review Only: settings that are important during migration but are not changed automatically.
- Skip: settings outside this script's scope.

The script does not manage certificates, private keys, IIS bindings, firewall/NAT, load balancers, DNS records, ASA credential deployment, SPNs, or Extended Protection changes.

License
-------
MIT License
Copyright (c) 2026 Ceyhun Kirmizitas

Caution
-------
Use this script at your own risk.
Review and test it in your environment before production use.
The author is not responsible for any issues, outages, or data loss resulting from its use.

.LINK
https://ceyhunkirmizitas.net

.LINK
https://github.com/Ceyhun-Kirmizitas

.LINK
https://www.linkedin.com/in/ceyhun-kirmizitas/
#>

[CmdletBinding(DefaultParameterSetName = 'Configure')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Interactive')]
    [switch]$Interactive,

    [Parameter(Mandatory = $true, ParameterSetName = 'Review')]
    [switch]$Review,

    [Parameter(Mandatory = $true, ParameterSetName = 'Backup')]
    [switch]$Backup,

    [Parameter(Mandatory = $true, ParameterSetName = 'Restore')]
    [switch]$Restore,

    [Parameter(Mandatory = $true, ParameterSetName = 'Clone')]
    [ValidateNotNullOrEmpty()]
    [string]$SourceServer,

    [Parameter(Mandatory = $true, ParameterSetName = 'Review')]
    [Parameter(Mandatory = $true, ParameterSetName = 'Backup')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Configure')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')]
    [ValidateNotNullOrEmpty()]
    [string[]]$Server,

    [Parameter(Mandatory = $true, ParameterSetName = 'Clone')]
    [ValidateNotNullOrEmpty()]
    [string[]]$TargetServer,

    [Parameter(Mandatory = $true, ParameterSetName = 'Restore')]
    [ValidateNotNullOrEmpty()]
    [string]$BackupFile,

    [Parameter(Mandatory = $false, ParameterSetName = 'Backup')]
    [string]$BackupPath,

    [Parameter(Mandatory = $false, ParameterSetName = 'Clone')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Restore')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Review')]
    [switch]$IncludeAuthentication,

    [Parameter(Mandatory = $false, ParameterSetName = 'Clone')]
    [switch]$CompareOnly,

    [Parameter(Mandatory = $false, ParameterSetName = 'Clone')]
    [switch]$ApplyChanges,

    [Parameter(Mandatory = $false, ParameterSetName = 'Configure')]
    [string]$InternalNamespace,

    [Parameter(Mandatory = $false, ParameterSetName = 'Configure')]
    [string]$ExternalNamespace,

    [Parameter(Mandatory = $false, ParameterSetName = 'Configure')]
    [switch]$ClearExternalUrls,

    [Parameter(Mandatory = $false, ParameterSetName = 'Clone')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Configure')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Restore')]
    [switch]$IncludePowerShellUrls,

    [Parameter(Mandatory = $false, ParameterSetName = 'Clone')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Restore')]
    [switch]$IncludeOutlookAnywhereSslRequirements,

    [Parameter(Mandatory = $false, ParameterSetName = 'Configure')]
    [Nullable[bool]]$OutlookAnywhereInternalClientsRequireSsl,

    [Parameter(Mandatory = $false, ParameterSetName = 'Configure')]
    [Nullable[bool]]$OutlookAnywhereExternalClientsRequireSsl,

    [Parameter(Mandatory = $false, ParameterSetName = 'Configure')]
    [ValidateSet('Basic','Ntlm','Negotiate')]
    [string]$OutlookAnywhereDefaultAuthenticationMethod,

    [Parameter(Mandatory = $false, ParameterSetName = 'Configure')]
    [string]$AutodiscoverSCPNamespace,

    [Parameter(Mandatory = $false, ParameterSetName = 'Clone')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Configure')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Review')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Restore')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')]
    [string]$OutputFile,

    [switch]$Help
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$script:ScriptBaseName = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$script:PublicScriptName = 'ExchangeURLManager.ps1'
$script:PublicScriptVersion = '1.0'

if ($Help) {
    $helpPages = @(
@"
ExchangeURLManager.ps1
Exchange Server Client Access URL / namespace manager

HELP 1/4 - MODES

  Interactive Configure:
    .\ExchangeURLManager.ps1 -Interactive

  Review:
    .\ExchangeURLManager.ps1 -Review -Server EX01,EX02

  Configure:
    .\ExchangeURLManager.ps1 -Server EX01,EX02 `
      -InternalNamespace mail.contoso.com `
      -ExternalNamespace mail.contoso.com

  Clone comparison (default, read-only):
    .\ExchangeURLManager.ps1 -SourceServer EX01 -TargetServer EX02,EX03

  Clone comparison (explicit):
    .\ExchangeURLManager.ps1 -SourceServer EX01 -TargetServer EX02,EX03 -CompareOnly

  Clone Apply:
    .\ExchangeURLManager.ps1 -SourceServer EX01 -TargetServer EX02,EX03 -ApplyChanges

  Backup:
    .\ExchangeURLManager.ps1 -Backup -Server EX01,EX02

  Restore:
    .\ExchangeURLManager.ps1 -Restore -BackupFile C:\Temp\EX01-ExchangeURL.json
"@,
@"
HELP 2/4 - INTERACTIVE CONFIGURE

  Start:
    .\ExchangeURLManager.ps1 -Interactive

  Configuration styles:
    1. Common namespace
    2. Select components
    3. Advanced per-component configuration

  Advanced selection:
    1. Internal only
    2. External only
    3. Both
    4. Skip

  Default:
    4 = Skip

  Value entry:
    Enter = keep current
    NONE  = clear the selected supported value

  Multi-server mode shows current values for each server.

  PowerShell URLs are never selected automatically.
"@,
@"
HELP 3/4 - SAFETY

  Review and Backup are read-only.

  Configure / Interactive / Restore and Clone with -ApplyChanges:
    Preview
      -> Confirm
      -> JSON pre-change backup
      -> Apply
      -> Verify

  Clone safety default:
    SourceServer + TargetServer = read-only comparison
    -CompareOnly                = explicit read-only intent
    -ApplyChanges               = explicit Clone Apply opt-in
    CompareOnly + ApplyChanges  = BLOCKER

  -OutputFile:
    Exports/reports only. No Exchange configuration is changed.

  Parameter-mode explicit opt-ins:
    -IncludeAuthentication
    -IncludePowerShellUrls
    -IncludeOutlookAnywhereSslRequirements

  Interactive uses explicit menu selections. PowerShell URLs are not selected by default.

  Review Only:
    ASA / Kerberos
    EWS MRSProxyEnabled
    Outlook Anywhere SSLOffloading
    PowerShell authentication
    PowerShell RequireSSL
    Extended Protection
"@,
@"
HELP 4/4 - EXAMPLES

  Interactive for selected servers:
    .\ExchangeURLManager.ps1 -Interactive -Server EX01,EX02

  Interactive command export:
    .\ExchangeURLManager.ps1 -Interactive -Server EX01,EX02 `
      -OutputFile C:\Temp\ExchangeURL-Commands.txt

  Clone comparison (default):
    .\ExchangeURLManager.ps1 -SourceServer EX01 `
      -TargetServer EX02,EX03 -IncludeAuthentication

  Clone comparison (explicit):
    .\ExchangeURLManager.ps1 -SourceServer EX01 `
      -TargetServer EX02,EX03 -CompareOnly -IncludeAuthentication

  Clone Apply:
    .\ExchangeURLManager.ps1 -SourceServer EX01 `
      -TargetServer EX02,EX03 -ApplyChanges -IncludeAuthentication

  Comparison console paging:
    Press ENTER to continue or Q to exit.
    Paging is disabled when -OutputFile is used or interactive input is unavailable.

  Full comment-based help:
    Get-Help .\ExchangeURLManager.ps1 -Full
"@
    )

    for ($page = 0; $page -lt $helpPages.Count; $page++) {
        $helpPages[$page] | Write-Host

        if ($page -lt ($helpPages.Count - 1)) {
            [void](Read-Host 'Press ENTER to continue')
            Write-Host
        }
    }

    return
}

# ---------------------------------------------------------------------------
# Startup Summary / Operator Context
# ---------------------------------------------------------------------------
function Show-FeedbackFooter {
    Write-Host ''
    Write-Host '############################################' -ForegroundColor Cyan
    Write-Host '# Feedback / Bugs' -ForegroundColor Cyan
    Write-Host '############################################' -ForegroundColor Cyan
    Write-Host 'For updates, feedback, bugs, and feature requests:'
    Write-Host 'https://github.com/Ceyhun-Kirmizitas' -ForegroundColor DarkGray
    Write-Host 'https://ceyhunkirmizitas.net' -ForegroundColor DarkGray
    Write-Host '############################################' -ForegroundColor Cyan
}


function Test-ConsolePagingAvailable {
    if (-not [Environment]::UserInteractive) {
        return $false
    }

    if ($Host.Name -ne 'ConsoleHost') {
        return $false
    }

    try {
        if ([Console]::IsInputRedirected) {
            return $false
        }
    }
    catch {
        return $false
    }

    return $true
}

function Read-ComparisonPageContinuation {
    while ($true) {
        $choice = Read-Host 'Press ENTER to continue or Q to exit'

        if ([string]::IsNullOrWhiteSpace($choice)) {
            Write-Host ''
            return $true
        }

        if ($choice.Trim() -match '^(?i:q|quit|exit)$') {
            return $false
        }

        Write-Warning 'Press ENTER to continue or Q to exit.'
    }
}

function Get-StartupContext {
    param(
        [Parameter(Mandatory = $true)][string]$ParameterSetName,
        [switch]$CompareOnly,
        [switch]$ApplyChanges,
        [AllowNull()][string]$OutputFile
    )

    $isCommandExport = -not [string]::IsNullOrWhiteSpace($OutputFile)

    switch ($ParameterSetName) {
        'Review' {
            return [PSCustomObject]@{
                Mode           = 'Configuration Review'
                Changes        = 'NONE'
                ChangesEnabled = $false
                Pause          = $false
                Description1   = 'The script reads supported Client Access URL/namespace configuration in this mode.'
                Description2   = 'No Exchange configuration changes will be made.'
            }
        }

        'Backup' {
            return [PSCustomObject]@{
                Mode           = 'Configuration Backup'
                Changes        = 'NONE'
                ChangesEnabled = $false
                Pause          = $false
                Description1   = 'The script reads supported Client Access configuration and creates JSON/TXT backup files.'
                Description2   = 'No Exchange configuration changes will be made.'
            }
        }

        'Interactive' {
            if ($isCommandExport) {
                return [PSCustomObject]@{
                    Mode           = 'Interactive Configure / Command Export'
                    Changes        = 'NONE'
                    ChangesEnabled = $false
                    Pause          = $true
                    Description1   = 'The script reads the current Client Access configuration, lets you select URL/hostname changes, and builds a Preview for command export.'
                    Description2   = 'No Exchange configuration changes will be made.'
                }
            }

            return [PSCustomObject]@{
                Mode           = 'Interactive Configure'
                Changes        = 'ENABLED'
                ChangesEnabled = $true
                Pause          = $true
                Description1   = 'The script reads the current Client Access configuration, lets you select URL/hostname changes, and shows a complete Preview.'
                Description2   = 'No changes are made until you confirm the plan. A successful pre-change backup is required before Apply starts.'
            }
        }

        'Clone' {
            if (-not $ApplyChanges) {
                return [PSCustomObject]@{
                    Mode           = $(if ($CompareOnly) { 'Compare Only' } else { 'Compare' })
                    Changes        = 'NONE'
                    ChangesEnabled = $false
                    Pause          = $false
                    Description1   = 'The script compares the source server with the target server(s).'
                    Description2   = 'No Exchange configuration changes will be made. Use -ApplyChanges explicitly to enable Clone Apply.'
                }
            }

            if ($isCommandExport) {
                return [PSCustomObject]@{
                    Mode           = 'Clone / Command Export'
                    Changes        = 'NONE'
                    ChangesEnabled = $false
                    Pause          = $false
                    Description1   = 'The script builds the source-to-target change plan and exports required commands.'
                    Description2   = 'No Exchange configuration changes will be made.'
                }
            }

            return [PSCustomObject]@{
                Mode           = 'Clone / Apply'
                Changes        = 'ENABLED'
                ChangesEnabled = $true
                Pause          = $true
                Description1   = 'The script compares the source server with the target server(s) and shows a Preview.'
                Description2   = 'Apply was explicitly enabled with -ApplyChanges and still requires confirmation and successful pre-change backup creation.'
            }
        }

        'Restore' {
            if ($isCommandExport) {
                return [PSCustomObject]@{
                    Mode           = 'Restore / Command Export'
                    Changes        = 'NONE'
                    ChangesEnabled = $false
                    Pause          = $false
                    Description1   = 'The script validates the backup and exports the required restore commands.'
                    Description2   = 'No Exchange configuration changes will be made.'
                }
            }

            return [PSCustomObject]@{
                Mode           = 'Restore'
                Changes        = 'ENABLED'
                ChangesEnabled = $true
                Pause          = $true
                Description1   = 'The script validates the backup and shows a same-server Restore Preview.'
                Description2   = 'Apply requires confirmation and successful pre-change backup creation.'
            }
        }

        default {
            if ($isCommandExport) {
                return [PSCustomObject]@{
                    Mode           = 'Configure / Command Export'
                    Changes        = 'NONE'
                    ChangesEnabled = $false
                    Pause          = $false
                    Description1   = 'The script builds the required URL/namespace change plan and exports commands.'
                    Description2   = 'No Exchange configuration changes will be made.'
                }
            }

            return [PSCustomObject]@{
                Mode           = 'Configure'
                Changes        = 'ENABLED'
                ChangesEnabled = $true
                Pause          = $true
                Description1   = 'The script builds the required URL/namespace change plan and shows a Preview.'
                Description2   = 'Apply requires confirmation and successful pre-change backup creation.'
            }
        }
    }
}

function Show-StartupSummary {
    param([Parameter(Mandatory = $true)]$Context)

    Write-Host ''
    Write-Host $script:PublicScriptName
    Write-Host 'Exchange Server Client Access URL and namespace configuration, review, backup, clone, and restore manager' -ForegroundColor Cyan
    Write-Host ''
    Write-Host 'Author  : Ceyhun Kirmizitas' -ForegroundColor Cyan
    Write-Host ("Version : {0}" -f $script:PublicScriptVersion) -ForegroundColor Cyan
    Write-Host ("Mode    : {0}" -f $Context.Mode) -ForegroundColor Cyan
    Write-Host ("Changes : {0}" -f $Context.Changes) -ForegroundColor $(if ($Context.ChangesEnabled) { 'Yellow' } else { 'Green' })
    Write-Host ''
    Write-Host $Context.Description1
    Write-Host $Context.Description2
    Write-Host ''
    Write-Host 'Compatibility note: Intended for Exchange Server 2016, Exchange Server 2019, and Exchange Server Subscription Edition.' -ForegroundColor Yellow
    Write-Host 'No hard Exchange-version gate is enforced.' -ForegroundColor Yellow
    Write-Host ''

    if ($Context.Pause) {
        while ($true) {
            $startChoice = Read-Host 'Press ENTER to start or Q to exit'

            if ([string]::IsNullOrWhiteSpace($startChoice)) {
                Write-Host ''
                break
            }

            if ($startChoice.Trim() -match '^(?i:q|quit|exit)$') {
                Write-Host ''
                Write-Host 'Exiting. No changes were made.' -ForegroundColor DarkCyan
                exit 0
            }

            Write-Warning 'Press ENTER to start or Q to exit.'
        }
    }
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

    # Microsoft's RemoteExchange.ps1 is not StrictMode-clean on every supported build.
    # Disable StrictMode only for the EMS bootstrap sequence, then restore it.
    Set-StrictMode -Off
    try {
        . $remoteExchange
        Connect-ExchangeServer -Auto -AllowClobber | Out-Null
    }
    finally {
        Set-StrictMode -Version 2.0
    }

    if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) { throw 'Exchange Management Shell could not be initialized.' }
}

# Implementation reference: CK-EUM-01
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
# Interactive Configure input layer
# ---------------------------------------------------------------------------
# Interactive mode collects operator intent only. It deliberately reuses the
# normal desired-state, change-plan, export, backup, Apply, and verification
# engines so interactive and parameter-based configuration cannot drift apart.

function New-InteractiveValueAction {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Keep','Set','Clear')]
        [string]$Action,
        $Value
    )

    return [PSCustomObject]@{
        Action = $Action
        Value  = $Value
    }
}

function Read-InteractiveMenuChoice {
    param(
        [Parameter(Mandatory = $true)][string]$Prompt,
        [Parameter(Mandatory = $true)][string[]]$ValidChoices,
        [Parameter(Mandatory = $true)][string]$DefaultChoice
    )

    while ($true) {
        $inputValue = Read-Host "$Prompt [$DefaultChoice]"
        if ([string]::IsNullOrWhiteSpace($inputValue)) { return $DefaultChoice }

        $choice = $inputValue.Trim()
        if ($ValidChoices -contains $choice) { return $choice }

        Write-Warning "Invalid selection. Choose: $($ValidChoices -join ', ')."
    }
}

function Read-InteractiveYesNo {
    param(
        [Parameter(Mandatory = $true)][string]$Prompt,
        [bool]$Default = $false
    )

    $defaultText = if ($Default) { 'Y' } else { 'N' }
    while ($true) {
        $value = Read-Host "$Prompt [Y/N] [$defaultText]"
        if ([string]::IsNullOrWhiteSpace($value)) { return $Default }

        switch -Regex ($value.Trim()) {
            '^(?i:y|yes)$' { return $true }
            '^(?i:n|no)$'  { return $false }
            default { Write-Warning 'Enter Y or N.' }
        }
    }
}

function Resolve-InteractiveServers {
    param([Parameter(Mandatory = $true)][string[]]$Servers)

    $resolved = New-Object System.Collections.ArrayList
    $invalid = New-Object System.Collections.ArrayList

    foreach ($server in $Servers) {
        try {
            $exchangeServer = Get-ExchangeServer -Identity $server -ErrorAction Stop
            $name = ([string]$exchangeServer.Name).Trim()
            if ([string]::IsNullOrWhiteSpace($name)) {
                throw "Exchange returned an empty server name."
            }

            if (-not (@($resolved | Where-Object { $_ -ieq $name }).Count -gt 0)) {
                [void]$resolved.Add($name)
            }
        }
        catch {
            [void]$invalid.Add([PSCustomObject]@{
                Server = $server
                Error  = $_.Exception.Message
            })
        }
    }

    return [PSCustomObject]@{
        Resolved = @($resolved)
        Invalid  = @($invalid)
    }
}

function Read-InteractiveServers {
    param([string[]]$ProvidedServers)

    if ($ProvidedServers -and $ProvidedServers.Count -gt 0) {
        $servers = @(Get-NormalizedServers -Servers $ProvidedServers)
        if ($servers.Count -eq 0) { throw 'At least one server is required.' }

        $validation = Resolve-InteractiveServers -Servers $servers
        if ($validation.Invalid.Count -gt 0) {
            $names = ($validation.Invalid | ForEach-Object { $_.Server }) -join ', '
            throw "The following server(s) could not be resolved as Exchange servers: $names"
        }

        return @($validation.Resolved)
    }

    $defaultServer = $env:COMPUTERNAME

    while ($true) {
        $inputValue = Read-Host "Server(s), comma-separated (Enter = $defaultServer)"
        $rawServers = if ([string]::IsNullOrWhiteSpace($inputValue)) {
            @($defaultServer)
        }
        else {
            @($inputValue -split ',' | ForEach-Object { $_.Trim() })
        }

        $servers = @(Get-NormalizedServers -Servers $rawServers)
        if ($servers.Count -eq 0) {
            Write-Warning 'Enter at least one Exchange server.'
            continue
        }

        Write-Host 'Validating Exchange server(s)...' -ForegroundColor DarkCyan
        $validation = Resolve-InteractiveServers -Servers $servers

        if ($validation.Invalid.Count -gt 0) {
            foreach ($item in $validation.Invalid) {
                Write-Warning ("Server '{0}' could not be resolved as an Exchange server." -f $item.Server)
            }
            Write-Warning 'Review the server name(s) and try again.'
            continue
        }

        Write-Host 'Server validation passed.' -ForegroundColor DarkCyan
        return @($validation.Resolved)
    }
}

function ConvertTo-InteractiveAbsoluteUrl {
    param([Parameter(Mandatory = $true)][string]$Value)

    $urlText = $Value.Trim()
    $uri = $null
    if (-not [Uri]::TryCreate($urlText, [UriKind]::Absolute, [ref]$uri) -or
        $uri.Scheme -notin @('http','https') -or
        [string]::IsNullOrWhiteSpace($uri.Host)) {
        throw "Invalid HTTP/HTTPS URL: $Value"
    }
    return $urlText
}

function ConvertTo-InteractiveNamespaceInput {
    param([Parameter(Mandatory = $true)][string]$Value)

    $namespace = $Value.Trim()
    if ([string]::IsNullOrWhiteSpace($namespace)) {
        throw 'Namespace cannot be blank.'
    }
    if ($namespace -match '^(?i:none|null|clear)$') {
        throw 'Enter a namespace/hostname. Clear keywords are not valid in this prompt.'
    }
    if ($namespace -match '[/\\?#]' -or $namespace -match '^[a-zA-Z][a-zA-Z0-9+.-]*://') {
        throw 'Enter a namespace/hostname only, not a full URL or path.'
    }
    if ([Uri]::CheckHostName($namespace) -ne [UriHostNameType]::Dns) {
        throw "Invalid namespace/hostname: $Value"
    }
    return $namespace.TrimEnd('.')
}

function Read-InteractiveUrlAction {
    param(
        [Parameter(Mandatory = $true)][string]$Label,
        [switch]$AllowClear
    )

    while ($true) {
        Write-Host ''
        Write-Host $Label -ForegroundColor Cyan
        Write-Host '  Enter : keep current'
        if ($AllowClear) { Write-Host '  NONE  : clear' }
        $value = Read-Host '  Value'

        if ([string]::IsNullOrWhiteSpace($value)) {
            return New-InteractiveValueAction -Action 'Keep'
        }
        if ($value.Trim() -match '^(?i:none|null|clear)$') {
            if ($AllowClear) { return New-InteractiveValueAction -Action 'Clear' }
            Write-Warning 'Clearing this value is not enabled in Interactive mode.'
            continue
        }

        try {
            $normalized = ConvertTo-InteractiveAbsoluteUrl -Value $value
            return New-InteractiveValueAction -Action 'Set' -Value $normalized
        }
        catch {
            Write-Warning $_.Exception.Message
        }
    }
}

function Read-InteractiveNamespaceUrlAction {
    param(
        [Parameter(Mandatory = $true)][string]$Label,
        [Parameter(Mandatory = $true)][string]$Path,
        [ValidateSet('http','https')]
        [string]$Scheme = 'https',
        [switch]$AllowClear
    )

    while ($true) {
        Write-Host ''
        Write-Host $Label -ForegroundColor Cyan
        Write-Host ("  URL path : {0}" -f $Path) -ForegroundColor DarkCyan
        Write-Host '  Enter    : keep current'
        if ($AllowClear) { Write-Host '  NONE     : clear' }
        $value = Read-Host '  Namespace'

        if ([string]::IsNullOrWhiteSpace($value)) {
            return New-InteractiveValueAction -Action 'Keep'
        }

        if ($value.Trim() -match '^(?i:none|null|clear)$') {
            if ($AllowClear) { return New-InteractiveValueAction -Action 'Clear' }
            Write-Warning 'Clearing this value is not enabled in Interactive mode.'
            continue
        }

        try {
            $namespace = ConvertTo-InteractiveNamespaceInput -Value $value
            $url = "$Scheme`://$namespace$Path"
            Write-Host ("  Generated URL: {0}" -f $url) -ForegroundColor DarkCyan
            return New-InteractiveValueAction -Action 'Set' -Value $url
        }
        catch {
            Write-Warning $_.Exception.Message
        }
    }
}

function Get-InteractiveCurrentUrlScheme {
    param(
        [Parameter(Mandatory = $true)][array]$States,
        [Parameter(Mandatory = $true)][string]$ComponentKey,
        [Parameter(Mandatory = $true)][string]$PropertyName
    )

    $schemes = @()

    foreach ($state in $States) {
        $value = Get-InteractiveStateCell -State $state -ComponentKey $ComponentKey -PropertyName $PropertyName
        if ([string]::IsNullOrWhiteSpace($value)) { continue }
        if ($value -in @('<Unable to Read>','<Not Available>','<null>','<empty>')) { continue }

        $uri = $null
        if ([Uri]::TryCreate([string]$value, [UriKind]::Absolute, [ref]$uri) -and
            $uri.Scheme -in @('http','https')) {
            $schemes += $uri.Scheme.ToLowerInvariant()
        }
    }

    $unique = @($schemes | Select-Object -Unique)
    if ($unique.Count -eq 1) { return $unique[0] }
    return $null
}

function Read-InteractiveUrlScheme {
    param(
        [AllowNull()][string]$CurrentScheme,
        [AllowNull()][string]$Label
    )

    $default = if ($CurrentScheme -in @('http','https')) { $CurrentScheme.ToUpperInvariant() } else { 'HTTPS' }
    $prompt = if ([string]::IsNullOrWhiteSpace($Label)) {
        "  Protocol [HTTP/HTTPS] [$default]"
    }
    else {
        "  $Label protocol [HTTP/HTTPS] [$default]"
    }

    while ($true) {
        $value = Read-Host $prompt
        if ([string]::IsNullOrWhiteSpace($value)) { return $default.ToLowerInvariant() }

        switch ($value.Trim().ToUpperInvariant()) {
            'HTTP'  { return 'http' }
            'HTTPS' { return 'https' }
            default { Write-Warning 'Enter HTTP or HTTPS.' }
        }
    }
}

function Read-InteractiveHostnameAction {
    param(
        [Parameter(Mandatory = $true)][string]$Label,
        [switch]$AllowClear
    )

    while ($true) {
        Write-Host ''
        Write-Host $Label -ForegroundColor Cyan
        Write-Host '  Enter : keep current'
        if ($AllowClear) { Write-Host '  NONE  : clear' }
        $value = Read-Host '  Value'

        if ([string]::IsNullOrWhiteSpace($value)) {
            return New-InteractiveValueAction -Action 'Keep'
        }
        if ($value.Trim() -match '^(?i:none|null|clear)$') {
            if ($AllowClear) { return New-InteractiveValueAction -Action 'Clear' }
            Write-Warning 'Clearing this value is not enabled in Interactive mode.'
            continue
        }

        try {
            $normalized = ConvertTo-InteractiveNamespaceInput -Value $value
            return New-InteractiveValueAction -Action 'Set' -Value $normalized
        }
        catch {
            Write-Warning $_.Exception.Message
        }
    }
}

function Read-InteractiveScpAction {
    while ($true) {
        Write-Host ''
        Write-Host 'New Autodiscover SCP namespace or URI' -ForegroundColor Cyan
        Write-Host '  Enter : keep current'
        Write-Host '  Example namespace: autodiscover.contoso.com'
        Write-Host '  Example URI      : https://autodiscover.contoso.com/Autodiscover/Autodiscover.xml'
        $value = Read-Host '  Value'

        if ([string]::IsNullOrWhiteSpace($value)) {
            return New-InteractiveValueAction -Action 'Keep'
        }

        try {
            if ($value.Trim() -match '^[a-zA-Z][a-zA-Z0-9+.-]*://') {
                $uri = ConvertTo-InteractiveAbsoluteUrl -Value $value
                return New-InteractiveValueAction -Action 'Set' -Value $uri
            }

            $namespace = Normalize-Namespace -Value $value
            return New-InteractiveValueAction -Action 'Set' -Value "https://$namespace/Autodiscover/Autodiscover.xml"
        }
        catch {
            Write-Warning $_.Exception.Message
        }
    }
}

function Get-InteractiveServerState {
    param([Parameter(Mandatory = $true)][string]$Server)

    $state = [ordered]@{
        Server            = $Server
        ResolvedServer    = $null
        ResolutionError   = $null
        Components        = @{}
        OutlookAny        = $null
        OutlookAnyError   = $null
        ClientAccess      = $null
        ClientAccessError = $null
    }

    try {
        $exchangeServer = Get-ExchangeServer -Identity $Server -ErrorAction Stop
        $state.ResolvedServer = [string]$exchangeServer.Name
    }
    catch {
        $state.ResolutionError = $_.Exception.Message
        return [PSCustomObject]$state
    }

    foreach ($entry in $VirtualDirectoryMap) {
        if (-not $entry.Path) { continue }

        try {
            $component = Get-DefaultVirtualDirectory -CommandName $entry.GetCmd -Server $state.ResolvedServer
            $state.Components[$entry.Key] = [PSCustomObject]@{ Success = $true; Object = $component; Error = $null }
        }
        catch {
            $state.Components[$entry.Key] = [PSCustomObject]@{ Success = $false; Object = $null; Error = $_.Exception.Message }
        }
    }

    try {
        $state.OutlookAny = Get-DefaultVirtualDirectory -CommandName 'Get-OutlookAnywhere' -Server $state.ResolvedServer
    }
    catch {
        $state.OutlookAnyError = $_.Exception.Message
    }

    try {
        $state.ClientAccess = Get-ClientAccessService -Identity $state.ResolvedServer -ErrorAction Stop
    }
    catch {
        $state.ClientAccessError = $_.Exception.Message
    }

    return [PSCustomObject]$state
}

function Get-InteractiveStateCell {
    param(
        [Parameter(Mandatory = $true)]$State,
        [Parameter(Mandatory = $true)][string]$ComponentKey,
        [Parameter(Mandatory = $true)][string]$PropertyName
    )

    if ($State.ResolutionError) { return '<Unable to Read>' }

    $object = $null
    if ($ComponentKey -eq 'OutlookAny') {
        if ($State.OutlookAnyError -or $null -eq $State.OutlookAny) { return '<Unable to Read>' }
        $object = $State.OutlookAny
    }
    elseif ($ComponentKey -eq 'ClientAccess') {
        if ($State.ClientAccessError -or $null -eq $State.ClientAccess) { return '<Unable to Read>' }
        $object = $State.ClientAccess
    }
    else {
        if (-not $State.Components.ContainsKey($ComponentKey)) { return '<Unable to Read>' }
        $componentState = $State.Components[$ComponentKey]
        if (-not $componentState.Success -or $null -eq $componentState.Object) { return '<Unable to Read>' }
        $object = $componentState.Object
    }

    $property = $object.PSObject.Properties[$PropertyName]
    if ($null -eq $property) { return '<Not Available>' }
    return ConvertTo-DisplayValue -Value $property.Value
}

function Write-InteractiveComponentHeader {
    param([Parameter(Mandatory = $true)][string]$Title)

    Write-Host ''
    Write-Host ('#' * 78) -ForegroundColor DarkCyan
    Write-Host ("# {0}" -f $Title) -ForegroundColor Cyan
    Write-Host ('#' * 78) -ForegroundColor DarkCyan
    Write-Host ''
}

function Show-InteractiveDualValues {
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][array]$States,
        [Parameter(Mandatory = $true)][string]$ComponentKey,
        [Parameter(Mandatory = $true)][string]$InternalProperty,
        [Parameter(Mandatory = $true)][string]$ExternalProperty
    )

    Write-InteractiveComponentHeader -Title $Title
    Write-Host 'Current values:' -ForegroundColor DarkCyan
    Write-Host ''

    foreach ($state in $States) {
        $internalValue = Get-InteractiveStateCell -State $state -ComponentKey $ComponentKey -PropertyName $InternalProperty
        $externalValue = Get-InteractiveStateCell -State $state -ComponentKey $ComponentKey -PropertyName $ExternalProperty

        Write-Host ("[{0}]" -f $state.Server) -ForegroundColor Cyan
        Write-Host ("{0} : {1}" -f $InternalProperty, $internalValue)
        Write-Host ("{0} : {1}" -f $ExternalProperty, $externalValue)
        Write-Host ''
    }
}

function Show-InteractiveDualValuesBody {
    param(
        [Parameter(Mandatory = $true)][array]$States,
        [Parameter(Mandatory = $true)][string]$ComponentKey,
        [Parameter(Mandatory = $true)][string]$InternalProperty,
        [Parameter(Mandatory = $true)][string]$ExternalProperty
    )

    Write-Host 'Current values:' -ForegroundColor DarkCyan
    Write-Host ''

    foreach ($state in $States) {
        $internalValue = Get-InteractiveStateCell -State $state -ComponentKey $ComponentKey -PropertyName $InternalProperty
        $externalValue = Get-InteractiveStateCell -State $state -ComponentKey $ComponentKey -PropertyName $ExternalProperty

        Write-Host ("[{0}]" -f $state.Server) -ForegroundColor Cyan
        Write-Host ("{0} : {1}" -f $InternalProperty, $internalValue)
        Write-Host ("{0} : {1}" -f $ExternalProperty, $externalValue)
        Write-Host ''
    }
}

function Show-InteractiveSingleValue {
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][array]$States,
        [Parameter(Mandatory = $true)][string]$ComponentKey,
        [Parameter(Mandatory = $true)][string]$PropertyName
    )

    Write-InteractiveComponentHeader -Title $Title
    Write-Host 'Current values:' -ForegroundColor DarkCyan
    Write-Host ''

    foreach ($state in $States) {
        $value = Get-InteractiveStateCell -State $state -ComponentKey $ComponentKey -PropertyName $PropertyName

        Write-Host ("[{0}]" -f $state.Server) -ForegroundColor Cyan
        Write-Host ("{0} : {1}" -f $PropertyName, $value)
        Write-Host ''
    }
}

function Read-InteractiveSideChoice {
    Write-Host ''
    Write-Host 'Change:'
    Write-Host '  1. Internal only'
    Write-Host '  2. External only'
    Write-Host '  3. Both'
    Write-Host '  4. Skip'
    return Read-InteractiveMenuChoice -Prompt 'Select' -ValidChoices @('1','2','3','4') -DefaultChoice '4'
}

function Read-InteractiveCommonNamespaces {
    param(
        [bool]$IncludeCommonNamespaces = $true,
        [bool]$IncludeScp = $true
    )

    Write-Host ''
    Write-Host '[Common namespace values]' -ForegroundColor Cyan
    Write-Host 'Enter namespace/hostname only; URL paths are generated automatically.' -ForegroundColor DarkCyan
    Write-Host 'Example: mail.contoso.com -> https://mail.contoso.com/owa' -ForegroundColor DarkCyan
    Write-Host ''

    $values = [ordered]@{}
    $prompts = @()
    if ($IncludeCommonNamespaces) {
        $prompts += @('Internal namespace','External namespace')
    }
    if ($IncludeScp) {
        $prompts += 'Autodiscover SCP namespace'
    }

    foreach ($prompt in $prompts) {
        while ($true) {
            try {
                $values[$prompt] = ConvertTo-InteractiveNamespaceInput -Value ([string](Read-Host $prompt))
                break
            }
            catch {
                Write-Warning $_.Exception.Message
            }
        }
    }

    [pscustomobject]@{
        InternalNamespace     = if ($IncludeCommonNamespaces) { $values['Internal namespace'] } else { $null }
        ExternalNamespace     = if ($IncludeCommonNamespaces) { $values['External namespace'] } else { $null }
        AutodiscoverNamespace = if ($IncludeScp) { $values['Autodiscover SCP namespace'] } else { $null }
        ClearExternalUrls     = $false
    }
}

function Get-InteractiveComponentMenu {
    return @(
        [PSCustomObject]@{ Number = 1; Key = 'Owa';             Name = 'OWA' }
        [PSCustomObject]@{ Number = 2; Key = 'Ecp';             Name = 'ECP' }
        [PSCustomObject]@{ Number = 3; Key = 'Ews';             Name = 'EWS' }
        [PSCustomObject]@{ Number = 4; Key = 'Mapi';            Name = 'MAPI' }
        [PSCustomObject]@{ Number = 5; Key = 'ActiveSync';      Name = 'ActiveSync' }
        [PSCustomObject]@{ Number = 6; Key = 'Oab';             Name = 'OAB' }
        [PSCustomObject]@{ Number = 7; Key = 'OutlookAny';      Name = 'Outlook Anywhere (RPC over HTTP)' }
        [PSCustomObject]@{ Number = 8; Key = 'AutodiscoverSCP'; Name = 'Autodiscover SCP' }
        [PSCustomObject]@{ Number = 9; Key = 'PowerShell';      Name = 'PowerShell' }
    )
}

function Read-InteractiveComponentSelection {
    $menu = @(Get-InteractiveComponentMenu)
    Write-Host ''
    Write-Host 'Components:' -ForegroundColor Cyan
    foreach ($item in $menu) {
        Write-Host ("  {0}. {1}" -f $item.Number, $item.Name)
    }

    $default = '1,2,3,4,5,6,7,8'
    while ($true) {
        Write-Host ''
        Write-Host 'Select components by number, separated by commas.' -ForegroundColor Cyan
        Write-Host 'Example: 1,3,5' -ForegroundColor DarkGray
        Write-Host ''
        Write-Host 'Press ENTER to select all standard components (1-8).' -ForegroundColor Cyan
        Write-Host 'PowerShell is not selected by default.' -ForegroundColor Yellow
        Write-Host ''
        $inputValue = Read-Host 'Selection'
        if ([string]::IsNullOrWhiteSpace($inputValue)) { $inputValue = $default }

        $parts = @($inputValue -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        $selected = New-Object System.Collections.ArrayList
        $valid = $true

        foreach ($part in $parts) {
            $number = 0
            if (-not [int]::TryParse($part, [ref]$number)) { $valid = $false; break }
            $match = @($menu | Where-Object { $_.Number -eq $number })
            if ($match.Count -ne 1) { $valid = $false; break }
            if (-not (@($selected | Where-Object { $_.Key -eq $match[0].Key }).Count -gt 0)) {
                [void]$selected.Add($match[0])
            }
        }

        if ($valid -and $selected.Count -gt 0) { return @($selected) }
        Write-Warning 'Enter one or more valid component numbers separated by commas.'
    }
}

function New-InteractiveSpec {
    return [PSCustomObject]@{
        ComponentActions    = [ordered]@{}
        ScpAction           = $null
        Style               = $null
        PowerShellSelected  = $false
    }
}

function Add-InteractiveCommonActions {
    param(
        [Parameter(Mandatory = $true)]$Spec,
        [Parameter(Mandatory = $true)]$Namespaces,
        [Parameter(Mandatory = $true)][array]$SelectedComponents,
        [bool]$IncludeScp,
        [ValidateSet('http','https')]
        [string]$PowerShellInternalScheme = 'http',
        [ValidateSet('http','https')]
        [string]$PowerShellExternalScheme = 'http'
    )

    foreach ($selected in $SelectedComponents) {
        if ($selected.Key -eq 'OutlookAny') {
            $Spec.ComponentActions['OutlookAny'] = [ordered]@{
                InternalHostname = New-InteractiveValueAction -Action 'Set' -Value $Namespaces.InternalNamespace
                ExternalHostname = if ($Namespaces.ClearExternalUrls) {
                    New-InteractiveValueAction -Action 'Clear'
                }
                else {
                    New-InteractiveValueAction -Action 'Set' -Value $Namespaces.ExternalNamespace
                }
            }
            continue
        }

        $entry = $VirtualDirectoryMap | Where-Object { $_.Key -eq $selected.Key } | Select-Object -First 1
        if ($null -eq $entry -or -not $entry.Path) { continue }

        if ($entry.Key -eq 'PowerShell') {
            $Spec.ComponentActions[$entry.Key] = [ordered]@{
                InternalUrl = New-InteractiveValueAction -Action 'Set' -Value "$PowerShellInternalScheme`://$($Namespaces.InternalNamespace)$($entry.Path)"
                ExternalUrl = if ($Namespaces.ClearExternalUrls) {
                    New-InteractiveValueAction -Action 'Clear'
                }
                else {
                    New-InteractiveValueAction -Action 'Set' -Value "$PowerShellExternalScheme`://$($Namespaces.ExternalNamespace)$($entry.Path)"
                }
            }
            $Spec.PowerShellSelected = $true
            continue
        }

        $Spec.ComponentActions[$entry.Key] = [ordered]@{
            InternalUrl = New-InteractiveValueAction -Action 'Set' -Value "https://$($Namespaces.InternalNamespace)$($entry.Path)"
            ExternalUrl = if ($Namespaces.ClearExternalUrls) {
                New-InteractiveValueAction -Action 'Clear'
            }
            else {
                New-InteractiveValueAction -Action 'Set' -Value "https://$($Namespaces.ExternalNamespace)$($entry.Path)"
            }
        }
    }

    if ($IncludeScp) {
        $Spec.ScpAction = New-InteractiveValueAction -Action 'Set' -Value "https://$($Namespaces.AutodiscoverNamespace)/Autodiscover/Autodiscover.xml"
    }
}

function Show-InteractivePowerShellSafetyNote {
    Write-Host ''
    Write-Host 'PowerShell' -ForegroundColor Cyan
    Write-Host '----------' -ForegroundColor Cyan
    Write-Host 'Note: The "/PowerShell" virtual directory is changed only when explicitly selected.' -ForegroundColor DarkYellow
    Write-Host 'Microsoft recommends modifying it only at the request of Customer Service and Support.' -ForegroundColor DarkYellow
    Write-Host 'PowerShell commonly uses HTTP; keep an existing HTTPS configuration unless you intend to change it.' -ForegroundColor DarkYellow
    Write-Host 'Reference: https://learn.microsoft.com/powershell/module/exchangepowershell/set-powershellvirtualdirectory?view=exchange-ps' -ForegroundColor DarkGray
    Write-Host ''
}

function Read-InteractivePowerShellProtocols {
    Write-Host ''
    Write-Host '[PowerShell protocol]' -ForegroundColor Cyan
    Write-Host 'Select the protocol for the PowerShell Internal and External URLs.' -ForegroundColor DarkCyan

    $internalScheme = Read-InteractiveUrlScheme -CurrentScheme 'http' -Label 'Internal'
    $externalScheme = Read-InteractiveUrlScheme -CurrentScheme 'http' -Label 'External'

    return [PSCustomObject]@{
        Internal = $internalScheme
        External = $externalScheme
    }
}

function Read-InteractiveAdvancedSpec {
    param([Parameter(Mandatory = $true)][array]$States)

    $spec = New-InteractiveSpec
    $spec.Style = 'Advanced per-component configuration'

    foreach ($entry in @($VirtualDirectoryMap | Where-Object { $_.Path -and $_.Key -ne 'PowerShell' })) {
        Show-InteractiveDualValues -Title $entry.Name -States $States -ComponentKey $entry.Key -InternalProperty 'InternalUrl' -ExternalProperty 'ExternalUrl'
        $choice = Read-InteractiveSideChoice
        if ($choice -eq '4') { continue }

        $actions = [ordered]@{}
        if ($choice -in @('1','3')) {
            $actions['InternalUrl'] = Read-InteractiveNamespaceUrlAction -Label 'New Internal namespace' -Path $entry.Path -AllowClear
        }
        if ($choice -in @('2','3')) {
            $actions['ExternalUrl'] = Read-InteractiveNamespaceUrlAction -Label 'New External namespace' -Path $entry.Path -AllowClear
        }
        $spec.ComponentActions[$entry.Key] = $actions
    }

    Show-InteractiveDualValues -Title 'Outlook Anywhere (RPC over HTTP)' -States $States -ComponentKey 'OutlookAny' -InternalProperty 'InternalHostname' -ExternalProperty 'ExternalHostname'
    $choice = Read-InteractiveSideChoice
    if ($choice -ne '4') {
        $actions = [ordered]@{}
        if ($choice -in @('1','3')) {
            # InternalHostname clearing is intentionally not enabled until it is
            # validated as a supported field scenario across supported builds.
            $actions['InternalHostname'] = Read-InteractiveHostnameAction -Label 'New InternalHostname'
        }
        if ($choice -in @('2','3')) {
            $actions['ExternalHostname'] = Read-InteractiveHostnameAction -Label 'New ExternalHostname' -AllowClear
        }
        $spec.ComponentActions['OutlookAny'] = $actions
    }

    Show-InteractiveSingleValue -Title 'Autodiscover SCP' -States $States -ComponentKey 'ClientAccess' -PropertyName 'AutoDiscoverServiceInternalUri'
    Write-Host 'Configure this component?' -ForegroundColor DarkCyan
    Write-Host ''
    if (Read-InteractiveYesNo -Prompt 'Change Autodiscover SCP?' -Default:$false) {
        $spec.ScpAction = Read-InteractiveScpAction
    }

    Write-InteractiveComponentHeader -Title 'PowerShell'
    Write-Host 'Note: The "/PowerShell" virtual directory is changed only when explicitly selected.' -ForegroundColor DarkYellow
    Write-Host 'Microsoft recommends modifying it only at the request of Customer Service and Support.' -ForegroundColor DarkYellow
    Write-Host 'PowerShell commonly uses HTTP; keep an existing HTTPS configuration unless you intend to change it.' -ForegroundColor DarkYellow
    Write-Host 'Reference: https://learn.microsoft.com/powershell/module/exchangepowershell/set-powershellvirtualdirectory?view=exchange-ps' -ForegroundColor DarkGray
    Write-Host ''

    if (Read-InteractiveYesNo -Prompt 'Configure PowerShell?' -Default:$false) {
        Write-Host ''
        Show-InteractiveDualValuesBody -States $States -ComponentKey 'PowerShell' -InternalProperty 'InternalUrl' -ExternalProperty 'ExternalUrl'
        $choice = Read-InteractiveSideChoice
        if ($choice -ne '4') {
            $actions = [ordered]@{}

            if ($choice -in @('1','3')) {
                $internalScheme = Get-InteractiveCurrentUrlScheme -States $States -ComponentKey 'PowerShell' -PropertyName 'InternalUrl'
                if ($null -eq $internalScheme) { $internalScheme = 'http' }

                Write-Host ''
                Write-Host 'New Internal URL' -ForegroundColor Cyan
                $internalScheme = Read-InteractiveUrlScheme -CurrentScheme $internalScheme -Label 'Internal'
                $actions['InternalUrl'] = Read-InteractiveNamespaceUrlAction `
                    -Label 'New Internal namespace' `
                    -Path '/powershell' `
                    -Scheme $internalScheme `
                    -AllowClear
            }

            if ($choice -in @('2','3')) {
                $externalScheme = Get-InteractiveCurrentUrlScheme -States $States -ComponentKey 'PowerShell' -PropertyName 'ExternalUrl'
                if ($null -eq $externalScheme) { $externalScheme = 'http' }

                Write-Host ''
                Write-Host 'New External URL' -ForegroundColor Cyan
                $externalScheme = Read-InteractiveUrlScheme -CurrentScheme $externalScheme -Label 'External'
                $actions['ExternalUrl'] = Read-InteractiveNamespaceUrlAction `
                    -Label 'New External namespace' `
                    -Path '/powershell' `
                    -Scheme $externalScheme `
                    -AllowClear
            }

            $spec.ComponentActions['PowerShell'] = $actions
            $spec.PowerShellSelected = $true
        }
    }

    return $spec
}

function Read-InteractiveConfigureSpec {
    param([Parameter(Mandatory = $true)][string[]]$Targets)

    Write-Host ''
    Write-Host 'Choose how to configure URLs and hostnames:' -ForegroundColor Cyan
    Write-Host ''
    Write-Host '  1. All standard components - common namespaces'
    Write-Host '     Enter one Internal namespace, one External namespace,'
    Write-Host '     and the Autodiscover SCP namespace.'
    Write-Host '     Applies to OWA, ECP, EWS, MAPI, ActiveSync, OAB,'
    Write-Host '     and Outlook Anywhere (RPC over HTTP).'
    Write-Host ''
    Write-Host '  2. Selected components - common namespaces'
    Write-Host '     Choose the components first, then enter the common'
    Write-Host '     Internal and External namespaces for those components.'
    Write-Host ''
    Write-Host '  3. Configure each component separately'
    Write-Host '     Shows current values for every target server.'
    Write-Host '     For each component, choose Internal, External, Both, or Skip.'
    Write-Host ''
    Write-Host 'Note: The "/PowerShell" virtual directory URLs are not changed unless explicitly selected.' -ForegroundColor DarkYellow
    Write-Host 'Microsoft recommends modifying this virtual directory only at the request of Microsoft Customer Service and Support.' -ForegroundColor DarkYellow
    Write-Host ''
    $style = Read-InteractiveMenuChoice -Prompt 'Select configuration method' -ValidChoices @('1','2','3') -DefaultChoice '1'

    if ($style -eq '3') {
        Write-Host ''
        Write-Host 'Reading current Exchange values... This may take a while. Please wait.' -ForegroundColor DarkCyan
        $states = @(
            foreach ($target in $Targets) {
                Get-InteractiveServerState -Server $target
            }
        )
        return Read-InteractiveAdvancedSpec -States $states
    }

    $spec = New-InteractiveSpec
    $spec.Style = if ($style -eq '1') { 'Common namespace' } else { 'Select components' }

    if ($style -eq '1') {
        $namespaces = Read-InteractiveCommonNamespaces -IncludeCommonNamespaces:$true -IncludeScp:$true
        $selected = @(Get-InteractiveComponentMenu | Where-Object { $_.Key -ne 'PowerShell' })
        $powerShellProtocols = $null

        Show-InteractivePowerShellSafetyNote
        if (Read-InteractiveYesNo -Prompt 'Include "/PowerShell" virtual directory URLs?' -Default:$false) {
            $selected += @(Get-InteractiveComponentMenu | Where-Object { $_.Key -eq 'PowerShell' })
            $powerShellProtocols = Read-InteractivePowerShellProtocols
        }

        if ($null -ne $powerShellProtocols) {
            Add-InteractiveCommonActions `
                -Spec $spec `
                -Namespaces $namespaces `
                -SelectedComponents $selected `
                -IncludeScp:$true `
                -PowerShellInternalScheme $powerShellProtocols.Internal `
                -PowerShellExternalScheme $powerShellProtocols.External
        }
        else {
            Add-InteractiveCommonActions -Spec $spec -Namespaces $namespaces -SelectedComponents $selected -IncludeScp:$true
        }
    }
    else {
        # In "Selected components" mode the operator chooses the complete scope
        # from one numbered list, including Autodiscover SCP.
        $selected = @(Read-InteractiveComponentSelection)
        $includeScp = (@($selected | Where-Object { $_.Key -eq 'AutodiscoverSCP' }).Count -gt 0)
        $includeCommonNamespaces = (@($selected | Where-Object { $_.Key -ne 'AutodiscoverSCP' }).Count -gt 0)

        $namespaces = Read-InteractiveCommonNamespaces `
            -IncludeCommonNamespaces:$includeCommonNamespaces `
            -IncludeScp:$includeScp

        $powerShellSelected = (@($selected | Where-Object { $_.Key -eq 'PowerShell' }).Count -gt 0)
        if ($powerShellSelected) {
            Show-InteractivePowerShellSafetyNote
            $powerShellProtocols = Read-InteractivePowerShellProtocols
            Add-InteractiveCommonActions `
                -Spec $spec `
                -Namespaces $namespaces `
                -SelectedComponents $selected `
                -IncludeScp:$includeScp `
                -PowerShellInternalScheme $powerShellProtocols.Internal `
                -PowerShellExternalScheme $powerShellProtocols.External
        }
        else {
            Add-InteractiveCommonActions -Spec $spec -Namespaces $namespaces -SelectedComponents $selected -IncludeScp:$includeScp
        }
    }

    return $spec
}

function Resolve-InteractiveActionValue {
    param(
        [Parameter(Mandatory = $true)]$CurrentObject,
        [Parameter(Mandatory = $true)][string]$PropertyName,
        [Parameter(Mandatory = $true)]$Action
    )

    switch ($Action.Action) {
        'Set'   { return $Action.Value }
        'Clear' { return $null }
        'Keep'  {
            $property = $CurrentObject.PSObject.Properties[$PropertyName]
            if ($null -eq $property) {
                throw "Cannot keep current value because property '$PropertyName' is not available."
            }
            return $property.Value
        }
        default { throw "Unsupported Interactive action '$($Action.Action)'." }
    }
}

function New-InteractiveDesiredConfiguration {
    param(
        [Parameter(Mandatory = $true)]$TargetSnapshot,
        [Parameter(Mandatory = $true)]$Spec
    )

    $components = New-Object System.Collections.ArrayList

    foreach ($componentKey in @($Spec.ComponentActions.Keys)) {
        $actions = $Spec.ComponentActions[$componentKey]

        if ($componentKey -eq 'OutlookAny') {
            $currentObject = $TargetSnapshot.OutlookAny
            $values = [ordered]@{}
            foreach ($propertyName in $actions.Keys) {
                $values[$propertyName] = Resolve-InteractiveActionValue -CurrentObject $currentObject -PropertyName $propertyName -Action $actions[$propertyName]
            }

            [void]$components.Add([PSCustomObject]@{
                Key = 'OutlookAny'
                Name = 'Outlook Anywhere (RPC over HTTP)'
                Values = $values
                Policy = 'Apply'
                ReviewNote = $null
                DefaultAuthenticationMethod = $null
            })
            continue
        }

        $entry = $VirtualDirectoryMap | Where-Object { $_.Key -eq $componentKey } | Select-Object -First 1
        if ($null -eq $entry) { throw "Unsupported Interactive component '$componentKey'." }

        $currentObject = Get-SnapshotComponent -Snapshot $TargetSnapshot -Key $componentKey
        $values = [ordered]@{}
        foreach ($propertyName in $actions.Keys) {
            $values[$propertyName] = Resolve-InteractiveActionValue -CurrentObject $currentObject -PropertyName $propertyName -Action $actions[$propertyName]
        }

        $reviewNote = if ($componentKey -eq 'PowerShell') {
            'Only PowerShell InternalUrl/ExternalUrl are included. PowerShell authentication, RequireSSL, and Extended Protection remain Review Only.'
        }
        else { $null }

        [void]$components.Add([PSCustomObject]@{
            Key = $componentKey
            Name = $entry.Name
            Values = $values
            Policy = 'Apply'
            ReviewNote = $reviewNote
        })
    }

    if ($null -ne $Spec.ScpAction) {
        $scpValue = Resolve-InteractiveActionValue -CurrentObject $TargetSnapshot.ClientAccess -PropertyName 'AutoDiscoverServiceInternalUri' -Action $Spec.ScpAction
        [void]$components.Add([PSCustomObject]@{
            Key = 'ClientAccess'
            Name = 'Autodiscover SCP'
            Values = [ordered]@{ AutoDiscoverServiceInternalUri = $scpValue }
            Policy = 'Apply'
            ReviewNote = $null
        })
    }

    return [PSCustomObject]@{
        Target = $TargetSnapshot.Server
        Components = @($components)
    }
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

function Get-AlternateServiceAccountStatus {
    param([Parameter(Mandatory = $true)][string]$Server)

    try {
        $clientAccessWithASA = Get-ClientAccessService -Identity $Server -IncludeAlternateServiceAccountCredentialStatus -ErrorAction Stop
        $info = Get-AlternateServiceAccountInfo -ClientAccessService $clientAccessWithASA
        return [PSCustomObject]@{
            Account = $info.Account
            Status  = $info.Status
            Detail  = $null
        }
    }
    catch {
        # ASA is Review Only. An optional ASA-status read failure must not block
        # normal Client Access discovery or unrelated URL/namespace operations.
        Write-Warning "[$Server] ASA/Kerberos status could not be read. Continuing because ASA is Review Only."
        return [PSCustomObject]@{
            Account = 'Unknown'
            Status  = 'Unable to Read'
            Detail  = $_.Exception.Message
        }
    }
}

function Get-ExchangeClientAccessSnapshot {
    param([Parameter(Mandatory = $true)][string]$Server)

    $exchangeServer = Get-ExchangeServer -Identity $Server -ErrorAction Stop

    # Normal Client Access discovery must not depend on ASA registry/status data.
    $clientAccess = Get-ClientAccessService -Identity $exchangeServer.Name -ErrorAction Stop
    $asa = Get-AlternateServiceAccountStatus -Server $exchangeServer.Name

    $data = [ordered]@{
        Server         = $Server.Trim()
        ExchangeServer = $exchangeServer
        ClientAccess   = $clientAccess
        ASA            = $asa
        OutlookAny     = Get-DefaultVirtualDirectory -CommandName 'Get-OutlookAnywhere' -Server $exchangeServer.Name
    }

    foreach ($entry in $VirtualDirectoryMap) {
        $data[$entry.Key] = Get-DefaultVirtualDirectory -CommandName $entry.GetCmd -Server $exchangeServer.Name
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

    if ($null -eq $Value) { return '<empty>' }
    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        return (@($Value) | ForEach-Object { ConvertTo-ComparableValue -Value $_ } | Sort-Object) -join '|'
    }

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return '<empty>' }

    $uri = $null
    if ([Uri]::TryCreate($text, [UriKind]::Absolute, [ref]$uri) -and $uri.Scheme -in @('http','https')) {
        return $uri.AbsoluteUri.TrimEnd('/')
    }

    return $text
}

function Get-SnapshotComponent {
    param(
        [Parameter(Mandatory = $true)]$Snapshot,
        [Parameter(Mandatory = $true)][string]$Key
    )

    return $Snapshot.PSObject.Properties[$Key].Value
}

# ---------------------------------------------------------------------------
# Review report and versioned JSON backup
# ---------------------------------------------------------------------------
$script:BackupSchemaName = 'ExchangeURLManagerBackup'
$script:BackupSchemaVersion = 1

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
    $asaDetail = $Snapshot.ASA.PSObject.Properties['Detail']
    if ($asaDetail -and -not [string]::IsNullOrWhiteSpace([string]$asaDetail.Value)) {
        [void]$Lines.Add('Review  = ASA/Kerberos status could not be read. Review this setting manually if required.')
    }
    [void]$Lines.Add('')

    [void]$Lines.Add('[Outlook Anywhere (RPC over HTTP)]')
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


function Get-ReviewVisibleProperties {
    param(
        [Parameter(Mandatory = $true)]$Entry,
        [switch]$IncludeAuthentication
    )

    $names = New-Object System.Collections.ArrayList

    if ($Entry.Path) {
        [void]$names.Add('InternalUrl')
        [void]$names.Add('ExternalUrl')
    }

    if ($IncludeAuthentication) {
        foreach ($name in $Entry.AuthProperties) {
            [void]$names.Add($name)
        }
    }

    foreach ($name in $Entry.ReviewProperties) {
        [void]$names.Add($name)
    }

    return @($names | Select-Object -Unique)
}

function Get-ReviewDataForServer {
    param(
        [Parameter(Mandatory = $true)][string]$Server,
        [switch]$IncludeAuthentication
    )

    $groups = [ordered]@{}

    try {
        $exchangeServer = Get-ExchangeServer -Identity $Server -ErrorAction Stop
        $resolvedServer = ([string]$exchangeServer.Name).Trim()
    }
    catch {
        return [PSCustomObject]@{
            Server       = $Server
            Resolved     = $false
            ResolveError = $_.Exception.Message
            Groups       = $groups
        }
    }

    foreach ($entry in $VirtualDirectoryMap) {
        $propertyNames = @(Get-ReviewVisibleProperties -Entry $entry -IncludeAuthentication:$IncludeAuthentication)
        if ($propertyNames.Count -eq 0) { continue }

        $values = [ordered]@{}
        $errorText = $null
        try {
            $component = Get-DefaultVirtualDirectory -CommandName $entry.GetCmd -Server $resolvedServer
            foreach ($propertyName in $propertyNames) {
                $property = $component.PSObject.Properties[$propertyName]
                if ($property) {
                    $values[$propertyName] = ConvertTo-DisplayValue $property.Value
                }
                else {
                    $values[$propertyName] = '<Not Available>'
                }
            }
        }
        catch {
            $errorText = $_.Exception.Message
        }

        $reviewGroupName = if ($entry.Key -eq 'Autodiscover') { 'Autodiscover Virtual Directory' } else { $entry.Name }
        $groups[$reviewGroupName] = [PSCustomObject]@{
            Values = $values
            Error  = $errorText
        }
    }

    # Autodiscover SCP is a namespace/URL setting and is always visible.
    $scpValues = [ordered]@{}
    $scpError = $null
    try {
        $clientAccess = Get-ClientAccessService -Identity $resolvedServer -ErrorAction Stop
        $scpProperty = $clientAccess.PSObject.Properties['AutoDiscoverServiceInternalUri']
        $scpValues['AutoDiscoverServiceInternalUri'] = if ($scpProperty) {
            ConvertTo-DisplayValue $scpProperty.Value
        }
        else {
            '<Not Available>'
        }
    }
    catch {
        $scpError = $_.Exception.Message
    }
    $groups['Autodiscover SCP'] = [PSCustomObject]@{
        Values = $scpValues
        Error  = $scpError
    }

    # ASA/Kerberos is visibility/review only and remains visible by default.
    $asa = Get-AlternateServiceAccountStatus -Server $resolvedServer
    $asaValues = [ordered]@{
        Account = ConvertTo-DisplayValue $asa.Account
        Status  = ConvertTo-DisplayValue $asa.Status
    }
    $asaDetail = $asa.PSObject.Properties['Detail']
    if ($asaDetail -and -not [string]::IsNullOrWhiteSpace([string]$asaDetail.Value)) {
        $asaValues['Review'] = 'ASA/Kerberos status could not be read. Other Review checks continue.'
    }
    $groups['Alternate Service Account / Kerberos'] = [PSCustomObject]@{
        Values = $asaValues
        Error  = $null
    }

    # Outlook Anywhere has URL/hostname and SSL visibility by default.
    # Authentication is shown only with -IncludeAuthentication.
    $oaValues = [ordered]@{}
    $oaError = $null
    try {
        $outlookAnywhere = Get-DefaultVirtualDirectory -CommandName 'Get-OutlookAnywhere' -Server $resolvedServer

        $oaProperties = New-Object System.Collections.ArrayList
        foreach ($propertyName in @(
            'InternalHostname',
            'ExternalHostname',
            'InternalClientsRequireSsl',
            'ExternalClientsRequireSsl'
        )) {
            [void]$oaProperties.Add($propertyName)
        }

        if ($IncludeAuthentication) {
            foreach ($propertyName in @(
                'InternalClientAuthenticationMethod',
                'ExternalClientAuthenticationMethod',
                'IISAuthenticationMethods'
            )) {
                [void]$oaProperties.Add($propertyName)
            }
        }

        [void]$oaProperties.Add('SSLOffloading')

        foreach ($propertyName in @($oaProperties)) {
            $property = $outlookAnywhere.PSObject.Properties[$propertyName]
            if ($property) {
                $oaValues[$propertyName] = ConvertTo-DisplayValue $property.Value
            }
            else {
                $oaValues[$propertyName] = '<Not Available>'
            }
        }
    }
    catch {
        $oaError = $_.Exception.Message
    }

    $groups['Outlook Anywhere (RPC over HTTP)'] = [PSCustomObject]@{
        Values = $oaValues
        Error  = $oaError
    }

    return [PSCustomObject]@{
        Server       = $resolvedServer
        Resolved     = $true
        ResolveError = $null
        Groups       = $groups
    }
}

function Add-ReviewGroupSection {
    param(
        [Parameter(Mandatory = $true)][System.Collections.ArrayList]$Lines,
        [Parameter(Mandatory = $true)][string]$GroupName,
        [Parameter(Mandatory = $true)][array]$ReviewData
    )

    $available = @(
        foreach ($item in $ReviewData) {
            if ($item.Resolved -and $item.Groups.Contains($GroupName)) {
                $item
            }
        }
    )
    if ($available.Count -eq 0) { return }

    [void]$Lines.Add('')
    [void]$Lines.Add('##############################################################################')
    [void]$Lines.Add(("# {0}" -f $GroupName))
    [void]$Lines.Add('##############################################################################')
    [void]$Lines.Add('')

    $showServerHeading = ($ReviewData.Count -gt 1)

    foreach ($item in $ReviewData) {
        if (-not $item.Resolved -or -not $item.Groups.Contains($GroupName)) { continue }

        if ($showServerHeading) {
            [void]$Lines.Add(("[{0}]" -f $item.Server))
        }

        $group = $item.Groups[$GroupName]
        if (-not [string]::IsNullOrWhiteSpace([string]$group.Error)) {
            [void]$Lines.Add(("Unable to Read = {0}" -f $group.Error))
        }
        else {
            $names = @($group.Values.Keys)
            $width = 0
            foreach ($name in $names) {
                if ($name.Length -gt $width) { $width = $name.Length }
            }

            foreach ($name in $names) {
                [void]$Lines.Add(("{0} = {1}" -f $name.PadRight($width), $group.Values[$name]))
            }
        }

        [void]$Lines.Add('')
    }
}

function Get-GroupedReviewLines {
    param(
        [Parameter(Mandatory = $true)][array]$ReviewData,
        [switch]$IncludeAuthentication
    )

    $lines = New-Object System.Collections.ArrayList

    [void]$lines.Add('Current configuration')
    [void]$lines.Add('=====================')
    [void]$lines.Add('')

    $failed = @($ReviewData | Where-Object { -not $_.Resolved })
    if ($failed.Count -gt 0) {
        [void]$lines.Add('##############################################################################')
        [void]$lines.Add('# Server read failures')
        [void]$lines.Add('##############################################################################')
        [void]$lines.Add('')
        foreach ($item in $failed) {
            [void]$lines.Add(("[{0}]" -f $item.Server))
            [void]$lines.Add(("BLOCKER = Exchange server could not be resolved: {0}" -f $item.ResolveError))
            [void]$lines.Add('')
        }
    }

    $groupOrder = @(
        'OWA',
        'ECP',
        'EWS',
        'MAPI',
        'ActiveSync',
        'OAB',
        'Outlook Anywhere (RPC over HTTP)',
        'Autodiscover Virtual Directory',
        'Autodiscover SCP',
        'PowerShell',
        'Alternate Service Account / Kerberos'
    )

    foreach ($groupName in $groupOrder) {
        Add-ReviewGroupSection -Lines $lines -GroupName $groupName -ReviewData $ReviewData
    }

    return @($lines)
}


function Write-ReviewConsoleLine {
    param([AllowEmptyString()][string]$Line)

    if ([string]::IsNullOrEmpty($Line)) {
        Write-Host ''
        return
    }

    if ($Line -eq 'Current configuration' -or
        $Line -eq '=====================' -or
        $Line -match '^#{10,}$' -or
        $Line -match '^# .+$') {
        Write-Host $Line -ForegroundColor Cyan
        return
    }

    if ($Line -match '^\[[^\]]+\]$') {
        Write-Host $Line -ForegroundColor Yellow
        return
    }

    if ($Line -match '^BLOCKER\s*=') {
        Write-Host $Line -ForegroundColor Red
        return
    }

    if ($Line -match '^Unable to Read\s*=') {
        Write-Host $Line -ForegroundColor Red
        return
    }

    if ($Line -match '^(?<Label>.+?)\s+=\s+(?<Value>.*)$') {
        $label = $matches['Label']
        $value = $matches['Value']

        Write-Host $label -ForegroundColor Gray -NoNewline
        Write-Host ' = ' -ForegroundColor DarkGray -NoNewline

        if ($value -match '^<(Unable to Read|Not Available)>$') {
            Write-Host $value -ForegroundColor Red
        }
        elseif ($value -match '^<(null|empty)>$' -or $value -eq 'Not Configured') {
            Write-Host $value -ForegroundColor DarkYellow
        }
        else {
            Write-Host $value -ForegroundColor White
        }
        return
    }

    Write-Host $Line
}

function Get-ReviewLinesForServer {
    param([Parameter(Mandatory = $true)][string]$Server)

    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add(("=== SERVER: {0} ===" -f $Server))
    [void]$lines.Add('')

    try {
        $exchangeServer = Get-ExchangeServer -Identity $Server -ErrorAction Stop
        $resolvedServer = [string]$exchangeServer.Name
    }
    catch {
        [void]$lines.Add(("BLOCKER = Exchange server could not be resolved: {0}" -f $_.Exception.Message))
        [void]$lines.Add('')
        return @($lines)
    }

    foreach ($entry in $VirtualDirectoryMap) {
        [void]$lines.Add(("[{0}]" -f $entry.Name))
        try {
            $component = Get-DefaultVirtualDirectory -CommandName $entry.GetCmd -Server $resolvedServer
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
                    [void]$lines.Add(("{0} = {1}" -f $propertyName, (ConvertTo-DisplayValue $property.Value)))
                }
            }
        }
        catch {
            [void]$lines.Add(("Unable to Read = {0}" -f $_.Exception.Message))
        }
        [void]$lines.Add('')
    }

    [void]$lines.Add('[Autodiscover SCP]')
    try {
        $clientAccess = Get-ClientAccessService -Identity $resolvedServer -ErrorAction Stop
        $scpProperty = $clientAccess.PSObject.Properties['AutoDiscoverServiceInternalUri']
        if ($scpProperty) {
            [void]$lines.Add(("AutoDiscoverServiceInternalUri = {0}" -f (ConvertTo-DisplayValue $scpProperty.Value)))
        }
        else {
            [void]$lines.Add('AutoDiscoverServiceInternalUri = <Not Available>')
        }
    }
    catch {
        [void]$lines.Add(("Unable to Read = {0}" -f $_.Exception.Message))
    }
    [void]$lines.Add('')

    [void]$lines.Add('[Alternate Service Account / Kerberos]')
    $asa = Get-AlternateServiceAccountStatus -Server $resolvedServer
    [void]$lines.Add(("Account = {0}" -f (ConvertTo-DisplayValue $asa.Account)))
    [void]$lines.Add(("Status  = {0}" -f (ConvertTo-DisplayValue $asa.Status)))
    if ($asa.Status -eq 'Unable to Read') {
        [void]$lines.Add('Review  = ASA/Kerberos status could not be read. Other Review checks continue.')
    }
    [void]$lines.Add('')

    [void]$lines.Add('[Outlook Anywhere (RPC over HTTP)]')
    try {
        $outlookAnywhere = Get-DefaultVirtualDirectory -CommandName 'Get-OutlookAnywhere' -Server $resolvedServer
        foreach ($propertyName in @(
            'InternalHostname','ExternalHostname','InternalClientsRequireSsl','ExternalClientsRequireSsl',
            'InternalClientAuthenticationMethod','ExternalClientAuthenticationMethod','IISAuthenticationMethods','SSLOffloading'
        )) {
            $property = $outlookAnywhere.PSObject.Properties[$propertyName]
            if ($property) {
                [void]$lines.Add(("{0} = {1}" -f $propertyName, (ConvertTo-DisplayValue $property.Value)))
            }
        }
    }
    catch {
        [void]$lines.Add(("Unable to Read = {0}" -f $_.Exception.Message))
    }
    [void]$lines.Add('')

    return @($lines)
}

function ConvertTo-BackupValue {
    param($Value)

    if ($null -eq $Value) { return $null }
    if ($Value -is [bool] -or
        $Value -is [byte] -or
        $Value -is [int16] -or
        $Value -is [int32] -or
        $Value -is [int64] -or
        $Value -is [decimal] -or
        $Value -is [double] -or
        $Value -is [single]) {
        return $Value
    }

    if ($Value -is [Uri]) { return [string]$Value.AbsoluteUri }

    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $items = @($Value | ForEach-Object { ConvertTo-BackupValue -Value $_ })
        # PowerShell function output normally unwraps a one-item collection.
        # Preserve the array so JSON multivalue properties keep a stable type.
        return ,$items
    }

    return [string]$Value
}

function New-ExchangeURLManagerBackupDocument {
    param(
        [Parameter(Mandatory = $true)]$Snapshot,
        [Parameter(Mandatory = $true)][ValidateSet('Manual','PreChange')][string]$BackupType,
        [AllowNull()][string]$Context
    )

    $virtualDirectories = New-Object System.Collections.ArrayList
    foreach ($entry in $VirtualDirectoryMap) {
        $component = Get-SnapshotComponent -Snapshot $Snapshot -Key $entry.Key
        $urls = [ordered]@{}
        if ($entry.Path) {
            foreach ($name in @('InternalUrl','ExternalUrl')) {
                $property = $component.PSObject.Properties[$name]
                if ($property) { $urls[$name] = ConvertTo-BackupValue -Value $property.Value }
            }
        }

        $authentication = [ordered]@{}
        foreach ($name in $entry.AuthProperties) {
            $property = $component.PSObject.Properties[$name]
            if ($property) { $authentication[$name] = ConvertTo-BackupValue -Value $property.Value }
        }

        $reviewOnly = [ordered]@{}
        foreach ($name in $entry.ReviewProperties) {
            $property = $component.PSObject.Properties[$name]
            if ($property) { $reviewOnly[$name] = ConvertTo-BackupValue -Value $property.Value }
        }

        $identityProperty = $component.PSObject.Properties['Identity']
        [void]$virtualDirectories.Add([PSCustomObject][ordered]@{
            Key            = $entry.Key
            Name           = $entry.Name
            Identity       = if ($identityProperty) { ConvertTo-BackupValue -Value $identityProperty.Value } else { $null }
            Urls           = [PSCustomObject]$urls
            Authentication = [PSCustomObject]$authentication
            ReviewOnly     = [PSCustomObject]$reviewOnly
        })
    }

    $serverObject = $Snapshot.ExchangeServer
    $serverNameProperty = $serverObject.PSObject.Properties['Name']
    if ($null -eq $serverNameProperty -or [string]::IsNullOrWhiteSpace([string]$serverNameProperty.Value)) {
        throw 'Exchange server snapshot is missing Name.'
    }
    $serverFqdnProperty = $serverObject.PSObject.Properties['Fqdn']
    $serverVersionProperty = $serverObject.PSObject.Properties['AdminDisplayVersion']
    $serverEditionProperty = $serverObject.PSObject.Properties['Edition']
    $serverRoleProperty = $serverObject.PSObject.Properties['ServerRole']
    $serverMetadata = [ordered]@{
        Name                = [string]$serverNameProperty.Value
        Fqdn                = if ($serverFqdnProperty) { ConvertTo-BackupValue -Value $serverFqdnProperty.Value } else { $null }
        AdminDisplayVersion = if ($serverVersionProperty) { ConvertTo-BackupValue -Value $serverVersionProperty.Value } else { $null }
        Edition             = if ($serverEditionProperty) { ConvertTo-BackupValue -Value $serverEditionProperty.Value } else { $null }
        ServerRole          = if ($serverRoleProperty) { ConvertTo-BackupValue -Value $serverRoleProperty.Value } else { $null }
    }

    $scpProperty = $Snapshot.ClientAccess.PSObject.Properties['AutoDiscoverServiceInternalUri']
    if ($null -eq $scpProperty) { throw 'Client Access snapshot is missing AutoDiscoverServiceInternalUri.' }
    $clientAccess = [ordered]@{
        AutoDiscoverServiceInternalUri = ConvertTo-BackupValue -Value $scpProperty.Value
    }

    $asa = [ordered]@{
        Account = ConvertTo-BackupValue -Value $Snapshot.ASA.Account
        Status  = ConvertTo-BackupValue -Value $Snapshot.ASA.Status
    }

    $outlookAnywhere = [ordered]@{}
    foreach ($name in @(
        'InternalHostname','ExternalHostname','InternalClientsRequireSsl','ExternalClientsRequireSsl',
        'InternalClientAuthenticationMethod','ExternalClientAuthenticationMethod','IISAuthenticationMethods','SSLOffloading'
    )) {
        $property = $Snapshot.OutlookAny.PSObject.Properties[$name]
        if ($property) { $outlookAnywhere[$name] = ConvertTo-BackupValue -Value $property.Value }
    }

    return [PSCustomObject][ordered]@{
        Schema = [PSCustomObject][ordered]@{
            Name    = $script:BackupSchemaName
            Version = $script:BackupSchemaVersion
        }
        Backup = [PSCustomObject][ordered]@{
            Type          = $BackupType
            CreatedUtc    = (Get-Date).ToUniversalTime().ToString('o')
            ScriptName    = $script:PublicScriptName
            ScriptVersion = $script:PublicScriptVersion
            Context       = $Context
        }
        Server = [PSCustomObject]$serverMetadata
        Configuration = [PSCustomObject][ordered]@{
            VirtualDirectories = @($virtualDirectories)
            ClientAccess       = [PSCustomObject]$clientAccess
            ASA                = [PSCustomObject]$asa
            OutlookAnywhere    = [PSCustomObject]$outlookAnywhere
        }
    }
}

function Get-DefaultBackupDirectory {
    $root = Join-Path -Path $PSScriptRoot -ChildPath 'ConfigBackups'
    if (-not (Test-Path -LiteralPath $root -PathType Container)) {
        New-Item -Path $root -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }
    return [System.IO.Path]::GetFullPath($root)
}

function Get-UniqueBackupFilePath {
    param(
        [Parameter(Mandatory = $true)][string]$Directory,
        [Parameter(Mandatory = $true)][string]$Server,
        [Parameter(Mandatory = $true)][ValidateSet('Manual','PreChange')][string]$BackupType
    )

    if (-not (Test-Path -LiteralPath $Directory -PathType Container)) {
        New-Item -Path $Directory -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }

    $safeServer = ($Server -replace '[^A-Za-z0-9_.-]', '_')
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $candidate = Join-Path -Path $Directory -ChildPath ("ExchangeURLManager-{0}-{1}-{2}.json" -f $safeServer, $BackupType, $timestamp)
    $suffix = 2
    while (Test-Path -LiteralPath $candidate) {
        $candidate = Join-Path -Path $Directory -ChildPath ("ExchangeURLManager-{0}-{1}-{2}-{3}.json" -f $safeServer, $BackupType, $timestamp, $suffix)
        $suffix++
    }
    return [System.IO.Path]::GetFullPath($candidate)
}

function Resolve-ManualBackupPath {
    param(
        [Parameter(Mandatory = $true)][string]$Server,
        [int]$ServerCount,
        [AllowNull()][string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return Get-UniqueBackupFilePath -Directory (Get-DefaultBackupDirectory) -Server $Server -BackupType 'Manual'
    }

    $candidate = $Path.Trim()
    if (-not [System.IO.Path]::IsPathRooted($candidate)) {
        $candidate = Join-Path -Path (Get-Location).Path -ChildPath $candidate
    }

    if (Test-Path -LiteralPath $candidate -PathType Container) {
        return Get-UniqueBackupFilePath -Directory ([System.IO.Path]::GetFullPath($candidate)) -Server $Server -BackupType 'Manual'
    }

    if (Test-Path -LiteralPath $candidate -PathType Leaf) {
        throw "BackupPath points to an existing file: $candidate"
    }

    $extension = [System.IO.Path]::GetExtension($candidate)
    if (-not [string]::IsNullOrWhiteSpace($extension) -and $extension -ne '.json') {
        throw 'BackupPath must be a directory or a .json filename.'
    }

    if ($extension -eq '.json') {
        if ($ServerCount -ne 1) {
            throw 'BackupPath can be a .json filename only when backing up one server.'
        }

        $parent = Split-Path -Parent $candidate
        if ([string]::IsNullOrWhiteSpace($parent)) { $parent = (Get-Location).Path }
        if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
            New-Item -Path $parent -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
        $candidate = [System.IO.Path]::GetFullPath($candidate)
        if (Test-Path -LiteralPath $candidate) { throw "Backup file already exists: $candidate" }
        return $candidate
    }

    if (-not (Test-Path -LiteralPath $candidate -PathType Container)) {
        New-Item -Path $candidate -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }

    return Get-UniqueBackupFilePath -Directory ([System.IO.Path]::GetFullPath($candidate)) -Server $Server -BackupType 'Manual'
}


function Format-HumanReadableBackupValue {
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return '<null>' }

    if ($Value -is [System.Array] -or $Value -is [System.Collections.IList]) {
        $items = @($Value | ForEach-Object {
            if ($null -eq $_) { '<null>' }
            elseif ([string]::IsNullOrEmpty([string]$_)) { '<empty>' }
            else { [string]$_ }
        })
        if ($items.Count -eq 0) { return '<empty>' }
        return ($items -join ', ')
    }

    $stringValue = [string]$Value
    if ([string]::IsNullOrEmpty($stringValue)) { return '<empty>' }
    return $stringValue
}

function Convert-BackupDocumentToHumanReadableText {
    param([Parameter(Mandatory = $true)]$Document)

    $lines = New-Object System.Collections.ArrayList

    [void]$lines.Add('ExchangeURLManager Configuration Backup')
    [void]$lines.Add('=======================================')
    [void]$lines.Add('')
    [void]$lines.Add(('Backup Type    : {0}' -f (Format-HumanReadableBackupValue $Document.Backup.Type)))
    [void]$lines.Add(('Created UTC    : {0}' -f (Format-HumanReadableBackupValue $Document.Backup.CreatedUtc)))
    [void]$lines.Add(('Context        : {0}' -f (Format-HumanReadableBackupValue $Document.Backup.Context)))
    [void]$lines.Add(('Script         : {0} {1}' -f (Format-HumanReadableBackupValue $Document.Backup.ScriptName), (Format-HumanReadableBackupValue $Document.Backup.ScriptVersion)))
    [void]$lines.Add('')
    [void]$lines.Add(('Server         : {0}' -f (Format-HumanReadableBackupValue $Document.Server.Name)))
    [void]$lines.Add(('FQDN           : {0}' -f (Format-HumanReadableBackupValue $Document.Server.Fqdn)))
    [void]$lines.Add(('Exchange       : {0}' -f (Format-HumanReadableBackupValue $Document.Server.AdminDisplayVersion)))
    [void]$lines.Add(('Edition        : {0}' -f (Format-HumanReadableBackupValue $Document.Server.Edition)))
    [void]$lines.Add(('Server Role    : {0}' -f (Format-HumanReadableBackupValue $Document.Server.ServerRole)))
    [void]$lines.Add('')

    foreach ($component in @($Document.Configuration.VirtualDirectories)) {
        [void]$lines.Add(('[{0}]' -f $component.Name))
        if ($component.Identity) {
            [void]$lines.Add(('Identity = {0}' -f (Format-HumanReadableBackupValue $component.Identity)))
        }

        foreach ($property in @($component.Urls.PSObject.Properties)) {
            [void]$lines.Add(('{0} = {1}' -f $property.Name, (Format-HumanReadableBackupValue $property.Value)))
        }

        foreach ($property in @($component.Authentication.PSObject.Properties)) {
            $authSuffix = if ($component.Key -eq 'PowerShell') { ' [Review Only]' } else { '' }
            [void]$lines.Add(('{0} = {1}{2}' -f $property.Name, (Format-HumanReadableBackupValue $property.Value), $authSuffix))
        }

        foreach ($property in @($component.ReviewOnly.PSObject.Properties)) {
            [void]$lines.Add(('{0} = {1} [Review Only]' -f $property.Name, (Format-HumanReadableBackupValue $property.Value)))
        }

        [void]$lines.Add('')
    }

    [void]$lines.Add('[Autodiscover SCP]')
    [void]$lines.Add(('AutoDiscoverServiceInternalUri = {0}' -f (Format-HumanReadableBackupValue $Document.Configuration.ClientAccess.AutoDiscoverServiceInternalUri)))
    [void]$lines.Add('')

    [void]$lines.Add('[Alternate Service Account / Kerberos]')
    [void]$lines.Add(('Account = {0}' -f (Format-HumanReadableBackupValue $Document.Configuration.ASA.Account)))
    [void]$lines.Add(('Status  = {0}' -f (Format-HumanReadableBackupValue $Document.Configuration.ASA.Status)))
    [void]$lines.Add('')

    [void]$lines.Add('[Outlook Anywhere (RPC over HTTP)]')
    foreach ($property in @($Document.Configuration.OutlookAnywhere.PSObject.Properties)) {
        $suffix = if ($property.Name -eq 'SSLOffloading') { ' [Review Only]' } else { '' }
        [void]$lines.Add(('{0} = {1}{2}' -f $property.Name, (Format-HumanReadableBackupValue $property.Value), $suffix))
    }

    [void]$lines.Add('')
    [void]$lines.Add('Restore source: JSON companion only. This TXT file is human-readable reference output.')
    [void]$lines.Add('')

    return ($lines -join [Environment]::NewLine)
}

function Export-HumanReadableBackupCompanion {
    param(
        [Parameter(Mandatory = $true)]$Document,
        [Parameter(Mandatory = $true)][string]$JsonPath
    )

    $txtPath = [System.IO.Path]::ChangeExtension($JsonPath, '.txt')

    if (Test-Path -LiteralPath $txtPath) {
        throw "Human-readable backup file already exists: $txtPath"
    }

    $content = Convert-BackupDocumentToHumanReadableText -Document $Document
    Set-Content -LiteralPath $txtPath -Value $content -Encoding UTF8 -ErrorAction Stop
    return $txtPath
}

function Export-ExchangeURLManagerBackup {
    param(
        [Parameter(Mandatory = $true)]$Snapshot,
        [Parameter(Mandatory = $true)][ValidateSet('Manual','PreChange')][string]$BackupType,
        [AllowNull()][string]$Path,
        [AllowNull()][string]$Context,
        [int]$ServerCount = 1
    )

    $destination = if ($BackupType -eq 'PreChange') {
        Get-UniqueBackupFilePath -Directory (Get-DefaultBackupDirectory) -Server ([string]$Snapshot.ExchangeServer.Name) -BackupType 'PreChange'
    }
    else {
        Resolve-ManualBackupPath -Server ([string]$Snapshot.ExchangeServer.Name) -ServerCount $ServerCount -Path $Path
    }

    $document = New-ExchangeURLManagerBackupDocument -Snapshot $Snapshot -BackupType $BackupType -Context $Context
    $json = $document | ConvertTo-Json -Depth 12
    Set-Content -LiteralPath $destination -Value $json -Encoding UTF8 -ErrorAction Stop

    # JSON is the canonical Restore artifact and remains fail-closed.
    # The TXT companion is human-readable only. TXT creation failure must not
    # invalidate a successfully created JSON backup or block Apply.
    try {
        $txtPath = Export-HumanReadableBackupCompanion -Document $document -JsonPath $destination
        Write-Host ("Human-readable backup created: {0}" -f $txtPath) -ForegroundColor Green
    }
    catch {
        Write-Warning ("JSON backup was created successfully, but the human-readable TXT companion could not be created: {0}" -f $_.Exception.Message)
    }

    return $destination
}

function Export-PreChangeBackups {
    param(
        [Parameter(Mandatory = $true)][string[]]$Targets,
        [AllowNull()][string]$Context
    )

    Write-Host ''
    Write-Host 'Creating pre-change backup(s) before Apply...' -ForegroundColor Cyan
    Write-Host 'JSON is the Restore backup; a human-readable TXT companion is also created.' -ForegroundColor DarkCyan
    Write-Host 'Apply will start only after all required JSON pre-change backups are created successfully.' -ForegroundColor DarkCyan

    $paths = New-Object System.Collections.ArrayList
    foreach ($target in $Targets) {
        Write-Host ("[{0}] Reading current configuration for backup..." -f $target) -ForegroundColor DarkCyan
        $freshSnapshot = Get-ExchangeClientAccessSnapshot -Server $target

        Write-Host ("[{0}] Writing pre-change JSON backup and TXT companion..." -f $target) -ForegroundColor DarkCyan
        $path = Export-ExchangeURLManagerBackup -Snapshot $freshSnapshot -BackupType 'PreChange' -Context $Context

        Write-Host ("[{0}] Pre-change backup created: {1}" -f $target, $path) -ForegroundColor Green
        [void]$paths.Add($path)
    }

    Write-Host ("Pre-change backup completed for {0} server(s)." -f $paths.Count) -ForegroundColor Green
    return @($paths)
}

function Import-ExchangeURLManagerBackup {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Backup file was not found: $Path"
    }

    try {
        $document = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        throw "Backup file is not valid JSON: $($_.Exception.Message)"
    }

    if ($null -eq $document -or $document -is [System.Array]) {
        throw 'Backup JSON root must be one object.'
    }

    $schemaProperty = $document.PSObject.Properties['Schema']
    if ($null -eq $schemaProperty -or $null -eq $schemaProperty.Value) {
        throw 'Backup is missing Schema.'
    }

    $schemaNameProperty = $document.Schema.PSObject.Properties['Name']
    $schemaVersionProperty = $document.Schema.PSObject.Properties['Version']
    if ($null -eq $schemaNameProperty -or $null -eq $schemaVersionProperty) {
        throw 'Backup schema metadata is incomplete.'
    }

    $schemaVersion = 0
    if (-not [int]::TryParse([string]$schemaVersionProperty.Value, [ref]$schemaVersion)) {
        throw 'Backup schema version is invalid.'
    }

    if ([string]$schemaNameProperty.Value -ne $script:BackupSchemaName -or
        $schemaVersion -ne $script:BackupSchemaVersion) {
        throw "Unsupported backup schema. Expected $($script:BackupSchemaName) version $($script:BackupSchemaVersion)."
    }

    $backupMetadataProperty = $document.PSObject.Properties['Backup']
    if ($null -eq $backupMetadataProperty -or $null -eq $backupMetadataProperty.Value) {
        throw 'Backup is missing Backup metadata.'
    }
    foreach ($metadataName in @('Type','CreatedUtc','ScriptName','ScriptVersion')) {
        $metadataProperty = $document.Backup.PSObject.Properties[$metadataName]
        if ($null -eq $metadataProperty -or [string]::IsNullOrWhiteSpace([string]$metadataProperty.Value)) {
            throw "Backup metadata is missing $metadataName."
        }
    }

    $backupType = [string]$document.Backup.Type
    if ($backupType -notin @('Manual','PreChange')) {
        throw "Backup metadata Type is invalid: $backupType"
    }

    $createdUtc = [DateTimeOffset]::MinValue
    if (-not [DateTimeOffset]::TryParse([string]$document.Backup.CreatedUtc, [ref]$createdUtc)) {
        throw 'Backup metadata CreatedUtc is invalid.'
    }

    $scriptNameProperty = $document.Backup.PSObject.Properties['ScriptName']
    if ($null -eq $scriptNameProperty -or [string]$scriptNameProperty.Value -ne $script:PublicScriptName) {
        throw "Backup was not created by $($script:PublicScriptName)."
    }

    $serverProperty = $document.PSObject.Properties['Server']
    if ($null -eq $serverProperty -or $null -eq $serverProperty.Value) {
        throw 'Backup is missing Server metadata.'
    }
    $serverNameProperty = $document.Server.PSObject.Properties['Name']
    if ($null -eq $serverNameProperty -or [string]::IsNullOrWhiteSpace([string]$serverNameProperty.Value)) {
        throw 'Backup is missing Server.Name.'
    }
    $backupServerName = [string]$serverNameProperty.Value
    if ($backupServerName -notmatch '^[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?$') {
        throw "Backup Server.Name is invalid: $backupServerName"
    }

    $serverFqdnProperty = $document.Server.PSObject.Properties['Fqdn']
    if ($null -eq $serverFqdnProperty) {
        throw 'Backup Server metadata is missing Fqdn.'
    }
    $backupFqdn = [string]$serverFqdnProperty.Value
    if (-not [string]::IsNullOrWhiteSpace($backupFqdn) -and
        [Uri]::CheckHostName($backupFqdn) -ne [UriHostNameType]::Dns) {
        throw "Backup Server.Fqdn is invalid: $backupFqdn"
    }

    $configurationProperty = $document.PSObject.Properties['Configuration']
    if ($null -eq $configurationProperty -or $null -eq $configurationProperty.Value) {
        throw 'Backup is missing Configuration data.'
    }

    foreach ($requiredName in @('VirtualDirectories','ClientAccess','OutlookAnywhere','ASA')) {
        $property = $document.Configuration.PSObject.Properties[$requiredName]
        if ($null -eq $property -or $null -eq $property.Value) {
            throw "Backup is missing Configuration.$requiredName."
        }
    }

    $requiredKeys = @('Owa','Ecp','Ews','Mapi','ActiveSync','Oab','PowerShell','Autodiscover')
    $items = @($document.Configuration.VirtualDirectories)
    if ($items.Count -ne $requiredKeys.Count) {
        throw "Backup must contain exactly $($requiredKeys.Count) supported virtual-directory entries."
    }
    foreach ($item in $items) {
        $keyProperty = $item.PSObject.Properties['Key']
        if ($null -eq $keyProperty -or [string]::IsNullOrWhiteSpace([string]$keyProperty.Value) -or
            [string]$keyProperty.Value -notin $requiredKeys) {
            throw 'Backup contains an invalid or unsupported virtual-directory key.'
        }
    }
    foreach ($key in $requiredKeys) {
        $matches = @($items | Where-Object { [string]$_.Key -eq $key })
        if ($matches.Count -ne 1) {
            throw "Backup must contain exactly one virtual-directory entry for '$key'."
        }

        $item = $matches[0]
        foreach ($containerName in @('Urls','Authentication','ReviewOnly')) {
            $containerProperty = $item.PSObject.Properties[$containerName]
            if ($null -eq $containerProperty -or $null -eq $containerProperty.Value) {
                throw "Backup virtual-directory entry '$key' is missing $containerName."
            }
        }

        $mapEntry = $VirtualDirectoryMap | Where-Object { $_.Key -eq $key } | Select-Object -First 1
        if ($mapEntry.Path) {
            foreach ($urlPropertyName in @('InternalUrl','ExternalUrl')) {
                if ($null -eq $item.Urls.PSObject.Properties[$urlPropertyName]) {
                    throw "Backup virtual-directory entry '$key' is missing $urlPropertyName."
                }
            }
        }
    }

    if ($null -eq $document.Configuration.ClientAccess.PSObject.Properties['AutoDiscoverServiceInternalUri']) {
        throw 'Backup is missing Configuration.ClientAccess.AutoDiscoverServiceInternalUri.'
    }

    foreach ($propertyName in @('Account','Status')) {
        if ($null -eq $document.Configuration.ASA.PSObject.Properties[$propertyName]) {
            throw "Backup is missing Configuration.ASA.$propertyName."
        }
    }

    foreach ($propertyName in @('InternalHostname','ExternalHostname')) {
        if ($null -eq $document.Configuration.OutlookAnywhere.PSObject.Properties[$propertyName]) {
            throw "Backup is missing Configuration.OutlookAnywhere.$propertyName."
        }
    }

    return $document
}

function Get-BackupVirtualDirectory {
    param(
        [Parameter(Mandatory = $true)]$BackupDocument,
        [Parameter(Mandatory = $true)][string]$Key
    )

    $matches = @($BackupDocument.Configuration.VirtualDirectories | Where-Object { [string]$_.Key -eq $Key })
    if ($matches.Count -ne 1) { throw "Backup virtual-directory entry '$Key' is missing or duplicated." }
    return $matches[0]
}

function Assert-SameServerRestore {
    param([Parameter(Mandatory = $true)]$BackupDocument)

    $backupServer = [string]$BackupDocument.Server.Name
    $liveServers = @(Get-ExchangeServer -Identity $backupServer -ErrorAction Stop)
    if ($liveServers.Count -ne 1) {
        throw "Backup server '$backupServer' did not resolve to exactly one Exchange server."
    }
    $liveServer = $liveServers[0]
    $liveNameProperty = $liveServer.PSObject.Properties['Name']
    if ($null -eq $liveNameProperty -or [string]::IsNullOrWhiteSpace([string]$liveNameProperty.Value)) {
        throw "Live Exchange server '$backupServer' did not return Name."
    }

    if (-not ([string]$liveNameProperty.Value).Equals($backupServer, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Backup server '$backupServer' does not match live Exchange server '$($liveNameProperty.Value)'. Cross-server Restore is not supported."
    }

    $backupFqdnProperty = $BackupDocument.Server.PSObject.Properties['Fqdn']
    $liveFqdnProperty = $liveServer.PSObject.Properties['Fqdn']
    if ($backupFqdnProperty -and $liveFqdnProperty -and
        -not [string]::IsNullOrWhiteSpace([string]$backupFqdnProperty.Value) -and
        -not ([string]$liveFqdnProperty.Value).Equals([string]$backupFqdnProperty.Value, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Backup FQDN '$($backupFqdnProperty.Value)' does not match live Exchange server FQDN '$($liveFqdnProperty.Value)'. Cross-server Restore is not supported."
    }

    return [string]$liveNameProperty.Value
}

# ---------------------------------------------------------------------------
# Build Desired Configuration
# ---------------------------------------------------------------------------
function New-DesiredConfiguration {
    param(
        [Parameter(Mandatory = $true)]$TargetSnapshot,
        $SourceSnapshot,
        [string]$InternalNamespace,
        [AllowNull()][string]$ExternalNamespace,
        [string]$AutodiscoverNamespace,
        [bool]$CopyAuthentication = $false,
        [bool]$IncludePowerShellUrls = $false,
        [bool]$IncludeOutlookAnywhereSslRequirements = $false,
        [bool]$CompareOutlookAnywhereSslRequirements = $false,
        [switch]$ClearExternalUrls,
        [Nullable[bool]]$OutlookAnywhereInternalClientsRequireSsl,
        [Nullable[bool]]$OutlookAnywhereExternalClientsRequireSsl,
        [string]$OutlookAnywhereDefaultAuthenticationMethod
    )

    $components = New-Object System.Collections.ArrayList

    foreach ($entry in $VirtualDirectoryMap) {
        $sourceObject = if ($SourceSnapshot) { Get-SnapshotComponent -Snapshot $SourceSnapshot -Key $entry.Key } else { $null }

        # URL values are handled separately from authentication so PowerShell URL
        # opt-in never causes PowerShell authentication/RequireSSL to be applied.
        if ($entry.Path) {
            $urlValues = [ordered]@{}
            if ($SourceSnapshot) {
                $urlValues['InternalUrl'] = $sourceObject.InternalUrl
                $urlValues['ExternalUrl'] = $sourceObject.ExternalUrl
            }
            else {
                $urlValues['InternalUrl'] = "https://$InternalNamespace$($entry.Path)"
                $urlValues['ExternalUrl'] = if ($ClearExternalUrls) { $null } else { "https://$ExternalNamespace$($entry.Path)" }
            }

            if ($entry.Key -eq 'PowerShell') {
                if ($IncludePowerShellUrls) {
                    $urlPolicy = 'Apply'
                    $urlReview = $null
                }
                else {
                    $urlPolicy = 'ReviewOnly'
                    $urlReview = 'PowerShell URL management requires explicit -IncludePowerShellUrls.'
                }
            }
            else {
                $urlPolicy = 'Apply'
                $urlReview = $null
            }

            [void]$components.Add([PSCustomObject]@{
                Key = $entry.Key
                Name = $entry.Name
                Values = $urlValues
                Policy = $urlPolicy
                ReviewNote = $urlReview
            })
        }

        if ($SourceSnapshot -and $CopyAuthentication -and $entry.AuthProperties.Count -gt 0) {
            $authValues = [ordered]@{}
            foreach ($propertyName in $entry.AuthProperties) {
                $property = $sourceObject.PSObject.Properties[$propertyName]
                if ($property) { $authValues[$propertyName] = $property.Value }
            }

            if ($authValues.Count -gt 0) {
                $authPolicy = if ($entry.Key -eq 'PowerShell') { 'ReviewOnly' } else { 'Apply' }
                $authReview = if ($entry.Key -eq 'PowerShell') {
                    'PowerShell authentication is Review Only and is not cloned automatically.'
                }
                else { $null }

                [void]$components.Add([PSCustomObject]@{
                    Key = $entry.Key
                    Name = "$($entry.Name) Authentication"
                    Values = $authValues
                    Policy = $authPolicy
                    ReviewNote = $authReview
                })
            }
        }

        # Migration-sensitive properties remain Review Only.
        if ($SourceSnapshot -and $entry.ReviewProperties.Count -gt 0) {
            $reviewValues = [ordered]@{}
            foreach ($propertyName in $entry.ReviewProperties) {
                $property = $sourceObject.PSObject.Properties[$propertyName]
                if ($property) { $reviewValues[$propertyName] = $property.Value }
            }

            if ($reviewValues.Count -gt 0) {
                [void]$components.Add([PSCustomObject]@{
                    Key = $entry.Key
                    Name = "$($entry.Name) (Review Only)"
                    Values = $reviewValues
                    Policy = 'ReviewOnly'
                    ReviewNote = 'Migration-sensitive setting. Review the target design before changing it.'
                })
            }
        }
    }

    $scpValue = if ($SourceSnapshot) {
        $SourceSnapshot.ClientAccess.AutoDiscoverServiceInternalUri
    }
    else {
        "https://$AutodiscoverNamespace/Autodiscover/Autodiscover.xml"
    }
    [void]$components.Add([PSCustomObject]@{
        Key = 'ClientAccess'
        Name = 'Autodiscover SCP'
        Values = [ordered]@{ AutoDiscoverServiceInternalUri = $scpValue }
        Policy = 'Apply'
        ReviewNote = $null
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
    $oaSslReviewValues = [ordered]@{}
    $oaDefaultAuth = $null

    if ($SourceSnapshot) {
        foreach ($propertyName in @('InternalHostname','ExternalHostname')) {
            $property = $SourceSnapshot.OutlookAny.PSObject.Properties[$propertyName]
            if ($property) { $oaValues[$propertyName] = $property.Value }
        }

        if ($IncludeOutlookAnywhereSslRequirements) {
            foreach ($propertyName in @('InternalClientsRequireSsl','ExternalClientsRequireSsl')) {
                $property = $SourceSnapshot.OutlookAny.PSObject.Properties[$propertyName]
                if ($property) { $oaValues[$propertyName] = $property.Value }
            }
        }
        elseif ($CompareOutlookAnywhereSslRequirements) {
            foreach ($propertyName in @('InternalClientsRequireSsl','ExternalClientsRequireSsl')) {
                $property = $SourceSnapshot.OutlookAny.PSObject.Properties[$propertyName]
                if ($property) { $oaSslReviewValues[$propertyName] = $property.Value }
            }
        }

        if ($CopyAuthentication) {
            foreach ($propertyName in @('InternalClientAuthenticationMethod','ExternalClientAuthenticationMethod','IISAuthenticationMethods')) {
                $property = $SourceSnapshot.OutlookAny.PSObject.Properties[$propertyName]
                if ($property) { $oaValues[$propertyName] = $property.Value }
            }
        }
    }
    else {
        $oaValues['InternalHostname'] = $InternalNamespace
        $oaValues['ExternalHostname'] = if ($ClearExternalUrls) { $null } else { $ExternalNamespace }

        if ($null -ne $OutlookAnywhereInternalClientsRequireSsl) {
            $oaValues['InternalClientsRequireSsl'] = [bool]$OutlookAnywhereInternalClientsRequireSsl
        }
        if ($null -ne $OutlookAnywhereExternalClientsRequireSsl) {
            $oaValues['ExternalClientsRequireSsl'] = [bool]$OutlookAnywhereExternalClientsRequireSsl
        }

        if (-not [string]::IsNullOrWhiteSpace($OutlookAnywhereDefaultAuthenticationMethod)) {
            $oaDefaultAuth = $OutlookAnywhereDefaultAuthenticationMethod
            $oaValues['InternalClientAuthenticationMethod'] = $oaDefaultAuth
            $oaValues['ExternalClientAuthenticationMethod'] = $oaDefaultAuth
            $oaValues['IISAuthenticationMethods'] = @($oaDefaultAuth)
        }
    }

    [void]$components.Add([PSCustomObject]@{
        Key = 'OutlookAny'
        Name = 'Outlook Anywhere (RPC over HTTP)'
        Values = $oaValues
        Policy = 'Apply'
        ReviewNote = $null
        DefaultAuthenticationMethod = $oaDefaultAuth
    })

    if ($SourceSnapshot -and $oaSslReviewValues.Count -gt 0) {
        [void]$components.Add([PSCustomObject]@{
            Key = 'OutlookAny'
            Name = 'Outlook Anywhere (RPC over HTTP) (Review Only)'
            Values = $oaSslReviewValues
            Policy = 'ReviewOnly'
            ReviewNote = 'Clone Apply preserves target SSL requirement values unless -IncludeOutlookAnywhereSslRequirements is explicitly supplied.'
        })
    }

    if ($SourceSnapshot) {
        $sslOffloadProperty = $SourceSnapshot.OutlookAny.PSObject.Properties['SSLOffloading']
        if ($sslOffloadProperty) {
            [void]$components.Add([PSCustomObject]@{
                Key = 'OutlookAny'
                Name = 'Outlook Anywhere (RPC over HTTP) (Review Only)'
                Values = [ordered]@{ SSLOffloading = $sslOffloadProperty.Value }
                Policy = 'ReviewOnly'
                ReviewNote = 'Review the load balancer/TLS topology before changing SSL offloading. This setting is not changed by the script.'
            })
        }
    }

    return [PSCustomObject]@{ Target = $TargetSnapshot.Server; Components = @($components) }
}

function New-RestoreDesiredConfiguration {
    param(
        [Parameter(Mandatory = $true)]$TargetSnapshot,
        [Parameter(Mandatory = $true)]$BackupDocument,
        [bool]$CopyAuthentication = $false,
        [bool]$IncludePowerShellUrls = $false,
        [bool]$IncludeOutlookAnywhereSslRequirements = $false
    )

    $components = New-Object System.Collections.ArrayList

    foreach ($entry in $VirtualDirectoryMap) {
        $backupItem = Get-BackupVirtualDirectory -BackupDocument $BackupDocument -Key $entry.Key

        if ($entry.Path) {
            if ($null -eq $backupItem.Urls) { throw "Backup is missing $($entry.Name) URL data." }
            $internalProperty = $backupItem.Urls.PSObject.Properties['InternalUrl']
            $externalProperty = $backupItem.Urls.PSObject.Properties['ExternalUrl']
            if ($null -eq $internalProperty -or $null -eq $externalProperty) {
                throw "Backup is missing $($entry.Name) InternalUrl/ExternalUrl."
            }

            $urlPolicy = if ($entry.Key -eq 'PowerShell' -and -not $IncludePowerShellUrls) { 'ReviewOnly' } else { 'Apply' }
            $urlReview = if ($entry.Key -eq 'PowerShell' -and -not $IncludePowerShellUrls) {
                'PowerShell URL restore requires explicit -IncludePowerShellUrls.'
            }
            elseif ($entry.Key -eq 'PowerShell') {
                'Only PowerShell InternalUrl/ExternalUrl are included. PowerShell authentication, RequireSSL, and Extended Protection remain Review Only.'
            }
            else { $null }

            [void]$components.Add([PSCustomObject]@{
                Key = $entry.Key
                Name = $entry.Name
                Values = [ordered]@{
                    InternalUrl = $internalProperty.Value
                    ExternalUrl = $externalProperty.Value
                }
                Policy = $urlPolicy
                ReviewNote = $urlReview
            })
        }

        if ($null -eq $backupItem.Authentication) { throw "Backup is missing $($entry.Name) authentication data." }
        $authValues = [ordered]@{}
        foreach ($propertyName in $entry.AuthProperties) {
            $property = $backupItem.Authentication.PSObject.Properties[$propertyName]
            if ($property) { $authValues[$propertyName] = $property.Value }
        }

        if ($authValues.Count -gt 0) {
            $authPolicy = if ($entry.Key -eq 'PowerShell' -or -not $CopyAuthentication) { 'ReviewOnly' } else { 'Apply' }
            $authReview = if ($entry.Key -eq 'PowerShell') {
                'PowerShell authentication is Review Only and is not restored automatically.'
            }
            elseif (-not $CopyAuthentication) {
                'Authentication exists in the backup but restore requires explicit -IncludeAuthentication.'
            }
            else { $null }

            [void]$components.Add([PSCustomObject]@{
                Key = $entry.Key
                Name = "$($entry.Name) Authentication"
                Values = $authValues
                Policy = $authPolicy
                ReviewNote = $authReview
            })
        }

        if ($null -eq $backupItem.ReviewOnly) { throw "Backup is missing $($entry.Name) Review Only data." }
        $reviewValues = [ordered]@{}
        foreach ($propertyName in $entry.ReviewProperties) {
            $property = $backupItem.ReviewOnly.PSObject.Properties[$propertyName]
            if ($property) { $reviewValues[$propertyName] = $property.Value }
        }
        if ($reviewValues.Count -gt 0) {
            [void]$components.Add([PSCustomObject]@{
                Key = $entry.Key
                Name = "$($entry.Name) (Review Only)"
                Values = $reviewValues
                Policy = 'ReviewOnly'
                ReviewNote = 'Review Only value from backup. It is not restored automatically.'
            })
        }
    }

    $scpProperty = $BackupDocument.Configuration.ClientAccess.PSObject.Properties['AutoDiscoverServiceInternalUri']
    if ($null -eq $scpProperty) { throw 'Backup is missing AutoDiscoverServiceInternalUri.' }
    [void]$components.Add([PSCustomObject]@{
        Key = 'ClientAccess'
        Name = 'Autodiscover SCP'
        Values = [ordered]@{ AutoDiscoverServiceInternalUri = $scpProperty.Value }
        Policy = 'Apply'
        ReviewNote = $null
    })

    [void]$components.Add([PSCustomObject]@{
        Key = 'ASA'
        Name = 'Alternate Service Account (Kerberos)'
        Values = [ordered]@{
            Account = $BackupDocument.Configuration.ASA.Account
            Status  = $BackupDocument.Configuration.ASA.Status
        }
        Policy = 'ReviewOnly'
        ReviewNote = 'ASA/Kerberos backup data is visibility only. Credentials are not restored by this script.'
    })

    $oa = $BackupDocument.Configuration.OutlookAnywhere
    $oaValues = [ordered]@{}
    foreach ($propertyName in @('InternalHostname','ExternalHostname')) {
        $property = $oa.PSObject.Properties[$propertyName]
        if ($null -eq $property) { throw "Backup is missing Outlook Anywhere $propertyName." }
        $oaValues[$propertyName] = $property.Value
    }

    $oaSslValues = [ordered]@{}
    foreach ($propertyName in @('InternalClientsRequireSsl','ExternalClientsRequireSsl')) {
        $property = $oa.PSObject.Properties[$propertyName]
        if ($property) { $oaSslValues[$propertyName] = $property.Value }
    }
    if ($IncludeOutlookAnywhereSslRequirements) {
        if ($oaSslValues.Count -ne 2) {
            throw 'Backup is missing one or more Outlook Anywhere SSL requirement values.'
        }
        foreach ($propertyName in $oaSslValues.Keys) { $oaValues[$propertyName] = $oaSslValues[$propertyName] }
    }

    $oaAuthValues = [ordered]@{}
    foreach ($propertyName in @('InternalClientAuthenticationMethod','ExternalClientAuthenticationMethod','IISAuthenticationMethods')) {
        $property = $oa.PSObject.Properties[$propertyName]
        if ($property) { $oaAuthValues[$propertyName] = $property.Value }
    }
    if ($CopyAuthentication) {
        if ($oaAuthValues.Count -ne 3) {
            throw 'Backup is missing one or more Outlook Anywhere authentication values.'
        }
        foreach ($propertyName in $oaAuthValues.Keys) { $oaValues[$propertyName] = $oaAuthValues[$propertyName] }
    }

    [void]$components.Add([PSCustomObject]@{
        Key = 'OutlookAny'
        Name = 'Outlook Anywhere (RPC over HTTP)'
        Values = $oaValues
        Policy = 'Apply'
        ReviewNote = $null
        DefaultAuthenticationMethod = $null
    })

    if (-not $IncludeOutlookAnywhereSslRequirements -and $oaSslValues.Count -gt 0) {
        [void]$components.Add([PSCustomObject]@{
            Key = 'OutlookAny'
            Name = 'Outlook Anywhere (RPC over HTTP) SSL Requirements'
            Values = $oaSslValues
            Policy = 'ReviewOnly'
            ReviewNote = 'SSL requirement values exist in the backup but restore requires explicit -IncludeOutlookAnywhereSslRequirements.'
        })
    }

    if (-not $CopyAuthentication -and $oaAuthValues.Count -gt 0) {
        [void]$components.Add([PSCustomObject]@{
            Key = 'OutlookAny'
            Name = 'Outlook Anywhere (RPC over HTTP) Authentication'
            Values = $oaAuthValues
            Policy = 'ReviewOnly'
            ReviewNote = 'Authentication values exist in the backup but restore requires explicit -IncludeAuthentication.'
        })
    }

    $sslOffloadProperty = $oa.PSObject.Properties['SSLOffloading']
    if ($sslOffloadProperty) {
        [void]$components.Add([PSCustomObject]@{
            Key = 'OutlookAny'
            Name = 'Outlook Anywhere (RPC over HTTP) (Review Only)'
            Values = [ordered]@{ SSLOffloading = $sslOffloadProperty.Value }
            Policy = 'ReviewOnly'
            ReviewNote = 'Review the load balancer/TLS topology before changing SSL offloading. This setting is not restored by the script.'
        })
    }

    return [PSCustomObject]@{
        Target = $TargetSnapshot.Server
        Components = @($components)
    }
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
            $currentProperty = $currentObject.PSObject.Properties[$propertyName]
            $currentValue = if ($currentProperty) { $currentProperty.Value } else { $null }
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
# Compare Clone Source and Target Values
# ---------------------------------------------------------------------------
# Default Clone comparison deliberately builds a complete comparison list, including
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
            $targetProperty = $targetObject.PSObject.Properties[$propertyName]
            $targetValue = if ($targetProperty) { $targetProperty.Value } else { $null }
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

function Get-ComparisonDisplayComponentName {
    param([AllowNull()][string]$Component)

    $name = $Component -replace ' \(Review Only\)$',''

    switch ($name) {
        'OWA Authentication'                              { return 'OWA' }
        'ECP Authentication'                              { return 'ECP' }
        'EWS Authentication'                              { return 'EWS' }
        'MAPI Authentication'                             { return 'MAPI' }
        'ActiveSync Authentication'                       { return 'ActiveSync' }
        'OAB Authentication'                              { return 'OAB' }
        'Outlook Anywhere (RPC over HTTP) Authentication' { return 'Outlook Anywhere (RPC over HTTP)' }
        'Autodiscover Authentication'                     { return 'Autodiscover Virtual Directory' }
        'PowerShell Authentication'                       { return 'PowerShell' }
        default                                           { return $name }
    }
}

function Get-ComparisonSummaryComponentName {
    param([AllowNull()][string]$Component)

    $name = Get-ComparisonDisplayComponentName -Component $Component

    switch ($name) {
        'Outlook Anywhere (RPC over HTTP)'         { return 'Outlook Anywhere' }
        'Alternate Service Account (Kerberos)'    { return 'ASA / Kerberos' }
        default                                   { return $name }
    }
}

function Get-ConsoleRenderedLineCount {
    param(
        [AllowNull()][string]$Text,
        [Parameter(Mandatory = $true)][int]$Width
    )

    if ($Width -lt 1) {
        return 1
    }

    if ([string]::IsNullOrEmpty($Text)) {
        return 1
    }

    return [Math]::Max(1, [int][Math]::Ceiling(([double]$Text.Length) / [double]$Width))
}

function Show-ComparisonPlan {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Comparisons,
        [Parameter(Mandatory = $true)][string]$SourceServer,
        [AllowEmptyCollection()][array]$TargetFailures = @(),
        [switch]$EnablePaging
    )

    Write-Host ''
    Write-Host 'Comparison' -ForegroundColor Cyan
    Write-Host '----------' -ForegroundColor Cyan

    $targets = @($Comparisons.Target | Select-Object -Unique)
    $pagingEnabled = $EnablePaging -and (Test-ConsolePagingAvailable)

    $pageCapacity = 0
    $consoleWidth = 120
    if ($pagingEnabled) {
        try {
            $pageCapacity = [Math]::Max(18, ([Console]::WindowHeight - 4))
            $consoleWidth = [Math]::Max(40, ([Console]::WindowWidth - 1))
        }
        catch {
            $pageCapacity = 24
            $consoleWidth = 120
        }
    }

    $settingGroups = @(
        $Comparisons |
            Group-Object Component,Setting |
            ForEach-Object {
                $first = $_.Group | Select-Object -First 1
                [PSCustomObject]@{
                    Component        = $first.Component
                    DisplayComponent = Get-ComparisonDisplayComponentName -Component $first.Component
                    Setting          = $first.Setting
                    Policy           = $first.Policy
                    ReviewNote       = $first.ReviewNote
                    Source           = $first.Source
                    Items            = @($_.Group)
                }
            } |
            Sort-Object `
                @{ Expression = { Get-ComponentDisplayOrder -Component $_.DisplayComponent } }, `
                DisplayComponent, `
                @{ Expression = { Get-SettingDisplayOrder -Setting $_.Setting } }, `
                Setting
    )

    $serverLabelWidth = $SourceServer.Length
    foreach ($target in $targets) {
        if (([string]$target).Length -gt $serverLabelWidth) {
            $serverLabelWidth = ([string]$target).Length
        }
    }

    $pageLinesUsed = 3
    $separator = '##############################################################################'
    $componentNames = @($settingGroups.DisplayComponent | Select-Object -Unique)

    foreach ($componentName in $componentNames) {
        $componentItems = @($settingGroups | Where-Object { $_.DisplayComponent -eq $componentName })
        $headerLines = @('', $separator, ("# {0}" -f $componentName), $separator, '')
        $headerRowCount = 0
        foreach ($line in $headerLines) {
            $headerRowCount += Get-ConsoleRenderedLineCount -Text $line -Width $consoleWidth
        }

        # Start a component on the current page only if the heading and first
        # complete setting block fit. Large components may continue on later
        # pages, but an individual setting block is never split.
        $firstSettingRows = 0
        if ($componentItems.Count -gt 0) {
            $first = $componentItems[0]
            $firstSettingRows += Get-ConsoleRenderedLineCount -Text $first.Setting -Width $consoleWidth
            if ($first.Policy -eq 'ReviewOnly') {
                $firstSettingRows += Get-ConsoleRenderedLineCount -Text '  Policy : Review Only' -Width $consoleWidth
            }
            if ($first.ReviewNote) {
                $firstSettingRows += Get-ConsoleRenderedLineCount -Text ("  Review : {0}" -f $first.ReviewNote) -Width $consoleWidth
            }
            $firstSettingRows += Get-ConsoleRenderedLineCount -Text ("  {0} : {1}" -f $SourceServer.PadRight($serverLabelWidth), $first.Source) -Width $consoleWidth
            foreach ($target in $targets) {
                $item = $first.Items | Where-Object { $_.Target -eq $target } | Select-Object -First 1
                if (-not $item) { continue }
                $targetLabel = ([string]$target).PadRight($serverLabelWidth)
                $firstSettingRows += Get-ConsoleRenderedLineCount -Text ("  {0} : {1}" -f $targetLabel, $item.TargetValue) -Width $consoleWidth
                $firstSettingRows += Get-ConsoleRenderedLineCount -Text ("  {0} : {1}" -f 'Status'.PadRight($serverLabelWidth), $item.Status) -Width $consoleWidth
            }
            $firstSettingRows += 1
        }

        if (
            $pagingEnabled -and
            $pageLinesUsed -gt 3 -and
            ($pageLinesUsed + $headerRowCount + $firstSettingRows) -gt $pageCapacity
        ) {
            if (-not (Read-ComparisonPageContinuation)) {
                return $false
            }
            $pageLinesUsed = 0
        }

        foreach ($line in $headerLines) {
            if ($line -eq $separator -or $line -like '# *') {
                Write-Host $line -ForegroundColor Cyan
            }
            else {
                Write-Host $line
            }
        }
        $pageLinesUsed += $headerRowCount

        $settingIndex = 0
        foreach ($group in $componentItems) {
            $settingIndex++
            $settingRows = 0
            $settingRows += Get-ConsoleRenderedLineCount -Text $group.Setting -Width $consoleWidth

            if ($group.Policy -eq 'ReviewOnly') {
                $settingRows += Get-ConsoleRenderedLineCount -Text '  Policy : Review Only' -Width $consoleWidth
            }

            if ($group.ReviewNote) {
                $settingRows += Get-ConsoleRenderedLineCount -Text ("  Review : {0}" -f $group.ReviewNote) -Width $consoleWidth
            }

            $sourceLabel = $SourceServer.PadRight($serverLabelWidth)
            $settingRows += Get-ConsoleRenderedLineCount -Text ("  {0} : {1}" -f $sourceLabel, $group.Source) -Width $consoleWidth

            foreach ($target in $targets) {
                $item = $group.Items | Where-Object { $_.Target -eq $target } | Select-Object -First 1
                if (-not $item) { continue }

                $targetLabel = ([string]$target).PadRight($serverLabelWidth)
                $settingRows += Get-ConsoleRenderedLineCount -Text ("  {0} : {1}" -f $targetLabel, $item.TargetValue) -Width $consoleWidth
                $settingRows += Get-ConsoleRenderedLineCount -Text ("  {0} : {1}" -f 'Status'.PadRight($serverLabelWidth), $item.Status) -Width $consoleWidth
            }

            $settingRows += 1

            if (
                $pagingEnabled -and
                $pageLinesUsed -gt 0 -and
                ($pageLinesUsed + $settingRows) -gt $pageCapacity
            ) {
                if (-not (Read-ComparisonPageContinuation)) {
                    return $false
                }

                $continuedHeader = @(
                    $separator,
                    ("# {0} (continued)" -f $componentName),
                    $separator,
                    ''
                )
                $pageLinesUsed = 0
                foreach ($line in $continuedHeader) {
                    Write-Host $line -ForegroundColor Cyan
                    $pageLinesUsed += Get-ConsoleRenderedLineCount -Text $line -Width $consoleWidth
                }
            }

            Write-Host $group.Setting -ForegroundColor White

            if ($group.Policy -eq 'ReviewOnly') {
                Write-Host '  Policy : Review Only' -ForegroundColor Yellow
            }

            if ($group.ReviewNote) {
                Write-Host ("  Review : {0}" -f $group.ReviewNote) -ForegroundColor DarkYellow
            }

            Write-Host ("  {0} : {1}" -f $sourceLabel, $group.Source) -ForegroundColor Gray

            foreach ($target in $targets) {
                $item = $group.Items | Where-Object { $_.Target -eq $target } | Select-Object -First 1
                if (-not $item) { continue }

                $targetLabel = ([string]$target).PadRight($serverLabelWidth)
                if ($item.Status -eq 'Different') {
                    Write-Host ("  {0} : {1}" -f $targetLabel, $item.TargetValue) -ForegroundColor Red
                    Write-Host ("  {0} : Different" -f 'Status'.PadRight($serverLabelWidth)) -ForegroundColor Red
                }
                else {
                    Write-Host ("  {0} : {1}" -f $targetLabel, $item.TargetValue) -ForegroundColor Gray
                    Write-Host ("  {0} : Same" -f 'Status'.PadRight($serverLabelWidth)) -ForegroundColor Green
                }
            }

            Write-Host ''
            $pageLinesUsed += $settingRows
        }
    }

    # Build a compact per-target Summary and aggregate by the final display
    # component name so URL/auth/Review Only values for one service appear once.
    $summaryModels = @()
    foreach ($target in $targets) {
        $targetItems = @($Comparisons | Where-Object { $_.Target -eq $target })
        $differentItems = @($targetItems | Where-Object { $_.Status -eq 'Different' })
        $sameItems = @($targetItems | Where-Object { $_.Status -eq 'Same' })

        $differentGroups = @(
            $differentItems |
                Group-Object { Get-ComparisonSummaryComponentName -Component $_.Component } |
                ForEach-Object {
                    [PSCustomObject]@{
                        Component = [string]$_.Name
                        Settings  = @($_.Group | Sort-Object @{ Expression = { Get-SettingDisplayOrder -Setting $_.Setting } }, Setting | ForEach-Object { $_.Setting })
                    }
                } |
                Sort-Object @{ Expression = { Get-ComponentDisplayOrder -Component (($_.Component -replace '^Outlook Anywhere$','Outlook Anywhere (RPC over HTTP)') -replace '^ASA / Kerberos$','Alternate Service Account (Kerberos)') } }, Component
        )

        $sameGroups = @(
            $sameItems |
                Group-Object { Get-ComparisonSummaryComponentName -Component $_.Component } |
                ForEach-Object {
                    [PSCustomObject]@{
                        Component = [string]$_.Name
                        Settings  = @($_.Group | Sort-Object @{ Expression = { Get-SettingDisplayOrder -Setting $_.Setting } }, Setting | ForEach-Object { $_.Setting })
                    }
                } |
                Sort-Object @{ Expression = { Get-ComponentDisplayOrder -Component (($_.Component -replace '^Outlook Anywhere$','Outlook Anywhere (RPC over HTTP)') -replace '^ASA / Kerberos$','Alternate Service Account (Kerberos)') } }, Component
        )

        $summaryModels += [PSCustomObject]@{
            Target           = $target
            ComparedCount    = $targetItems.Count
            DifferentCount   = $differentItems.Count
            SameCount        = $sameItems.Count
            DifferentGroups  = $differentGroups
            SameGroups       = $sameGroups
        }
    }

    $summaryHeader = @('Summary', '-------', '')
    $summaryHeaderRows = 0
    foreach ($line in $summaryHeader) {
        $summaryHeaderRows += Get-ConsoleRenderedLineCount -Text $line -Width $consoleWidth
    }

    if (
        $pagingEnabled -and
        $pageLinesUsed -gt 0 -and
        ($pageLinesUsed + $summaryHeaderRows) -gt $pageCapacity
    ) {
        if (-not (Read-ComparisonPageContinuation)) {
            return $false
        }
        $pageLinesUsed = 0
    }

    Write-Host 'Summary' -ForegroundColor Cyan
    Write-Host '-------' -ForegroundColor Cyan
    Write-Host ''
    $pageLinesUsed += $summaryHeaderRows

    foreach ($model in $summaryModels) {
        $metadataLines = @(
            ("Source server     : {0}" -f $SourceServer),
            ("Target server     : {0}" -f $model.Target),
            ("Compared settings : {0}" -f $model.ComparedCount),
            ("Different         : {0}" -f $model.DifferentCount),
            ("Same              : {0}" -f $model.SameCount),
            ''
        )

        $metadataRows = 0
        foreach ($line in $metadataLines) {
            $metadataRows += Get-ConsoleRenderedLineCount -Text $line -Width $consoleWidth
        }

        if (
            $pagingEnabled -and
            $pageLinesUsed -gt 0 -and
            ($pageLinesUsed + $metadataRows) -gt $pageCapacity
        ) {
            if (-not (Read-ComparisonPageContinuation)) {
                return $false
            }
            $pageLinesUsed = 0
        }

        Write-Host $metadataLines[0] -ForegroundColor Gray
        Write-Host $metadataLines[1] -ForegroundColor Gray
        Write-Host $metadataLines[2] -ForegroundColor Gray
        Write-Host $metadataLines[3] -ForegroundColor Red
        Write-Host $metadataLines[4] -ForegroundColor Green
        Write-Host ''
        $pageLinesUsed += $metadataRows

        # Different section: paginate between complete component lines and
        # account for console wrapping of long setting lists.
        if ($model.DifferentGroups.Count -gt 0) {
            $differentWidth = ($model.DifferentGroups | ForEach-Object { $_.Component.Length } | Measure-Object -Maximum).Maximum
            $differentHeading = 'Different:'
            $headingRows = Get-ConsoleRenderedLineCount -Text $differentHeading -Width $consoleWidth

            if ($pagingEnabled -and ($pageLinesUsed + $headingRows) -gt $pageCapacity) {
                if (-not (Read-ComparisonPageContinuation)) { return $false }
                $pageLinesUsed = 0
            }

            Write-Host $differentHeading -ForegroundColor Red
            $pageLinesUsed += $headingRows

            $differentIndex = 0
            foreach ($group in $model.DifferentGroups) {
                $differentIndex++
                $line = "  {0} : {1}" -f $group.Component.PadRight($differentWidth), ($group.Settings -join ', ')
                $rowCount = Get-ConsoleRenderedLineCount -Text $line -Width $consoleWidth

                if ($pagingEnabled -and ($pageLinesUsed + $rowCount) -gt $pageCapacity) {
                    if (-not (Read-ComparisonPageContinuation)) { return $false }
                    Write-Host 'Different (continued):' -ForegroundColor Red
                    $pageLinesUsed = Get-ConsoleRenderedLineCount -Text 'Different (continued):' -Width $consoleWidth
                }

                Write-Host $line -ForegroundColor Red
                $pageLinesUsed += $rowCount
            }
        }
        else {
            $line = 'Different: None'
            $rows = Get-ConsoleRenderedLineCount -Text $line -Width $consoleWidth
            if ($pagingEnabled -and ($pageLinesUsed + $rows) -gt $pageCapacity) {
                if (-not (Read-ComparisonPageContinuation)) { return $false }
                $pageLinesUsed = 0
            }
            Write-Host $line -ForegroundColor Green
            $pageLinesUsed += $rows
        }

        Write-Host ''
        $pageLinesUsed += 1

        # Same section: same paging rules as Different.
        if ($model.SameGroups.Count -gt 0) {
            $sameWidth = ($model.SameGroups | ForEach-Object { $_.Component.Length } | Measure-Object -Maximum).Maximum
            $sameHeading = 'Same:'
            $headingRows = Get-ConsoleRenderedLineCount -Text $sameHeading -Width $consoleWidth

            if ($pagingEnabled -and ($pageLinesUsed + $headingRows) -gt $pageCapacity) {
                if (-not (Read-ComparisonPageContinuation)) { return $false }
                $pageLinesUsed = 0
            }

            Write-Host $sameHeading -ForegroundColor Green
            $pageLinesUsed += $headingRows

            foreach ($group in $model.SameGroups) {
                $line = "  {0} : {1}" -f $group.Component.PadRight($sameWidth), ($group.Settings -join ', ')
                $rowCount = Get-ConsoleRenderedLineCount -Text $line -Width $consoleWidth

                if ($pagingEnabled -and ($pageLinesUsed + $rowCount) -gt $pageCapacity) {
                    if (-not (Read-ComparisonPageContinuation)) { return $false }
                    Write-Host 'Same (continued):' -ForegroundColor Green
                    $pageLinesUsed = Get-ConsoleRenderedLineCount -Text 'Same (continued):' -Width $consoleWidth
                }

                Write-Host $line -ForegroundColor Green
                $pageLinesUsed += $rowCount
            }
        }
        else {
            $line = 'Same: None'
            $rows = Get-ConsoleRenderedLineCount -Text $line -Width $consoleWidth
            if ($pagingEnabled -and ($pageLinesUsed + $rows) -gt $pageCapacity) {
                if (-not (Read-ComparisonPageContinuation)) { return $false }
                $pageLinesUsed = 0
            }
            Write-Host $line -ForegroundColor Red
            $pageLinesUsed += $rows
        }

        Write-Host ''
        $pageLinesUsed += 1
    }

    foreach ($failure in @($TargetFailures)) {
        $failureLines = @(
            ("Source server : {0}" -f $SourceServer),
            ("Target server : {0}" -f $failure.Target),
            ("Status        : {0}" -f $failure.Status),
            ("Detail        : {0}" -f $failure.Detail),
            ''
        )

        $failureRows = 0
        foreach ($line in $failureLines) {
            $failureRows += Get-ConsoleRenderedLineCount -Text $line -Width $consoleWidth
        }

        if (
            $pagingEnabled -and
            $pageLinesUsed -gt 0 -and
            ($pageLinesUsed + $failureRows) -gt $pageCapacity
        ) {
            if (-not (Read-ComparisonPageContinuation)) {
                return $false
            }
            $pageLinesUsed = 0
        }

        Write-Host $failureLines[0] -ForegroundColor Gray
        Write-Host $failureLines[1] -ForegroundColor Gray
        Write-Host $failureLines[2] -ForegroundColor Red
        Write-Host $failureLines[3] -ForegroundColor DarkYellow
        Write-Host ''
        $pageLinesUsed += $failureRows
    }

    $completionFooterLineCount = 11
    if (
        $pagingEnabled -and
        $pageLinesUsed -gt 0 -and
        ($pageLinesUsed + $completionFooterLineCount) -gt $pageCapacity
    ) {
        if (-not (Read-ComparisonPageContinuation)) {
            return $false
        }
    }

    return $true
}

# Default Clone comparison + OutputFile is a reporting combination, not a command-export
# combination. Keeping this function separate from Export-CommandPlan makes it
# impossible for a read-only comparison to accidentally emit executable Set-* commands.
function Export-ComparisonReport {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Comparisons,
        [Parameter(Mandatory = $true)][string[]]$Targets,
        [Parameter(Mandatory = $true)][string]$SourceServer,
        [AllowEmptyCollection()][array]$TargetFailures = @(),
        [switch]$IncludeAuthentication
    )

    $lines = New-Object System.Collections.ArrayList
    foreach ($line in @(
        '# ExchangeURLManager.ps1 comparison report'
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

    $settingGroups = @(
        $Comparisons |
            Group-Object Component,Setting |
            ForEach-Object {
                $first = $_.Group | Select-Object -First 1
                [PSCustomObject]@{
                    Component        = $first.Component
                    DisplayComponent = Get-ComparisonDisplayComponentName -Component $first.Component
                    Setting          = $first.Setting
                    Policy           = $first.Policy
                    ReviewNote       = $first.ReviewNote
                    Source           = $first.Source
                    Items            = @($_.Group)
                }
            } |
            Sort-Object `
                @{ Expression = { Get-ComponentDisplayOrder -Component $_.DisplayComponent } }, `
                DisplayComponent, `
                @{ Expression = { Get-SettingDisplayOrder -Setting $_.Setting } }, `
                Setting
    )

    $serverLabelWidth = $SourceServer.Length
    foreach ($target in $Targets) {
        if (([string]$target).Length -gt $serverLabelWidth) {
            $serverLabelWidth = ([string]$target).Length
        }
    }

    $currentComponent = $null
    foreach ($group in $settingGroups) {
        if ($currentComponent -ne $group.DisplayComponent) {
            $currentComponent = $group.DisplayComponent
            [void]$lines.Add('##############################################################################')
            [void]$lines.Add(("# {0}" -f $currentComponent))
            [void]$lines.Add('##############################################################################')
            [void]$lines.Add('')
        }

        [void]$lines.Add($group.Setting)
        if ($group.Policy -eq 'ReviewOnly') {
            [void]$lines.Add('  Policy : Review Only')
        }
        if ($group.ReviewNote) {
            [void]$lines.Add(("  Review : {0}" -f $group.ReviewNote))
        }

        $sourceLabel = $SourceServer.PadRight($serverLabelWidth)
        [void]$lines.Add(("  {0} : {1}" -f $sourceLabel, $group.Source))

        foreach ($target in $Targets) {
            $item = $group.Items | Where-Object { $_.Target -eq $target } | Select-Object -First 1
            if (-not $item) { continue }

            $targetLabel = ([string]$target).PadRight($serverLabelWidth)
            [void]$lines.Add(("  {0} : {1}" -f $targetLabel, $item.TargetValue))
            [void]$lines.Add(("  {0} : {1}" -f 'Status'.PadRight($serverLabelWidth), $item.Status))
        }

        [void]$lines.Add('')
    }

    [void]$lines.Add('Summary')
    [void]$lines.Add('-------')

    foreach ($target in $Targets) {
        $targetItems = @($Comparisons | Where-Object { $_.Target -eq $target })
        $differentItems = @($targetItems | Where-Object { $_.Status -eq 'Different' })
        $sameItems = @($targetItems | Where-Object { $_.Status -eq 'Same' })

        [void]$lines.Add('')
        [void]$lines.Add(("Source server     : {0}" -f $SourceServer))
        [void]$lines.Add(("Target server     : {0}" -f $target))
        [void]$lines.Add(("Compared settings : {0}" -f $targetItems.Count))
        [void]$lines.Add(("Different         : {0}" -f $differentItems.Count))
        [void]$lines.Add(("Same              : {0}" -f $sameItems.Count))
        [void]$lines.Add('')

        $differentGroups = @(
            $differentItems |
                Group-Object { Get-ComparisonSummaryComponentName -Component $_.Component } |
                ForEach-Object {
                    [PSCustomObject]@{
                        Component = [string]$_.Name
                        Settings  = @($_.Group | Sort-Object @{ Expression = { Get-SettingDisplayOrder -Setting $_.Setting } }, Setting | ForEach-Object { $_.Setting })
                    }
                } |
                Sort-Object `
                    @{ Expression = { Get-ComponentDisplayOrder -Component (($_.Component -replace '^Outlook Anywhere$','Outlook Anywhere (RPC over HTTP)') -replace '^ASA / Kerberos$','Alternate Service Account (Kerberos)') } }, `
                    Component
        )

        if ($differentGroups.Count -gt 0) {
            [void]$lines.Add('Different:')
            $differentWidth = ($differentGroups | ForEach-Object { $_.Component.Length } | Measure-Object -Maximum).Maximum
            foreach ($group in $differentGroups) {
                [void]$lines.Add(("  {0} : {1}" -f $group.Component.PadRight($differentWidth), ($group.Settings -join ', ')))
            }
        }
        else {
            [void]$lines.Add('Different: None')
        }

        [void]$lines.Add('')

        $sameGroups = @(
            $sameItems |
                Group-Object { Get-ComparisonSummaryComponentName -Component $_.Component } |
                ForEach-Object {
                    [PSCustomObject]@{
                        Component = [string]$_.Name
                        Settings  = @($_.Group | Sort-Object @{ Expression = { Get-SettingDisplayOrder -Setting $_.Setting } }, Setting | ForEach-Object { $_.Setting })
                    }
                } |
                Sort-Object `
                    @{ Expression = { Get-ComponentDisplayOrder -Component (($_.Component -replace '^Outlook Anywhere$','Outlook Anywhere (RPC over HTTP)') -replace '^ASA / Kerberos$','Alternate Service Account (Kerberos)') } }, `
                    Component
        )

        if ($sameGroups.Count -gt 0) {
            [void]$lines.Add('Same:')
            $sameWidth = ($sameGroups | ForEach-Object { $_.Component.Length } | Measure-Object -Maximum).Maximum
            foreach ($group in $sameGroups) {
                [void]$lines.Add(("  {0} : {1}" -f $group.Component.PadRight($sameWidth), ($group.Settings -join ', ')))
            }
        }
        else {
            [void]$lines.Add('Same: None')
        }

        [void]$lines.Add('')
    }

    foreach ($failure in @($TargetFailures)) {
        [void]$lines.Add(("Source server : {0}" -f $SourceServer))
        [void]$lines.Add(("Target server : {0}" -f $failure.Target))
        [void]$lines.Add(("Status        : {0}" -f $failure.Status))
        [void]$lines.Add(("Detail        : {0}" -f $failure.Detail))
        [void]$lines.Add('')
    }

    Set-Content -Path $Path -Value $lines -Encoding UTF8
}

function Get-ComponentDisplayOrder {
    param([AllowNull()][string]$Component)

    switch ($Component) {
        'OWA'                                      { return 10 }
        'ECP'                                      { return 20 }
        'EWS'                                      { return 30 }
        'MAPI'                                     { return 40 }
        'ActiveSync'                               { return 50 }
        'OAB'                                      { return 60 }
        'Outlook Anywhere (RPC over HTTP)'         { return 70 }
        'Autodiscover Virtual Directory'            { return 75 }
        'Autodiscover SCP'                         { return 80 }
        'PowerShell'                               { return 90 }
        'Alternate Service Account (Kerberos)'     { return 100 }
        default                                    { return 110 }
    }
}

function Get-SettingDisplayOrder {
    param([AllowNull()][string]$Setting)

    switch ($Setting) {
        'InternalUrl'                          { return 10 }
        'InternalHostname'                     { return 10 }
        'AutoDiscoverServiceInternalUri'       { return 10 }
        'ExternalUrl'                          { return 20 }
        'ExternalHostname'                     { return 20 }
        'InternalClientsRequireSsl'            { return 30 }
        'ExternalClientsRequireSsl'            { return 40 }
        'InternalClientAuthenticationMethod'   { return 50 }
        'ExternalClientAuthenticationMethod'   { return 60 }
        'IISAuthenticationMethods'             { return 70 }
        'BasicAuthentication'                  { return 80 }
        'DigestAuthentication'                 { return 90 }
        'WindowsAuthentication'                { return 100 }
        'FormsAuthentication'                  { return 110 }
        'AdfsAuthentication'                   { return 120 }
        'OAuthAuthentication'                  { return 130 }
        'WSSecurityAuthentication'             { return 140 }
        'CertificateAuthentication'            { return 150 }
        'BasicAuthEnabled'                     { return 160 }
        'WindowsAuthEnabled'                   { return 170 }
        'ClientCertAuth'                       { return 180 }
        'LogonFormat'                          { return 190 }
        'DefaultDomain'                        { return 200 }
        'MRSProxyEnabled'                      { return 210 }
        'RequireSSL'                           { return 220 }
        'SSLOffloading'                        { return 230 }
        default                                { return 300 }
    }
}

function Show-ChangePlan {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Changes)

    if ($Changes.Count -eq 0) {
        return
    }

    Write-Host ''
    Write-Host 'Preview' -ForegroundColor Cyan
    Write-Host '-------' -ForegroundColor Cyan
    $currentTarget = $null
    foreach ($change in ($Changes | Sort-Object `
        Target, `
        @{ Expression = { Get-ComponentDisplayOrder -Component $_.Component } }, `
        Component, `
        @{ Expression = { Get-SettingDisplayOrder -Setting $_.Setting } }, `
        Setting)) {
        if ($currentTarget -ne $change.Target) {
            $currentTarget = $change.Target
            Write-Host ''
            Write-Host "[$currentTarget]" -ForegroundColor Yellow
        }
        $displayComponent = ($change.Component -replace ' \(Review Only\)$','')
        Write-Host ("{0} - {1}" -f $displayComponent, $change.Setting)
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
        [Parameter(Mandatory = $true)][string]$ComponentKey,
        [string]$PropertyName
    )
    $matches = @($DesiredConfiguration.Components | Where-Object { $_.Key -eq $ComponentKey })
    if (-not [string]::IsNullOrWhiteSpace($PropertyName)) {
        $propertyMatches = @($matches | Where-Object { $_.PSObject.Properties[$PropertyName] })
        if ($propertyMatches.Count -gt 0) { $matches = $propertyMatches }
    }
    $apply = @($matches | Where-Object { $_.Policy -eq 'Apply' } | Select-Object -First 1)
    if ($apply.Count -gt 0) { return $apply[0] }
    return $matches | Select-Object -First 1
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
        $candidate = "ExchangeURLManager-Commands-{0}.txt" -f (Get-Date -Format 'yyyyMMdd-HHmmss')
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

    if (Test-Path -LiteralPath $candidate) {
        throw "OutputFile already exists: $candidate"
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
        $value = $Parameters[$name]

        # Common switch parameters must use colon syntax when an explicit
        # Boolean value is emitted. Example: -Confirm:$false.
        if ($name -eq 'Confirm' -and $value -is [bool]) {
            [void]$parts.Add(("-{0}:{1}" -f $name, (ConvertTo-PowerShellLiteral -Value $value)))
            continue
        }

        [void]$parts.Add("-$name")
        [void]$parts.Add((ConvertTo-PowerShellLiteral -Value $value))
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

    $changes = @(Get-PlanForComponent -Plan $Plan -ComponentKey $MapEntry.Key)
    if ($changes.Count -eq 0) { return $null }

    $targetObject = Get-SnapshotComponent -Snapshot $TargetSnapshot -Key $MapEntry.Key
    $params = [ordered]@{ Identity = $targetObject.Identity }
    foreach ($change in $changes) { $params[$change.Setting] = $change.NewValue }

    # MAPI requires IISAuthenticationMethods when URL values are changed.
    # Preserve the target value unless authentication is explicitly part of the change plan.
    if ($MapEntry.Key -eq 'Mapi' -and ($changes.Setting -contains 'InternalUrl' -or $changes.Setting -contains 'ExternalUrl')) {
        if (-not $params.Contains('IISAuthenticationMethods')) {
            $iisAuthProperty = $targetObject.PSObject.Properties['IISAuthenticationMethods']
            if ($null -eq $iisAuthProperty -or $null -eq $iisAuthProperty.Value) {
                throw "Cannot determine MAPI IISAuthenticationMethods for $($TargetSnapshot.Server)."
            }
            $params['IISAuthenticationMethods'] = $iisAuthProperty.Value
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

    $changes = @(Get-PlanForComponent -Plan $Plan -ComponentKey 'OutlookAny')
    if ($changes.Count -eq 0) { return $null }

    $desired = Get-DesiredComponent -DesiredConfiguration $DesiredConfiguration -ComponentKey 'OutlookAny' -PropertyName 'DefaultAuthenticationMethod'
    $params = [ordered]@{ Identity = $TargetSnapshot.OutlookAny.Identity }
    foreach ($change in $changes) { $params[$change.Setting] = $change.NewValue }

    $internalHostnameChanged = ($changes.Setting -contains 'InternalHostname')
    $externalHostnameChanged = ($changes.Setting -contains 'ExternalHostname')

    $defaultAuthProperty = $desired.PSObject.Properties['DefaultAuthenticationMethod']
    $defaultAuth = if ($defaultAuthProperty) { [string]$defaultAuthProperty.Value } else { $null }

    if (-not [string]::IsNullOrWhiteSpace($defaultAuth)) {
        # DefaultAuthenticationMethod updates internal/external/IIS auth together.
        # Do not combine it with the individual authentication parameters.
        foreach ($name in @('InternalClientAuthenticationMethod','ExternalClientAuthenticationMethod','IISAuthenticationMethods')) {
            if ($params.Contains($name)) { $params.Remove($name) }
        }
        $params['DefaultAuthenticationMethod'] = $defaultAuth
    }

    # Live validation confirmed that Set-OutlookAnywhere requires companion
    # parameters when hostname values are changed. Add only the companion values
    # required for the side being changed and preserve the target value unless the
    # plan already contains an explicit value.
    if ($internalHostnameChanged -and -not $params.Contains('InternalClientsRequireSsl')) {
        $property = $TargetSnapshot.OutlookAny.PSObject.Properties['InternalClientsRequireSsl']
        if ($null -eq $property -or $null -eq $property.Value) {
            throw "Cannot determine Outlook Anywhere (RPC over HTTP) InternalClientsRequireSsl for $($TargetSnapshot.Server)."
        }
        $params['InternalClientsRequireSsl'] = $property.Value
    }

    if ($externalHostnameChanged) {
        if (-not $params.Contains('ExternalClientsRequireSsl')) {
            $property = $TargetSnapshot.OutlookAny.PSObject.Properties['ExternalClientsRequireSsl']
            if ($null -eq $property -or $null -eq $property.Value) {
                throw "Cannot determine Outlook Anywhere (RPC over HTTP) ExternalClientsRequireSsl for $($TargetSnapshot.Server)."
            }
            $params['ExternalClientsRequireSsl'] = $property.Value
        }

        # ExternalHostname additionally requires either DefaultAuthenticationMethod
        # or ExternalClientAuthenticationMethod. If DefaultAuthenticationMethod was
        # explicitly requested above, do not add the individual auth parameter.
        if ([string]::IsNullOrWhiteSpace($defaultAuth) -and -not $params.Contains('ExternalClientAuthenticationMethod')) {
            $property = $TargetSnapshot.OutlookAny.PSObject.Properties['ExternalClientAuthenticationMethod']
            if ($null -eq $property -or $null -eq $property.Value) {
                throw "Cannot determine Outlook Anywhere (RPC over HTTP) ExternalClientAuthenticationMethod for $($TargetSnapshot.Server)."
            }
            $params['ExternalClientAuthenticationMethod'] = $property.Value
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
    $adfsEnabled = $false
    foreach ($component in @($DesiredConfiguration.Components | Where-Object { $_.Policy -eq 'Apply' -and $_.Key -in @('Owa','Ecp') })) {
        if ($component.Values.Contains('AdfsAuthentication')) {
            $adfsEnabled = $adfsEnabled -or [bool]$component.Values['AdfsAuthentication']
        }
    }

    # OWA/ECP are paired. Emit ECP first when ADFS is enabled so both sides stay aligned.
    $orderedKeys = if ($adfsEnabled) { @('Ecp','Owa') } else { @('Owa','Ecp') }
    $entries = @(
        foreach ($key in $orderedKeys) { $VirtualDirectoryMap | Where-Object { $_.Key -eq $key } | Select-Object -First 1 }
        $VirtualDirectoryMap | Where-Object { $_.Key -notin @('Owa','Ecp') }
    )

    foreach ($entry in $entries) {
        $changes = @(Get-PlanForComponent -Plan $applyPlan -ComponentKey $entry.Key)
        if ($changes.Count -eq 0) { continue }
        $params = Get-VirtualDirectoryCommandParameters -MapEntry $entry -TargetSnapshot $TargetSnapshot -Plan $applyPlan
        # The Manager already obtained one explicit operator confirmation before Apply.
        # Suppress per-cmdlet native confirmation prompts to avoid partial Apply when
        # an operator answers No/No to All part-way through the operation list.
        $params['Confirm'] = $false
        [void]$operations.Add([PSCustomObject]@{ CommandName = $entry.SetCmd; Parameters = $params; Changes = $changes; ReviewNote = $null })
    }

    $scpChanges = @(Get-PlanForComponent -Plan $applyPlan -ComponentKey 'ClientAccess')
    if ($scpChanges.Count -gt 0) {
        [void]$operations.Add([PSCustomObject]@{
            CommandName = 'Set-ClientAccessService'
            Parameters  = [ordered]@{ Identity = $TargetSnapshot.Server; AutoDiscoverServiceInternalUri = $scpChanges[0].NewValue; Confirm = $false }
            Changes     = $scpChanges
            ReviewNote  = $null
        })
    }

    $oaChanges = @(Get-PlanForComponent -Plan $applyPlan -ComponentKey 'OutlookAny')
    if ($oaChanges.Count -gt 0) {
        $oaParams = Get-OutlookAnywhereCommandParameters -TargetSnapshot $TargetSnapshot -DesiredConfiguration $DesiredConfiguration -Plan $applyPlan
        $oaParams['Confirm'] = $false
        [void]$operations.Add([PSCustomObject]@{
            CommandName = 'Set-OutlookAnywhere'
            Parameters  = $oaParams
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
        [string]$Mode = 'Configure',
        [string]$SourceDescription,
        [switch]$IncludeAuthentication
    )

    $lines = New-Object System.Collections.ArrayList
    foreach ($line in @(
        '# ExchangeURLManager.ps1 generated command file'
        '# Tool Author : Ceyhun Kirmizitas'
        "# Generated   : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        "# Operator    : $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
        "# Mode        : $Mode"
        "# Source      : $(if ($SourceDescription) { $SourceDescription } elseif ($SourceServer) { $SourceServer } else { '<none>' })"
        "# Targets     : $($Targets -join ', ')"
        "# Authentication: $(if ($IncludeAuthentication) { 'Included by explicit request' } else { 'Not changed' })"
        '# WARNING: Review every command before running it in Exchange Management Shell.'
        '# ExchangeURLManager.ps1 did not apply any Exchange changes.'
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
# Apply Readiness
# ---------------------------------------------------------------------------
# Some Exchange Set-* cmdlets can display an additional native confirmation
# when a proposed URL/hostname cannot be resolved. That prompt is separate from
# the manager's own Preview/Confirm flow and can leave a partially applied plan
# if the operator stops midway. Detect obvious unresolved hosts before Apply.
function Get-HostnameFromPlannedChange {
    param([Parameter(Mandatory = $true)]$Change)

    if ($Change.Policy -ne 'Apply') { return $null }
    if ($null -eq $Change.NewValue) { return $null }

    # Live Exchange validation showed that Set-WebServicesVirtualDirectory
    # raises a separate native continuation prompt when an InternalUrl or
    # ExternalUrl host cannot be resolved, even with Confirm:$false. Do not
    # generalize this blocker to unrelated URL/hostname settings that did not
    # raise that prompt during live validation.
    if ([string]$Change.Component -ne 'EWS') { return $null }
    if ([string]$Change.Setting -notin @('InternalUrl','ExternalUrl')) { return $null }

    $uri = $null
    if ([Uri]::TryCreate([string]$Change.NewValue, [UriKind]::Absolute, [ref]$uri)) {
        return $uri.DnsSafeHost.TrimEnd('.')
    }

    return $null
}

function Test-PlannedHostnameResolution {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Changes)

    $results = New-Object System.Collections.ArrayList
    $seen = @{}

    foreach ($change in $Changes) {
        $hostName = Get-HostnameFromPlannedChange -Change $change
        if ([string]::IsNullOrWhiteSpace($hostName)) { continue }

        $key = $hostName.ToLowerInvariant()
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true

        $ip = $null
        if ([System.Net.IPAddress]::TryParse($hostName, [ref]$ip)) {
            [void]$results.Add([PSCustomObject]@{
                HostName  = $hostName
                Status    = 'Resolved'
                Addresses = $hostName
                Detail    = 'IP address literal'
            })
            continue
        }

        try {
            $addresses = @([System.Net.Dns]::GetHostAddresses($hostName))
            if ($addresses.Count -eq 0) {
                throw "No addresses were returned."
            }

            [void]$results.Add([PSCustomObject]@{
                HostName  = $hostName
                Status    = 'Resolved'
                Addresses = (($addresses | ForEach-Object { $_.IPAddressToString } | Select-Object -Unique) -join ', ')
                Detail    = $null
            })
        }
        catch {
            [void]$results.Add([PSCustomObject]@{
                HostName  = $hostName
                Status    = 'BLOCKER'
                Addresses = $null
                Detail    = $_.Exception.Message
            })
        }
    }

    return @($results)
}

function Test-ApplyReadiness {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$ActionableChanges,
        [switch]$ReportOnly
    )

    $dnsResults = @(Test-PlannedHostnameResolution -Changes $ActionableChanges)
    if ($dnsResults.Count -eq 0) { return $true }

    Write-Host ''
    Write-Host 'Apply readiness' -ForegroundColor Cyan
    Write-Host '---------------' -ForegroundColor Cyan
    Write-Host 'Checking DNS resolution for proposed EWS URL hostnames...' -ForegroundColor DarkCyan

    foreach ($item in $dnsResults) {
        if ($item.Status -eq 'Resolved') {
            Write-Host ("  PASS    {0} -> {1}" -f $item.HostName, $item.Addresses) -ForegroundColor Green
        }
        else {
            Write-Host ("  BLOCKER {0} - cannot be resolved" -f $item.HostName) -ForegroundColor Red
        }
    }

    $blocked = @($dnsResults | Where-Object { $_.Status -eq 'BLOCKER' })
    if ($blocked.Count -eq 0) {
        Write-Host 'Apply readiness check passed.' -ForegroundColor Green
        return $true
    }

    Write-Host ''
    if ($ReportOnly) {
        Write-Warning 'Automatic Apply would be blocked because one or more proposed EWS URL hostnames cannot be resolved.'
        Write-Warning 'Set-WebServicesVirtualDirectory displays additional native confirmation prompts for unresolved URL hosts, even with Confirm:$false.'
        Write-Host 'Command export will continue because -OutputFile does not apply Exchange configuration changes.' -ForegroundColor Yellow
        Write-Host 'Review the generated command carefully before running it manually.' -ForegroundColor Yellow
        return $false
    }

    Write-Warning 'Automatic Apply is blocked because one or more proposed EWS URL hostnames cannot be resolved.'
    Write-Warning 'Set-WebServicesVirtualDirectory displays additional native confirmation prompts for unresolved URL hosts, even with Confirm:$false. This can interrupt the plan and leave a partially applied configuration.'
    Write-Host 'Create or verify the required DNS records, then run the Preview again.' -ForegroundColor Yellow
    Write-Host 'If unresolved EWS values are intentional, use -OutputFile and review/run the generated commands manually.' -ForegroundColor Yellow
    Write-Host 'No Exchange configuration changes were applied.' -ForegroundColor Green
    return $false
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

    $results = New-Object System.Collections.ArrayList
    foreach ($targetServer in $Targets) {
        $targetPlan = @($Changes | Where-Object { $_.Target -eq $targetServer })
        if ($targetPlan.Count -eq 0) { continue }
        Write-Host "Configuring $targetServer..." -ForegroundColor Cyan
        try {
            # Review Only items are intentionally excluded from Apply. They were
            # already displayed in the preview and are never changed automatically.
            $operations = @(New-TargetOperations -TargetSnapshot $TargetSnapshots[$targetServer] -DesiredConfiguration $DesiredByTarget[$targetServer] -Plan $targetPlan)
            foreach ($operation in $operations) {
                Invoke-SetCommand -CommandName $operation.CommandName -Parameters $operation.Parameters
            }
            [void]$results.Add([PSCustomObject]@{ Server = $targetServer; Status = 'Applied'; Detail = $null })
        }
        catch {
            [void]$results.Add([PSCustomObject]@{ Server = $targetServer; Status = 'Failed'; Detail = $_.Exception.Message })
            Write-Warning "[$targetServer] Apply failed. Remaining operations for this target were skipped."
        }
    }
    return @($results)
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
            Write-Host ("[{0}] Reading configuration for verification..." -f $targetServer) -ForegroundColor DarkCyan
            $snapshot = Get-ExchangeClientAccessSnapshot -Server $targetServer
            $remaining = @(New-ChangePlan -TargetSnapshot $snapshot -DesiredConfiguration $DesiredByTarget[$targetServer] -ApplyOnly)
            [void]$results.Add([PSCustomObject]@{
                Server           = $targetServer
                Status           = $(if ($remaining.Count -eq 0) { 'Verified' } else { 'Mismatch' })
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

function Get-NormalizedServers {
    param([string[]]$Servers)

    $seen = @{}
    $count = 0

    foreach ($server in $Servers) {
        if ($null -eq $server) { continue }

        $normalized = ([string]$server).Trim()
        if ([string]::IsNullOrWhiteSpace($normalized)) { continue }
        if ($seen.ContainsKey($normalized)) { continue }

        $seen[$normalized] = $true
        $count++

        # Emit plain strings to the pipeline. Callers wrap the function call
        # in @() when they need a stable collection.
        Write-Output $normalized
    }

    if ($count -eq 0) {
        throw 'At least one server is required.'
    }
}

# Main
$isCloneComparison = ($PSCmdlet.ParameterSetName -eq 'Clone' -and -not $ApplyChanges)
$includePowerShellUrlsRequested = [bool]$IncludePowerShellUrls.IsPresent
$includeAuthenticationRequested = [bool]$IncludeAuthentication.IsPresent
$includeOutlookAnywhereSslRequirementsRequested = [bool]$IncludeOutlookAnywhereSslRequirements.IsPresent

if ($PSCmdlet.ParameterSetName -eq 'Clone' -and $CompareOnly -and $ApplyChanges) {
    Write-Host ''
    Write-Host 'BLOCKER: -CompareOnly and -ApplyChanges cannot be used together.' -ForegroundColor Red
    Write-Host 'Choose either read-only comparison or explicit Clone Apply.' -ForegroundColor Yellow
    Write-Host ''
    Write-Host 'No Exchange configuration changes were made.' -ForegroundColor Green
    Write-Host ''
    Show-FeedbackFooter
    return
}

$startupContext = Get-StartupContext `
    -ParameterSetName $PSCmdlet.ParameterSetName `
    -CompareOnly:$CompareOnly `
    -ApplyChanges:$ApplyChanges `
    -OutputFile $OutputFile

Show-StartupSummary -Context $startupContext

Write-Host 'Checking Exchange Management Shell...' -ForegroundColor DarkCyan
Initialize-ExchangeShell
Write-Host 'Exchange Management Shell ready.' -ForegroundColor DarkCyan
Write-Host ''

if ($PSCmdlet.ParameterSetName -eq 'Review') {
    $targets = @(Get-NormalizedServers -Servers $Server)
    $reviewData = New-Object System.Collections.ArrayList
    $reportLines = New-Object System.Collections.ArrayList

    [void]$reportLines.Add('ExchangeURLManager.ps1 review report')
    [void]$reportLines.Add(("Generated : {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')))
    [void]$reportLines.Add('Mode      : Review (read-only)')
    [void]$reportLines.Add(("Authentication details : {0}" -f $(if ($IncludeAuthentication) { 'Included' } else { 'Hidden by default; use -IncludeAuthentication to display' })))
    [void]$reportLines.Add('')

    Write-Host "Servers: $($targets -join ', ')" -ForegroundColor Cyan
    Write-Host ("Authentication details : {0}" -f $(if ($IncludeAuthentication) { 'Included' } else { 'Hidden (use -IncludeAuthentication to display)' })) -ForegroundColor Cyan
    Write-Host 'Reading current Exchange configuration... This may take a while. Please wait.' -ForegroundColor DarkCyan

    foreach ($target in $targets) {
        Write-Host ("[{0}] Reading current Exchange configuration... Please wait." -f $target) -ForegroundColor DarkCyan
        [void]$reviewData.Add((Get-ReviewDataForServer -Server $target -IncludeAuthentication:$IncludeAuthentication))
    }

    $groupedLines = @(Get-GroupedReviewLines -ReviewData @($reviewData) -IncludeAuthentication:$IncludeAuthentication)

    Write-Host ''
    foreach ($line in $groupedLines) {
        Write-ReviewConsoleLine -Line $line
        [void]$reportLines.Add($line)
    }

    if (-not [string]::IsNullOrWhiteSpace($OutputFile)) {
        $outputPath = Resolve-OutputFilePath -Path $OutputFile
        Set-Content -LiteralPath $outputPath -Value @($reportLines) -Encoding UTF8 -ErrorAction Stop
        Write-Host ''
        Write-Host "Review report created: $outputPath" -ForegroundColor Green
    }

    Write-Host ''
    Write-Host 'Review completed.' -ForegroundColor Green
    Write-Host 'No Exchange configuration changes were made.' -ForegroundColor DarkGray
    Show-FeedbackFooter
    return
}

if ($PSCmdlet.ParameterSetName -eq 'Backup') {
    $targets = @(Get-NormalizedServers -Servers $Server)
    $results = New-Object System.Collections.ArrayList

    Write-Host "Servers: $($targets -join ', ')" -ForegroundColor Cyan

    foreach ($target in $targets) {
        try {
            Write-Host ("[{0}] Reading current configuration for backup..." -f $target) -ForegroundColor DarkCyan
            $snapshot = Get-ExchangeClientAccessSnapshot -Server $target
            Write-Host ("[{0}] Writing manual JSON backup and TXT companion..." -f $target) -ForegroundColor DarkCyan
            $path = Export-ExchangeURLManagerBackup -Snapshot $snapshot -BackupType 'Manual' -Path $BackupPath -Context 'Manual Backup' -ServerCount $targets.Count
            Write-Host ("[{0}] Backup created: {1}" -f $snapshot.Server, $path) -ForegroundColor Green
            [void]$results.Add([PSCustomObject]@{ Server = $snapshot.Server; Status = 'Backed up'; Path = $path; Detail = $null })
        }
        catch {
            [void]$results.Add([PSCustomObject]@{ Server = $target; Status = 'Failed'; Path = $null; Detail = $_.Exception.Message })
            Write-Warning "[$target] Backup failed: $($_.Exception.Message)"
        }
    }

    Write-Host ''
    Write-Host 'Backup summary' -ForegroundColor Cyan
    Write-Host '--------------' -ForegroundColor Cyan
    $results | Format-Table Server,Status,Path -AutoSize
    if (@($results | Where-Object { $_.Status -ne 'Backed up' }).Count -gt 0) {
        Write-Warning 'One or more server backups failed. Successful backup files were kept.'
    }
    Show-FeedbackFooter
    return
}

$sourceSnapshot = $null
$restoreDocument = $null
$restoreBackupPath = $null
$interactiveSpec = $null
$targets = @()
$targetSnapshots = @{}
$desiredByTarget = @{}
$allChanges = New-Object System.Collections.ArrayList
$allComparisons = New-Object System.Collections.ArrayList
$comparisonFailures = New-Object System.Collections.ArrayList

if ($PSCmdlet.ParameterSetName -eq 'Interactive') {
    $targets = @(Read-InteractiveServers -ProvidedServers $Server)

    Write-Host "Servers: $($targets -join ', ')" -ForegroundColor Cyan
    Write-Host 'No Exchange configuration changes are made while collecting Interactive input.' -ForegroundColor DarkCyan

    $interactiveSpec = Read-InteractiveConfigureSpec -Targets $targets

    Write-Host ''
    Write-Host "Interactive style : $($interactiveSpec.Style)" -ForegroundColor Cyan
    $selectedNames = New-Object System.Collections.ArrayList
    foreach ($key in @($interactiveSpec.ComponentActions.Keys)) {
        if ($key -eq 'OutlookAny') {
            [void]$selectedNames.Add('Outlook Anywhere (RPC over HTTP)')
        }
        else {
            $entry = $VirtualDirectoryMap | Where-Object { $_.Key -eq $key } | Select-Object -First 1
            if ($entry) { [void]$selectedNames.Add($entry.Name) }
        }
    }
    if ($null -ne $interactiveSpec.ScpAction) { [void]$selectedNames.Add('Autodiscover SCP') }
    Write-Host "Selected          : $(if ($selectedNames.Count -gt 0) { $selectedNames -join ', ' } else { '<none>' })" -ForegroundColor Cyan

    if ($interactiveSpec.PowerShellSelected) {
        Write-Warning 'PowerShell InternalUrl/ExternalUrl are selected. PowerShell authentication, RequireSSL, and Extended Protection remain Review Only.'
    }

}
elseif ($PSCmdlet.ParameterSetName -eq 'Clone') {
    $SourceServer = $SourceServer.Trim()
    $targets = @(Get-NormalizedServers -Servers $TargetServer)

    if ($targets -contains $SourceServer) {
        Write-Host ("BLOCKER: Source server {0} cannot also be listed in -TargetServer." -f $SourceServer) -ForegroundColor Red
        Write-Host ("Remove {0} from -TargetServer and run again." -f $SourceServer) -ForegroundColor Yellow
        Write-Host ''
        Write-Host 'No Exchange configuration changes were made.' -ForegroundColor DarkGray
        Show-FeedbackFooter
        return
    }

    Write-Host "Source server : $SourceServer" -ForegroundColor Cyan
    $targetLabel = if ($targets.Count -eq 1) { 'Target server ' } else { 'Target servers' }
    Write-Host ("{0}: {1}" -f $targetLabel, ($targets -join ', ')) -ForegroundColor Cyan
    Write-Host "Authentication : $(if ($includeAuthenticationRequested) { 'Included' } else { 'Not included' })" -ForegroundColor Cyan

    if ($isCloneComparison) {
        Write-Host 'Comparison    : Read-only' -ForegroundColor Cyan
    }
    else {
        Write-Host "Outlook Anywhere (RPC over HTTP) SSL settings: $(if ($includeOutlookAnywhereSslRequirementsRequested) { 'Included by explicit request' } else { 'Preserve target values' })" -ForegroundColor Cyan
        Write-Host "PowerShell URLs: $(if ($includePowerShellUrlsRequested) { 'Included by explicit request' } else { 'Review Only' })" -ForegroundColor $(if ($includePowerShellUrlsRequested) { 'Yellow' } else { 'Cyan' })

        if ($includePowerShellUrlsRequested) {
            Write-Host 'PowerShell InternalUrl and ExternalUrl will be included in Apply.' -ForegroundColor Yellow
            Write-Host 'PowerShell authentication, RequireSSL, and Extended Protection remain Review Only.' -ForegroundColor DarkYellow
        }
    }

    Write-Host ''
    Write-Host 'Reading source Exchange values... This may take a while. Please wait.' -ForegroundColor DarkCyan

    try {
        $sourceSnapshot = Get-ExchangeClientAccessSnapshot -Server $SourceServer
    }
    catch {
        Write-Host ("BLOCKER: Unable to read source Exchange server {0}." -f $SourceServer) -ForegroundColor Red
        Write-Host ("Detail  : {0}" -f $_.Exception.Message) -ForegroundColor DarkYellow
        if ($isCloneComparison) {
            Write-Host 'Comparison cannot continue without a valid source server.' -ForegroundColor Yellow
        }
        else {
            Write-Host 'Clone cannot continue without a valid source server.' -ForegroundColor Yellow
        }
        Write-Host ''
        Write-Host 'No Exchange configuration changes were made.' -ForegroundColor DarkGray
        Show-FeedbackFooter
        return
    }
}
elseif ($PSCmdlet.ParameterSetName -eq 'Restore') {
    Write-Host ''
    Write-Host 'Loading and validating Restore backup...' -ForegroundColor DarkCyan
    $restoreBackupPath = [System.IO.Path]::GetFullPath($BackupFile)
    $restoreDocument = Import-ExchangeURLManagerBackup -Path $restoreBackupPath
    $restoreTarget = Assert-SameServerRestore -BackupDocument $restoreDocument
    $targets = @($restoreTarget)

    Write-Host "Backup file   : $restoreBackupPath" -ForegroundColor Cyan
    Write-Host "Target server : $restoreTarget" -ForegroundColor Cyan
    Write-Host "Authentication : $(if ($IncludeAuthentication) { 'Included by explicit request' } else { 'Not changed' })" -ForegroundColor Cyan
    Write-Host "Outlook Anywhere (RPC over HTTP) SSL settings: $(if ($includeOutlookAnywhereSslRequirementsRequested) { 'Included by explicit request' } else { 'Preserve target values' })" -ForegroundColor Cyan
    Write-Host "PowerShell URLs: $(if ($includePowerShellUrlsRequested) { 'Included by explicit request' } else { 'Review Only' })" -ForegroundColor $(if ($includePowerShellUrlsRequested) { 'Yellow' } else { 'Cyan' })
    if ($includePowerShellUrlsRequested) {
        Write-Host 'PowerShell InternalUrl and ExternalUrl will be included in Restore.' -ForegroundColor Yellow
        Write-Host 'PowerShell authentication, RequireSSL, and Extended Protection remain Review Only.' -ForegroundColor DarkYellow
    }
}
else {
    $targets = if ($Server) { @(Get-NormalizedServers -Servers $Server) } else { @($env:COMPUTERNAME) }

    if ([string]::IsNullOrWhiteSpace($InternalNamespace)) {
        $InternalNamespace = Normalize-Namespace -Value (Read-Host 'Internal namespace (example: mail.contoso.com)')
    }
    else { $InternalNamespace = Normalize-Namespace -Value $InternalNamespace }

    if ($ClearExternalUrls -and -not [string]::IsNullOrWhiteSpace($ExternalNamespace)) {
        throw 'Use either -ExternalNamespace or -ClearExternalUrls, not both.'
    }

    if ($ClearExternalUrls) {
        $ExternalNamespace = $null
    }
    elseif ([string]::IsNullOrWhiteSpace($ExternalNamespace)) {
        $externalInput = Read-Host "External namespace ($InternalNamespace, NONE = clear external URLs)"
        if ($externalInput -match '^(?i:none|null|clear)$') {
            $ClearExternalUrls = $true
            $ExternalNamespace = $null
        }
        else {
            $ExternalNamespace = if ([string]::IsNullOrWhiteSpace($externalInput)) { $InternalNamespace } else { Normalize-Namespace -Value $externalInput }
        }
    }
    else { $ExternalNamespace = Normalize-Namespace -Value $ExternalNamespace }

    $suggestedAutodiscover = Get-SuggestedAutodiscoverNamespace -ClientAccessNamespace $InternalNamespace
    if ([string]::IsNullOrWhiteSpace($AutodiscoverSCPNamespace)) {
        $autodiscoverInput = Read-Host "Autodiscover SCP namespace (Enter = $suggestedAutodiscover)"
        $AutodiscoverSCPNamespace = if ([string]::IsNullOrWhiteSpace($autodiscoverInput)) { $suggestedAutodiscover } else { Normalize-Namespace -Value $autodiscoverInput }
    }
    else { $AutodiscoverSCPNamespace = Normalize-Namespace -Value $AutodiscoverSCPNamespace }

    Write-Host "Internal namespace        : $InternalNamespace" -ForegroundColor Cyan
    Write-Host "External namespace        : $(if ($ClearExternalUrls) { '<clear external URLs>' } else { $ExternalNamespace })" -ForegroundColor Cyan
    Write-Host "Autodiscover SCP namespace: $AutodiscoverSCPNamespace" -ForegroundColor Cyan
    Write-Host "Servers                   : $($targets -join ', ')" -ForegroundColor Cyan
    Write-Host "PowerShell URLs           : $(if ($includePowerShellUrlsRequested) { 'Included by explicit request' } else { 'Review Only' })" -ForegroundColor $(if ($includePowerShellUrlsRequested) { 'Yellow' } else { 'Cyan' })
    Write-Host "Outlook Anywhere (RPC over HTTP) Internal SSL: $(if ($null -ne $OutlookAnywhereInternalClientsRequireSsl) { [bool]$OutlookAnywhereInternalClientsRequireSsl } else { '<preserve current>' })" -ForegroundColor Cyan
    Write-Host "Outlook Anywhere (RPC over HTTP) External SSL: $(if ($null -ne $OutlookAnywhereExternalClientsRequireSsl) { [bool]$OutlookAnywhereExternalClientsRequireSsl } else { '<preserve current>' })" -ForegroundColor Cyan
    Write-Host "Outlook Anywhere (RPC over HTTP) Default Auth: $(if ($OutlookAnywhereDefaultAuthenticationMethod) { $OutlookAnywhereDefaultAuthenticationMethod } else { '<not changed>' })" -ForegroundColor Cyan
}

Write-Host ''
if ($isCloneComparison) {
    Write-Host 'Preparing comparison...' -ForegroundColor DarkCyan
    Write-Host 'Reading current target Exchange values... This may take a while. Please wait.' -ForegroundColor DarkCyan
}
elseif ($PSCmdlet.ParameterSetName -eq 'Interactive') {
    Write-Host 'Preparing Preview...' -ForegroundColor DarkCyan
    Write-Host 'Refreshing current Exchange values... This may take a while. Please wait.' -ForegroundColor DarkCyan
}
else {
    Write-Host 'Preparing Preview...' -ForegroundColor DarkCyan
    Write-Host 'Reading current Exchange values... This may take a while. Please wait.' -ForegroundColor DarkCyan
}

foreach ($target in $targets) {
    $readAction = if ($PSCmdlet.ParameterSetName -eq 'Interactive') { 'Refreshing' } else { 'Reading' }
    Write-Host ("[{0}] {1} current configuration..." -f $target, $readAction) -ForegroundColor DarkCyan

    if ($isCloneComparison) {
        try {
            $snapshot = Get-ExchangeClientAccessSnapshot -Server $target
            $targetSnapshots[$target] = $snapshot

            $desired = New-DesiredConfiguration `
                -TargetSnapshot $snapshot `
                -SourceSnapshot $sourceSnapshot `
                -CopyAuthentication $includeAuthenticationRequested `
                -IncludePowerShellUrls $includePowerShellUrlsRequested `
                -IncludeOutlookAnywhereSslRequirements $includeOutlookAnywhereSslRequirementsRequested `
                -CompareOutlookAnywhereSslRequirements:$true

            $desiredByTarget[$target] = $desired

            foreach ($comparison in @(New-ComparisonPlan -TargetSnapshot $snapshot -DesiredConfiguration $desired)) {
                [void]$allComparisons.Add($comparison)
            }
        }
        catch {
            [void]$comparisonFailures.Add([PSCustomObject]@{
                Target = $target
                Status = 'Unable to Read'
                Detail = $_.Exception.Message
            })

            Write-Host ("[{0}] BLOCKER: Unable to read or compare target Exchange configuration." -f $target) -ForegroundColor Red
            Write-Host ("[{0}] Detail : {1}" -f $target, $_.Exception.Message) -ForegroundColor DarkYellow
            Write-Host ("[{0}] Continuing with available comparison results." -f $target) -ForegroundColor Yellow
        }

        continue
    }

    $snapshot = Get-ExchangeClientAccessSnapshot -Server $target
    $targetSnapshots[$target] = $snapshot

    $desired = if ($PSCmdlet.ParameterSetName -eq 'Interactive') {
        New-InteractiveDesiredConfiguration -TargetSnapshot $snapshot -Spec $interactiveSpec
    }
    elseif ($PSCmdlet.ParameterSetName -eq 'Clone') {
        New-DesiredConfiguration -TargetSnapshot $snapshot -SourceSnapshot $sourceSnapshot -CopyAuthentication $includeAuthenticationRequested -IncludePowerShellUrls $includePowerShellUrlsRequested -IncludeOutlookAnywhereSslRequirements $includeOutlookAnywhereSslRequirementsRequested
    }
    elseif ($PSCmdlet.ParameterSetName -eq 'Restore') {
        New-RestoreDesiredConfiguration -TargetSnapshot $snapshot -BackupDocument $restoreDocument -CopyAuthentication $includeAuthenticationRequested -IncludePowerShellUrls $includePowerShellUrlsRequested -IncludeOutlookAnywhereSslRequirements $includeOutlookAnywhereSslRequirementsRequested
    }
    else {
        New-DesiredConfiguration -TargetSnapshot $snapshot -InternalNamespace $InternalNamespace -ExternalNamespace $ExternalNamespace -AutodiscoverNamespace $AutodiscoverSCPNamespace -IncludePowerShellUrls $includePowerShellUrlsRequested -ClearExternalUrls:$ClearExternalUrls -OutlookAnywhereInternalClientsRequireSsl $OutlookAnywhereInternalClientsRequireSsl -OutlookAnywhereExternalClientsRequireSsl $OutlookAnywhereExternalClientsRequireSsl -OutlookAnywhereDefaultAuthenticationMethod $OutlookAnywhereDefaultAuthenticationMethod
    }

    $desiredByTarget[$target] = $desired

    if ($PSCmdlet.ParameterSetName -eq 'Clone' -and $ApplyChanges -and $includePowerShellUrlsRequested) {
        $powerShellUrlComponent = @(
            $desired.Components |
                Where-Object {
                    $_.Key -eq 'PowerShell' -and
                    $_.Values.Contains('InternalUrl') -and
                    $_.Values.Contains('ExternalUrl')
                }
        ) | Select-Object -First 1

        if (-not $powerShellUrlComponent -or $powerShellUrlComponent.Policy -ne 'Apply') {
            throw 'Internal safety guard: -IncludePowerShellUrls was explicitly requested, but PowerShell InternalUrl/ExternalUrl were not marked Apply-capable.'
        }
    }

    foreach ($change in @(New-ChangePlan -TargetSnapshot $snapshot -DesiredConfiguration $desired)) {
        [void]$allChanges.Add($change)
    }
}

if ($isCloneComparison) {
    $enableComparisonPaging = [string]::IsNullOrWhiteSpace($OutputFile)
    $comparisonDisplayCompleted = Show-ComparisonPlan -Comparisons @($allComparisons) -SourceServer $SourceServer -TargetFailures @($comparisonFailures) -EnablePaging:$enableComparisonPaging

    if (-not $comparisonDisplayCompleted) {
        Write-Host ''
        Write-Host 'Comparison display stopped by operator.' -ForegroundColor DarkCyan
        Write-Host 'No Exchange configuration changes were made.' -ForegroundColor DarkGray
        Show-FeedbackFooter
        return
    }

    if (-not [string]::IsNullOrWhiteSpace($OutputFile)) {
        $outputPath = Resolve-OutputFilePath -Path $OutputFile
        Export-ComparisonReport -Path $outputPath -Comparisons @($allComparisons) -Targets $targets -SourceServer $SourceServer -TargetFailures @($comparisonFailures) -IncludeAuthentication:$IncludeAuthentication
        Write-Host ''
        Write-Host "Comparison report created: $outputPath" -ForegroundColor Green
    }

    Write-Host ''
    if ($comparisonFailures.Count -gt 0) {
        Write-Host 'Comparison completed with one or more target read failures.' -ForegroundColor Yellow
    }
    else {
        Write-Host 'Comparison completed.' -ForegroundColor Green
    }
    Write-Host 'No Exchange configuration changes were made.' -ForegroundColor DarkGray
    Show-FeedbackFooter
    return
}

Show-ChangePlan -Changes @($allChanges)

if ($PSCmdlet.ParameterSetName -eq 'Configure' -and -not [string]::IsNullOrWhiteSpace($OutlookAnywhereDefaultAuthenticationMethod)) {
    Write-Host ''
    Write-Host 'Outlook Anywhere (RPC over HTTP) authentication note' -ForegroundColor Yellow
    Write-Host '  DefaultAuthenticationMethod updates InternalClientAuthenticationMethod, ExternalClientAuthenticationMethod, and IISAuthenticationMethods together.' -ForegroundColor Yellow
}

if ($allChanges.Count -eq 0) {
    Write-Host 'No changes are required.' -ForegroundColor Green
    Write-Host 'Configuration is already in the desired state.' -ForegroundColor Green
    Show-FeedbackFooter
    return
}

$actionableChanges = @($allChanges | Where-Object { $_.Policy -eq 'Apply' })

if (-not [string]::IsNullOrWhiteSpace($OutputFile)) {
    # Report readiness findings even though command export is non-destructive.
    # BLOCKERs prevent automatic Apply, but they do not prevent -OutputFile.
    [void](Test-ApplyReadiness -ActionableChanges $actionableChanges -ReportOnly)

    $outputPath = Resolve-OutputFilePath -Path $OutputFile
    $modeName = $PSCmdlet.ParameterSetName
    $sourceDescription = if ($modeName -eq 'Restore') { $restoreBackupPath } elseif ($modeName -eq 'Clone') { $SourceServer } elseif ($modeName -eq 'Interactive') { '<Interactive Configure>' } else { '<Configure mode>' }
    Export-CommandPlan -Path $outputPath -Changes @($allChanges) -TargetSnapshots $targetSnapshots -DesiredByTarget $desiredByTarget -Targets $targets -SourceServer $SourceServer -Mode $modeName -SourceDescription $sourceDescription -IncludeAuthentication:$IncludeAuthentication

    Write-Host ''
    Write-Host "Command file created: $outputPath" -ForegroundColor Green
    Write-Host 'No Exchange configuration changes were applied. Review the TXT file and run the required commands manually.' -ForegroundColor Yellow
    return
}

if ($actionableChanges.Count -eq 0) {
    Write-Host 'No automatic changes are required. Review Only differences were shown above.' -ForegroundColor Yellow
    Show-FeedbackFooter
    return
}

# Defense in depth: Clone must never reach an Apply-capable path without the
# operator explicitly supplying -ApplyChanges.
if ($PSCmdlet.ParameterSetName -eq 'Clone' -and (-not $ApplyChanges -or $CompareOnly)) {
    throw 'Internal safety guard: Clone Apply requires explicit -ApplyChanges and cannot run with -CompareOnly.'
}

if (-not (Test-ApplyReadiness -ActionableChanges $actionableChanges)) {
    Show-FeedbackFooter
    return
}

$confirmation = Read-Host 'Continue and apply these changes? [Y/N]'
if ($confirmation -notmatch '^[Yy]$') {
    Write-Host 'Operation cancelled. No Exchange configuration changes were applied.' -ForegroundColor Yellow
    Show-FeedbackFooter
    return
}

$backupContext = if ($PSCmdlet.ParameterSetName -eq 'Interactive') {
    'Pre-change before Interactive Configure'
}
elseif ($PSCmdlet.ParameterSetName -eq 'Clone') {
    "Pre-change before Clone from $SourceServer"
}
elseif ($PSCmdlet.ParameterSetName -eq 'Restore') {
    "Pre-change before Restore from $restoreBackupPath"
}
else {
    'Pre-change before Configure'
}

$preChangeBackups = @(Export-PreChangeBackups -Targets $targets -Context $backupContext)
Write-Host ''
Write-Host 'Pre-change JSON backup(s) created:' -ForegroundColor Green
foreach ($path in $preChangeBackups) { Write-Host "  $path" -ForegroundColor Green }

Write-Host ''
Write-Host 'Applying changes...' -ForegroundColor Cyan
$applyResults = @(Invoke-ConfigurationPlan -Changes @($allChanges) -TargetSnapshots $targetSnapshots -DesiredByTarget $desiredByTarget -Targets $targets)

Write-Host ''
Write-Host 'Verifying changes... This may take a while. Please wait.' -ForegroundColor DarkCyan
$verification = @(Test-ConfigurationPlan -DesiredByTarget $desiredByTarget -Targets $targets)

Write-Host ''
Write-Host 'Apply summary' -ForegroundColor Cyan
Write-Host '-------------' -ForegroundColor Cyan
$applyResults | Format-Table Server,Status,Detail -AutoSize

Write-Host ''
Write-Host 'Verification summary' -ForegroundColor Cyan
Write-Host '--------------------' -ForegroundColor Cyan
$verification | Format-Table Server,Status,RemainingChanges -AutoSize

if (@($applyResults | Where-Object { $_.Status -ne 'Applied' }).Count -gt 0 -or
    @($verification | Where-Object { $_.Status -ne 'Verified' }).Count -gt 0) {
    throw 'One or more targets could not be fully verified after configuration. Review the summary above.'
}
else {
    Write-Host 'Configuration completed and verified.' -ForegroundColor Green
    Show-FeedbackFooter
}
