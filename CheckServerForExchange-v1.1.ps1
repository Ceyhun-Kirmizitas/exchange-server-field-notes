<#
.SYNOPSIS
Checks whether a Windows Server is ready for an Exchange Server Subscription Edition Mailbox role installation.

.DESCRIPTION
CheckServerForExchange.ps1 is intentionally read-only for server configuration.
It collects the current Windows Server state and reports PASS / BLOCKER / REVIEW results.
It never changes Windows Features, services, page file settings, network settings, registry configuration, regional settings, time zone,
security settings, or Exchange configuration.

The script checks the operating system, Windows PowerShell, domain and Active Directory reachability, Active Directory forest functional
level, primary DNS suffix/FQDN, network basics, IPv4/IPv6 state, Windows time synchronization, pending reboot state, memory/CPU and storage
guidance, page file configuration, .NET Framework, required Windows Features, Remote Registry, Visual C++ prerequisites, UCMA 4.0,
IIS URL Rewrite, Credential Guard, power plan, regional settings, time zone, and IE Enhanced Security Configuration (ESC) separately for
Administrators and Users.

When multiple servers are checked, the script also compares regional settings and time zone values across the servers so peer-server
differences are visible immediately.

Page file guidance is calculated from the server's installed RAM. Microsoft Exchange Server guidance is to set the paging file minimum and
maximum to the same value: 25% of installed memory.

For IPv6, the script treats the normal Windows default behavior as supported. If the environment specifically requires IPv4 preference,
the report also recognizes the Microsoft-documented optional setting:
  Registry path : HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters
  Value name    : DisabledComponents
  Type          : REG_DWORD
  Value         : 0x20 (decimal 32)
A restart is required after changing this value. The script does not make the change.

.PARAMETER Server
Optional server name or names to check. If omitted, the script checks the local server.
Remote checks use Windows PowerShell Remoting / WinRM with the current credentials. Each server is evaluated independently.
A remote connection failure is reported as BLOCKER for that server and does not stop checks for the remaining servers.

.PARAMETER OutputFile
Optional TXT report path. For multiple servers, all server results are written to one report. No server configuration is changed.

.PARAMETER Help
Displays a short usage guide and exits without running readiness checks.

.EXAMPLE
.\CheckServerForExchange-v1.1.ps1
Runs all checks and displays the results.

.EXAMPLE
.\CheckServerForExchange-v1.1.ps1 -Server EXSE01
Checks one remote server.

.EXAMPLE
.\CheckServerForExchange-v1.1.ps1 -Server EXSE01,EXSE02,EXSE03
Checks multiple remote servers. A failure on one server does not stop the remaining checks.

.EXAMPLE
.\CheckServerForExchange-v1.1.ps1 -Server EXSE01,EXSE02 -OutputFile C:\Temp\ExchangeServerCheck.txt
Checks multiple servers and writes all results to one TXT report.

.NOTES
Author  : Ceyhun Kirmizitas
Version : 1.1
Date    : 23/09/2026
Scope   : Exchange Server Subscription Edition Mailbox server readiness check
Shell   : Windows PowerShell 5.1
Mode    : Read-only server configuration check

Field design basis
------------------
Built around practical Exchange 2016 -> Exchange SE side-by-side migration work. Regional consistency is highlighted because differences
between newly built Exchange servers can create difficult-to-diagnose behavior. Environment-specific network and security settings are
reported rather than changed.

Microsoft guidance used by this script includes:
- Exchange Server 2019 and SE prerequisites:
  https://learn.microsoft.com/en-us/exchange/plan-and-deploy/prerequisites
- Exchange Server 2019 and SE system requirements:
  https://learn.microsoft.com/en-us/exchange/plan-and-deploy/system-requirements
- Exchange Server supportability matrix:
  https://learn.microsoft.com/en-us/exchange/plan-and-deploy/supportability-matrix
- Exchange Setup primary DNS suffix readiness:
  https://learn.microsoft.com/en-us/exchange/plan-and-deploy/deployment-ref/ms-exch-setupreadiness-fqdnmissing
- Windows IPv6 configuration guidance:
  https://learn.microsoft.com/en-us/troubleshoot/windows-server/networking/configure-ipv6-in-windows
- Windows IE Enhanced Security Configuration:
  https://learn.microsoft.com/en-us/windows-hardware/customize/desktop/unattend/microsoft-windows-ie-esc

Change log
----------
1.1 - 23/09/2026
- Standardized the canonical filename as CheckServerForExchange-v1.1.ps1.
- Added primary DNS suffix/FQDN validation.
- Added Active Directory forest functional level validation.
- Added Windows Time service/source visibility.
- Added CPU socket visibility against Exchange hardware guidance.
- Improved prerequisite package detection and reports the detected package versions.
- Corrected IPv6 evaluation so normal Windows default preference is not reported as a review item.
- Added automatic cross-server regional/time-zone consistency comparison.
- Removed hard-coded v1.0 labels from generated reports and console output.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false, DontShow = $true)]
    [bool]$InternalLocal = $false,

    [Parameter(Mandatory = $false)]
    [string[]]$Server,

    [Parameter(Mandatory = $false)]
    [string]$OutputFile,

    [Parameter(Mandatory = $false)]
    [switch]$Help
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$script:ScriptBaseName = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)

$script:ScriptVersion = '1.1'

if ($Help) {
    @"
CheckServerForExchange-v1.1.ps1
Exchange Server SE Windows Server readiness check

COMMON USAGE
  Check the local server:
    .\CheckServerForExchange-v1.1.ps1

  Check one remote server:
    .\CheckServerForExchange-v1.1.ps1 -Server EXSE01

  Check multiple servers and compare peer settings:
    .\CheckServerForExchange-v1.1.ps1 -Server EXSE01,EXSE02

  Save one combined TXT report:
    .\CheckServerForExchange-v1.1.ps1 -Server EXSE01,EXSE02 -OutputFile C:\Temp\ExchangeServerCheck.txt

RESULTS
  PASS     Expected / ready
  BLOCKER  Must be addressed before Exchange installation
  REVIEW   Validate against the intended environment/design

NOTES
  - Readiness check only. It does not change Windows or Exchange configuration.
  - Remote checks require PowerShell Remoting / WinRM.
  - For full help:
      Get-Help .\CheckServerForExchange-v1.1.ps1 -Full
"@ | Write-Host
    return
}

# Exchange SE Mailbox role Windows Features from current Microsoft prerequisites.
$script:DesktopExperienceFeatures = @(
    'Server-Media-Foundation',
    'NET-Framework-45-Core',
    'NET-Framework-45-ASPNET',
    'NET-WCF-HTTP-Activation45',
    'NET-WCF-Pipe-Activation45',
    'NET-WCF-TCP-Activation45',
    'NET-WCF-TCP-PortSharing45',
    'RPC-over-HTTP-proxy',
    'RSAT-Clustering',
    'RSAT-Clustering-CmdInterface',
    'RSAT-Clustering-Mgmt',
    'RSAT-Clustering-PowerShell',
    'WAS-Process-Model',
    'Web-Asp-Net45',
    'Web-Basic-Auth',
    'Web-Client-Auth',
    'Web-Digest-Auth',
    'Web-Dir-Browsing',
    'Web-Dyn-Compression',
    'Web-Http-Errors',
    'Web-Http-Logging',
    'Web-Http-Redirect',
    'Web-Http-Tracing',
    'Web-ISAPI-Ext',
    'Web-ISAPI-Filter',
    'Web-Metabase',
    'Web-Mgmt-Console',
    'Web-Mgmt-Service',
    'Web-Net-Ext45',
    'Web-Request-Monitor',
    'Web-Server',
    'Web-Stat-Compression',
    'Web-Static-Content',
    'Web-Windows-Auth',
    'Web-WMI',
    'Windows-Identity-Foundation',
    'RSAT-ADDS'
)

$script:ServerCoreFeatures = @(
    'Server-Media-Foundation',
    'NET-Framework-45-Core',
    'NET-Framework-45-ASPNET',
    'NET-WCF-HTTP-Activation45',
    'NET-WCF-Pipe-Activation45',
    'NET-WCF-TCP-Activation45',
    'NET-WCF-TCP-PortSharing45',
    'RPC-over-HTTP-proxy',
    'RSAT-Clustering',
    'RSAT-Clustering-CmdInterface',
    'RSAT-Clustering-PowerShell',
    'WAS-Process-Model',
    'Web-Asp-Net45',
    'Web-Basic-Auth',
    'Web-Client-Auth',
    'Web-Digest-Auth',
    'Web-Dir-Browsing',
    'Web-Dyn-Compression',
    'Web-Http-Errors',
    'Web-Http-Logging',
    'Web-Http-Redirect',
    'Web-Http-Tracing',
    'Web-ISAPI-Ext',
    'Web-ISAPI-Filter',
    'Web-Metabase',
    'Web-Mgmt-Service',
    'Web-Net-Ext45',
    'Web-Request-Monitor',
    'Web-Server',
    'Web-Stat-Compression',
    'Web-Static-Content',
    'Web-Windows-Auth',
    'Web-WMI',
    'RSAT-ADDS'
)

# ---------------------------------------------------------------------------
# Generic helpers
# ---------------------------------------------------------------------------
function Test-IsAdministrator {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch { return $false }
}

function ConvertTo-DisplayText {
    param($Value)

    if ($null -eq $Value) { return '<Not Set>' }
    if ($Value -is [System.Array]) {
        if ($Value.Count -eq 0) { return '<None>' }
        return (@($Value | ForEach-Object { [string]$_ }) -join ', ')
    }
    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return '<Empty>' }
    return $text
}

function New-CheckResult {
    param(
        [Parameter(Mandatory = $true)][string]$Category,
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][ValidateSet('PASS','BLOCKER','REVIEW')][string]$Status,
        $Current,
        $Expected,
        [string]$Message = ''
    )

    [PSCustomObject]@{
        Category = $Category
        Name     = $Name
        Status   = $Status
        Current  = ConvertTo-DisplayText $Current
        Expected = ConvertTo-DisplayText $Expected
        Message  = $Message
    }
}

function Resolve-OutputFilePath {
    param([Parameter(Mandatory = $true)][string]$Path)

    $expanded = [Environment]::ExpandEnvironmentVariables($Path)
    if (-not [System.IO.Path]::IsPathRooted($expanded)) {
        $expanded = Join-Path (Get-Location).Path $expanded
    }

    $directory = Split-Path -Parent $expanded
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    return [System.IO.Path]::GetFullPath($expanded)
}

function Get-InstalledApplications {
    $paths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    $items = New-Object System.Collections.Generic.List[object]
    foreach ($path in $paths) {
        try {
            foreach ($item in @(Get-ItemProperty -Path $path -ErrorAction SilentlyContinue)) {
                if (-not [string]::IsNullOrWhiteSpace([string]$item.DisplayName)) {
                    [void]$items.Add([PSCustomObject]@{
                        DisplayName    = [string]$item.DisplayName
                        DisplayVersion = [string]$item.DisplayVersion
                        Publisher      = [string]$item.Publisher
                        UninstallKey   = [string]$item.PSChildName
                    })
                }
            }
        }
        catch { }
    }
    return @($items)
}


function Get-InstalledApplicationText {
    param(
        [Parameter(Mandatory = $true)]$State,
        [Parameter(Mandatory = $true)][string]$Pattern
    )

    $matches = @(
        $State.InstalledApps |
            Where-Object { $_.DisplayName -match $Pattern } |
            Sort-Object DisplayName,DisplayVersion -Unique
    )

    if ($matches.Count -eq 0) { return $null }

    return (($matches | ForEach-Object {
        if ([string]::IsNullOrWhiteSpace([string]$_.DisplayVersion)) {
            [string]$_.DisplayName
        }
        else {
            "{0} [{1}]" -f $_.DisplayName,$_.DisplayVersion
        }
    }) -join '; ')
}

function Get-ADFunctionalLevelInfo {
    $forestMode = $null
    $domainMode = $null
    $errorText = $null

    try {
        $forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
        $forestMode = [string]$forest.ForestMode
    }
    catch {
        $errorText = $_.Exception.Message
    }

    try {
        $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
        $domainMode = [string]$domain.DomainMode
    }
    catch {
        if ([string]::IsNullOrWhiteSpace($errorText)) {
            $errorText = $_.Exception.Message
        }
    }

    [PSCustomObject]@{
        ForestMode = $forestMode
        DomainMode = $domainMode
        Error      = $errorText
    }
}

function Get-TimeSyncInfo {
    $serviceState = 'Unknown'
    $source = $null
    $errorText = $null

    try {
        $service = Get-CimInstance Win32_Service -Filter "Name='W32Time'" -ErrorAction Stop
        $serviceState = "{0} / {1}" -f $service.StartMode,$service.State
    }
    catch {
        $errorText = $_.Exception.Message
    }

    try {
        $sourceOutput = @(& w32tm.exe /query /source 2>&1)
        if ($LASTEXITCODE -eq 0 -and $sourceOutput.Count -gt 0) {
            $source = (($sourceOutput | ForEach-Object { [string]$_ }) -join ' ').Trim()
        }
        elseif ([string]::IsNullOrWhiteSpace($errorText)) {
            $errorText = (($sourceOutput | ForEach-Object { [string]$_ }) -join ' ').Trim()
        }
    }
    catch {
        if ([string]::IsNullOrWhiteSpace($errorText)) {
            $errorText = $_.Exception.Message
        }
    }

    [PSCustomObject]@{
        ServiceState = $serviceState
        Source       = $source
        Error        = $errorText
    }
}

function Get-DotNetFrameworkInfo {
    $release = $null
    try {
        $release = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -Name Release -ErrorAction Stop).Release
    }
    catch { }

    $version = 'Not detected'
    if ($null -ne $release) {
        if ([int64]$release -ge 533320) { $version = '4.8.1' }
        elseif ([int64]$release -ge 528040) { $version = '4.8' }
        else { $version = "Pre-4.8 (Release $release)" }
    }

    [PSCustomObject]@{
        Release = $release
        Version = $version
    }
}

function Get-PendingRebootInfo {
    $reasons = New-Object System.Collections.Generic.List[string]

    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') {
        [void]$reasons.Add('Component Based Servicing')
    }
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') {
        [void]$reasons.Add('Windows Update')
    }
    try {
        $sessionManager = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction Stop
        if ($null -ne $sessionManager.PendingFileRenameOperations) { [void]$reasons.Add('Pending File Rename Operations') }
    }
    catch { }
    try {
        $computerName = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName' -ErrorAction Stop
        $activeName = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName' -ErrorAction Stop
        if ([string]$computerName.ComputerName -ne [string]$activeName.ComputerName) { [void]$reasons.Add('Computer rename') }
    }
    catch { }

    [PSCustomObject]@{
        Pending = ($reasons.Count -gt 0)
        Reasons = @($reasons)
    }
}

function Get-PowerPlanInfo {
    try {
        $text = (& powercfg.exe /GetActiveScheme 2>$null | Out-String).Trim()
        $guid = $null
        $name = $text
        if ($text -match '([0-9a-fA-F-]{36})') { $guid = $matches[1].ToLowerInvariant() }
        if ($text -match '\(([^)]+)\)') { $name = $matches[1] }
        return [PSCustomObject]@{ Name = $name; Guid = $guid; Raw = $text }
    }
    catch {
        return [PSCustomObject]@{ Name = 'Unknown'; Guid = $null; Raw = $_.Exception.Message }
    }
}

function Get-ActiveDirectorySiteInfo {
    $site = $null
    $dc = $null
    $gc = $null
    $siteError = $null
    $dcError = $null

    try {
        $siteOutput = @(& nltest.exe /dsgetsite 2>&1)
        if ($LASTEXITCODE -eq 0 -and $siteOutput.Count -gt 0) {
            $site = ([string]$siteOutput[0]).Trim()
        }
        else { $siteError = ($siteOutput -join ' ').Trim() }
    }
    catch { $siteError = $_.Exception.Message }

    try {
        $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        if ($cs.PartOfDomain -and -not [string]::IsNullOrWhiteSpace([string]$cs.Domain)) {
            $dcOutput = @(& nltest.exe "/dsgetdc:$($cs.Domain)" /WRITABLE 2>&1)
            if ($LASTEXITCODE -eq 0) {
                $dcLine = @($dcOutput | Where-Object { [string]$_ -match '(?i)DC:' } | Select-Object -First 1)
                if ($dcLine.Count -gt 0) { $dc = ([string]$dcLine[0] -replace '^\s*DC:\s*','').Trim('\ ') }
            }
            else { $dcError = ($dcOutput -join ' ').Trim() }

            $gcOutput = @(& nltest.exe "/dsgetdc:$($cs.Domain)" /GC /WRITABLE 2>&1)
            if ($LASTEXITCODE -eq 0) {
                $gcLine = @($gcOutput | Where-Object { [string]$_ -match '(?i)DC:' } | Select-Object -First 1)
                if ($gcLine.Count -gt 0) { $gc = ([string]$gcLine[0] -replace '^\s*DC:\s*','').Trim('\ ') }
            }
        }
    }
    catch { $dcError = $_.Exception.Message }

    [PSCustomObject]@{
        Site      = $site
        DC        = $dc
        GC        = $gc
        SiteError = $siteError
        DCError   = $dcError
    }
}

function Get-NetworkSummary {
    $items = New-Object System.Collections.Generic.List[object]
    try {
        # Include active team/vNIC interfaces too. -Physical can hide the interface that actually owns the server IP configuration.
        $adapters = @(Get-NetAdapter -ErrorAction Stop | Where-Object { $_.Status -eq 'Up' })
        foreach ($adapter in $adapters) {
            $ipConfig = Get-NetIPConfiguration -InterfaceIndex $adapter.ifIndex -ErrorAction SilentlyContinue
            $ipv4Binding = Get-NetAdapterBinding -Name $adapter.Name -ComponentID ms_tcpip -ErrorAction SilentlyContinue
            $ipv6Binding = Get-NetAdapterBinding -Name $adapter.Name -ComponentID ms_tcpip6 -ErrorAction SilentlyContinue
            $ipv4If = Get-NetIPInterface -InterfaceIndex $adapter.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue
            $dnsClient = Get-DnsClient -InterfaceIndex $adapter.ifIndex -ErrorAction SilentlyContinue

            $ipv4 = @()
            $ipv6 = @()
            $gateway = @()
            $dnsServers = @()
            $dnsSuffix = $null
            if ($ipConfig) {
                $ipv4 = @($ipConfig.IPv4Address | ForEach-Object { $_.IPAddress })
                $ipv6 = @($ipConfig.IPv6Address | Where-Object { $_.IPAddress -notlike 'fe80:*' } | ForEach-Object { $_.IPAddress })
                $gateway = @($ipConfig.IPv4DefaultGateway | ForEach-Object { $_.NextHop })
                if ($ipConfig.DNSServer) { $dnsServers = @($ipConfig.DNSServer.ServerAddresses) }
                if ($dnsClient) { $dnsSuffix = [string]$dnsClient.ConnectionSpecificSuffix }
            }

            [void]$items.Add([PSCustomObject]@{
                Name           = $adapter.Name
                InterfaceIndex = $adapter.ifIndex
                MacAddress     = $adapter.MacAddress
                IPv4           = $ipv4
                IPv6           = $ipv6
                Gateway        = $gateway
                DnsServers     = $dnsServers
                IPv4Enabled    = [bool]($ipv4Binding -and $ipv4Binding.Enabled)
                IPv6Enabled    = [bool]($ipv6Binding -and $ipv6Binding.Enabled)
                Dhcp           = if ($ipv4If) { [string]$ipv4If.Dhcp } else { 'Unknown' }
                DnsSuffix      = $dnsSuffix
            })
        }
    }
    catch { }
    return @($items)
}

function Get-IPv6PolicyInfo {
    $path = 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters'
    $exists = $false
    $value = 0
    try {
        $property = Get-ItemProperty -Path $path -Name DisabledComponents -ErrorAction Stop
        $exists = $true
        $value = [uint32]$property.DisabledComponents
    }
    catch { }

    $meaning = switch ($value) {
        0   { 'Default IPv6 behavior' }
        32  { 'Prefer IPv4 over IPv6 (0x20)' }
        255 { 'IPv6 disabled by policy (0xFF)' }
        default { ('Custom IPv6 policy (0x{0:X2})' -f $value) }
    }

    [PSCustomObject]@{
        RegistryValueExists = $exists
        DisabledComponents  = $value
        HexValue            = ('0x{0:X2}' -f $value)
        Meaning             = $meaning
    }
}

function Get-CredentialGuardInfo {
    $lsaCfgFlags = $null
    try {
        $lsaCfgFlags = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa' -Name LsaCfgFlags -ErrorAction Stop).LsaCfgFlags
    }
    catch { }

    $running = @()
    try {
        $deviceGuard = Get-CimInstance -ClassName Win32_DeviceGuard -Namespace root\Microsoft\Windows\DeviceGuard -ErrorAction Stop
        $running = @($deviceGuard.SecurityServicesRunning)
    }
    catch { }

    $enabled = (($null -ne $lsaCfgFlags -and [int]$lsaCfgFlags -ne 0) -or ($running -contains 1))
    [PSCustomObject]@{
        Enabled                 = $enabled
        LsaCfgFlags             = $lsaCfgFlags
        SecurityServicesRunning = $running
    }
}

function Get-IEEscInfo {
    param([string]$InstallationType)

    if ($InstallationType -match '(?i)core') {
        return [PSCustomObject]@{
            Applicable = $false
            AdministratorState = 'N/A (Server Core)'
            UserState          = 'N/A (Server Core)'
            AdministratorValue = $null
            UserValue          = $null
        }
    }

    $adminPath = 'HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}'
    $userPath  = 'HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}'

    $adminValue = $null
    $userValue = $null
    try { $adminValue = (Get-ItemProperty -LiteralPath $adminPath -Name IsInstalled -ErrorAction Stop).IsInstalled } catch { }
    try { $userValue = (Get-ItemProperty -LiteralPath $userPath -Name IsInstalled -ErrorAction Stop).IsInstalled } catch { }

    $adminState = if ($null -eq $adminValue) { 'Unknown' } elseif ([int]$adminValue -eq 1) { 'On' } elseif ([int]$adminValue -eq 0) { 'Off' } else { "Unknown ($adminValue)" }
    $userState  = if ($null -eq $userValue)  { 'Unknown' } elseif ([int]$userValue -eq 1)  { 'On' } elseif ([int]$userValue -eq 0)  { 'Off' } else { "Unknown ($userValue)" }

    [PSCustomObject]@{
        Applicable = $true
        AdministratorState = $adminState
        UserState          = $userState
        AdministratorValue = $adminValue
        UserValue          = $userValue
    }
}

function Get-ServerState {
    $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
    $registryOs = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop
    $pending = Get-PendingRebootInfo
    $network = @(Get-NetworkSummary)
    $ipv6Policy = Get-IPv6PolicyInfo
    $dotNet = Get-DotNetFrameworkInfo
    $apps = @(Get-InstalledApplications)
    $ad = Get-ActiveDirectorySiteInfo
    $power = Get-PowerPlanInfo
    $credentialGuard = Get-CredentialGuardInfo
    $ieEsc = Get-IEEscInfo -InstallationType ([string]$registryOs.InstallationType)
    $adFunctionalLevel = Get-ADFunctionalLevelInfo
    $timeSync = Get-TimeSyncInfo

    $primaryDnsSuffix = $null
    try {
        $primaryDnsSuffix = [System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties().DomainName
    }
    catch { }

    $features = @()
    try {
        Import-Module ServerManager -ErrorAction Stop
        $features = @(Get-WindowsFeature -ErrorAction Stop)
    }
    catch { }

    $remoteRegistry = $null
    try { $remoteRegistry = Get-CimInstance Win32_Service -Filter "Name='RemoteRegistry'" -ErrorAction Stop } catch { }

    $pageFileSettings = @()
    $pageFileUsage = @()
    try { $pageFileSettings = @(Get-CimInstance Win32_PageFileSetting -ErrorAction Stop) } catch { }
    try { $pageFileUsage = @(Get-CimInstance Win32_PageFileUsage -ErrorAction Stop) } catch { }

    $fixedDisks = @()
    try {
        $fixedDisks = @(Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction Stop | ForEach-Object {
            [PSCustomObject]@{
                DeviceID  = $_.DeviceID
                FileSystem = $_.FileSystem
                SizeGB    = if ($_.Size) { [math]::Round($_.Size / 1GB, 1) } else { 0 }
                FreeGB    = if ($_.FreeSpace) { [math]::Round($_.FreeSpace / 1GB, 1) } else { 0 }
            }
        })
    }
    catch { }

    $systemLocale = $null
    $culture = $null
    $uiCulture = $null
    $timeZone = $null
    try { $systemLocale = Get-WinSystemLocale } catch { }
    try { $culture = Get-Culture } catch { }
    try { $uiCulture = Get-UICulture } catch { }
    try { $timeZone = Get-TimeZone } catch { }
    [PSCustomObject]@{
        Timestamp            = Get-Date
        ComputerName         = [string]$env:COMPUTERNAME
        ComputerSystem       = $cs
        OperatingSystem      = $os
        RegistryOS           = $registryOs
        InstalledApps        = $apps
        WindowsFeatures      = $features
        RemoteRegistry       = $remoteRegistry
        PendingReboot        = $pending
        Network              = $network
        IPv6Policy           = $ipv6Policy
        DotNet               = $dotNet
        AD                    = $ad
        PowerPlan            = $power
        CredentialGuard      = $credentialGuard
        IEEsc                 = $ieEsc
        ADFunctionalLevel     = $adFunctionalLevel
        TimeSync              = $timeSync
        PrimaryDnsSuffix      = $primaryDnsSuffix
        PageFileSettings     = $pageFileSettings
        PageFileUsage        = $pageFileUsage
        FixedDisks           = $fixedDisks
        SystemLocale         = $systemLocale
        Culture              = $culture
        UICulture            = $uiCulture
        TimeZone             = $timeZone
        IsAdministrator      = Test-IsAdministrator
        AutomaticManagedPagefile = [bool]$cs.AutomaticManagedPagefile
    }
}

# ---------------------------------------------------------------------------
# Check evaluation
# ---------------------------------------------------------------------------
function Get-OsYear {
    param($State)
    $caption = [string]$State.OperatingSystem.Caption
    foreach ($year in @('2025','2022','2019')) {
        if ($caption -match $year) { return $year }
    }
    return 'Unknown'
}

function Get-RequiredWindowsFeatures {
    param($State)
    $installationType = [string]$State.RegistryOS.InstallationType
    if ($installationType -match '(?i)core') { return @($script:ServerCoreFeatures) }
    return @($script:DesktopExperienceFeatures)
}

function Get-ExchangePreparationChecks {
    param([Parameter(Mandatory = $true)]$State)

    $results = New-Object System.Collections.Generic.List[object]
    $osYear = Get-OsYear -State $State
    $caption = [string]$State.OperatingSystem.Caption
    $edition = [string]$State.RegistryOS.EditionID
    $installType = [string]$State.RegistryOS.InstallationType
    $architecture = [string]$State.OperatingSystem.OSArchitecture

    if ($State.IsAdministrator) {
        [void]$results.Add((New-CheckResult -Category 'Host' -Name 'Administrator' -Status PASS -Current 'Elevated' -Expected 'Run from an elevated Windows PowerShell session'))
    }
    else {
        [void]$results.Add((New-CheckResult -Category 'Host' -Name 'Administrator' -Status BLOCKER -Current 'Not elevated' -Expected 'Run as Administrator' -Message 'Some prerequisite and registry checks may be incomplete without elevation.'))
    }

    $psVersion = [string]$PSVersionTable.PSVersion
    $psEdition = if ($PSVersionTable.ContainsKey('PSEdition')) { [string]$PSVersionTable.PSEdition } else { 'Desktop' }
    if ($PSVersionTable.PSVersion.Major -eq 5 -and $psEdition -eq 'Desktop') {
        [void]$results.Add((New-CheckResult -Category 'Host' -Name 'Windows PowerShell' -Status PASS -Current "$psVersion / $psEdition" -Expected 'Windows PowerShell 5.1 (Windows-included version)'))
    }
    else {
        [void]$results.Add((New-CheckResult -Category 'Host' -Name 'Windows PowerShell' -Status BLOCKER -Current "$psVersion / $psEdition" -Expected 'Windows PowerShell 5.1 (Windows-included version)' -Message 'Run the check from Windows PowerShell 5.1. Exchange Server uses the Windows-included Windows PowerShell version.'))
    }

    $supportedOs = ($osYear -in @('2019','2022','2025')) -and ($edition -match '(?i)standard|datacenter') -and ($architecture -match '64')
    if ($supportedOs) {
        [void]$results.Add((New-CheckResult -Category 'Operating System' -Name 'Windows Server' -Status PASS -Current "$caption / $edition / $installType / $architecture" -Expected 'Windows Server 2019, 2022, or 2025 Standard/Datacenter x64'))
    }
    else {
        [void]$results.Add((New-CheckResult -Category 'Operating System' -Name 'Windows Server' -Status BLOCKER -Current "$caption / $edition / $installType / $architecture" -Expected 'Windows Server 2019, 2022, or 2025 Standard/Datacenter x64' -Message 'Exchange Server SE OS supportability check failed.'))
    }

    if ($State.ComputerSystem.PartOfDomain) {
        [void]$results.Add((New-CheckResult -Category 'Active Directory' -Name 'Domain Membership' -Status PASS -Current $State.ComputerSystem.Domain -Expected 'Domain member'))
    }
    else {
        [void]$results.Add((New-CheckResult -Category 'Active Directory' -Name 'Domain Membership' -Status BLOCKER -Current 'Workgroup' -Expected 'Domain member' -Message 'Join the server to the target Active Directory domain before Exchange installation.'))
    }

    $primaryDnsSuffix = [string]$State.PrimaryDnsSuffix
    if (-not [string]::IsNullOrWhiteSpace($primaryDnsSuffix)) {
        $fqdn = "{0}.{1}" -f $State.ComputerName,$primaryDnsSuffix
        [void]$results.Add((New-CheckResult -Category 'Active Directory' -Name 'Primary DNS Suffix / FQDN' -Status PASS -Current $fqdn -Expected 'Primary DNS suffix configured before Exchange installation'))
    }
    else {
        [void]$results.Add((New-CheckResult -Category 'Active Directory' -Name 'Primary DNS Suffix / FQDN' -Status BLOCKER -Current '<Missing>' -Expected 'Primary DNS suffix configured before Exchange installation' -Message 'Exchange Setup requires a valid server FQDN. Do not rename the server or change its primary DNS suffix after Exchange is installed.'))
    }

    $forestMode = [string]$State.ADFunctionalLevel.ForestMode
    if ($forestMode -match 'Windows2012R2Forest|Windows2016Forest') {
        [void]$results.Add((New-CheckResult -Category 'Active Directory' -Name 'Forest Functional Level' -Status PASS -Current $forestMode -Expected 'Windows Server 2012 R2 or Windows Server 2016 forest functional level'))
    }
    elseif (-not [string]::IsNullOrWhiteSpace($forestMode)) {
        [void]$results.Add((New-CheckResult -Category 'Active Directory' -Name 'Forest Functional Level' -Status BLOCKER -Current $forestMode -Expected 'Windows Server 2012 R2 or Windows Server 2016 forest functional level' -Message 'Exchange Server SE supports the forest functional levels listed in the current Exchange supportability matrix.'))
    }
    else {
        [void]$results.Add((New-CheckResult -Category 'Active Directory' -Name 'Forest Functional Level' -Status REVIEW -Current $State.ADFunctionalLevel.Error -Expected 'Windows Server 2012 R2 or Windows Server 2016 forest functional level' -Message 'The forest functional level could not be determined from this server.'))
    }

    if ([int]$State.ComputerSystem.DomainRole -ge 4) {
        [void]$results.Add((New-CheckResult -Category 'Active Directory' -Name 'Server Role' -Status BLOCKER -Current 'Domain Controller' -Expected 'Member Server' -Message 'Use a member server for Exchange.'))
    }
    else {
        [void]$results.Add((New-CheckResult -Category 'Active Directory' -Name 'Server Role' -Status PASS -Current 'Member Server' -Expected 'Member Server'))
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$State.AD.Site)) {
        [void]$results.Add((New-CheckResult -Category 'Active Directory' -Name 'AD Site' -Status PASS -Current $State.AD.Site -Expected 'Resolved AD site'))
    }
    else {
        [void]$results.Add((New-CheckResult -Category 'Active Directory' -Name 'AD Site' -Status BLOCKER -Current $State.AD.SiteError -Expected 'Resolved AD site' -Message 'Verify AD subnet/site mapping and domain connectivity.'))
    }

    if ($State.AD.DC -and $State.AD.GC) {
        [void]$results.Add((New-CheckResult -Category 'Active Directory' -Name 'Writable DC / GC' -Status PASS -Current ("DC={0}; GC={1}" -f $State.AD.DC,$State.AD.GC) -Expected 'Writable DC and GC reachable'))
    }
    else {
        [void]$results.Add((New-CheckResult -Category 'Active Directory' -Name 'Writable DC / GC' -Status BLOCKER -Current ("DC={0}; GC={1}; {2}" -f $State.AD.DC,$State.AD.GC,$State.AD.DCError) -Expected 'Writable DC and GC reachable'))
    }

    $dnsServers = @($State.Network | ForEach-Object { $_.DnsServers } | Where-Object { $_ } | Select-Object -Unique)
    $dhcpAdapters = @($State.Network | Where-Object { $_.Dhcp -eq 'Enabled' } | ForEach-Object { $_.Name })
    if ($dnsServers.Count -gt 0) {
        [void]$results.Add((New-CheckResult -Category 'Network' -Name 'DNS Client' -Status PASS -Current ($dnsServers -join ', ') -Expected 'Internal AD-capable DNS servers' -Message 'Verify the addresses against the customer AD/DNS design.'))
    }
    else {
        [void]$results.Add((New-CheckResult -Category 'Network' -Name 'DNS Client' -Status BLOCKER -Current '<None>' -Expected 'At least one internal AD DNS server'))
    }

    if ($dhcpAdapters.Count -gt 0) {
        [void]$results.Add((New-CheckResult -Category 'Network' -Name 'IPv4 Addressing' -Status REVIEW -Current ("DHCP: {0}" -f ($dhcpAdapters -join ', ')) -Expected 'Stable server addressing' -Message 'Confirm whether DHCP/reservation is intentional for this Exchange server.'))
    }
    else {
        [void]$results.Add((New-CheckResult -Category 'Network' -Name 'IPv4 Addressing' -Status PASS -Current 'No active adapter reports DHCP' -Expected 'Stable server addressing'))
    }

    $ipv4Enabled = @($State.Network | Where-Object { $_.IPv4Enabled }).Count -gt 0
    [void]$results.Add((New-CheckResult -Category 'Network' -Name 'IPv4 Binding' -Status $(if ($ipv4Enabled) { 'PASS' } else { 'BLOCKER' }) -Current $(if ($ipv4Enabled) { 'Enabled' } else { 'Not detected/enabled' }) -Expected 'Enabled'))

    $ipv6DisabledAdapters = @($State.Network | Where-Object { -not $_.IPv6Enabled } | ForEach-Object { $_.Name })
    if ($ipv6DisabledAdapters.Count -eq 0) {
        [void]$results.Add((New-CheckResult -Category 'Network' -Name 'IPv6 Adapter Binding' -Status PASS -Current 'Enabled on active adapters' -Expected 'Keep IPv6 bound/enabled'))
    }
    else {
        [void]$results.Add((New-CheckResult -Category 'Network' -Name 'IPv6 Adapter Binding' -Status REVIEW -Current ("Disabled: {0}" -f ($ipv6DisabledAdapters -join ', ')) -Expected 'Keep IPv6 bound/enabled' -Message 'Microsoft does not recommend unbinding IPv6.'))
    }

    $dcValue = [uint32]$State.IPv6Policy.DisabledComponents
    $ipv4PreferenceRegistry = 'Optional: HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters\DisabledComponents (REG_DWORD) = 0x20 / decimal 32'
    if ($dcValue -eq 0) {
        [void]$results.Add((New-CheckResult -Category 'Network' -Name 'IPv6 Preference Policy' -Status PASS -Current "$($State.IPv6Policy.HexValue) - $($State.IPv6Policy.Meaning)" -Expected 'Windows default IPv6 behavior is supported' -Message 'No IPv4 preference override is required by Exchange.'))
    }
    elseif ($dcValue -eq 32) {
        [void]$results.Add((New-CheckResult -Category 'Network' -Name 'IPv6 Preference Policy' -Status PASS -Current "$($State.IPv6Policy.HexValue) - $($State.IPv6Policy.Meaning)" -Expected $ipv4PreferenceRegistry -Message 'Supported optional configuration when the environment intentionally prefers IPv4 while keeping IPv6 enabled.'))
    }
    elseif ($dcValue -eq 255) {
        [void]$results.Add((New-CheckResult -Category 'Network' -Name 'IPv6 Preference Policy' -Status REVIEW -Current "$($State.IPv6Policy.HexValue) - $($State.IPv6Policy.Meaning)" -Expected 'Default IPv6 behavior, or 0x20 when IPv4 preference is intentionally required' -Message 'IPv6 is disabled by policy. Confirm the registry value and adapter bindings are intentionally and consistently configured before Exchange installation.'))
    }
    else {
        [void]$results.Add((New-CheckResult -Category 'Network' -Name 'IPv6 Preference Policy' -Status REVIEW -Current "$($State.IPv6Policy.HexValue) - $($State.IPv6Policy.Meaning)" -Expected 'Default IPv6 behavior, or optional 0x20 when IPv4 preference is required' -Message 'A custom IPv6 policy is configured. Review it against the Windows IPv6 guidance before Exchange installation.'))
    }

    $timeSource = [string]$State.TimeSync.Source
    $timeState = [string]$State.TimeSync.ServiceState
    if (-not [string]::IsNullOrWhiteSpace($timeSource) -and $timeSource -notmatch '(?i)Local CMOS Clock|Free-running System Clock') {
        [void]$results.Add((New-CheckResult -Category 'Operating System' -Name 'Windows Time' -Status PASS -Current ("{0}; Source={1}" -f $timeState,$timeSource) -Expected 'Running with a valid domain/environment time source'))
    }
    else {
        $timeCurrent = if ([string]::IsNullOrWhiteSpace($timeSource)) { "{0}; {1}" -f $timeState,$State.TimeSync.Error } else { "{0}; Source={1}" -f $timeState,$timeSource }
        [void]$results.Add((New-CheckResult -Category 'Operating System' -Name 'Windows Time' -Status REVIEW -Current $timeCurrent -Expected 'Running with a valid domain/environment time source' -Message 'Verify time synchronization before Exchange installation. Kerberos and Exchange authentication are time-sensitive.'))
    }

    if ($State.PendingReboot.Pending) {
        [void]$results.Add((New-CheckResult -Category 'Operating System' -Name 'Pending reboot' -Status BLOCKER -Current ($State.PendingReboot.Reasons -join ', ') -Expected 'No pending reboot' -Message 'Restart the server and rerun the check before Exchange Setup.'))
    }
    else {
        [void]$results.Add((New-CheckResult -Category 'Operating System' -Name 'Pending reboot' -Status PASS -Current 'No' -Expected 'No pending reboot'))
    }

    $ramGB = [math]::Round($State.ComputerSystem.TotalPhysicalMemory / 1GB, 1)
    if ($ramGB -ge 128) {
        [void]$results.Add((New-CheckResult -Category 'Hardware' -Name 'Memory' -Status PASS -Current "$ramGB GB" -Expected '128 GB minimum recommended for Mailbox role'))
    }
    else {
        [void]$results.Add((New-CheckResult -Category 'Hardware' -Name 'Memory' -Status REVIEW -Current "$ramGB GB" -Expected '128 GB minimum recommended for Mailbox role' -Message 'Validate sizing for the intended workload.'))
    }

    $cpuSockets = [int]$State.ComputerSystem.NumberOfProcessors
    $logicalProcessors = [int]$State.ComputerSystem.NumberOfLogicalProcessors
    if ($cpuSockets -le 2) {
        [void]$results.Add((New-CheckResult -Category 'Hardware' -Name 'CPU Sockets' -Status PASS -Current ("Sockets={0}; LogicalProcessors={1}" -f $cpuSockets,$logicalProcessors) -Expected 'Up to 2 processor sockets recommended'))
    }
    else {
        [void]$results.Add((New-CheckResult -Category 'Hardware' -Name 'CPU Sockets' -Status REVIEW -Current ("Sockets={0}; LogicalProcessors={1}" -f $cpuSockets,$logicalProcessors) -Expected 'Up to 2 processor sockets recommended' -Message 'Review the physical/virtual CPU topology and Exchange sizing design.'))
    }

    $systemDrive = [string]$env:SystemDrive
    $systemDisk = @($State.FixedDisks | Where-Object { $_.DeviceID -ieq $systemDrive } | Select-Object -First 1)
    if ($systemDisk.Count -gt 0) {
        $diskText = "$($systemDisk[0].FreeGB) GB free / $($systemDisk[0].FileSystem)"
        if ($systemDisk[0].FreeGB -lt 0.2) {
            [void]$results.Add((New-CheckResult -Category 'Storage' -Name 'System Drive Free Space' -Status BLOCKER -Current $diskText -Expected 'At least 200 MB free on system drive'))
        }
        elseif ($systemDisk[0].FreeGB -lt 30) {
            [void]$results.Add((New-CheckResult -Category 'Storage' -Name 'System Drive Free Space' -Status REVIEW -Current $diskText -Expected '30 GB free if Exchange binaries will be installed on this drive'))
        }
        else {
            [void]$results.Add((New-CheckResult -Category 'Storage' -Name 'System Drive Free Space' -Status PASS -Current $diskText -Expected 'At least 30 GB when used as Exchange installation drive'))
        }
        [void]$results.Add((New-CheckResult -Category 'Storage' -Name 'System Drive File System' -Status $(if ($systemDisk[0].FileSystem -ieq 'NTFS') { 'PASS' } else { 'BLOCKER' }) -Current $systemDisk[0].FileSystem -Expected 'NTFS'))
    }
    else {
        [void]$results.Add((New-CheckResult -Category 'Storage' -Name 'System Drive' -Status BLOCKER -Current 'Not detected' -Expected 'System drive information available'))
    }

    $desiredPageFileMB = [int][math]::Round(($State.ComputerSystem.TotalPhysicalMemory / 1MB) * 0.25, 0)
    $desiredPageFileGB = [math]::Round($desiredPageFileMB / 1024, 1)
    $pageSetting = @($State.PageFileSettings)
    $pageCurrent = if ($pageSetting.Count -eq 0) {
        "Automatic=$($State.AutomaticManagedPagefile); no Win32_PageFileSetting object detected"
    }
    else {
        "Automatic=$($State.AutomaticManagedPagefile); " + (($pageSetting | ForEach-Object { "{0}: Initial={1} MB Maximum={2} MB" -f $_.Name,$_.InitialSize,$_.MaximumSize }) -join '; ')
    }
    $pageMatches = (-not $State.AutomaticManagedPagefile -and $pageSetting.Count -eq 1 -and [int]$pageSetting[0].InitialSize -eq $desiredPageFileMB -and [int]$pageSetting[0].MaximumSize -eq $desiredPageFileMB)
    [void]$results.Add((New-CheckResult -Category 'Operating System' -Name 'Recommended Page File' -Status $(if ($pageMatches) { 'PASS' } else { 'REVIEW' }) -Current $pageCurrent -Expected ("Initial=Maximum={0} MB ({1} GB), 25% of installed RAM" -f $desiredPageFileMB,$desiredPageFileGB) -Message 'Microsoft Exchange guidance: Initial and Maximum should be the same value: 25% of installed RAM. This script only reports the recommendation.'))

    $dotNetSupported = $false
    $dotNetRecommended = $false
    if ($osYear -eq '2019') {
        $dotNetSupported = ($State.DotNet.Release -and [int64]$State.DotNet.Release -ge 528040 -and [int64]$State.DotNet.Release -lt 533320)
        $dotNetRecommended = $dotNetSupported
    }
    elseif ($osYear -in @('2022','2025')) {
        $dotNetSupported = ($State.DotNet.Release -and [int64]$State.DotNet.Release -ge 528040)
        $dotNetRecommended = ($State.DotNet.Release -and [int64]$State.DotNet.Release -ge 533320)
    }

    if ($dotNetSupported) {
        $message = if ($dotNetRecommended) { '' } else { '.NET Framework 4.8 is supported; 4.8.1 is recommended on Windows Server 2022/2025.' }
        [void]$results.Add((New-CheckResult -Category 'Prerequisites' -Name '.NET Framework' -Status PASS -Current $State.DotNet.Version -Expected $(if ($osYear -eq '2019') { '4.8' } else { '4.8 or 4.8.1 (4.8.1 recommended)' }) -Message $message))
    }
    else {
        [void]$results.Add((New-CheckResult -Category 'Prerequisites' -Name '.NET Framework' -Status BLOCKER -Current $State.DotNet.Version -Expected $(if ($osYear -eq '2019') { '4.8' } else { '4.8 or 4.8.1 (4.8.1 recommended)' }) -Message 'Install a supported .NET Framework version before Exchange Setup.'))
    }

    $requiredFeatures = @(Get-RequiredWindowsFeatures -State $State)
    if (@($State.WindowsFeatures).Count -eq 0) {
        [void]$results.Add((New-CheckResult -Category 'Prerequisites' -Name 'Windows Features' -Status BLOCKER -Current 'Unable to query Get-WindowsFeature' -Expected 'Required Exchange SE Mailbox Windows Features'))
    }
    else {
        $missingFeatures = @()
        foreach ($featureName in $requiredFeatures) {
            $feature = @($State.WindowsFeatures | Where-Object { $_.Name -eq $featureName } | Select-Object -First 1)
            if ($feature.Count -eq 0 -or -not $feature[0].Installed) { $missingFeatures += $featureName }
        }
        if ($missingFeatures.Count -eq 0) {
            [void]$results.Add((New-CheckResult -Category 'Prerequisites' -Name 'Windows Features' -Status PASS -Current 'All required Exchange SE Mailbox features installed' -Expected 'Complete'))
        }
        else {
            [void]$results.Add((New-CheckResult -Category 'Prerequisites' -Name 'Windows Features' -Status BLOCKER -Current ("Missing: {0}" -f ($missingFeatures -join ', ')) -Expected 'All required Exchange SE Mailbox Windows Features installed' -Message 'Exchange Setup can install supported Windows components by using the prerequisite option or /InstallWindowsComponents.'))
        }
    }

    if ($State.RemoteRegistry -and $State.RemoteRegistry.StartMode -eq 'Auto') {
        [void]$results.Add((New-CheckResult -Category 'Prerequisites' -Name 'Remote Registry' -Status PASS -Current "$($State.RemoteRegistry.StartMode) / $($State.RemoteRegistry.State)" -Expected 'Startup type Automatic (not Disabled)'))
    }
    else {
        $rrCurrent = if ($State.RemoteRegistry) { "$($State.RemoteRegistry.StartMode) / $($State.RemoteRegistry.State)" } else { 'Service not found' }
        [void]$results.Add((New-CheckResult -Category 'Prerequisites' -Name 'Remote Registry' -Status BLOCKER -Current $rrCurrent -Expected 'Startup type Automatic (not Disabled)' -Message 'Microsoft Exchange prerequisites require the Remote Registry service startup type to be Automatic.'))
    }

    $vc2012Text = Get-InstalledApplicationText -State $State -Pattern 'Microsoft Visual C\+\+ 2012.*x64'
    $vc2012Installed = -not [string]::IsNullOrWhiteSpace($vc2012Text)
    [void]$results.Add((New-CheckResult -Category 'Prerequisites' -Name 'Visual C++ 2012 x64' -Status $(if ($vc2012Installed) { 'PASS' } else { 'BLOCKER' }) -Current $(if ($vc2012Installed) { $vc2012Text } else { 'Not detected' }) -Expected 'Installed' -Message $(if ($vc2012Installed) { '' } else { 'Install the Microsoft Visual C++ 2012 x64 Redistributable before Exchange Setup.' })))

    $vc2013Text = Get-InstalledApplicationText -State $State -Pattern 'Microsoft Visual C\+\+ 2013.*x64'
    $vc2013Installed = -not [string]::IsNullOrWhiteSpace($vc2013Text)
    [void]$results.Add((New-CheckResult -Category 'Prerequisites' -Name 'Visual C++ 2013 x64' -Status $(if ($vc2013Installed) { 'PASS' } else { 'BLOCKER' }) -Current $(if ($vc2013Installed) { $vc2013Text } else { 'Not detected' }) -Expected 'Installed' -Message $(if ($vc2013Installed) { '' } else { 'Install the Microsoft Visual C++ 2013 x64 Redistributable before Exchange Setup.' })))

    $ucmaText = Get-InstalledApplicationText -State $State -Pattern 'Unified Communications Managed API 4\.0|UCMA 4\.0'
    $ucmaInstalled = -not [string]::IsNullOrWhiteSpace($ucmaText)
    [void]$results.Add((New-CheckResult -Category 'Prerequisites' -Name 'Unified Communications Managed API 4.0' -Status $(if ($ucmaInstalled) { 'PASS' } else { 'BLOCKER' }) -Current $(if ($ucmaInstalled) { $ucmaText } else { 'Not detected' }) -Expected 'Installed' -Message $(if ($ucmaInstalled) { '' } else { 'Install UCMA 4.0 before Exchange Setup. The installer is also available in the UCMARedist folder on Exchange media.' })))

    $urlRewriteText = Get-InstalledApplicationText -State $State -Pattern 'IIS URL Rewrite Module|URL Rewrite Module'
    $urlRewriteInstalled = -not [string]::IsNullOrWhiteSpace($urlRewriteText)
    [void]$results.Add((New-CheckResult -Category 'Prerequisites' -Name 'IIS URL Rewrite Module 2' -Status $(if ($urlRewriteInstalled) { 'PASS' } else { 'BLOCKER' }) -Current $(if ($urlRewriteInstalled) { $urlRewriteText } else { 'Not detected' }) -Expected 'Installed' -Message $(if ($urlRewriteInstalled) { '' } else { 'Install IIS URL Rewrite Module before Exchange Setup.' })))

    if ($State.CredentialGuard.Enabled) {
        [void]$results.Add((New-CheckResult -Category 'Security' -Name 'Credential Guard' -Status BLOCKER -Current ("Enabled (LsaCfgFlags={0}; Running={1})" -f $State.CredentialGuard.LsaCfgFlags,($State.CredentialGuard.SecurityServicesRunning -join ',')) -Expected 'Disabled for Exchange Server' -Message 'Review the organization security/GPO configuration before Exchange installation.'))
    }
    else {
        [void]$results.Add((New-CheckResult -Category 'Security' -Name 'Credential Guard' -Status PASS -Current 'Not detected as enabled' -Expected 'Disabled for Exchange Server'))
    }

    $powerGuid = [string]$State.PowerPlan.Guid
    if ($powerGuid -eq '8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c') {
        [void]$results.Add((New-CheckResult -Category 'Performance' -Name 'Power Plan' -Status PASS -Current $State.PowerPlan.Name -Expected 'High performance'))
    }
    else {
        [void]$results.Add((New-CheckResult -Category 'Performance' -Name 'Power Plan' -Status REVIEW -Current $State.PowerPlan.Name -Expected 'Review for Exchange workload / CPU throttling'))
    }

    $systemLocaleName = if ($State.SystemLocale) { [string]$State.SystemLocale.Name } else { '<Unknown>' }
    $cultureName = if ($State.Culture) { [string]$State.Culture.Name } else { '<Unknown>' }
    $uiCultureName = if ($State.UICulture) { [string]$State.UICulture.Name } else { '<Unknown>' }
    $timeZoneName = if ($State.TimeZone) { [string]$State.TimeZone.Id } else { '<Unknown>' }
    $shortDate = if ($State.Culture) { [string]$State.Culture.DateTimeFormat.ShortDatePattern } else { '<Unknown>' }
    $decimalSeparator = if ($State.Culture) { [string]$State.Culture.NumberFormat.NumberDecimalSeparator } else { '<Unknown>' }

    [void]$results.Add((New-CheckResult -Category 'Regional Consistency' -Name 'System Locale' -Status REVIEW -Current $systemLocaleName -Expected 'Must be consistent across the new Exchange servers' -Message 'Compare this value across every new Exchange server before production traffic.'))
    [void]$results.Add((New-CheckResult -Category 'Regional Consistency' -Name 'Current Culture' -Status REVIEW -Current "$cultureName; ShortDate=$shortDate; Decimal=$decimalSeparator" -Expected 'Must be consistent across the new Exchange servers'))
    [void]$results.Add((New-CheckResult -Category 'Regional Consistency' -Name 'UI Culture' -Status REVIEW -Current $uiCultureName -Expected 'Review for consistency across the new Exchange servers'))
    [void]$results.Add((New-CheckResult -Category 'Regional Consistency' -Name 'Time Zone' -Status REVIEW -Current $timeZoneName -Expected 'Must be consistent with the deployment design'))

    if ($State.IEEsc.Applicable) {
        [void]$results.Add((New-CheckResult -Category 'Security' -Name 'IE ESC - Administrators' -Status REVIEW -Current $State.IEEsc.AdministratorState -Expected 'Environment decision; Windows Server default is On' -Message 'Exchange installation does not require IE ESC to be disabled. Current state is reported separately for Administrators.'))
        [void]$results.Add((New-CheckResult -Category 'Security' -Name 'IE ESC - Users' -Status REVIEW -Current $State.IEEsc.UserState -Expected 'Environment decision; Windows Server default is On' -Message 'Exchange installation does not require IE ESC to be disabled. Current state is reported separately for Users.'))
    }
    else {
        [void]$results.Add((New-CheckResult -Category 'Security' -Name 'IE ESC - Administrators' -Status REVIEW -Current $State.IEEsc.AdministratorState -Expected 'Not applicable to Server Core'))
        [void]$results.Add((New-CheckResult -Category 'Security' -Name 'IE ESC - Users' -Status REVIEW -Current $State.IEEsc.UserState -Expected 'Not applicable to Server Core'))
    }

    return @($results)
}

function Get-CrossServerConsistencyChecks {
    param([Parameter(Mandatory = $true)][object[]]$Results)

    $connected = @($Results | Where-Object { $_.ConnectionStatus -eq 'Connected' -and $null -ne $_.State })
    if ($connected.Count -lt 2) { return @() }

    $checks = New-Object System.Collections.Generic.List[object]

    $comparisons = @(
        [PSCustomObject]@{
            Name = 'System Locale'
            Values = @($connected | ForEach-Object {
                $value = if ($_.State.SystemLocale) { [string]$_.State.SystemLocale.Name } else { '<Unknown>' }
                "{0}={1}" -f $_.Server,$value
            })
            Unique = @($connected | ForEach-Object {
                if ($_.State.SystemLocale) { [string]$_.State.SystemLocale.Name } else { '<Unknown>' }
            } | Select-Object -Unique)
        },
        [PSCustomObject]@{
            Name = 'Current Culture'
            Values = @($connected | ForEach-Object {
                if ($_.State.Culture) {
                    $culture = $_.State.Culture
                    "{0}={1};ShortDate={2};Decimal={3}" -f $_.Server,$culture.Name,$culture.DateTimeFormat.ShortDatePattern,$culture.NumberFormat.NumberDecimalSeparator
                }
                else {
                    "{0}=<Unknown>" -f $_.Server
                }
            })
            Unique = @($connected | ForEach-Object {
                if ($_.State.Culture) {
                    $culture = $_.State.Culture
                    "{0}|{1}|{2}" -f $culture.Name,$culture.DateTimeFormat.ShortDatePattern,$culture.NumberFormat.NumberDecimalSeparator
                }
                else {
                    '<Unknown>'
                }
            } | Select-Object -Unique)
        },
        [PSCustomObject]@{
            Name = 'UI Culture'
            Values = @($connected | ForEach-Object {
                $value = if ($_.State.UICulture) { [string]$_.State.UICulture.Name } else { '<Unknown>' }
                "{0}={1}" -f $_.Server,$value
            })
            Unique = @($connected | ForEach-Object {
                if ($_.State.UICulture) { [string]$_.State.UICulture.Name } else { '<Unknown>' }
            } | Select-Object -Unique)
        },
        [PSCustomObject]@{
            Name = 'Time Zone'
            Values = @($connected | ForEach-Object {
                $value = if ($_.State.TimeZone) { [string]$_.State.TimeZone.Id } else { '<Unknown>' }
                "{0}={1}" -f $_.Server,$value
            })
            Unique = @($connected | ForEach-Object {
                if ($_.State.TimeZone) { [string]$_.State.TimeZone.Id } else { '<Unknown>' }
            } | Select-Object -Unique)
        }
    )

    foreach ($comparison in $comparisons) {
        $status = if ($comparison.Unique.Count -eq 1) { 'PASS' } else { 'REVIEW' }
        $message = if ($status -eq 'PASS') {
            'Values are consistent across the checked servers.'
        }
        else {
            'Peer-server values differ. Confirm whether the difference is intentional before Exchange installation or production traffic.'
        }

        [void]$checks.Add((New-CheckResult -Category 'Cross-Server Consistency' -Name $comparison.Name -Status $status -Current ($comparison.Values -join ' | ') -Expected 'Consistent across peer Exchange servers' -Message $message))
    }

    return @($checks)
}

# ---------------------------------------------------------------------------
# Display / report helpers
# ---------------------------------------------------------------------------
function Get-StatusColor {
    param([string]$Status)
    switch ($Status) {
        'PASS'    { 'Green' }
        'BLOCKER' { 'Red' }
        'REVIEW'  { 'Cyan' }
        default   { 'Gray' }
    }
}

function Show-CheckResults {
    param([Parameter(Mandatory = $true)][object[]]$Checks)

    $categories = @($Checks.Category | Select-Object -Unique)
    foreach ($category in $categories) {
        Write-Host ''
        Write-Host "[$category]" -ForegroundColor White
        foreach ($check in @($Checks | Where-Object { $_.Category -eq $category })) {
            $color = Get-StatusColor -Status $check.Status
            Write-Host ("{0,-9} {1}" -f $check.Status,$check.Name) -ForegroundColor $color
            Write-Host ("  Current : {0}" -f $check.Current) -ForegroundColor DarkGray
            if ($check.Expected -ne '<Not Set>') { Write-Host ("  Expected: {0}" -f $check.Expected) -ForegroundColor DarkGray }
            if (-not [string]::IsNullOrWhiteSpace($check.Message)) { Write-Host ("  Note    : {0}" -f $check.Message) -ForegroundColor DarkGray }
        }
    }

    $pass = @($Checks | Where-Object { $_.Status -eq 'PASS' }).Count
    $blocker = @($Checks | Where-Object { $_.Status -eq 'BLOCKER' }).Count
    $review = @($Checks | Where-Object { $_.Status -eq 'REVIEW' }).Count
    Write-Host ''
    Write-Host ("Summary: PASS={0}  BLOCKER={1}  REVIEW={2}" -f $pass,$blocker,$review) -ForegroundColor White
}

function Get-CheckReportLines {
    param(
        [Parameter(Mandatory = $true)]$State,
        [Parameter(Mandatory = $true)][object[]]$Checks
    )

    $lines = New-Object System.Collections.Generic.List[string]
    [void]$lines.Add(("{0}.ps1  Version {1}" -f $script:ScriptBaseName,$script:ScriptVersion))
    [void]$lines.Add(('Report Time    : {0}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')))
    [void]$lines.Add(('Computer       : {0}' -f $State.ComputerName))
    [void]$lines.Add(('OS             : {0}' -f $State.OperatingSystem.Caption))
    [void]$lines.Add(('OS Build       : {0}' -f $State.OperatingSystem.BuildNumber))
    [void]$lines.Add(('Installation   : {0}' -f $State.RegistryOS.InstallationType))
    [void]$lines.Add(('Domain         : {0}' -f $State.ComputerSystem.Domain))
    [void]$lines.Add(('Primary DNS    : {0}' -f $State.PrimaryDnsSuffix))
    [void]$lines.Add(('AD Site        : {0}' -f $State.AD.Site))
    [void]$lines.Add(('Forest Level   : {0}' -f $State.ADFunctionalLevel.ForestMode))
    [void]$lines.Add('Mode           : Read-only server configuration check')
    [void]$lines.Add('')

    foreach ($category in @($Checks.Category | Select-Object -Unique)) {
        [void]$lines.Add("[$category]")
        foreach ($check in @($Checks | Where-Object { $_.Category -eq $category })) {
            [void]$lines.Add(("{0,-9} {1}" -f $check.Status,$check.Name))
            [void]$lines.Add(("  Current : {0}" -f $check.Current))
            [void]$lines.Add(("  Expected: {0}" -f $check.Expected))
            if (-not [string]::IsNullOrWhiteSpace($check.Message)) { [void]$lines.Add(("  Note    : {0}" -f $check.Message)) }
        }
        [void]$lines.Add('')
    }

    $pass = @($Checks | Where-Object { $_.Status -eq 'PASS' }).Count
    $blocker = @($Checks | Where-Object { $_.Status -eq 'BLOCKER' }).Count
    $review = @($Checks | Where-Object { $_.Status -eq 'REVIEW' }).Count
    [void]$lines.Add(("Summary: PASS={0}  BLOCKER={1}  REVIEW={2}" -f $pass,$blocker,$review))
    return @($lines)
}

function Export-CheckReport {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)]$State,
        [Parameter(Mandatory = $true)][object[]]$Checks
    )

    $resolvedPath = Resolve-OutputFilePath -Path $Path
    $lines = @(Get-CheckReportLines -State $State -Checks $Checks)
    Set-Content -LiteralPath $resolvedPath -Value $lines -Encoding UTF8 -ErrorAction Stop
    return $resolvedPath
}

# ---------------------------------------------------------------------------
# Multi-server orchestration helpers
# ---------------------------------------------------------------------------
function Test-IsLocalTarget {
    param([Parameter(Mandatory = $true)][string]$ComputerName)

    $name = $ComputerName.Trim()
    if ([string]::IsNullOrWhiteSpace($name)) { return $true }
    if ($name -in @('.','localhost','127.0.0.1','::1')) { return $true }
    if ($name -ieq $env:COMPUTERNAME) { return $true }

    try {
        $localFqdn = [System.Net.Dns]::GetHostEntry($env:COMPUTERNAME).HostName
        if (-not [string]::IsNullOrWhiteSpace($localFqdn) -and $name -ieq $localFqdn) { return $true }
    }
    catch { }

    return $false
}

function Invoke-LocalServerCheck {
    $state = Get-ServerState
    $checks = @(Get-ExchangePreparationChecks -State $state)

    return [PSCustomObject]@{
        Server           = [string]$state.ComputerName
        ConnectionStatus = 'Connected'
        Error            = ''
        State            = $state
        Checks           = $checks
    }
}

function New-ConnectionFailureResult {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [Parameter(Mandatory = $true)][string]$ErrorMessage
    )

    $check = [PSCustomObject]@{
        Category = 'Connection'
        Name     = 'Connection Failed'
        Status   = 'BLOCKER'
        Current  = $ErrorMessage
        Expected = 'PowerShell Remoting / WinRM reachable with the current credentials'
        Message  = 'The server was skipped. Other requested servers will continue.'
    }

    return [PSCustomObject]@{
        Server           = $ComputerName
        ConnectionStatus = 'Failed'
        Error            = $ErrorMessage
        State            = $null
        Checks           = @($check)
    }
}

function Invoke-ServerCheck {
    param([Parameter(Mandatory = $true)][string]$ComputerName)

    if (Test-IsLocalTarget -ComputerName $ComputerName) {
        return Invoke-LocalServerCheck
    }

    try {
        $target = $ComputerName
        $remoteResult = Invoke-Command -ComputerName $target -FilePath $PSCommandPath -ArgumentList @($true) -ErrorAction Stop
        $firstResult = @($remoteResult | Select-Object -First 1)
        if ($firstResult.Count -eq 0) {
            return New-ConnectionFailureResult -ComputerName $ComputerName -ErrorMessage 'Remote command completed without returning a check result.'
        }
        return $firstResult[0]
    }
    catch {
        return New-ConnectionFailureResult -ComputerName $ComputerName -ErrorMessage $_.Exception.Message
    }
}

function Show-ServerCheckResult {
    param([Parameter(Mandatory = $true)]$Result)

    Write-Host ''
    Write-Host '==================================================' -ForegroundColor DarkGray
    Write-Host ("Server: {0}" -f $Result.Server) -ForegroundColor Yellow
    Write-Host '==================================================' -ForegroundColor DarkGray

    Show-CheckResults -Checks @($Result.Checks)
}

function Get-MultiServerReportLines {
    param(
        [Parameter(Mandatory = $true)][object[]]$Results,
        [object[]]$ComparisonChecks = @()
    )

    $lines = New-Object System.Collections.Generic.List[string]
    [void]$lines.Add(("{0}.ps1  Version {1}" -f $script:ScriptBaseName,$script:ScriptVersion))
    [void]$lines.Add(('Report Time : {0}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')))
    [void]$lines.Add('Mode        : Read-only server configuration check')
    [void]$lines.Add(('Servers     : {0}' -f (@($Results | ForEach-Object { $_.Server }) -join ', ')))
    [void]$lines.Add('')

    foreach ($result in $Results) {
        [void]$lines.Add('==================================================')
        [void]$lines.Add(("Server: {0}" -f $result.Server))
        [void]$lines.Add('==================================================')

        if ($result.ConnectionStatus -eq 'Connected' -and $null -ne $result.State) {
            foreach ($line in @(Get-CheckReportLines -State $result.State -Checks @($result.Checks))) {
                [void]$lines.Add($line)
            }
        }
        else {
            [void]$lines.Add('[Connection]')
            [void]$lines.Add('BLOCKER   Connection Failed')
            [void]$lines.Add(('  Current : {0}' -f $result.Error))
            [void]$lines.Add('  Expected: PowerShell Remoting / WinRM reachable with the current credentials')
            [void]$lines.Add('  Note    : The server was skipped. Other requested servers continued.')
            [void]$lines.Add('')
            [void]$lines.Add('Summary: PASS=0  BLOCKER=1  REVIEW=0')
        }
        [void]$lines.Add('')
    }

    if (@($ComparisonChecks).Count -gt 0) {
        [void]$lines.Add('==================================================')
        [void]$lines.Add('Cross-Server Consistency')
        [void]$lines.Add('==================================================')
        foreach ($check in $ComparisonChecks) {
            [void]$lines.Add(("{0,-9} {1}" -f $check.Status,$check.Name))
            [void]$lines.Add(("  Current : {0}" -f $check.Current))
            [void]$lines.Add(("  Expected: {0}" -f $check.Expected))
            if (-not [string]::IsNullOrWhiteSpace($check.Message)) {
                [void]$lines.Add(("  Note    : {0}" -f $check.Message))
            }
        }
        [void]$lines.Add('')
    }

    return @($lines)
}

function Export-MultiServerReport {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][object[]]$Results,
        [object[]]$ComparisonChecks = @()
    )

    $resolvedPath = Resolve-OutputFilePath -Path $Path
    $lines = @(Get-MultiServerReportLines -Results $Results -ComparisonChecks $ComparisonChecks)
    Set-Content -LiteralPath $resolvedPath -Value $lines -Encoding UTF8 -ErrorAction Stop
    return $resolvedPath
}

# Internal remoting entry point. It returns structured data and writes no report/configuration changes.
if ($InternalLocal) {
    Invoke-LocalServerCheck
    return
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
Write-Host ''
Write-Host ("{0}.ps1" -f $script:ScriptBaseName) -ForegroundColor Yellow
Write-Host 'Exchange Server SE Mailbox server readiness check' -ForegroundColor DarkCyan
Write-Host ("Version       : {0}" -f $script:ScriptVersion) -ForegroundColor Cyan
Write-Host 'Mode          : Read-only' -ForegroundColor Cyan
Write-Host ''

$targets = @()
if (@($Server).Count -eq 0) {
    $targets = @($env:COMPUTERNAME)
}
else {
    $targets = @($Server | ForEach-Object { if ($null -ne $_) { $_.Trim() } } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
}

if ($targets.Count -eq 0) {
    $targets = @($env:COMPUTERNAME)
}

$results = New-Object System.Collections.Generic.List[object]
foreach ($target in $targets) {
    $result = Invoke-ServerCheck -ComputerName $target
    [void]$results.Add($result)
    Show-ServerCheckResult -Result $result
}

$comparisonChecks = @(Get-CrossServerConsistencyChecks -Results @($results))
if ($comparisonChecks.Count -gt 0) {
    Write-Host ''
    Write-Host '==================================================' -ForegroundColor DarkGray
    Write-Host 'Cross-Server Consistency' -ForegroundColor Yellow
    Write-Host '==================================================' -ForegroundColor DarkGray
    Show-CheckResults -Checks $comparisonChecks
}

if (-not [string]::IsNullOrWhiteSpace($OutputFile)) {
    $outputPath = Export-MultiServerReport -Path $OutputFile -Results @($results) -ComparisonChecks $comparisonChecks
    Write-Host ''
    Write-Host ("Report exported: {0}" -f $outputPath) -ForegroundColor Green
}

Write-Host ''
Write-Host 'Check complete. No server configuration changes were made.' -ForegroundColor Cyan
