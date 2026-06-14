#requires -modules ActiveDirectory

<#
.PARAMETER IncludeDCHealth
    When set (default), collects per-DC hardware/performance/disk/service/replication
    inventory for the "Domain Controller Health" tab via Get-DCInventory.ps1
    (agentless, CIM/WinRM). Use -IncludeDCHealth:$false to skip this collection on very
    large environments or when WinRM is not available, producing the Forest Overview
    tab only.

.PARAMETER DCHealthThrottleLimit
    Maximum number of domain controllers queried in parallel for the DC Health tab.
    Defaults to 10. Increase for large environments with many DCs.

.PARAMETER DCHealthCredential
    Optional credential used for the remote WinRM connections made by
    Get-DCInventory.ps1.
#>

param(
    [bool]$IncludeDCHealth = $true,
    [int]$DCHealthThrottleLimit = 10,
    [System.Management.Automation.PSCredential]$DCHealthCredential
)

Import-Module ActiveDirectory -ErrorAction Stop

$ReportPath = "C:\Temp\Enterprise_AD_Operational_Dashboard.html"

if (!(Test-Path "C:\Temp")) {
    New-Item -Path "C:\Temp" -ItemType Directory | Out-Null
}

Write-Host "Collecting Enterprise Active Directory information..." -ForegroundColor Cyan

function ConvertTo-HtmlSafe {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ""
    }

    return [System.Security.SecurityElement]::Escape($Value)
}

function ConvertTo-JavaScriptString {
    param([string]$Value)

    if ($null -eq $Value) { return "" }

    return $Value `
        -replace "\\", "\\" `
        -replace "'", "\'" `
        -replace "`r", "" `
        -replace "`n", "\n"
}

function Convert-TimeSpanToReadable {
    param($Value)

    if ($null -eq $Value) {
        return "Not Configured"
    }

    if ($Value -is [TimeSpan]) {
        return "$([math]::Abs($Value.Days)) Days"
    }

    return $Value.ToString()
}

function Get-CleanADSiteName {
    param([string]$SiteValue)

    if ([string]::IsNullOrWhiteSpace($SiteValue)) {
        return "No site assigned"
    }

    if ($SiteValue -match "^CN=([^,]+),") {
        return $matches[1]
    }

    return $SiteValue
}

function Get-ADGroupMemberDetailsSafe {
    param(
        [Parameter(Mandatory)]
        [string]$Identity
    )

    try {
        $Members = @(Get-ADGroupMember -Identity $Identity -Recursive -ErrorAction Stop |
            Sort-Object Name |
            Select-Object Name, SamAccountName, objectClass)

        if ($Members.Count -eq 0) {
            return [pscustomobject]@{
                Count   = 0
                Tooltip = "<div class='tooltip-line'>No members found</div>"
            }
        }

        $Tooltip = ($Members | ForEach-Object {
            $Name = ConvertTo-HtmlSafe $_.Name
            $Sam  = ConvertTo-HtmlSafe $_.SamAccountName
            $Type = ConvertTo-HtmlSafe $_.objectClass

            if ($Sam) {
                "<div class='tooltip-line'><b>$Name</b><span>$Sam ($Type)</span></div>"
            }
            else {
                "<div class='tooltip-line'><b>$Name</b><span>$Type</span></div>"
            }
        }) -join ""

        return [pscustomobject]@{
            Count   = $Members.Count
            Tooltip = $Tooltip
        }
    }
    catch {
        return [pscustomobject]@{
            Count   = "N/A"
            Tooltip = "<div class='tooltip-line'>Unable to retrieve members</div>"
        }
    }
}

function Get-PrivilegedGroupCsv {
    param(
        [Parameter(Mandatory)]
        [string]$Identity,

        [Parameter(Mandatory)]
        [string]$DomainName
    )

    $ExportData = @(Get-ADGroupMember -Identity $Identity -Recursive -ErrorAction SilentlyContinue |
        ForEach-Object {

            $DisplayName = $null

            if ($_.objectClass -eq "user") {
                try {
                    $User = Get-ADUser -Identity $_.SamAccountName -Properties DisplayName -ErrorAction Stop
                    $DisplayName = $User.DisplayName
                }
                catch {
                    $DisplayName = $_.Name
                }
            }
            else {
                $DisplayName = $_.Name
            }

            [pscustomobject]@{
                SamAccountName = $_.SamAccountName
                DisplayName    = $DisplayName
                DomainName     = $DomainName
            }
        })

    return $ExportData | ConvertTo-Csv -NoTypeInformation
}

$Domain  = Get-ADDomain
$Forest  = Get-ADForest
$RootDSE = Get-ADRootDSE

$DomainName            = $Domain.DNSRoot
$ForestRootDomain      = $Forest.RootDomain
$NetBIOSName           = $Domain.NetBIOSName
$ForestFunctionalLevel = $Forest.ForestMode
$DomainFunctionalLevel = $Domain.DomainMode
$GeneratedOn           = Get-Date -Format "dd-MMM-yyyy hh:mm:ss tt"

$AllDomains  = @($Forest.Domains)
$DomainCount = $AllDomains.Count

$DomainTooltip = ""

if ($DomainCount -gt 1) {
    foreach ($DomainEntry in $AllDomains) {
        if ($DomainEntry -eq $ForestRootDomain) {
            $Type = "Root Domain"
        }
        else {
            $Type = "Child Domain"
        }

        $DomainTooltip += "<div class='tooltip-line'><b>$(ConvertTo-HtmlSafe $DomainEntry)</b><span>$Type</span></div>"
    }

    $DomainTooltip += "<div class='tooltip-line'><b>Forest Structure</b><span>Parent - Child Domain Forest</span></div>"
}
else {
    $DomainTooltip += "<div class='tooltip-line'><b>$(ConvertTo-HtmlSafe $ForestRootDomain)</b><span>Single Domain Forest</span></div>"
}

try {
    $ConfigNC = $RootDSE.configurationNamingContext
    $Partitions = Get-ADObject "CN=Partitions,$ConfigNC" -Properties uPNSuffixes
    $UPNSuffixes = @($Partitions.uPNSuffixes)

    if ($UPNSuffixes.Count -eq 0) {
        $UPNSuffixes = @($DomainName)
    }

    $UPNSuffixesText = ($UPNSuffixes -join ", ")

    $UPNSuffixTooltip = ($UPNSuffixes | ForEach-Object {
        "<div class='tooltip-line'><b>$(ConvertTo-HtmlSafe $_)</b><span>Configured UPN suffix</span></div>"
    }) -join ""
}
catch {
    $UPNSuffixesText = $DomainName
    $UPNSuffixTooltip = "<div class='tooltip-line'><b>$DomainName</b><span>Default domain suffix</span></div>"
}

try {
    $RecycleBinFeature = Get-ADOptionalFeature -Identity "Recycle Bin Feature" -ErrorAction Stop

    if ($RecycleBinFeature.EnabledScopes.Count -gt 0) {
        $ADRecycleBinStatus = "Enabled"
        $RecycleBinHealth = "Good"
    }
    else {
        $ADRecycleBinStatus = "Disabled"
        $RecycleBinHealth = "Warning"
    }
}
catch {
    $ADRecycleBinStatus = "N/A"
    $RecycleBinHealth = "Unknown"
}

try {
    $DirectoryServicePath = "CN=Directory Service,CN=Windows NT,CN=Services,$ConfigNC"
    $DirectoryService = Get-ADObject -Identity $DirectoryServicePath -Properties tombstoneLifetime -ErrorAction Stop

    if ($DirectoryService.tombstoneLifetime) {
        $TombstoneLifetime = "$($DirectoryService.tombstoneLifetime) Days"
    }
    else {
        $TombstoneLifetime = "Default / Not Set"
    }
}
catch {
    $TombstoneLifetime = "N/A"
}

$DomainControllers = @(Get-ADDomainController -Filter *)
$DomainControllerCount = $DomainControllers.Count

# ---------------------------------------------------------------------------
# Lightweight DC ICMP Ping Check
# ---------------------------------------------------------------------------

$DCPingResults = @(
    foreach ($DC in $DomainControllers) {
        $PingStatus = Test-Connection -ComputerName $DC.HostName -Count 1 -Quiet -ErrorAction SilentlyContinue

        [pscustomobject]@{
            DomainController = $DC.HostName
            IPv4Address      = $DC.IPv4Address
            Site             = $DC.Site
            ICMPStatus       = if ($PingStatus) { "Reachable" } else { "Not Responding" }
        }
    }
)

$DCPingFailedCount = @($DCPingResults | Where-Object { $_.ICMPStatus -ne "Reachable" }).Count

if ($DCPingFailedCount -gt 0) {
    $DCPingHealth = "Warning"
    $DCPingMessage = "$DCPingFailedCount DC(s) not responding to ICMP"
}
else {
    $DCPingHealth = "Good"
    $DCPingMessage = "All DCs responding to ICMP"
}

$DCPingTooltip = ($DCPingResults | ForEach-Object {
    "<div class='tooltip-line'><b>$($_.DomainController)</b><span>$($_.ICMPStatus) | IP: $($_.IPv4Address) | Site: $($_.Site)</span></div>"
}) -join ""

$DCPingCsvJs = ConvertTo-JavaScriptString (
    (
        $DCPingResults |
        ConvertTo-Csv -NoTypeInformation
    ) -join "`n"
)

$DomainControllersCsv = $DomainControllers |
    Select-Object `
        @{Name="DomainController";Expression={$_.HostName}},
        IPv4Address,
        @{Name="IsGC";Expression={$_.IsGlobalCatalog}},
        OperatingSystem,
        Site |
    ConvertTo-Csv -NoTypeInformation

$DomainControllersCsvJs = ConvertTo-JavaScriptString ($DomainControllersCsv -join "`n")

try {
    $ADSites = @(Get-ADReplicationSite -Filter * -ErrorAction Stop)
    $ADSiteCount = $ADSites.Count

    $ADSubnets = @(Get-ADReplicationSubnet -Filter * -Properties Site, Location, Description -ErrorAction Stop)
    $ADSubnetCount = $ADSubnets.Count

    $SitesWithNoDCs = @(
        foreach ($Site in $ADSites) {
            $DCsInSite = @($DomainControllers | Where-Object { $_.Site -eq $Site.Name })

            if ($DCsInSite.Count -eq 0) {
                $Site
            }
        }
    )

    $SitesWithNoDCCount = $SitesWithNoDCs.Count

    if ($SitesWithNoDCCount -gt 0) {
        $SitesNoDCHealth = "Warning"
    }
        else {
        $SitesNoDCHealth = "Good"
    }


    $ADSitesTooltip = ($ADSites | Sort-Object Name | ForEach-Object {
        "<div class='tooltip-line'><b>$(ConvertTo-HtmlSafe $_.Name)</b><span>Active Directory Site</span></div>"
    }) -join ""

    if (!$ADSitesTooltip) {
        $ADSitesTooltip = "<div class='tooltip-line'>No AD sites found</div>"
    }

    $SitesWithNoDCTooltip = ($SitesWithNoDCs | Sort-Object Name | ForEach-Object {
        "<div class='tooltip-line'><b>$(ConvertTo-HtmlSafe $_.Name)</b><span>No domain controller found in this site</span></div>"
    }) -join ""

    if (!$SitesWithNoDCTooltip) {
        $SitesWithNoDCTooltip = "<div class='tooltip-line'>All AD sites have at least one domain controller</div>"
    }

    $ADSubnetsTooltip = ($ADSubnets | Sort-Object Name | ForEach-Object {
        $SubnetName = ConvertTo-HtmlSafe $_.Name
        $SiteName = ConvertTo-HtmlSafe (Get-CleanADSiteName $_.Site)

        "<div class='tooltip-line'><b>$SubnetName</b><span>Site: $SiteName</span></div>"
    }) -join ""

    if (!$ADSubnetsTooltip) {
        $ADSubnetsTooltip = "<div class='tooltip-line'>No AD subnets found</div>"
    }

    $ADSitesCsvJs = ConvertTo-JavaScriptString (
        (
            $ADSites |
            Sort-Object Name |
            Select-Object Name, Description |
            ConvertTo-Csv -NoTypeInformation
        ) -join "`n"
    )

    $SitesWithNoDCsCsvJs = ConvertTo-JavaScriptString (
        (
            $SitesWithNoDCs |
            Sort-Object Name |
            Select-Object Name, Description |
            ConvertTo-Csv -NoTypeInformation
        ) -join "`n"
    )

    $ADSubnetsCsvJs = ConvertTo-JavaScriptString (
        (
            $ADSubnets |
            Sort-Object Name |
            Select-Object `
                Name,
                @{Name="Site";Expression={Get-CleanADSiteName $_.Site}},
                Location,
                Description |
            ConvertTo-Csv -NoTypeInformation
        ) -join "`n"
    )
}
catch {
    $ADSites = $null
    $ADSubnets = $null
    $SitesWithNoDCs = @()

    $ADSiteCount = "N/A"
    $SitesWithNoDCCount = "N/A"
    $ADSubnetCount = "N/A"

    $ADSitesTooltip = "<div class='tooltip-line'>Unable to read AD sites</div>"
    $SitesWithNoDCTooltip = "<div class='tooltip-line'>Unable to calculate sites with no DCs</div>"
    $ADSubnetsTooltip = "<div class='tooltip-line'>Unable to read AD subnets</div>"

    $ADSitesCsvJs = ""
    $SitesWithNoDCsCsvJs = ""
    $ADSubnetsCsvJs = ""
}

$FSMORoles = [ordered]@{
    "Schema Master"         = $Forest.SchemaMaster
    "Domain Naming Master"  = $Forest.DomainNamingMaster
    "PDC Emulator"          = $Domain.PDCEmulator
    "RID Master"            = $Domain.RIDMaster
    "Infrastructure Master" = $Domain.InfrastructureMaster
}

$FSMORoleHolders = @($FSMORoles.Values | Select-Object -Unique)
$FSMORoleHolderCount = $FSMORoleHolders.Count

$FSMOTooltip = ($FSMORoles.GetEnumerator() | ForEach-Object {
    "<div class='tooltip-line'><b>$(ConvertTo-HtmlSafe $_.Key)</b><span>$(ConvertTo-HtmlSafe $_.Value)</span></div>"
}) -join ""

try {
    $PasswordPolicy = Get-ADDefaultDomainPasswordPolicy -ErrorAction Stop

    $PasswordPolicySummary = "Min $($PasswordPolicy.MinPasswordLength) | Max $([math]::Abs($PasswordPolicy.MaxPasswordAge.Days)) Days"

    $PasswordPolicyTooltip = @"
<div class='tooltip-line'><b>Minimum Password Length</b><span>$($PasswordPolicy.MinPasswordLength)</span></div>
<div class='tooltip-line'><b>Maximum Password Age</b><span>$(Convert-TimeSpanToReadable $PasswordPolicy.MaxPasswordAge)</span></div>
<div class='tooltip-line'><b>Minimum Password Age</b><span>$(Convert-TimeSpanToReadable $PasswordPolicy.MinPasswordAge)</span></div>
<div class='tooltip-line'><b>Password History Count</b><span>$($PasswordPolicy.PasswordHistoryCount)</span></div>
<div class='tooltip-line'><b>Complexity Enabled</b><span>$($PasswordPolicy.ComplexityEnabled)</span></div>
<div class='tooltip-line'><b>Reversible Encryption Enabled</b><span>$($PasswordPolicy.ReversibleEncryptionEnabled)</span></div>
<div class='tooltip-line'><b>Lockout Threshold</b><span>$($PasswordPolicy.LockoutThreshold)</span></div>
<div class='tooltip-line'><b>Lockout Duration</b><span>$(Convert-TimeSpanToReadable $PasswordPolicy.LockoutDuration)</span></div>
<div class='tooltip-line'><b>Lockout Observation Window</b><span>$(Convert-TimeSpanToReadable $PasswordPolicy.LockoutObservationWindow)</span></div>
"@
}
catch {
    $PasswordPolicySummary = "N/A"
    $PasswordPolicyTooltip = "<div class='tooltip-line'>Unable to read password policy</div>"
}

$DomainAdmins     = Get-ADGroupMemberDetailsSafe -Identity "Domain Admins"
$EnterpriseAdmins = Get-ADGroupMemberDetailsSafe -Identity "Enterprise Admins"
$SchemaAdmins     = Get-ADGroupMemberDetailsSafe -Identity "Schema Admins"

$DomainAdminsCsvJs = ConvertTo-JavaScriptString ((Get-PrivilegedGroupCsv -Identity "Domain Admins" -DomainName $DomainName) -join "`n")
$EnterpriseAdminsCsvJs = ConvertTo-JavaScriptString ((Get-PrivilegedGroupCsv -Identity "Enterprise Admins" -DomainName $DomainName) -join "`n")
$SchemaAdminsCsvJs = ConvertTo-JavaScriptString ((Get-PrivilegedGroupCsv -Identity "Schema Admins" -DomainName $DomainName) -join "`n")

try {
    Import-Module DnsServer -ErrorAction Stop

    $DnsZones = @(Get-DnsServerZone -ErrorAction Stop)

    $DnsZoneCount = $DnsZones.Count
    $ADIntegratedDnsZoneCount = @($DnsZones | Where-Object { $_.IsDsIntegrated -eq $true }).Count
    $StandaloneDnsZoneCount = @($DnsZones | Where-Object { $_.IsDsIntegrated -ne $true }).Count

    $DnsZonesCsvJs = ConvertTo-JavaScriptString (
        (
            $DnsZones |
            Select-Object ZoneName, ZoneType, IsDsIntegrated, ReplicationScope, DynamicUpdate |
            ConvertTo-Csv -NoTypeInformation
        ) -join "`n"
    )
}
catch {
    $DnsZones = $null
    $DnsZoneCount = "N/A"
    $ADIntegratedDnsZoneCount = "N/A"
    $StandaloneDnsZoneCount = "N/A"
    $DnsZonesCsvJs = ""
}

try {
    Import-Module GroupPolicy -ErrorAction Stop

    $AllGPOs = @(Get-GPO -All)
    $GPOCount = $AllGPOs.Count

    $DisabledGPOs = @($AllGPOs | Where-Object {
        $_.GpoStatus -ne "AllSettingsEnabled"
    })

    $DisabledGPOCount = $DisabledGPOs.Count

    $UnlinkedGPOs = @(
        foreach ($GPO in $AllGPOs) {
            try {
                [xml]$Report = Get-GPOReport -Guid $GPO.Id -ReportType Xml

                if ($Report.GPO.LinksTo.Count -eq 0) {
                    $GPO
                }
            }
            catch {}
        }
    )

    $UnlinkedGPOCount = $UnlinkedGPOs.Count

    $GpoCsvJs = ConvertTo-JavaScriptString (
        (
            $AllGPOs |
            Select-Object DisplayName, Id, Owner, GpoStatus, CreationTime, ModificationTime |
            ConvertTo-Csv -NoTypeInformation
        ) -join "`n"
    )

    $UnlinkedGpoCsvJs = ConvertTo-JavaScriptString (
        (
            $UnlinkedGPOs |
            Select-Object DisplayName, Id, Owner, CreationTime, ModificationTime |
            ConvertTo-Csv -NoTypeInformation
        ) -join "`n"
    )
}
catch {
    $GPOCount = "N/A"
    $DisabledGPOCount = "N/A"
    $UnlinkedGPOCount = "N/A"
    $GpoCsvJs = ""
    $UnlinkedGpoCsvJs = ""
}

$TotalUsers = @(Get-ADUser -LDAPFilter "(objectCategory=person)" -ResultSetSize $null).Count

$ActiveUsers = @(Get-ADUser `
    -LDAPFilter "(&(objectCategory=person)(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))" `
    -ResultSetSize $null).Count

$DisabledUsers = @(Get-ADUser `
    -LDAPFilter "(&(objectCategory=person)(objectClass=user)(userAccountControl:1.2.840.113556.1.4.803:=2))" `
    -ResultSetSize $null).Count

$TotalGroups = @(Get-ADGroup -LDAPFilter "(objectClass=group)" -ResultSetSize $null).Count

$SecurityGroups = @(Get-ADGroup `
    -LDAPFilter "(&(objectClass=group)(groupType:1.2.840.113556.1.4.803:=2147483648))" `
    -ResultSetSize $null).Count

$DistributionGroups = @(Get-ADGroup `
    -LDAPFilter "(&(objectClass=group)(!(groupType:1.2.840.113556.1.4.803:=2147483648)))" `
    -ResultSetSize $null).Count

$TotalComputers = @(Get-ADComputer -LDAPFilter "(objectCategory=computer)" -ResultSetSize $null).Count

$ActiveComputers = @(Get-ADComputer `
    -LDAPFilter "(&(objectCategory=computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))" `
    -ResultSetSize $null).Count

$DisabledComputers = @(Get-ADComputer `
    -LDAPFilter "(&(objectCategory=computer)(userAccountControl:1.2.840.113556.1.4.803:=2))" `
    -ResultSetSize $null).Count

try {
    $Trusts = @(Get-ADTrust -Filter * -ErrorAction Stop)
}
catch {
    $Trusts = $null
}

# ---------------------------------------------------------------------------
# Domain Controller Health tab - data collection
#
# DC list/sidebar data (name, site, GC flag, OS, FSMO roles) is always built
# from $DomainControllers above. Per-DC hardware/performance/disk/service/
# replication detail is collected agentlessly via Get-DCInventory.ps1
# (CIM/WinRM) when -IncludeDCHealth is set (default). The DC list is built
# dynamically for however many DCs exist - nothing is capped or hardcoded -
# the sidebar list simply scrolls once it exceeds the visible area.
# ---------------------------------------------------------------------------

$FSMOByServer = @{}

foreach ($RoleEntry in $FSMORoles.GetEnumerator()) {
    $RoleServer = $RoleEntry.Value

    if (![string]::IsNullOrWhiteSpace($RoleServer)) {
        if (-not $FSMOByServer.ContainsKey($RoleServer)) {
            $FSMOByServer[$RoleServer] = @()
        }

        $FSMOByServer[$RoleServer] += $RoleEntry.Key
    }
}

$DCListData = @(
    $DomainControllers | ForEach-Object {
        $ShortRoles = @($FSMOByServer[$_.HostName] | ForEach-Object {
            switch ($_) {
                "Schema Master"         { "Schema Master" }
                "Domain Naming Master"  { "Domain Naming Master" }
                "PDC Emulator"          { "PDC Emulator" }
                "RID Master"            { "RID Master" }
                "Infrastructure Master" { "Infrastructure Master" }
                default                 { $_ }
            }
        })

        [ordered]@{
            Name            = $_.HostName
            IPv4Address     = $_.IPv4Address
            Site            = Get-CleanADSiteName $_.Site
            IsGlobalCatalog = [bool]$_.IsGlobalCatalog
            OperatingSystem = $_.OperatingSystem
            FSMORoles       = $ShortRoles
        }
    }
)

$DCInventory = @()

if ($IncludeDCHealth) {
    $CollectorScriptPath = Join-Path -Path $PSScriptRoot -ChildPath "Get-DCInventory.ps1"

    if (Test-Path $CollectorScriptPath) {
        try {
            . $CollectorScriptPath

            $InventoryParams = @{
                ComputerName  = @($DomainControllers.HostName)
                ThrottleLimit = $DCHealthThrottleLimit
            }

            if ($DCHealthCredential) {
                $InventoryParams.Credential = $DCHealthCredential
            }

            $DCInventory = @(Get-AllDCInventory @InventoryParams)
        }
        catch {
            Write-Host "Warning: DC Health collection failed - $($_.Exception.Message)" -ForegroundColor Yellow
            $DCInventory = @()
        }
    }
    else {
        Write-Host "Warning: Get-DCInventory.ps1 not found alongside this script - Domain Controller Health tab will show DC list only." -ForegroundColor Yellow
    }
}

$DCListJson = ($DCListData | ConvertTo-Json -Depth 6 -Compress) -replace "</script>", "<\/script>"

if ($DCInventory.Count -gt 0) {
    $DCInventoryJson = ($DCInventory | ConvertTo-Json -Depth 8 -Compress) -replace "</script>", "<\/script>"
}
else {
    $DCInventoryJson = "[]"
}

$html = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Enterprise Active Directory Operational Dashboard</title>

<style>
    body {
        margin: 0;
        padding: 0;
        font-family: "Segoe UI", Arial, sans-serif;
        background: #f5f7fb;
        color: #1f2937;
    }

    .page-generated {
        width: 950px;
        margin: 18px auto 0 auto;
        text-align: right;
        font-size: 13px;
        color: #64748b;
    }

    .container {
        width: 92%;
        margin: 20px auto 30px auto;
    }

    .header {
        max-width: 950px;
        margin: 0 auto 18px auto;
        background: linear-gradient(90deg, #dbeafe, #eff6ff);
        padding: 22px;
        text-align: center;
        border-radius: 14px;
        border: 1px solid #bfdbfe;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.08);
    }

    .header-title {
        font-size: 30px;
        font-weight: 800;
        color: #0f172a;
        letter-spacing: 0.2px;
    }

    .config-section {
        max-width: 950px;
        margin: 0 auto 28px auto;
        background: linear-gradient(135deg, #eef4ff, #f6f9ff);
        border: 1px solid #dbe7ff;
        border-radius: 18px;
        padding: 22px;
        box-shadow: 0 8px 22px rgba(15, 23, 42, 0.07);
    }

    .config-title {
        font-size: 20px;
        font-weight: 800;
        color: #111827;
        border-left: 5px solid #2563eb;
        padding-left: 10px;
        margin-bottom: 20px;
    }

    .config-grid {
        display: grid;
        grid-template-columns: repeat(2, 1fr);
        gap: 16px;
    }

    .config-tile {
        position: relative;
        background: linear-gradient(180deg,#f8fbff,#eef5ff);
        border: 1px solid #e5e7eb;
        border-left: 5px solid #2563eb;
        border-radius: 12px;
        padding: 10px 16px;
        min-height: 75px;
        box-shadow: 0 6px 16px rgba(15, 23, 42, 0.06);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
        overflow: visible;
    }

    .config-tile:hover {
        transform: translateY(-4px);
        box-shadow: 0 14px 30px rgba(15, 23, 42, 0.14);
        z-index: 9999;
    }

    .config-label {
        font-size: 13px;
        color: #64748b;
        margin-bottom: 8px;
        font-weight: 600;
    }

    .config-value {
        font-size: 20px;
        color: #111827;
        font-weight: 800;
        word-break: break-word;
    }

    .config-value-small {
        font-size: 15px;
        color: #111827;
        font-weight: 800;
        line-height: 1.35;
        word-break: break-word;
    }

    .config-footer {
        font-size: 12px;
        color: #94a3b8;
        margin-top: 8px;
    }

    .dashboard {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 22px;
        max-width: 950px;
        margin: 0 auto;
        overflow: visible;
    }

    .tile {
        position: relative;
        background: #ffffff;
        color: #111827;
        border: 1px solid #e5e7eb;
        min-height: 118px;
        border-radius: 18px;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        padding: 18px;
        box-shadow: 0 8px 20px rgba(15, 23, 42, 0.08);
        transition: transform 0.2s ease, box-shadow 0.2s ease, border-color 0.2s ease;
        overflow: visible;
    }

    .tile:hover {
        transform: translateY(-5px);
        box-shadow: 0 14px 30px rgba(15, 23, 42, 0.14);
        border-color: #93c5fd;
        z-index: 9999;
    }

    .tile-top {
        display: flex;
        justify-content: space-between;
        align-items: flex-start;
        gap: 12px;
    }

    .tile-title {
        font-size: 14px;
        font-weight: 600;
        color: #475569;
        line-height: 1.35;
    }

    .tile-actions {
        display: flex;
        align-items: center;
        gap: 6px;
    }

    .tile-icon {
        width: 34px;
        height: 34px;
        border-radius: 12px;
        background: #eff6ff;
        color: #2563eb;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 18px;
        font-weight: 700;
    }

    .export-btn {
        width: 26px;
        height: 26px;
        border-radius: 8px;
        border: 1px solid #bfdbfe;
        background: #eff6ff;
        color: #2563eb;
        font-size: 14px;
        font-weight: 700;
        cursor: pointer;
        display: flex;
        align-items: center;
        justify-content: center;
    }

    .export-btn:hover {
        background: #dbeafe;
        border-color: #93c5fd;
    }

    .section-export-btn-inline {
        float: right;
        width: 26px;
        height: 26px;
        border-radius: 8px;
        border: 1px solid rgba(255,255,255,0.55);
        background: rgba(255,255,255,0.16);
        color: #ffffff;
        font-size: 14px;
        font-weight: 700;
        cursor: pointer;
        margin-top: -3px;
    }

    .section-export-btn-inline:hover {
        background: rgba(255,255,255,0.28);
        border-color: #ffffff;
    }

    .tile-value {
        font-size: 34px;
        font-weight: 800;
        color: #0f172a;
        line-height: 1;
    }

    .tile-value-small {
        font-size: 17px;
        font-weight: 800;
        color: #0f172a;
        line-height: 1.35;
    }

    .tile-footer {
        font-size: 12px;
        color: #94a3b8;
        margin-top: 8px;
    }

    .health {
        display: inline-block;
        margin-top: 8px;
        padding: 3px 9px;
        border-radius: 999px;
        font-size: 11px;
        font-weight: 700;
    }

    .health-good {
        background: #dcfce7;
        color: #166534;
    }

    .health-warning {
        background: #fef3c7;
        color: #92400e;
    }

    .health-critical {
        background: #fee2e2;
        color: #991b1b;
    }

    .health-unknown {
        background: #e5e7eb;
        color: #374151;
    }

    .tile.accent-blue {
        border-top: 4px solid #3b82f6;
    }

    .tile.accent-green {
        border-top: 4px solid #22c55e;
    }

    .tile.accent-orange {
        border-top: 4px solid #f59e0b;
    }

    .tile.accent-purple {
        border-top: 4px solid #8b5cf6;
    }

    .tile.accent-red {
        border-top: 4px solid #ef4444;
    }

    .tile.accent-cyan {
        border-top: 4px solid #06b6d4;
    }

    .tooltip-box {
        display: none;
        position: absolute;
        left: 20px;
        top: 112px;
        width: 430px;
        max-height: 360px;
        overflow-y: auto;
        background: #0f172a;
        color: #f8fafc;
        border-radius: 12px;
        padding: 14px 16px;
        font-size: 13px;
        line-height: 1.6;
        z-index: 99999;
        box-shadow: 0 16px 36px rgba(15, 23, 42, 0.35);
    }

    .tile:hover .tooltip-box,
    .config-tile:hover .tooltip-box {
        display: block;
    }

    .tooltip-line {
        padding: 5px 0;
        border-bottom: 1px solid rgba(255,255,255,0.12);
    }

    .tooltip-line:last-child {
        border-bottom: none;
    }

    .tooltip-line b {
        display: block;
        color: #ffffff;
        font-weight: 700;
    }

    .tooltip-line span {
        display: block;
        color: #cbd5e1;
        font-size: 12px;
    }

    .section-title {
        max-width: 950px;
        margin: 38px auto 15px auto;
        font-size: 20px;
        font-weight: 700;
        color: #111827;
        border-left: 5px solid #2563eb;
        padding-left: 10px;
        position: relative;
    }

    table {
        width: 950px;
        margin: 0 auto 25px auto;
        border-collapse: collapse;
        background: white;
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 8px 20px rgba(15, 23, 42, 0.08);
    }

    th {
        background: #2563eb;
        color: white;
        padding: 11px;
        text-align: left;
        font-size: 14px;
    }

    td {
        padding: 10px;
        border-bottom: 1px solid #e5e7eb;
        font-size: 13px;
    }

    tr:hover {
        background: #f8fafc;
    }

    .footer {
        margin-top: 35px;
        text-align: center;
        font-size: 12px;
        color: #64748b;
    }

    /* =====================================================================
       Header subtitle / footer copyright
       ===================================================================== */

    .header-subtitle {
        margin-top: 8px;
        font-size: 13px;
        font-weight: 500;
        color: #475569;
    }

    .footer-line {
        margin: 4px 0;
    }

    .footer-copyright {
        color: #94a3b8;
    }

    /* =====================================================================
       Tab navigation
       ===================================================================== */

    .tab-bar {
        max-width: 950px;
        margin: 0 auto 24px auto;
        display: flex;
        gap: 4px;
        background-color: #f3f2f1;
        border: 2px solid #ffffff;
        box-shadow: 0 4px 14px rgba(15, 23, 42, 0.06);
        border-radius: 10px;
        padding: 4px;
    }

    .tab-button {
        appearance: none;
        border: none;
        background: transparent;
        cursor: pointer;
        padding: 10px 18px;
        font-size: 15px;
        font-weight: 700;
        font-family: inherit;
        color: #616161;
        border-radius: 8px;
        flex: 1;
        transition: color 0.15s ease, background-color 0.15s ease;
    }

    .tab-button:hover:not(.active) {
        color: #201f1e;
        background-color: #e1dfdd;
    }

    .tab-button.active {
        color: #ffffff;
        background-color: #0078d4;
    }

    .tab-content {
        display: none;
    }

    .tab-content.active {
        display: block;
    }

    /* =====================================================================
       Domain Controller Health tab
       ===================================================================== */

    .dc-health-grid {
        display: grid;
        grid-template-columns: 260px minmax(0, 1fr);
        gap: 20px;
        max-width: 950px;
        margin: 0 auto;
        align-items: start;
    }

    .dc-sidebar {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 14px;
        padding: 14px;
        box-shadow: 0 8px 20px rgba(15, 23, 42, 0.06);
    }

    .dc-search-wrap {
        position: relative;
        margin-bottom: 10px;
    }

    .dc-search-icon {
        position: absolute;
        left: 10px;
        top: 50%;
        transform: translateY(-50%);
        color: #94a3b8;
        pointer-events: none;
        display: flex;
        align-items: center;
    }

    .dc-search-input {
        width: 100%;
        box-sizing: border-box;
        padding: 9px 10px 9px 32px;
        border: 1px solid #cbd5e1;
        border-radius: 8px;
        font-size: 13px;
        font-family: inherit;
        margin-bottom: 0;
    }

    .dc-search-input:focus {
        outline: none;
        border-color: #2563eb;
    }

    .dc-count {
        font-size: 11px;
        color: #94a3b8;
        margin-bottom: 8px;
    }

    .dc-list {
        display: flex;
        flex-direction: column;
        gap: 2px;
        max-height: 420px;
        overflow-y: auto;
        padding-right: 4px;
    }

    .dc-list::-webkit-scrollbar {
        width: 8px;
    }

    .dc-list::-webkit-scrollbar-track {
        background: #f1f5f9;
        border-radius: 8px;
    }

    .dc-list::-webkit-scrollbar-thumb {
        background: #cbd5e1;
        border-radius: 8px;
    }

    .dc-list::-webkit-scrollbar-thumb:hover {
        background: #94a3b8;
    }

    .dc-item {
        padding: 8px 10px;
        border-radius: 8px;
        font-size: 13px;
        color: #334155;
        cursor: pointer;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 6px;
    }

    .dc-item:hover {
        background: #f1f5f9;
    }

    .dc-item.active {
        background: #2563eb;
        color: #ffffff;
        font-weight: 700;
    }

    .dc-item .dc-item-dot {
        width: 8px;
        height: 8px;
        border-radius: 50%;
        flex-shrink: 0;
    }

    .dc-item-name {
        flex: 1;
        min-width: 0;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
    }

    .dc-item-dot.dot-good { background: #22c55e; }
    .dc-item-dot.dot-warning { background: #f59e0b; }
    .dc-item-dot.dot-critical { background: #ef4444; }
    .dc-item-dot.dot-unknown { background: #94a3b8; }

    .dc-empty {
        padding: 30px 10px;
        text-align: center;
        font-size: 12px;
        color: #94a3b8;
    }

    .dc-detail {
        min-height: 200px;
    }

    .dc-detail-header {
        display: flex;
        align-items: center;
        gap: 10px;
        flex-wrap: wrap;
        margin-bottom: 16px;
    }

    .dc-detail-title {
        font-size: 20px;
        font-weight: 800;
        color: #0f172a;
    }

    .dc-badge {
        background: #eff6ff;
        color: #2563eb;
        font-size: 11px;
        font-weight: 700;
        padding: 4px 10px;
        border-radius: 999px;
    }

    .dc-tile-grid {
        display: grid;
        grid-template-columns: repeat(2, 1fr);
        gap: 16px;
        margin-bottom: 16px;
    }

    .dc-tile {
        position: relative;
        background: #ffffff;
        border: 1px solid #eef1f5;
        border-radius: 14px;
        padding: 16px 18px;
        box-shadow: 0 4px 14px rgba(15, 23, 42, 0.045);
        transition: transform 0.2s ease, box-shadow 0.2s ease, border-color 0.2s ease;
    }

    .dc-tile:hover {
        transform: translateY(-4px);
        box-shadow: 0 16px 32px rgba(15, 23, 42, 0.12);
        border-color: #dbeafe;
        z-index: 9999;
    }

    .dc-tile-title {
        font-size: 13px;
        font-weight: 700;
        color: #1d4ed8;
        margin-bottom: 12px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 8px;
    }

    .dc-tile-title-left {
        display: flex;
        align-items: center;
    }

    .dc-tile-icon {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        margin-right: 8px;
        color: #94a3b8;
        flex-shrink: 0;
    }

    .dc-section-title {
        font-size: 12px;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        color: #94a3b8;
        margin: 0 0 10px 0;
        padding-bottom: 6px;
        border-bottom: 1px solid #e5e7eb;
        grid-column: 1 / -1;
    }

    .dc-section-title.dc-section-title-spaced {
        margin-top: 6px;
    }

    .dc-pill {
        font-size: 11px;
        font-weight: 800;
        padding: 3px 10px;
        border-radius: 999px;
        white-space: nowrap;
    }

    .dc-pill-good { background: #ecfdf5; color: #047857; }
    .dc-pill-warning { background: #fffbeb; color: #92400e; }
    .dc-pill-critical { background: #fef2f2; color: #991b1b; }
    .dc-pill-unknown { background: #f1f5f9; color: #64748b; }

    .dc-tile-tip {
        display: none;
        position: absolute;
        top: 100%;
        left: 16px;
        right: 16px;
        margin-top: 6px;
        background: #0f172a;
        color: #e2e8f0;
        font-size: 12px;
        font-weight: 400;
        line-height: 1.5;
        padding: 10px 12px;
        border-radius: 10px;
        box-shadow: 0 12px 24px rgba(15, 23, 42, 0.25);
        z-index: 20;
        text-align: left;
    }

    .dc-tile-tip b {
        color: #ffffff;
    }

    .dc-tile-tip a {
        color: #93c5fd;
        text-decoration: underline;
    }

    .dc-tile-tip a:hover {
        color: #bfdbfe;
    }

    .dc-tile:hover .dc-tile-tip {
        display: block;
    }

    .dc-kv-row {
        display: flex;
        justify-content: space-between;
        gap: 10px;
        font-size: 13px;
        padding: 4px 0;
        border-bottom: 1px solid #f1f5f9;
    }

    .dc-kv-row:last-child {
        border-bottom: none;
    }

    .dc-kv-row span:first-child {
        color: #64748b;
    }

    .dc-kv-row span:last-child {
        font-weight: 700;
        color: #0f172a;
        text-align: right;
    }

    .gauge-wrap {
        display: flex;
        align-items: center;
        gap: 12px;
    }

    .gauge-value {
        font-size: 28px;
        font-weight: 800;
        color: #0f172a;
        min-width: 70px;
    }

    .gauge-bar {
        flex: 1;
        height: 10px;
        border-radius: 999px;
        background: #e5e7eb;
        overflow: hidden;
    }

    .gauge-fill {
        height: 100%;
        border-radius: 999px;
    }

    .gauge-fill.gauge-good { background: #22c55e; }
    .gauge-fill.gauge-warning { background: #f59e0b; }
    .gauge-fill.gauge-critical { background: #ef4444; }

    .dc-disk-row {
        display: grid;
        grid-template-columns: 60px 1fr 90px 60px;
        gap: 10px;
        align-items: center;
        font-size: 12px;
        padding: 6px 0;
        border-bottom: 1px solid #f1f5f9;
    }

    .dc-disk-row:last-child {
        border-bottom: none;
    }

    .dc-disk-drive {
        font-weight: 800;
        color: #0f172a;
    }

    .dc-disk-role {
        color: #64748b;
        font-size: 11px;
    }

    .dc-disk-free {
        text-align: right;
        font-weight: 700;
        color: #0f172a;
    }

    .dc-service-grid {
        display: grid;
        grid-template-columns: repeat(2, 1fr);
        gap: 8px;
    }

    .dc-service-row {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 8px;
        font-size: 12px;
        padding: 6px 8px;
        border-radius: 8px;
        background: #f8fafc;
    }

    .dc-mini-table {
        width: 100%;
        border-collapse: collapse;
        margin: 0;
        border-radius: 10px;
        box-shadow: none;
    }

    .dc-mini-table th {
        background: #f1f5f9;
        color: #475569;
        font-size: 11px;
        text-transform: uppercase;
        letter-spacing: 0.03em;
        padding: 8px 10px;
        text-align: left;
    }

    .dc-mini-table td {
        font-size: 12px;
        padding: 8px 10px;
        border-bottom: 1px solid #f1f5f9;
    }

    .dc-tile.dc-tile-wide {
        grid-column: 1 / -1;
    }

    .dc-unreachable {
        padding: 40px 20px;
        text-align: center;
        color: #991b1b;
        background: #fef2f2;
        border: 1px solid #fecaca;
        border-radius: 14px;
        font-size: 13px;
    }

    @media only screen and (max-width: 900px) {
        .dashboard,
        .config-grid {
            grid-template-columns: 1fr;
        }

        table,
        .page-generated {
            width: 100%;
        }

        .header-title {
            font-size: 24px;
        }

        .tooltip-box {
            width: 280px;
        }

        .dc-health-grid {
            grid-template-columns: 1fr;
        }

        .dc-tile-grid,
        .dc-service-grid {
            grid-template-columns: 1fr;
        }
    }
</style>
</head>

<body>

<div class="page-generated">Generated On: $GeneratedOn</div>

<div class="container">

    <div class="header">
        <div class="header-title">Enterprise Active Directory Health Dashboard</div>
        <div class="header-subtitle">Forest: $ForestRootDomain &nbsp;|&nbsp; Domain: $DomainName &nbsp;|&nbsp; $DomainControllerCount Domain Controllers &nbsp;|&nbsp; Generated: $GeneratedOn</div>
    </div>

    <div class="tab-bar">
        <button class="tab-button active" id="tab-btn-forest" onclick="showTab('forest')">Forest Overview</button>
        <button class="tab-button" id="tab-btn-dchealth" onclick="showTab('dchealth')">Domain Controller Health</button>
        <button class="tab-button" id="tab-btn-deepinsights" onclick="showTab('deepinsights')">Deep Insights</button>
        <button class="tab-button" id="tab-btn-secposture" onclick="showTab('secposture')">Security Posture</button>
    </div>

    <div id="tab-forest" class="tab-content active">

    <div class="config-section">
        <div class="config-title">Forest & Domain Configuration</div>

        <div class="config-grid">

            <div class="config-tile">
                <div class="config-label">Forest Root Domain</div>
                <div class="config-value">$ForestRootDomain</div>
                <div class="config-footer">Forest root domain</div>
            </div>

            <div class="config-tile">
                <div class="config-label">Domains</div>
                <div class="config-value">$DomainCount</div>
                <div class="config-footer">Hover for details</div>
                <div class="tooltip-box">$DomainTooltip</div>
            </div>

            <div class="config-tile">
                <div class="config-label">NetBIOS Name</div>
                <div class="config-value">$NetBIOSName</div>
                <div class="config-footer">AD NetBIOS namespace</div>
            </div>

            <div class="config-tile">
                <div class="config-label">UPN Domain Suffixes</div>
                <div class="config-value-small">$UPNSuffixesText</div>
                <div class="config-footer">Additional UPN suffixes configured in forest</div>
                <div class="tooltip-box">$UPNSuffixTooltip</div>
            </div>

            <div class="config-tile">
                <div class="config-label">Forest Functional Level</div>
                <div class="config-value">$ForestFunctionalLevel</div>
                <div class="config-footer">Forest-wide functional mode</div>
            </div>

            <div class="config-tile">
                <div class="config-label">Domain Functional Level</div>
                <div class="config-value">$DomainFunctionalLevel</div>
                <div class="config-footer">Current domain functional mode</div>
            </div>

            <div class="config-tile">
                <div class="config-label">AD Recycle Bin Status</div>
                <div class="config-value">$ADRecycleBinStatus</div>
                <span class="health health-$($RecycleBinHealth.ToLower())">$RecycleBinHealth</span>
                <div class="config-footer">Deleted object recovery capability</div>
            </div>

            <div class="config-tile">
                <div class="config-label">Tombstone Lifetime</div>
                <div class="config-value">$TombstoneLifetime</div>
                <div class="config-footer">Deleted object lifetime</div>
            </div>

        </div>
    </div>

    <div class="section-title">Active Directory Operations</div>

    <div class="dashboard">

        <div class="tile accent-blue">
            <div class="tile-top">
                <div class="tile-title">Domain Controllers</div>
                <div class="tile-actions">
                    <button class="export-btn" title="Export DC ICMP Status" onclick="downloadCsv('DC_ICMP_Status.csv', dcPingCsv)">$([char]0x21E9)</button>
                    <div class="tile-icon">DC</div>
                </div>
            </div>
            <div>
                <div class="tile-value">$DomainControllerCount</div>
                <span class="health health-$($DCPingHealth.ToLower())">$DCPingHealth</span>
                <div class="tile-footer">$DCPingMessage</div>
            </div>
            <div class="tooltip-box">$DCPingTooltip</div>
        </div>

        <div class="tile accent-cyan">
            <div class="tile-top">
                <div class="tile-title">FSMO Role Holders</div>
                <div class="tile-icon">FS</div>
            </div>
            <div>
                <div class="tile-value">$FSMORoleHolderCount</div>
                <div class="tile-footer">Hover to view role holder names</div>
            </div>
            <div class="tooltip-box">$FSMOTooltip</div>
        </div>

        <div class="tile accent-green">
            <div class="tile-top">
                <div class="tile-title">Default Domain Password Policy</div>
                <div class="tile-icon">PW</div>
            </div>
            <div>
                <div class="tile-value-small">$PasswordPolicySummary</div>
                <div class="tile-footer">Hover to view full policy</div>
            </div>
            <div class="tooltip-box">$PasswordPolicyTooltip</div>
        </div>

        <div class="tile accent-green">
            <div class="tile-top">
                <div class="tile-title">AD Sites</div>
                <div class="tile-actions">
                    <button class="export-btn" title="Export AD Sites" onclick="downloadCsv('AD_Sites.csv', adSitesCsv)">$([char]0x21E9)</button>
                    <div class="tile-icon">ST</div>
                </div>
            </div>
            <div>
                <div class="tile-value">$ADSiteCount</div>
                <div class="tile-footer">Hover to view site names</div>
            </div>
            <div class="tooltip-box">$ADSitesTooltip</div>
        </div>

        <div class="tile accent-orange">
            <div class="tile-top">
                <div class="tile-title">Sites with No DCs</div>
                <div class="tile-actions">
                    <button class="export-btn" title="Export Sites with No DCs" onclick="downloadCsv('Sites_With_No_DCs.csv', sitesWithNoDCsCsv)">$([char]0x21E9)</button>
                    <div class="tile-icon">ND</div>
            </div>
        </div>
     <div>
        <div class="tile-value">$SitesWithNoDCCount</div>
        <span class="health health-$($SitesNoDCHealth.ToLower())">$SitesNoDCHealth</span>
        <div class="tile-footer">Sites without local domain controllers</div>
        </div>
        <div class="tooltip-box">$SitesWithNoDCTooltip</div>
    </div>

        <div class="tile accent-blue">
            <div class="tile-top">
                <div class="tile-title">Subnets</div>
                <div class="tile-actions">
                    <button class="export-btn" title="Export AD Subnets" onclick="downloadCsv('AD_Subnets.csv', adSubnetsCsv)">$([char]0x21E9)</button>
                    <div class="tile-icon">SN</div>
                </div>
            </div>
            <div>
                <div class="tile-value">$ADSubnetCount</div>
                <div class="tile-footer">Hover to view subnet to site mapping</div>
            </div>
            <div class="tooltip-box">$ADSubnetsTooltip</div>
        </div>

        <div class="tile accent-red">
            <div class="tile-top">
                <div class="tile-title">Domain Admins</div>
                <div class="tile-actions">
                    <button class="export-btn" title="Export Domain Admins" onclick="downloadCsv('Domain_Admins.csv', domainAdminsCsv)">$([char]0x21E9)</button>
                    <div class="tile-icon">DA</div>
                </div>
            </div>
            <div>
                <div class="tile-value">$($DomainAdmins.Count)</div>
                <div class="tile-footer">Privileged group members</div>
            </div>
            <div class="tooltip-box">$($DomainAdmins.Tooltip)</div>
        </div>

        <div class="tile accent-purple">
            <div class="tile-top">
                <div class="tile-title">Enterprise Admins</div>
                <div class="tile-actions">
                    <button class="export-btn" title="Export Enterprise Admins" onclick="downloadCsv('Enterprise_Admins.csv', enterpriseAdminsCsv)">$([char]0x21E9)</button>
                    <div class="tile-icon">EA</div>
                </div>
            </div>
            <div>
                <div class="tile-value">$($EnterpriseAdmins.Count)</div>
                <div class="tile-footer">Forest-level privileged members</div>
            </div>
            <div class="tooltip-box">$($EnterpriseAdmins.Tooltip)</div>
        </div>

        <div class="tile accent-orange">
            <div class="tile-top">
                <div class="tile-title">Schema Admins</div>
                <div class="tile-actions">
                    <button class="export-btn" title="Export Schema Admins" onclick="downloadCsv('Schema_Admins.csv', schemaAdminsCsv)">$([char]0x21E9)</button>
                    <div class="tile-icon">SA</div>
                </div>
            </div>
            <div>
                <div class="tile-value">$($SchemaAdmins.Count)</div>
                <div class="tile-footer">Schema modification members</div>
            </div>
            <div class="tooltip-box">$($SchemaAdmins.Tooltip)</div>
        </div>

        <div class="tile accent-orange">
            <div class="tile-top">
                <div class="tile-title">DNS Zones</div>
                <div class="tile-actions">
                    <button class="export-btn" title="Export DNS Zones" onclick="downloadCsv('DNS_Zones.csv', dnsZonesCsv)">$([char]0x21E9)</button>
                    <div class="tile-icon">DNS</div>
                </div>
            </div>
            <div>
                <div class="tile-value">$DnsZoneCount</div>
                <div class="tile-footer">Total DNS zones</div>
            </div>
        </div>

        <div class="tile accent-green">
            <div class="tile-top">
                <div class="tile-title">AD Integrated DNS Zones</div>
                <div class="tile-icon">ADI</div>
            </div>
            <div>
                <div class="tile-value">$ADIntegratedDnsZoneCount</div>
                <div class="tile-footer">Directory-integrated zones</div>
            </div>
        </div>

        <div class="tile accent-blue">
            <div class="tile-top">
                <div class="tile-title">Standalone DNS Zones</div>
                <div class="tile-icon">STD</div>
            </div>
            <div>
                <div class="tile-value">$StandaloneDnsZoneCount</div>
                <div class="tile-footer">File-backed / non-AD zones</div>
            </div>
        </div>

        <div class="tile accent-blue">
            <div class="tile-top">
                <div class="tile-title">GPO Count</div>
                <div class="tile-actions">
                    <button class="export-btn" title="Export GPO Inventory" onclick="downloadCsv('GPO_Inventory.csv', gpoCsv)">$([char]0x21E9)</button>
                    <div class="tile-icon">GP</div>
                </div>
            </div>
            <div>
                <div class="tile-value">$GPOCount</div>
                <div class="tile-footer">Group Policy Objects</div>
            </div>
        </div>

        <div class="tile accent-orange">
            <div class="tile-top">
                <div class="tile-title">Unlinked GPOs</div>
                <div class="tile-actions">
                    <button class="export-btn" title="Export Unlinked GPOs" onclick="downloadCsv('Unlinked_GPOs.csv', unlinkedGpoCsv)">$([char]0x21E9)</button>
                    <div class="tile-icon">UL</div>
                </div>
            </div>
            <div>
                <div class="tile-value">$UnlinkedGPOCount</div>
                <div class="tile-footer">No Site / Domain / OU links</div>
            </div>
        </div>

        <div class="tile accent-red">
            <div class="tile-top">
                <div class="tile-title">Disabled GPOs</div>
                <div class="tile-icon">DG</div>
            </div>
            <div>
                <div class="tile-value">$DisabledGPOCount</div>
                <div class="tile-footer">User / Computer settings disabled</div>
            </div>
        </div>

        <div class="tile accent-blue">
            <div class="tile-top">
                <div class="tile-title">Total User Objects</div>
                <div class="tile-icon">U</div>
            </div>
            <div>
                <div class="tile-value">$TotalUsers</div>
                <div class="tile-footer">All user objects</div>
            </div>
        </div>

        <div class="tile accent-green">
            <div class="tile-top">
                <div class="tile-title">Active User Objects</div>
                <div class="tile-icon">AU</div>
            </div>
            <div>
                <div class="tile-value">$ActiveUsers</div>
                <div class="tile-footer">Enabled users</div>
            </div>
        </div>

        <div class="tile accent-purple">
            <div class="tile-top">
                <div class="tile-title">Disabled User Objects</div>
                <div class="tile-icon">DU</div>
            </div>
            <div>
                <div class="tile-value">$DisabledUsers</div>
                <div class="tile-footer">Disabled users</div>
            </div>
        </div>

        <div class="tile accent-blue">
            <div class="tile-top">
                <div class="tile-title">Total Groups</div>
                <div class="tile-icon">G</div>
            </div>
            <div>
                <div class="tile-value">$TotalGroups</div>
                <div class="tile-footer">All group objects</div>
            </div>
        </div>

        <div class="tile accent-purple">
            <div class="tile-top">
                <div class="tile-title">Security Groups</div>
                <div class="tile-icon">SG</div>
            </div>
            <div>
                <div class="tile-value">$SecurityGroups</div>
                <div class="tile-footer">Security-enabled groups</div>
            </div>
        </div>

        <div class="tile accent-orange">
            <div class="tile-top">
                <div class="tile-title">Distribution Groups</div>
                <div class="tile-icon">DG</div>
            </div>
            <div>
                <div class="tile-value">$DistributionGroups</div>
                <div class="tile-footer">Mail distribution groups</div>
            </div>
        </div>

        <div class="tile accent-cyan">
            <div class="tile-top">
                <div class="tile-title">Total Computer Objects</div>
                <div class="tile-icon">C</div>
            </div>
            <div>
                <div class="tile-value">$TotalComputers</div>
                <div class="tile-footer">All computer objects</div>
            </div>
        </div>

        <div class="tile accent-green">
            <div class="tile-top">
                <div class="tile-title">Active Computer Objects</div>
                <div class="tile-icon">AC</div>
            </div>
            <div>
                <div class="tile-value">$ActiveComputers</div>
                <div class="tile-footer">Enabled computers</div>
            </div>
        </div>

        <div class="tile accent-red">
            <div class="tile-top">
                <div class="tile-title">Disabled Computer Objects</div>
                <div class="tile-icon">DC</div>
            </div>
            <div>
                <div class="tile-value">$DisabledComputers</div>
                <div class="tile-footer">Disabled computers</div>
            </div>
        </div>

    </div>

    <div class="section-title">FSMO Role Holders</div>

    <table>
        <tr>
            <th>Role</th>
            <th>Server</th>
        </tr>
        <tr>
            <td>Schema Master</td>
            <td>$($Forest.SchemaMaster)</td>
        </tr>
        <tr>
            <td>Domain Naming Master</td>
            <td>$($Forest.DomainNamingMaster)</td>
        </tr>
        <tr>
            <td>PDC Emulator</td>
            <td>$($Domain.PDCEmulator)</td>
        </tr>
        <tr>
            <td>RID Master</td>
            <td>$($Domain.RIDMaster)</td>
        </tr>
        <tr>
            <td>Infrastructure Master</td>
            <td>$($Domain.InfrastructureMaster)</td>
        </tr>
    </table>

    <div class="section-title">Domain Controllers</div>

    <table>
        <tr>
            <th>Domain Controller</th>
            <th>IPv4 Address</th>
            <th>IsGC</th>
            <th>Operating System</th>
            <th>
                Site
                <button class="section-export-btn-inline" title="Export Domain Controllers" onclick="downloadCsv('Domain_Controllers.csv', domainControllersCsv)">$([char]0x21E9)</button>
            </th>
        </tr>
"@

foreach ($DC in $DomainControllers) {
    $html += @"
        <tr>
            <td>$($DC.HostName)</td>
            <td>$($DC.IPv4Address)</td>
            <td>$($DC.IsGlobalCatalog)</td>
            <td>$($DC.OperatingSystem)</td>
            <td>$($DC.Site)</td>
        </tr>
"@
}

$html += @"
    </table>

    <div class="section-title">Trust Relationships</div>

    <table>
        <tr>
            <th>Source Domain</th>
            <th>Target Domain</th>
            <th>Direction</th>
            <th>Trust Type</th>
            <th>Transitive</th>
        </tr>
"@

if ($Trusts) {
    foreach ($Trust in $Trusts) {
        $IsTransitive = $Trust.DisallowTransivity -eq $false

        $html += @"
        <tr>
            <td>$DomainName</td>
            <td>$($Trust.Name)</td>
            <td>$($Trust.Direction)</td>
            <td>$($Trust.TrustType)</td>
            <td>$IsTransitive</td>
        </tr>
"@
    }
}
else {
    $html += @"
        <tr>
            <td colspan="5">No trust relationships found or unable to read trust information.</td>
        </tr>
"@
}

$html += @"
    </table>

    <div class="section-title">DNS Zones</div>

    <table>
        <tr>
            <th>Zone Name</th>
            <th>Zone Type</th>
            <th>Is AD Integrated</th>
            <th>Replication Scope</th>
            <th>
                Dynamic Update
                <button class="section-export-btn-inline" title="Export DNS Zones" onclick="downloadCsv('DNS_Zones.csv', dnsZonesCsv)">$([char]0x21E9)</button>
            </th>
        </tr>
"@

if ($DnsZones) {
    foreach ($Zone in $DnsZones) {
        $html += @"
        <tr>
            <td>$($Zone.ZoneName)</td>
            <td>$($Zone.ZoneType)</td>
            <td>$($Zone.IsDsIntegrated)</td>
            <td>$($Zone.ReplicationScope)</td>
            <td>$($Zone.DynamicUpdate)</td>
        </tr>
"@
    }
}
else {
    $html += @"
        <tr>
            <td colspan="5">DNS Server module unavailable or unable to read DNS zones.</td>
        </tr>
"@
}

$html += @"
    </table>

    </div>

    <div id="tab-dchealth" class="tab-content">

        <div class="dc-health-grid">

            <div class="dc-sidebar">
                <div class="dc-search-wrap">
                    <span class="dc-search-icon"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="7"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg></span>
                    <input type="text" id="dcSearchInput" class="dc-search-input" placeholder="Search domain controller..." onkeyup="filterDCList()" autocomplete="off">
                </div>
                <div class="dc-count" id="dcCount"></div>
                <div class="dc-list" id="dcList"></div>
            </div>

            <div class="dc-detail" id="dcDetail">
                <div class="dc-empty">Select a domain controller from the list to view its health details.</div>
            </div>

        </div>

    </div>

    <div id="tab-secposture" class="tab-content">

        <div class="dc-health-grid">

            <div class="dc-sidebar">
                <div class="dc-search-wrap">
                    <span class="dc-search-icon"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="7"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg></span>
                    <input type="text" id="secSearchInput" class="dc-search-input" placeholder="Search domain controller..." onkeyup="filterSecList()" autocomplete="off">
                </div>
                <div class="dc-count" id="secCount"></div>
                <div class="dc-list" id="secList"></div>
            </div>

            <div class="dc-detail" id="secDetail">
                <div class="dc-empty">Select a domain controller from the list to view its security posture.</div>
            </div>

        </div>

    </div>

    <div id="tab-deepinsights" class="tab-content">

        <div class="dc-health-grid">

            <div class="dc-sidebar">
                <div class="dc-search-wrap">
                    <span class="dc-search-icon"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="7"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg></span>
                    <input type="text" id="diSearchInput" class="dc-search-input" placeholder="Search domain controller..." onkeyup="filterDIList()" autocomplete="off">
                </div>
                <div class="dc-count" id="diCount"></div>
                <div class="dc-list" id="diList"></div>
            </div>

            <div class="dc-detail" id="diDetail">
                <div class="dc-empty">Select a domain controller from the list to view its deep insights.</div>
            </div>

        </div>

    </div>

    <div class="footer">
        <div class="footer-line">Enterprise Active Directory Health Dashboard | Generated using PowerShell</div>
        <div class="footer-line footer-copyright">&copy; $(Get-Date -Format yyyy) Azgar Mohammad. All rights reserved. | For internal IT operations use only.</div>
    </div>

</div>

<script id="dc-list-data" type="application/json">$DCListJson</script>
<script id="dc-inventory-data" type="application/json">$DCInventoryJson</script>

<script>
function downloadCsv(fileName, csvContent) {
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);

    link.setAttribute("href", url);
    link.setAttribute("download", fileName);
    link.style.visibility = "hidden";

    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    URL.revokeObjectURL(url);
}

const domainControllersCsv = '$DomainControllersCsvJs';
const dcPingCsv = '$DCPingCsvJs';
const adSitesCsv = '$ADSitesCsvJs';
const sitesWithNoDCsCsv = '$SitesWithNoDCsCsvJs';
const adSubnetsCsv = '$ADSubnetsCsvJs';
const dnsZonesCsv = '$DnsZonesCsvJs';
const gpoCsv = '$GpoCsvJs';
const unlinkedGpoCsv = '$UnlinkedGpoCsvJs';
const domainAdminsCsv = '$DomainAdminsCsvJs';
const enterpriseAdminsCsv = '$EnterpriseAdminsCsvJs';
const schemaAdminsCsv = '$SchemaAdminsCsvJs';

// =====================================================================
// Tab navigation
// =====================================================================

function showTab(tabName) {
    document.querySelectorAll('.tab-content').forEach(function (el) {
        el.classList.remove('active');
    });
    document.querySelectorAll('.tab-button').forEach(function (el) {
        el.classList.remove('active');
    });

    document.getElementById('tab-' + tabName).classList.add('active');
    document.getElementById('tab-btn-' + tabName).classList.add('active');
}

// =====================================================================
// Domain Controller Health tab
// =====================================================================

const dcListData = JSON.parse(document.getElementById('dc-list-data').textContent || '[]');
const dcInventoryData = JSON.parse(document.getElementById('dc-inventory-data').textContent || '[]');

const dcInventoryByName = {};
dcInventoryData.forEach(function (d) { dcInventoryByName[d.ComputerName] = d; });

let selectedDC = null;

function escapeHtml(value) {
    if (value === null || value === undefined) { return ''; }

    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function kvRow(label, value) {
    const display = (value === null || value === undefined || value === '') ? 'N/A' : value;
    return '<div class="dc-kv-row"><span>' + escapeHtml(label) + '</span><span>' + escapeHtml(display) + '</span></div>';
}

function gaugeClass(status) {
    return 'gauge-' + status;
}

function cpuStatus(pct) {
    if (pct >= 85) { return 'critical'; }
    if (pct >= 70) { return 'warning'; }
    return 'good';
}

function memStatus(pct) {
    if (pct >= 90) { return 'critical'; }
    if (pct >= 80) { return 'warning'; }
    return 'good';
}

function diskStatus(pct) {
    if (pct >= 90) { return 'critical'; }
    if (pct >= 80) { return 'warning'; }
    return 'good';
}

function dcPill(status) {
    const labels = { good: 'Good', warning: 'Warning', critical: 'Critical', unknown: 'Unknown' };
    return '<span class="dc-pill dc-pill-' + status + '">' + (labels[status] || 'Unknown') + '</span>';
}

function parsePhaseOffsetSeconds(text) {
    if (!text) { return null; }
    const m = String(text).match(/-?\d+(\.\d+)?/);
    if (!m) { return null; }
    return parseFloat(m[0]);
}

const TILE_ICONS = {
    server: '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="6" rx="1.5"/><rect x="3" y="14" width="18" height="6" rx="1.5"/><circle cx="7" cy="7" r="0.6" fill="currentColor" stroke="none"/><circle cx="7" cy="17" r="0.6" fill="currentColor" stroke="none"/></svg>',
    monitor: '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="12" rx="1.5"/><line x1="8" y1="20" x2="16" y2="20"/><line x1="12" y1="16" x2="12" y2="20"/></svg>',
    cpu: '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="6" y="6" width="12" height="12" rx="1.5"/><line x1="9" y1="1" x2="9" y2="6"/><line x1="15" y1="1" x2="15" y2="6"/><line x1="9" y1="18" x2="9" y2="23"/><line x1="15" y1="18" x2="15" y2="23"/><line x1="1" y1="9" x2="6" y2="9"/><line x1="1" y1="15" x2="6" y2="15"/><line x1="18" y1="9" x2="23" y2="9"/><line x1="18" y1="15" x2="23" y2="15"/></svg>',
    memory: '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><ellipse cx="12" cy="5" rx="8" ry="3"/><path d="M4 5v6c0 1.66 3.58 3 8 3s8-1.34 8-3V5"/><path d="M4 11v6c0 1.66 3.58 3 8 3s8-1.34 8-3v-6"/></svg>',
    disk: '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="2"/></svg>',
    network: '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><line x1="3" y1="12" x2="21" y2="12"/><path d="M12 3c2.5 2.5 4 6 4 9s-1.5 6.5-4 9c-2.5-2.5-4-6-4-9s1.5-6.5 4-9z"/></svg>',
    services: '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 11l3 3 8-8"/><path d="M21 12v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h11"/></svg>',
    clock: '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><polyline points="12 7 12 12 15.5 14"/></svg>',
    lock: '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="5" y="11" width="14" height="9" rx="1.5"/><path d="M8 11V7a4 4 0 0 1 8 0v4"/></svg>',
    refresh: '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12a9 9 0 1 1-2.64-6.36"/><polyline points="21 3 21 9 15 9"/></svg>',
    shield: '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3l8 4v5c0 5-3.5 8-8 9-4.5-1-8-4-8-9V7l8-4z"/></svg>',
    package: '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3l8 4.5v9L12 21l-8-4.5v-9L12 3z"/><path d="M12 12l8-4.5"/><path d="M12 12v9"/><path d="M12 12L4 7.5"/></svg>',
    checklist: '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 6h11"/><path d="M9 12h11"/><path d="M9 18h11"/><path d="M4.5 6l.75.75L6.5 5.5"/><path d="M4.5 12l.75.75L6.5 11.5"/><path d="M4.5 18l.75.75L6.5 17.5"/></svg>',
    database: '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><ellipse cx="12" cy="5" rx="8" ry="3"/><path d="M4 5v14c0 1.66 3.58 3 8 3s8-1.34 8-3V5"/><path d="M4 12c0 1.66 3.58 3 8 3s8-1.34 8-3"/></svg>',
    globe: '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M3 12h18"/><path d="M12 3a14 14 0 0 1 3.5 9 14 14 0 0 1-3.5 9 14 14 0 0 1-3.5-9 14 14 0 0 1 3.5-9z"/></svg>',
    alert: '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 9v4"/><path d="M12 17h.01"/><path d="M10.3 4.3 2.6 18a1.8 1.8 0 0 0 1.6 2.7h15.6a1.8 1.8 0 0 0 1.6-2.7L13.7 4.3a1.8 1.8 0 0 0-3.4 0z"/></svg>',
    login: '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4"/><polyline points="10 17 15 12 10 7"/><line x1="15" y1="12" x2="3" y2="12"/></svg>',
    broadcast: '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="10" x2="12" y2="21"/><path d="M8 14a4 4 0 0 1 8 0"/><path d="M5 11a8 8 0 0 1 14 0"/><circle cx="12" cy="6" r="1.5" fill="currentColor" stroke="none"/></svg>',
    plug: '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 3v4"/><path d="M15 3v4"/><rect x="6" y="7" width="12" height="7" rx="1.5"/><path d="M12 14v3"/><path d="M9 21a3 3 0 0 1 3-3 3 3 0 0 1 3 3"/></svg>'
};

function tileTitle(iconName, titleText, rightHtml) {
    const icon = TILE_ICONS[iconName] || '';
    return '<div class="dc-tile-title"><span class="dc-tile-title-left"><span class="dc-tile-icon">' + icon + '</span>' + escapeHtml(titleText) + '</span>' + (rightHtml || '') + '</div>';
}

function tileTip(html) {
    return '<div class="dc-tile-tip">' + html + '</div>';
}

function healthClass(status) {
    const s = (status || '').toLowerCase();
    if (s === 'ok' || s === 'good' || s === 'running') { return 'health-good'; }
    if (s === 'warning') { return 'health-warning'; }
    if (s === 'critical') { return 'health-critical'; }
    return 'health-unknown';
}

function dcStatusDot(dc) {
    const inv = dcInventoryByName[dc.Name];

    if (!inv || !inv.Reachable) { return 'dot-unknown'; }

    const rank = { good: 0, warning: 1, critical: 2 };
    let worst = 'good';

    function consider(status) {
        const s = (status || '').toLowerCase();
        if (rank[s] !== undefined && rank[s] > rank[worst]) { worst = s; }
    }

    consider(inv.LDAPS && inv.LDAPS.Status);
    consider(inv.TimeSync && inv.TimeSync.Status);

    if (inv.System && inv.System.PendingReboot) { consider('warning'); }

    (inv.Services || []).forEach(function (svc) {
        if (svc.Status !== 'Running' && svc.Status !== 'Not Found') { consider('critical'); }
    });

    return 'dot-' + worst;
}

function renderDCList(filterText) {
    const listEl = document.getElementById('dcList');
    const countEl = document.getElementById('dcCount');
    const filter = (filterText || '').trim().toLowerCase();

    const filtered = dcListData.filter(function (dc) {
        return !filter ||
            dc.Name.toLowerCase().indexOf(filter) !== -1 ||
            dc.Site.toLowerCase().indexOf(filter) !== -1;
    });

    countEl.textContent = filtered.length + ' of ' + dcListData.length + ' domain controllers';

    if (filtered.length === 0) {
        listEl.innerHTML = '<div class="dc-empty">No domain controllers match your search.</div>';
        return;
    }

    let html = '';

    filtered
        .slice()
        .sort(function (a, b) { return a.Name.localeCompare(b.Name); })
        .forEach(function (dc) {
            const activeClass = (dc.Name === selectedDC) ? ' active' : '';
            const dot = dcStatusDot(dc);

            html += '<div class="dc-item' + activeClass + '" onclick="selectDC(\'' + dc.Name.replace(/'/g, "\\'") + '\')">' +
                '<span class="dc-item-name" title="' + escapeHtml(dc.Name) + '">' + escapeHtml(dc.Name) + '</span>' +
                '<span class="dc-item-dot ' + dot + '"></span>' +
                '</div>';
        });

    listEl.innerHTML = html;
}

function filterDCList() {
    renderDCList(document.getElementById('dcSearchInput').value);
}

function selectDC(name) {
    selectedDC = name;
    renderDCList(document.getElementById('dcSearchInput').value);
    renderDCDetail(name);
}

function renderDCDetail(name) {
    const detailEl = document.getElementById('dcDetail');
    const meta = dcListData.find(function (d) { return d.Name === name; });
    const inv = dcInventoryByName[name];

    if (!meta) {
        detailEl.innerHTML = '<div class="dc-empty">Domain controller not found.</div>';
        return;
    }

    let badges = '<span class="dc-badge">' + escapeHtml(meta.Site) + '</span>';

    if (meta.IsGlobalCatalog) {
        badges += '<span class="dc-badge">Global Catalog</span>';
    }

    (meta.FSMORoles || []).forEach(function (role) {
        badges += '<span class="dc-badge">' + escapeHtml(role) + '</span>';
    });

    let html = '<div class="dc-detail-header"><div class="dc-detail-title">' + escapeHtml(meta.Name) + '</div>' + badges + '</div>';

    if (!inv || !inv.Reachable) {
        const reason = (inv && inv.Error) ? escapeHtml(inv.Error) : 'No inventory data was collected for this domain controller (DC Health collection may be disabled or WinRM is unreachable).';
        html += '<div class="dc-unreachable">Unable to collect live health data.<br>' + reason + '</div>';
        detailEl.innerHTML = html;
        return;
    }

    const hw   = inv.Hardware || {};
    const sys  = inv.System || {};
    const perf = inv.Performance || {};
    const ts   = inv.TimeSync || {};
    const ldaps = inv.LDAPS || {};

    html += '<div class="dc-tile-grid">';

    // ---- Section: Server & Hardware Health ----
    html += '<div class="dc-section-title">Server &amp; Hardware Health</div>';

    // Hardware inventory
    html += '<div class="dc-tile">' + tileTitle('server', 'Hardware Inventory') +
        kvRow('Manufacturer', hw.Manufacturer) +
        kvRow('Machine Type', hw.MachineType) +
        kvRow('Processor', hw.ProcessorName) +
        kvRow('CPU Sockets', hw.NumberOfProcessors) +
        kvRow('Cores / Logical CPUs', (hw.NumberOfCores || '?') + ' / ' + (hw.NumberOfLogicalProcessors || '?')) +
        kvRow('Memory Capacity', (hw.TotalMemoryGB != null ? hw.TotalMemoryGB + ' GB' : 'N/A')) +
        '</div>';

    // System / uptime / pending reboot
    const rebootBadge = sys.PendingReboot
        ? '<span class="health health-warning">Reboot Pending</span>'
        : '<span class="health health-good">No Reboot Pending</span>';

    html += '<div class="dc-tile">' + tileTitle('monitor', 'System &amp; Uptime') +
        kvRow('Operating System', sys.Caption) +
        kvRow('OS Build', sys.BuildNumber) +
        kvRow('Last Boot Time', sys.LastBootTime) +
        kvRow('Uptime', sys.UptimeReadable) +
        '<div class="dc-kv-row"><span>Pending Reboot</span><span>' + rebootBadge + '</span></div>' +
        (sys.PendingReboot ? kvRow('Reason', sys.PendingRebootReasons) : '') +
        '</div>';

    // CPU utilization
    const cpuPct = (perf.CPUUtilizationPercent != null) ? perf.CPUUtilizationPercent : 0;
    const cpuSt = cpuStatus(cpuPct);
    let cpuTip = '';

    if (cpuSt === 'warning') {
        cpuTip = tileTip('CPU usage is above 70%. <b>Recommendation:</b> sustained high CPU can slow down logons, LDAP queries and AD replication on this DC. Check for runaway processes or plan a capacity review.');
    }
    else if (cpuSt === 'critical') {
        cpuTip = tileTip('CPU usage is above 85%. <b>Recommendation:</b> this DC may be too slow to process logons and replication reliably. Investigate immediately and consider redistributing load to another DC.');
    }

    html += '<div class="dc-tile" id="dcTileCpu">' + tileTitle('cpu', 'CPU Utilization', dcPill(cpuSt)) +
        '<div class="gauge-wrap"><div class="gauge-value">' + cpuPct + '%</div>' +
        '<div class="gauge-bar"><div class="gauge-fill ' + gaugeClass(cpuSt) + '" style="width:' + Math.min(cpuPct, 100) + '%"></div></div></div>' +
        '<div class="dc-kv-row" style="margin-top:8px;"><span>Logical Processors</span><span>' + (hw.NumberOfLogicalProcessors || 'N/A') + '</span></div>' +
        cpuTip +
        '</div>';

    // Memory utilization
    const memPct = (perf.MemoryUtilizationPercent != null) ? perf.MemoryUtilizationPercent : 0;
    const memSt = memStatus(memPct);
    let memTip = '';

    if (memSt === 'warning') {
        memTip = tileTip('Memory usage is above 80%. <b>Recommendation:</b> as memory pressure grows, AD database caching shrinks and logon and query response times can increase. Keep an eye on this trend.');
    }
    else if (memSt === 'critical') {
        memTip = tileTip('Memory usage is above 90%. <b>Recommendation:</b> the DC may start paging heavily, which can seriously delay authentication and replication. Free up memory or add more RAM soon.');
    }

    html += '<div class="dc-tile" id="dcTileMemory">' + tileTitle('memory', 'Memory Utilization', dcPill(memSt)) +
        '<div class="gauge-wrap"><div class="gauge-value">' + memPct + '%</div>' +
        '<div class="gauge-bar"><div class="gauge-fill ' + gaugeClass(memSt) + '" style="width:' + Math.min(memPct, 100) + '%"></div></div></div>' +
        '<div class="dc-kv-row" style="margin-top:8px;"><span>Memory Used / Total</span><span>' + (perf.MemoryUsedGB != null ? perf.MemoryUsedGB + ' / ' + perf.MemoryTotalGB + ' GB' : 'N/A') + '</span></div>' +
        memTip +
        '</div>';

    // Disk capacity - dynamically discovered, no hardcoded drive letters
    html += '<div class="dc-tile dc-tile-wide">' + tileTitle('disk', 'Disk Capacity');

    let diskTipLines = '';

    if (inv.Disks && inv.Disks.length > 0) {
        html += '<table class="dc-mini-table"><tr><th>Drive</th><th>Description</th><th>Capacity</th><th>Used Space</th><th>Free Space</th><th>Free %</th></tr>';

        inv.Disks.forEach(function (d) {
            const usedGB = (d.SizeGB - d.FreeGB);
            const freePercent = (100 - d.UsedPercent);

            html += '<tr>' +
                '<td><b>' + escapeHtml(d.Drive) + '</b></td>' +
                '<td>' + escapeHtml(d.Role) + (d.Label ? ' (' + escapeHtml(d.Label) + ')' : '') + '</td>' +
                '<td>' + d.SizeGB + ' GB</td>' +
                '<td>' + usedGB.toFixed(2) + ' GB (' + d.UsedPercent + '%)</td>' +
                '<td>' + d.FreeGB + ' GB</td>' +
                '<td>' + freePercent.toFixed(1) + '%</td>' +
                '</tr>';

            const dSt = diskStatus(d.UsedPercent);

            if (dSt !== 'good') {
                diskTipLines += '<div>' + (dSt === 'critical' ? '<b>Critical:</b> ' : '<b>Warning:</b> ') +
                    'Drive ' + escapeHtml(d.Drive) + ' (' + escapeHtml(d.Role) + ') is ' + d.UsedPercent + '% used.</div>';
            }
        });

        html += '</table>';
    }
    else {
        html += kvRow('No disk data collected', '-');
    }

    if (diskTipLines) {
        html += tileTip(diskTipLines + '<div style="margin-top:6px;"><b>Recommendation:</b> if the volume hosting the AD database, SYSVOL, or logs fills up, AD can stop accepting writes and replication may fail. Free up space or extend the affected volume(s) soon.</div>');
    }

    html += '</div>';

    // Network adapters (includes IPv4/IPv6 status and DNS settings for the same interface)
    const eth = inv.EthernetSettings || {};

    html += '<div class="dc-tile dc-tile-wide">' + tileTitle('network', 'Network Adapters');

    if (inv.Network && inv.Network.length > 0) {
        html += '<table class="dc-mini-table"><tr><th>Adapter</th><th>IPv4</th><th>IPv6</th><th>Preferred DNS</th><th>Alternate DNS</th></tr>';

        inv.Network.forEach(function (n) {
            const dnsList = (n.DNSServers || '').split(',').map(function (s) { return s.trim(); }).filter(function (s) { return s; });
            const preferredDns = dnsList[0] || '-';
            const alternateDns = dnsList[1] || '-';

            const ipv4Display = n.IPAddress ? ('Enabled (' + n.IPAddress + ')') : 'Unknown';

            let ipv6Display = '-';

            if (eth.AdapterName && n.AdapterName === eth.AdapterName) {
                ipv6Display = (eth.IPv6Enabled === true) ? ('Enabled' + (eth.IPv6Address ? ' (' + eth.IPv6Address + ')' : ''))
                    : (eth.IPv6Enabled === false) ? 'Disabled' : 'Unknown';
            }

            html += '<tr><td>' + escapeHtml(n.AdapterName) + '</td><td>' + escapeHtml(ipv4Display) + '</td><td>' + escapeHtml(ipv6Display) + '</td><td>' + escapeHtml(preferredDns) + '</td><td>' + escapeHtml(alternateDns) + '</td></tr>';
        });

        html += '</table>';
    }
    else {
        html += kvRow('No network data collected', '-');
    }

    html += '</div>';

    // ---- Section: Active Directory Service Health ----
    html += '<div class="dc-section-title dc-section-title-spaced">Active Directory Service Health</div>';

    // AD service health
    html += '<div class="dc-tile dc-tile-wide">' + tileTitle('services', 'AD Service Health') + '<div class="dc-service-grid">';

    (inv.Services || []).forEach(function (svc) {
        const cls = svc.Status === 'Running' ? 'health-good' : (svc.Status === 'Not Found' ? 'health-unknown' : 'health-critical');
        html += '<div class="dc-service-row"><span>' + escapeHtml(svc.DisplayName) + '</span><span class="health ' + cls + '">' + escapeHtml(svc.Status) + '</span></div>';
    });

    html += '</div></div>';

    // Time synchronization
    const offsetSec = parsePhaseOffsetSeconds(ts.PhaseOffset);
    let tsTip = '';

    if (offsetSec !== null && Math.abs(offsetSec) > 300) {
        tsTip = tileTip('This DC\'s clock is off by more than 5 minutes, which exceeds the default Kerberos clock skew tolerance. <b>Recommendation:</b> logons, authentication and replication can start failing until the time sync issue is fixed.');
    }

    html += '<div class="dc-tile">' + tileTitle('clock', 'Time Synchronization', '<span class="health ' + healthClass(ts.Status) + '">' + escapeHtml(ts.Status || 'Unknown') + '</span>') +
        kvRow('Source', ts.Source) +
        kvRow('Stratum', ts.Stratum) +
        kvRow('Last Sync', ts.LastSyncTime) +
        kvRow('Phase Offset', ts.PhaseOffset) +
        tsTip +
        '</div>';

    // LDAPS certificate
    html += '<div class="dc-tile">' + tileTitle('lock', 'LDAPS Certificate', '<span class="health ' + healthClass(ldaps.Status) + '">' + escapeHtml(ldaps.Status || 'Unknown') + '</span>') +
        kvRow('Port 636', ldaps.PortOpen ? 'Open' : 'Closed') +
        kvRow('Subject', ldaps.CertSubject) +
        kvRow('Expires', ldaps.CertExpiry) +
        kvRow('Days to Expiry', ldaps.DaysToExpiry) +
        '</div>';

    // Replication partners
    html += '<div class="dc-tile dc-tile-wide">' + tileTitle('refresh', 'Replication Partners');

    if (inv.Replication && inv.Replication.length > 0) {
        html += '<table class="dc-mini-table"><tr><th>Partner</th><th>Partition</th><th>Last Success</th><th>Failures</th><th>Last Result</th></tr>';

        inv.Replication.forEach(function (r) {
            const failCls = (r.ConsecutiveFailures && r.ConsecutiveFailures !== '0') ? 'health-warning' : 'health-good';

            html += '<tr><td>' + escapeHtml(r.Partner) + '</td><td>' + escapeHtml(r.Partition) + '</td><td>' + escapeHtml(r.LastSuccess) + '</td>' +
                '<td><span class="health ' + failCls + '">' + escapeHtml(r.ConsecutiveFailures) + '</span></td><td>' + escapeHtml(r.LastResult) + '</td></tr>';
        });

        html += '</table>';
    }
    else {
        html += kvRow('No replication partner data collected', '-');
    }

    html += '</div>';

    html += '</div>'; // close dc-tile-grid

    detailEl.innerHTML = html;

    adjustSidebarHeight();
}

function adjustSidebarHeight() {
    const listEl = document.getElementById('dcList');
    const cpuTile = document.getElementById('dcTileCpu');

    if (!listEl || !cpuTile) { return; }

    const cpuRect = cpuTile.getBoundingClientRect();
    const listRect = listEl.getBoundingClientRect();
    const targetHeight = cpuRect.bottom - listRect.top;

    if (targetHeight > 100) {
        listEl.style.maxHeight = targetHeight + 'px';
    }
}

window.addEventListener('resize', function () {
    if (selectedDC) { adjustSidebarHeight(); }
    if (selectedDIDC) { adjustDIHeight(); }
});

function adjustDIHeight() {
    const listEl = document.getElementById('diList');
    const backupTile = document.getElementById('diTileAdBackup');

    if (!listEl || !backupTile) { return; }

    const tileRect = backupTile.getBoundingClientRect();
    const listRect = listEl.getBoundingClientRect();
    const targetHeight = tileRect.bottom - listRect.top;

    if (targetHeight > 100) {
        listEl.style.maxHeight = targetHeight + 'px';
    }
}

renderDCList('');

if (dcListData.length > 0) {
    selectDC(dcListData[0].Name);
}

// =====================================================================
// Security Posture tab
// =====================================================================

let selectedSecDC = null;

function boolLabel(v) {
    if (v === null || v === undefined) { return 'N/A'; }
    return v ? 'Yes' : 'No';
}

function smbStatus(smb) {
    if (!smb || smb.SMB1Enabled === null || smb.SMB1Enabled === undefined) { return 'unknown'; }
    if (smb.SMB1Enabled === true) { return 'critical'; }
    if (smb.SigningRequired !== true) { return 'warning'; }
    return 'good';
}

function tlsStatus(tlsArr) {
    if (!tlsArr || tlsArr.length === 0) { return 'unknown'; }

    function isEnabled(p) {
        if (p.Enabled === 0) { return false; }
        if (p.DisabledByDefault === 1 && p.Enabled !== 1) { return false; }
        return true;
    }

    const byProto = {};
    tlsArr.forEach(function (p) { byProto[p.Protocol] = p; });

    const riskyEnabled = ['SSL 2.0', 'SSL 3.0', 'TLS 1.0'].some(function (name) {
        return byProto[name] && isEnabled(byProto[name]);
    });

    if (riskyEnabled) { return 'critical'; }

    if (byProto['TLS 1.1'] && isEnabled(byProto['TLS 1.1'])) { return 'warning'; }

    return 'good';
}

function certPostureStatus(ldaps) {
    if (!ldaps || !ldaps.PortOpen) { return 'unknown'; }

    const days = ldaps.DaysToExpiry;

    if (days === null || days === undefined) { return 'unknown'; }
    if (days < 30) { return 'critical'; }
    if (days <= 60 || ldaps.SelfSigned === true || (ldaps.KeySizeBits && ldaps.KeySizeBits < 2048)) { return 'warning'; }
    return 'good';
}

function ldapSigningStatus(sig) {
    if (!sig || sig.LDAPServerIntegrity === null || sig.LDAPServerIntegrity === undefined) { return 'unknown'; }

    const integrity = sig.LDAPServerIntegrity;
    const channel = sig.LdapEnforceChannelBinding;

    if (integrity !== 2 || channel === 0 || channel === null || channel === undefined) { return 'critical'; }
    if (channel === 1) { return 'warning'; }
    return 'good';
}

function ntlmStatus(ntlm) {
    if (!ntlm || ntlm.LmCompatibilityLevel === null || ntlm.LmCompatibilityLevel === undefined) { return 'unknown'; }

    const level = ntlm.LmCompatibilityLevel;

    if (level <= 2) { return 'critical'; }

    const auditingEnabled = (ntlm.AuditReceivingNTLM && ntlm.AuditReceivingNTLM !== 0) ||
        (ntlm.AuditNTLMInDomain && ntlm.AuditNTLMInDomain !== 0);

    if (level === 5 && auditingEnabled) { return 'good'; }

    return 'warning';
}

function auditPolicyStatus(ap) {
    if (!ap) { return 'unknown'; }

    const vals = [ap.DirectoryServiceChanges, ap.CredentialValidation, ap.AuthorizationPolicyChange];

    if (vals.every(function (v) { return !v || v === 'Unknown'; })) { return 'unknown'; }

    if (ap.DirectoryServiceChanges === 'No Auditing' || ap.AuthorizationPolicyChange === 'No Auditing') {
        return 'critical';
    }

    const allCovered = vals.every(function (v) {
        return v && v.indexOf('Success') !== -1;
    });

    return allCovered ? 'good' : 'warning';
}

function netbiosStatus(netbios) {
    const adapters = (netbios && netbios.Adapters) || [];

    if (adapters.length === 0) { return 'unknown'; }
    if (adapters.some(function (a) { return a.Setting === 'Enabled'; })) { return 'critical'; }
    if (adapters.every(function (a) { return a.Setting === 'Disabled'; })) { return 'good'; }

    return 'warning';
}

function secTip(risk, recommendation, linkUrl, linkText) {
    return tileTip('<div><b>Risk:</b> ' + risk + '</div>' +
        '<div style="margin-top:4px;"><b>Recommendation:</b> ' + recommendation + '</div>' +
        '<div style="margin-top:4px;"><a href="' + linkUrl + '" target="_blank" rel="noopener noreferrer">' + (linkText || 'Microsoft documentation') + '</a></div>');
}

function secStatusDot(dc) {
    const inv = dcInventoryByName[dc.Name];

    if (!inv || !inv.Reachable || !inv.SecurityPosture) { return 'dot-unknown'; }

    const sp = inv.SecurityPosture;
    const rank = { good: 0, warning: 1, critical: 2 };
    let worst = 'good';

    function consider(status) {
        if (status === 'unknown') { return; }
        if (rank[status] !== undefined && rank[status] > rank[worst]) { worst = status; }
    }

    consider(smbStatus(sp.SMB));
    consider(tlsStatus(sp.TLS));
    consider(certPostureStatus(inv.LDAPS));
    consider(ldapSigningStatus(sp.LDAPSigning));
    consider(ntlmStatus(sp.NTLM));
    consider(auditPolicyStatus(sp.AuditPolicy));
    consider(netbiosStatus(sp.NetBIOS));

    return 'dot-' + worst;
}

function renderSecList(filterText) {
    const listEl = document.getElementById('secList');
    const countEl = document.getElementById('secCount');
    const filter = (filterText || '').trim().toLowerCase();

    const filtered = dcListData.filter(function (dc) {
        return !filter ||
            dc.Name.toLowerCase().indexOf(filter) !== -1 ||
            dc.Site.toLowerCase().indexOf(filter) !== -1;
    });

    countEl.textContent = filtered.length + ' of ' + dcListData.length + ' domain controllers';

    if (filtered.length === 0) {
        listEl.innerHTML = '<div class="dc-empty">No domain controllers match your search.</div>';
        return;
    }

    let html = '';

    filtered
        .slice()
        .sort(function (a, b) { return a.Name.localeCompare(b.Name); })
        .forEach(function (dc) {
            const activeClass = (dc.Name === selectedSecDC) ? ' active' : '';
            const dot = secStatusDot(dc);

            html += '<div class="dc-item' + activeClass + '" onclick="selectSecDC(\'' + dc.Name.replace(/'/g, "\\'") + '\')">' +
                '<span class="dc-item-name" title="' + escapeHtml(dc.Name) + '">' + escapeHtml(dc.Name) + '</span>' +
                '<span class="dc-item-dot ' + dot + '"></span>' +
                '</div>';
        });

    listEl.innerHTML = html;
}

function filterSecList() {
    renderSecList(document.getElementById('secSearchInput').value);
}

function selectSecDC(name) {
    selectedSecDC = name;
    renderSecList(document.getElementById('secSearchInput').value);
    renderSecDetail(name);
}

function renderSecDetail(name) {
    const detailEl = document.getElementById('secDetail');
    const meta = dcListData.find(function (d) { return d.Name === name; });
    const inv = dcInventoryByName[name];

    if (!meta) {
        detailEl.innerHTML = '<div class="dc-empty">Domain controller not found.</div>';
        return;
    }

    let badges = '<span class="dc-badge">' + escapeHtml(meta.Site) + '</span>';

    if (meta.IsGlobalCatalog) {
        badges += '<span class="dc-badge">Global Catalog</span>';
    }

    (meta.FSMORoles || []).forEach(function (role) {
        badges += '<span class="dc-badge">' + escapeHtml(role) + '</span>';
    });

    let html = '<div class="dc-detail-header"><div class="dc-detail-title">' + escapeHtml(meta.Name) + '</div>' + badges + '</div>';

    if (!inv || !inv.Reachable) {
        const reason = (inv && inv.Error) ? escapeHtml(inv.Error) : 'No inventory data was collected for this domain controller (collection may be disabled or WinRM is unreachable).';
        html += '<div class="dc-unreachable">Unable to collect live security posture data.<br>' + reason + '</div>';
        detailEl.innerHTML = html;
        return;
    }

    const sp = inv.SecurityPosture || {};
    const ldaps = inv.LDAPS || {};
    const sys = inv.System || {};

    html += '<div class="dc-tile-grid">';

    // 1. SMB protocol security
    const smb = sp.SMB || {};
    const smbSt = smbStatus(smb);
    let smbTip = '';

    if (smbSt === 'critical') {
        smbTip = secTip(
            'SMBv1 is enabled. SMBv1 has no protection against relay or tampering attacks and is a common ransomware entry point (e.g., EternalBlue/WannaCry-class exploits).',
            'Disable SMBv1 on this domain controller and require SMB signing.',
            'https://techcommunity.microsoft.com/blog/coreinfrastructureandsecurityblog/active-directory-hardening-series---part-6-%E2%80%93-enforcing-smb-signing/4272168',
            'AD Hardening Series &ndash; Part 6: Enforcing SMB Signing'
        );
    }
    else if (smbSt === 'warning') {
        smbTip = secTip(
            'SMB signing is not required. Without it, an attacker positioned on the network can tamper with or relay SMB traffic to/from this DC (NTLM relay attacks).',
            'Enable "Require SMB signing" via Group Policy or Set-SmbServerConfiguration -RequireSecuritySignature $true.',
            'https://techcommunity.microsoft.com/blog/coreinfrastructureandsecurityblog/active-directory-hardening-series---part-6-%E2%80%93-enforcing-smb-signing/4272168',
            'AD Hardening Series &ndash; Part 6: Enforcing SMB Signing'
        );
    }

    html += '<div class="dc-tile">' + tileTitle('lock', 'SMB Protocol Security', dcPill(smbSt)) +
        kvRow('SMBv1 Enabled', boolLabel(smb.SMB1Enabled)) +
        kvRow('Signing Enabled', boolLabel(smb.SigningEnabled)) +
        kvRow('Signing Required', boolLabel(smb.SigningRequired)) +
        smbTip +
        '</div>';

    // 2. TLS / Schannel
    const tlsArr = sp.TLS || [];
    const tlsSt = tlsStatus(tlsArr);
    let tlsTip = '';

    if (tlsSt === 'critical') {
        tlsTip = secTip(
            'A legacy protocol (SSL 2.0/3.0 or TLS 1.0) is enabled. These protocols have known cryptographic weaknesses and are no longer considered secure for LDAPS and other TLS-protected traffic.',
            'Disable SSL 2.0, SSL 3.0 and TLS 1.0 in Schannel and ensure only TLS 1.2 (and 1.3 where supported) remain enabled.',
            'https://techcommunity.microsoft.com/blog/microsoft-security-baselines/security-baseline-for-windows-server-2025-version-2602/4496468',
            'Windows Server 2025 Security Baseline'
        );
    }
    else if (tlsSt === 'warning') {
        tlsTip = secTip(
            'TLS 1.1 is enabled. TLS 1.1 lacks modern cipher suite support and is being phased out across the industry.',
            'Disable TLS 1.1 via Schannel registry settings, leaving TLS 1.2/1.3 as the only enabled protocols.',
            'https://techcommunity.microsoft.com/blog/microsoft-security-baselines/security-baseline-for-windows-server-2025-version-2602/4496468',
            'Windows Server 2025 Security Baseline'
        );
    }

    html += '<div class="dc-tile">' + tileTitle('shield', 'TLS / Schannel', dcPill(tlsSt));

    if (tlsArr.length > 0) {
        tlsArr.forEach(function (p) {
            const enabledLabel = (p.Enabled === null || p.Enabled === undefined) ? 'Default' : (p.Enabled === 1 ? 'Enabled' : 'Disabled');
            html += kvRow(p.Protocol, enabledLabel);
        });
    }
    else {
        html += kvRow('No Schannel data collected', '-');
    }

    html += tlsTip + '</div>';

    // 3. LDAPS certificate
    const certSt = certPostureStatus(ldaps);
    let certTip = '';

    if (certSt === 'critical') {
        certTip = secTip(
            'The LDAPS certificate is expired or expires in under 30 days, or port 636 is not responding. Once expired, LDAPS-dependent clients (and tools enforcing LDAP channel binding) will fail to bind.',
            'Renew/replace the LDAPS certificate now and confirm port 636 is listening with the new certificate.',
            'https://learn.microsoft.com/en-us/troubleshoot/windows-server/identity/enable-ldap-over-ssl-3rd-certification-authority',
            'Enable LDAP over SSL (LDAPS) &ndash; Microsoft Learn'
        );
    }
    else if (certSt === 'warning') {
        certTip = secTip(
            'The LDAPS certificate expires within 60 days, is self-signed, or uses a key shorter than 2048 bits, any of which weakens trust or risks an outage if not renewed in time.',
            'Issue a CA-signed certificate (>=2048-bit key) from an enterprise/trusted CA before the current certificate expires.',
            'https://learn.microsoft.com/en-us/troubleshoot/windows-server/identity/enable-ldap-over-ssl-3rd-certification-authority',
            'Enable LDAP over SSL (LDAPS) &ndash; Microsoft Learn'
        );
    }

    html += '<div class="dc-tile">' + tileTitle('lock', 'LDAPS Certificate', dcPill(certSt)) +
        kvRow('Port 636', ldaps.PortOpen ? 'Open' : 'Closed') +
        kvRow('Subject', ldaps.CertSubject) +
        kvRow('Issuer', ldaps.CertIssuer) +
        kvRow('Self-Signed', boolLabel(ldaps.SelfSigned)) +
        kvRow('Signature Algorithm', ldaps.SignatureAlgorithm) +
        kvRow('Key Size (bits)', ldaps.KeySizeBits) +
        kvRow('Valid From', ldaps.CertNotBefore) +
        kvRow('Expires', ldaps.CertExpiry) +
        kvRow('Days to Expiry', ldaps.DaysToExpiry) +
        certTip +
        '</div>';

    // 4. LDAP signing
    const sig = sp.LDAPSigning || {};
    const sigSt = ldapSigningStatus(sig);
    let sigTip = '';

    const integrityLabels = { 0: 'None', 1: 'Negotiate signing', 2: 'Require signing' };
    const channelLabels = { 0: 'Never', 1: 'When supported', 2: 'Always' };
    const integrityLabel = integrityLabels[sig.LDAPServerIntegrity];
    const channelLabel = channelLabels[sig.LdapEnforceChannelBinding];

    if (sigSt === 'critical') {
        sigTip = secTip(
            'LDAP signing is not required and/or channel binding is disabled. Unsigned LDAP traffic can be tampered with or relayed, enabling NTLM relay attacks against this domain controller.',
            'Set LDAP Server Integrity to "Require signing" (2) and LDAP channel binding to "Always" (2) via registry/GPO on all domain controllers.',
            'https://techcommunity.microsoft.com/blog/coreinfrastructureandsecurityblog/active-directory-hardening-series---part-3-%E2%80%93-enforcing-ldap-signing/4066233',
            'AD Hardening Series &ndash; Part 3: Enforcing LDAP Signing'
        );
    }
    else if (sigSt === 'warning') {
        sigTip = secTip(
            'LDAP channel binding is set to "When supported" rather than "Always", leaving a window where an unbound LDAP session could still be relayed by older clients.',
            'Move LDAP channel binding from "When supported" (1) to "Always" (2) once all clients support channel binding.',
            'https://techcommunity.microsoft.com/blog/coreinfrastructureandsecurityblog/active-directory-hardening-series---part-3-%E2%80%93-enforcing-ldap-signing/4066233',
            'AD Hardening Series &ndash; Part 3: Enforcing LDAP Signing'
        );
    }

    html += '<div class="dc-tile">' + tileTitle('shield', 'LDAP Signing', dcPill(sigSt)) +
        kvRow('LDAP Server Integrity', integrityLabel || 'N/A') +
        kvRow('Channel Binding', channelLabel || 'N/A') +
        sigTip +
        '</div>';

    // 5. NTLM hardening
    const ntlm = sp.NTLM || {};
    const ntlmSt = ntlmStatus(ntlm);
    let ntlmTip = '';

    const lmLabels = {
        0: 'Send LM & NTLM (no NTLMv2)', 1: 'Send LM & NTLM, use NTLMv2 if negotiated',
        2: 'Send NTLM only', 3: 'Send NTLMv2 only', 4: 'Send NTLMv2 only, refuse LM',
        5: 'Send NTLMv2 only, refuse LM & NTLM'
    };

    if (ntlmSt === 'critical') {
        ntlmTip = secTip(
            'This DC accepts LM or NTLMv1 authentication. These protocols use weak hashing and are vulnerable to offline cracking and pass-the-hash style attacks.',
            'Raise LM Compatibility Level to 5 ("Send NTLMv2 only, refuse LM & NTLM") after confirming no legacy clients depend on NTLMv1/LM.',
            'https://techcommunity.microsoft.com/blog/coreinfrastructureandsecurityblog/active-directory-hardening-series---part-8-%E2%80%93-disabling-ntlm/4485782',
            'AD Hardening Series &ndash; Part 8: Disabling NTLM'
        );
    }
    else if (ntlmSt === 'warning') {
        ntlmTip = secTip(
            'NTLMv2 is enforced but NTLM auditing is not fully enabled, or the compatibility level is not yet at its strictest setting, so residual NTLM usage cannot be measured before restricting it further.',
            'Enable NTLM auditing (Audit Receiving NTLM / Audit NTLM in Domain) to baseline remaining NTLM usage, then move toward LM Compatibility Level 5 and restricting NTLM.',
            'https://techcommunity.microsoft.com/blog/coreinfrastructureandsecurityblog/active-directory-hardening-series---part-8-%E2%80%93-disabling-ntlm/4485782',
            'AD Hardening Series &ndash; Part 8: Disabling NTLM'
        );
    }

    html += '<div class="dc-tile">' + tileTitle('package', 'NTLM Hardening', dcPill(ntlmSt)) +
        kvRow('LM Compatibility Level', (ntlm.LmCompatibilityLevel !== null && ntlm.LmCompatibilityLevel !== undefined) ? ntlm.LmCompatibilityLevel + ' - ' + (lmLabels[ntlm.LmCompatibilityLevel] || '') : 'N/A') +
        kvRow('Restrict Sending NTLM', ntlm.RestrictSendingNTLM) +
        kvRow('Audit Receiving NTLM', ntlm.AuditReceivingNTLM) +
        kvRow('Audit NTLM in Domain', ntlm.AuditNTLMInDomain) +
        ntlmTip +
        '</div>';

    // 6. Audit policy coverage
    const ap = sp.AuditPolicy || {};
    const apSt = auditPolicyStatus(ap);
    let apTip = '';

    if (apSt === 'critical') {
        apTip = secTip(
            'Directory Service Changes or Authorization Policy Change auditing is fully disabled. Without these, changes to AD objects, GPOs and permissions/privileges leave no audit trail, making it difficult to detect or investigate compromise or insider misuse.',
            'Enable success and failure auditing for "Directory Service Changes" and "Authorization Policy Change" via Advanced Audit Policy / GPO.',
            'https://techcommunity.microsoft.com/blog/microsoft-security-baselines/security-baseline-for-windows-server-2025-version-2602/4496468',
            'Windows Server 2025 Security Baseline'
        );
    }
    else if (apSt === 'warning') {
        apTip = secTip(
            'One or more recommended audit subcategories (Directory Service Changes, Credential Validation, Authorization Policy Change) is not fully configured, leaving gaps in the audit trail for authentication and AD object changes.',
            'Review Advanced Audit Policy Configuration and enable Success/Failure auditing for all three subcategories per the Microsoft security baseline.',
            'https://techcommunity.microsoft.com/blog/microsoft-security-baselines/security-baseline-for-windows-server-2025-version-2602/4496468',
            'Windows Server 2025 Security Baseline'
        );
    }

    html += '<div class="dc-tile">' + tileTitle('services', 'Audit Policy Coverage', dcPill(apSt)) +
        kvRow('Directory Service Changes', ap.DirectoryServiceChanges) +
        kvRow('Credential Validation', ap.CredentialValidation) +
        kvRow('Authorization Policy Change', ap.AuthorizationPolicyChange) +
        apTip +
        '</div>';

    // 7. NetBIOS over TCP/IP (NetBT)
    const netbios = sp.NetBIOS || {};
    const netbiosSt = netbiosStatus(netbios);
    let netbiosTip = '';

    if (netbiosSt === 'critical') {
        netbiosTip = secTip(
            'NetBIOS over TCP/IP is enabled on one or more network adapters. This allows legacy NBT-NS broadcast name resolution, which attackers can abuse for NBT-NS/LLMNR poisoning and credential relay attacks.',
            'Disable NetBIOS over TCP/IP on all network adapters (set "Disable NetBIOS over TCP/IP" in the adapter\'s WINS settings, or via DHCP scope option 46 / the NetbiosOptions registry value).',
            'https://learn.microsoft.com/en-us/troubleshoot/windows-server/networking/disable-netbios-tcp-ip-issues',
            'Disable NetBIOS over TCP/IP &ndash; Microsoft Learn'
        );
    }
    else if (netbiosSt === 'warning') {
        netbiosTip = secTip(
            'NetBIOS over TCP/IP is left at its default (DHCP-controlled) setting on one or more adapters, so it may still be active depending on DHCP scope options, leaving legacy NBT-NS name resolution exposed.',
            'Explicitly disable NetBIOS over TCP/IP on all network adapters rather than relying on the DHCP default.',
            'https://learn.microsoft.com/en-us/troubleshoot/windows-server/networking/disable-netbios-tcp-ip-issues',
            'Disable NetBIOS over TCP/IP &ndash; Microsoft Learn'
        );
    }

    html += '<div class="dc-tile">' + tileTitle('broadcast', 'NetBIOS over TCP/IP (NetBT)', dcPill(netbiosSt));

    if (netbios.Adapters && netbios.Adapters.length > 0) {
        netbios.Adapters.forEach(function (a) {
            const cls = (a.Setting === 'Disabled') ? 'health-good' : (a.Setting === 'Enabled') ? 'health-critical' : 'health-warning';
            html += '<div class="dc-kv-row"><span>' + escapeHtml(a.Adapter) + '</span><span class="health ' + cls + '">' + escapeHtml(a.Setting) + '</span></div>';
        });
    }
    else {
        html += kvRow('No NetBIOS adapter data collected', '-');
    }

    html += netbiosTip + '</div>';

    // 8. Recent hotfix installed updates (full-width, last)
    const hotfixes = sp.Hotfixes || [];
    const rebootPending = !!(sys && sys.PendingReboot);

    const hotfixRebootBadge = rebootPending
        ? '<span class="health health-warning">Reboot Pending</span>'
        : '<span class="health health-good">No Reboot Pending</span>';

    html += '<div class="dc-tile dc-tile-wide">' + tileTitle('refresh', 'Recent Hotfix Installed Updates',
        rebootPending ? '<span class="dc-pill dc-pill-critical">Reboot Pending</span>' : '');

    html += '<div class="dc-kv-row"><span>Pending Reboot</span><span>' + hotfixRebootBadge + '</span></div>';

    if (hotfixes.length > 0) {
        html += '<table class="dc-mini-table"><tr><th>Hotfix ID</th><th>Description</th><th>Installed On</th></tr>';

        hotfixes.forEach(function (h) {
            html += '<tr><td>' + escapeHtml(h.HotFixID) + '</td><td>' + escapeHtml(h.Description) + '</td><td>' + escapeHtml(h.InstalledOn) + '</td></tr>';
        });

        html += '</table>';
    }
    else {
        html += kvRow('No hotfix data collected', '-');
    }

    if (rebootPending) {
        html += '<div class="dc-kv-row" style="margin-top:8px;"><span>Pending Reboot Reason</span><span>' + escapeHtml(sys.PendingRebootReasons || 'Unknown') + '</span></div>' +
            '<div style="margin-top:8px; padding:10px 12px; border-radius:6px; background:#fffbeb; color:#92400e; font-size:12.5px;">' +
            'Restart <b>' + escapeHtml(meta.Name) + '</b> to complete installation of pending updates. A reboot is required before these changes take full effect.' +
            '</div>';
    }

    html += '</div>';

    html += '</div>'; // close dc-tile-grid

    detailEl.innerHTML = html;
}

renderSecList('');

if (dcListData.length > 0) {
    selectSecDC(dcListData[0].Name);
}

// =====================================================================
// Deep Insights tab
// =====================================================================

let selectedDIDC = null;

function dcDiagStatus(dd) {
    if (!dd) { return 'unknown'; }

    const vals = Object.keys(dd).map(function (k) { return dd[k]; });

    if (vals.every(function (v) { return v === 'Unknown'; })) { return 'unknown'; }
    if (vals.some(function (v) { return v === 'Failed'; })) { return 'critical'; }
    if (vals.some(function (v) { return v === 'Unknown'; })) { return 'warning'; }

    return 'good';
}

function dfsrStatus(arr, shares) {
    if (!arr || arr.length === 0) { return 'unknown'; }

    let result = 'good';

    if (arr.some(function (f) { return f.State === 'In Error'; })) { result = 'critical'; }
    else if (!arr.every(function (f) { return f.State === 'Normal'; })) { result = 'warning'; }

    if (shares && shares.NETLOGON && shares.NETLOGON.Status !== 'Shared') { result = 'critical'; }

    return result;
}

function backupStatus(ab) {
    if (!ab || !ab.Status) { return 'unknown'; }
    return ab.Status.toLowerCase();
}

function dnsHealthStatus(dns) {
    if (!dns || dns.SRVRecordRegistered === null || dns.SRVRecordRegistered === undefined) { return 'unknown'; }
    if (dns.SRVRecordRegistered === false) { return 'critical'; }
    if (dns.SelfResolvable === false) { return 'warning'; }
    return 'good';
}

function eventLogStatus(events) {
    if (!events) { return 'unknown'; }
    if (events.length === 0) { return 'good'; }
    if (events.length <= 2) { return 'warning'; }
    return 'critical';
}

function diStatusDot(dc) {
    const inv = dcInventoryByName[dc.Name];

    if (!inv || !inv.Reachable || !inv.DeepInsights) { return 'dot-unknown'; }

    const di = inv.DeepInsights;
    const rank = { good: 0, warning: 1, critical: 2 };
    let worst = 'good';

    function consider(status) {
        if (status === 'unknown') { return; }
        if (rank[status] !== undefined && rank[status] > rank[worst]) { worst = status; }
    }

    consider(dcDiagStatus(di.DCDiag));
    consider(dfsrStatus(di.DFSRBacklog));
    consider(backupStatus(inv.ADBackup));
    consider(dnsHealthStatus(di.DNSHealth));
    consider(eventLogStatus(di.EventLogs && di.EventLogs.DirectoryService));
    consider(eventLogStatus(di.EventLogs && di.EventLogs.DFSReplication));
    consider(eventLogStatus(di.EventLogs && di.EventLogs.DNSServer));

    return 'dot-' + worst;
}

function renderDIList(filterText) {
    const listEl = document.getElementById('diList');
    const countEl = document.getElementById('diCount');
    const filter = (filterText || '').trim().toLowerCase();

    const filtered = dcListData.filter(function (dc) {
        return !filter ||
            dc.Name.toLowerCase().indexOf(filter) !== -1 ||
            dc.Site.toLowerCase().indexOf(filter) !== -1;
    });

    countEl.textContent = filtered.length + ' of ' + dcListData.length + ' domain controllers';

    if (filtered.length === 0) {
        listEl.innerHTML = '<div class="dc-empty">No domain controllers match your search.</div>';
        return;
    }

    let html = '';

    filtered
        .slice()
        .sort(function (a, b) { return a.Name.localeCompare(b.Name); })
        .forEach(function (dc) {
            const activeClass = (dc.Name === selectedDIDC) ? ' active' : '';
            const dot = diStatusDot(dc);

            html += '<div class="dc-item' + activeClass + '" onclick="selectDIDC(\'' + dc.Name.replace(/'/g, "\\'") + '\')">' +
                '<span class="dc-item-name" title="' + escapeHtml(dc.Name) + '">' + escapeHtml(dc.Name) + '</span>' +
                '<span class="dc-item-dot ' + dot + '"></span>' +
                '</div>';
        });

    listEl.innerHTML = html;
}

function filterDIList() {
    renderDIList(document.getElementById('diSearchInput').value);
}

function selectDIDC(name) {
    selectedDIDC = name;
    renderDIList(document.getElementById('diSearchInput').value);
    renderDIDetail(name);
}

function renderEventLogTile(iconName, titleText, events) {
    const st = eventLogStatus(events);
    let tileHtml = '<div class="dc-tile dc-tile-wide">' + tileTitle(iconName, titleText, dcPill(st));

    if (events && events.length > 0) {
        tileHtml += '<table class="dc-mini-table"><tr><th>Time</th><th>Event ID</th><th>Source</th><th>Message</th></tr>';

        events.forEach(function (e) {
            tileHtml += '<tr><td>' + escapeHtml(e.Time) + '</td><td>' + escapeHtml(e.EventID) + '</td><td>' + escapeHtml(e.Source) + '</td><td>' + escapeHtml(e.Message) + '</td></tr>';
        });

        tileHtml += '</table>';
    }
    else {
        tileHtml += '<div class="dc-kv-row"><span>No recent error-level events found</span><span class="health health-good">Clear</span></div>';
    }

    tileHtml += '</div>';
    return tileHtml;
}

function renderDIDetail(name) {
    const detailEl = document.getElementById('diDetail');
    const meta = dcListData.find(function (d) { return d.Name === name; });
    const inv = dcInventoryByName[name];

    if (!meta) {
        detailEl.innerHTML = '<div class="dc-empty">Domain controller not found.</div>';
        return;
    }

    let badges = '<span class="dc-badge">' + escapeHtml(meta.Site) + '</span>';

    if (meta.IsGlobalCatalog) {
        badges += '<span class="dc-badge">Global Catalog</span>';
    }

    (meta.FSMORoles || []).forEach(function (role) {
        badges += '<span class="dc-badge">' + escapeHtml(role) + '</span>';
    });

    let html = '<div class="dc-detail-header"><div class="dc-detail-title">' + escapeHtml(meta.Name) + '</div>' + badges + '</div>';

    if (!inv || !inv.Reachable) {
        const reason = (inv && inv.Error) ? escapeHtml(inv.Error) : 'No inventory data was collected for this domain controller (collection may be disabled or WinRM is unreachable).';
        html += '<div class="dc-unreachable">Unable to collect live deep insights data.<br>' + reason + '</div>';
        detailEl.innerHTML = html;
        return;
    }

    const di = inv.DeepInsights || {};
    const ab = inv.ADBackup || {};

    html += '<div class="dc-tile-grid">';

    // 1. DCDIAG summary
    const dd = di.DCDiag || {};
    const ddSt = dcDiagStatus(dd);
    let ddTip = '';

    if (ddSt === 'critical') {
        ddTip = secTip(
            'One or more core DCDIAG tests failed (Connectivity, Replications, Advertising, FSMOCheck, KnowsOfRoleHolders, NetLogons, Services, or SysVolCheck). A failing test usually means this DC cannot be located, cannot replicate, or cannot authenticate clients correctly.',
            'Run "dcdiag /v" on this domain controller to see the detailed failure and resolve the underlying connectivity, replication, or service issue.',
            'https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/dcdiag',
            'dcdiag &ndash; Microsoft Learn'
        );
    }
    else if (ddSt === 'warning') {
        ddTip = secTip(
            'Some DCDIAG tests could not be evaluated, so a problem in those areas may be going undetected.',
            'Confirm dcdiag.exe is available and this account has rights to run it locally on the DC, then re-run collection.',
            'https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/dcdiag',
            'dcdiag &ndash; Microsoft Learn'
        );
    }

    html += '<div class="dc-tile">' + tileTitle('checklist', 'DCDIAG Summary', dcPill(ddSt)) + '<div class="dc-service-grid">';

    const ddKeys = Object.keys(dd);

    if (ddKeys.length > 0) {
        ddKeys.forEach(function (test) {
            const v = dd[test];
            const cls = (v === 'Passed') ? 'health-good' : (v === 'Failed') ? 'health-critical' : 'health-unknown';
            html += '<div class="dc-service-row"><span>' + escapeHtml(test) + '</span><span class="health ' + cls + '">' + escapeHtml(v) + '</span></div>';
        });
    }
    else {
        html += '<div class="dc-service-row"><span>No DCDIAG data collected</span><span class="health health-unknown">Unknown</span></div>';
    }

    html += '</div>' + ddTip + '</div>';

    // 2. SYSVOL / DFSR replication health
    const dfsrArr = di.DFSRBacklog || [];
    const shares = di.SharesHealth || {};
    const dfsrSt = dfsrStatus(dfsrArr, shares);
    let dfsrTip = '';

    const netlogonShared = !shares.NETLOGON || shares.NETLOGON.Status === 'Shared';

    if (!netlogonShared) {
        dfsrTip = secTip(
            'The NETLOGON share is not present on this domain controller. Clients authenticating against this DC will fail to download Group Policy objects and logon scripts.',
            'Check SYSVOL/DFSR replication health and the Netlogon service on this DC; the NETLOGON share is published from the same replicated SYSVOL data.',
            'https://learn.microsoft.com/en-us/troubleshoot/windows-server/networking/troubleshoot-missing-sysvol-and-netlogon-shares',
            'Troubleshoot missing SYSVOL/Netlogon shares &ndash; Microsoft Learn'
        );
    }
    else if (dfsrSt === 'critical') {
        dfsrTip = secTip(
            'One or more SYSVOL replicated folders are in an error state. SYSVOL replication failures can cause Group Policy and logon script inconsistencies between domain controllers.',
            'Investigate DFSR health on this DC (DFS Replication event log, dfsrdiag) and resolve the replicated folder error.',
            'https://learn.microsoft.com/en-us/troubleshoot/windows-server/networking/troubleshoot-missing-sysvol-and-netlogon-shares',
            'Troubleshoot missing SYSVOL/Netlogon shares &ndash; Microsoft Learn'
        );
    }
    else if (dfsrSt === 'warning') {
        dfsrTip = secTip(
            'SYSVOL replication is not in its normal steady state (e.g. initial sync or auto recovery), which can temporarily cause Group Policy or logon script inconsistencies until it completes.',
            'Monitor DFSR replication state; if it does not return to "Normal", investigate using dfsrdiag and the DFS Replication event log.',
            'https://learn.microsoft.com/en-us/troubleshoot/windows-server/networking/troubleshoot-missing-sysvol-and-netlogon-shares',
            'Troubleshoot missing SYSVOL/Netlogon shares &ndash; Microsoft Learn'
        );
    }

    html += '<div class="dc-tile">' + tileTitle('refresh', 'SYSVOL / DFSR Replication Health', dcPill(dfsrSt));

    if (dfsrArr.length > 0) {
        dfsrArr.forEach(function (f) {
            const cls = (f.State === 'Normal') ? 'health-good' : (f.State === 'In Error') ? 'health-critical' : 'health-warning';
            html += '<div class="dc-kv-row"><span>' + escapeHtml(f.ReplicatedFolderName) + '</span><span class="health ' + cls + '">' + escapeHtml(f.State) + '</span></div>';
        });
    }
    else {
        html += kvRow('No DFSR replicated folder data collected', '-');
    }

    if (shares.SYSVOL) {
        const sysvolCls = (shares.SYSVOL.Status === 'Shared') ? 'health-good' : 'health-critical';
        html += '<div class="dc-kv-row"><span>SYSVOL Share</span><span class="health ' + sysvolCls + '">' + escapeHtml(shares.SYSVOL.Status) + '</span></div>';
    }

    if (shares.NETLOGON) {
        const netlogonCls = (shares.NETLOGON.Status === 'Shared') ? 'health-good' : 'health-critical';
        html += '<div class="dc-kv-row"><span>NETLOGON Share</span><span class="health ' + netlogonCls + '">' + escapeHtml(shares.NETLOGON.Status) + '</span></div>';
    }

    html += dfsrTip + '</div>';

    // 3. AD backup recency
    const backupSt = backupStatus(ab);
    let backupTip = '';

    if (backupSt === 'critical') {
        backupTip = secTip(
            'This DC\'s AD database has not been backed up recently (or ever). Without a recent system state backup, recovering from accidental deletions, corruption, or a ransomware event could mean significant data loss.',
            'Schedule regular system state backups (e.g. Windows Server Backup) for this DC, well within the AD tombstone lifetime.',
            'https://learn.microsoft.com/en-us/services-hub/unified/health/remediation-steps-ad/investigate-why-active-directory-directory-partitions-are-not-backed-up-for-longer-than-half-the-tom',
            'AD partitions not backed up &ndash; Microsoft Learn'
        );
    }
    else if (backupSt === 'warning') {
        backupTip = secTip(
            'This DC\'s last backup is more than 7 days old, increasing the amount of data that could be lost if a restore is ever needed.',
            'Review the backup schedule for this DC and confirm system state backups are completing successfully.',
            'https://learn.microsoft.com/en-us/services-hub/unified/health/remediation-steps-ad/investigate-why-active-directory-directory-partitions-are-not-backed-up-for-longer-than-half-the-tom',
            'AD partitions not backed up &ndash; Microsoft Learn'
        );
    }

    html += '<div class="dc-tile" id="diTileAdBackup">' + tileTitle('database', 'AD Backup Recency', dcPill(backupSt)) +
        kvRow('Last Backup', ab.LastBackupTime) +
        kvRow('Days Since Backup', ab.DaysSinceBackup) +
        backupTip +
        '</div>';

    // 4. DNS health
    const dns = di.DNSHealth || {};
    const dnsSt = dnsHealthStatus(dns);
    let dnsTip = '';

    if (dnsSt === 'critical') {
        dnsTip = secTip(
            'This DC\'s LDAP SRV record was not found in DNS. Clients and other domain controllers may be unable to locate this DC for authentication and replication.',
            'Restart the Netlogon service on this DC to re-register its DNS records, and verify the DNS zone allows the required dynamic updates.',
            'https://learn.microsoft.com/en-us/troubleshoot/windows-server/networking/verify-srv-dns-records-have-been-created',
            'Verify SRV DNS records &ndash; Microsoft Learn'
        );
    }
    else if (dnsSt === 'warning') {
        dnsTip = secTip(
            'This DC\'s own hostname does not resolve via DNS, which can cause certificate validation, Kerberos, and replication issues.',
            'Verify this DC\'s DNS client settings and A/PTR records, and confirm its configured DNS servers are reachable and authoritative for the zone.',
            'https://learn.microsoft.com/en-us/troubleshoot/windows-server/networking/verify-srv-dns-records-have-been-created',
            'Verify SRV DNS records &ndash; Microsoft Learn'
        );
    }

    html += '<div class="dc-tile">' + tileTitle('globe', 'DNS Health', dcPill(dnsSt)) +
        kvRow('SRV Record Registered', boolLabel(dns.SRVRecordRegistered)) +
        kvRow('SRV Record Count', dns.SRVRecordCount) +
        kvRow('Self Resolvable', boolLabel(dns.SelfResolvable)) +
        dnsTip +
        '</div>';

    // 5. Recent interactive logons (full-width)
    const logons = di.RecentLogons || [];

    html += '<div class="dc-tile dc-tile-wide">' + tileTitle('login', 'Recent Interactive Logons (Last 5)');

    if (logons.length > 0) {
        html += '<table class="dc-mini-table"><tr><th>Time</th><th>Account</th><th>Logon Type</th><th>Source IP</th></tr>';

        logons.forEach(function (l) {
            html += '<tr><td>' + escapeHtml(l.Time) + '</td><td>' + escapeHtml(l.Account) + '</td><td>' + escapeHtml(l.LogonType) + '</td><td>' + escapeHtml(l.SourceIP) + '</td></tr>';
        });

        html += '</table>';
    }
    else {
        html += kvRow('No recent interactive logon events found', '-');
    }

    html += '</div>';

    // 6-8. Top 5 recent error-level events per log (full-width, last)
    const eventLogs = di.EventLogs || {};

    html += renderEventLogTile('alert', 'Directory Service Errors (Last 5)', eventLogs.DirectoryService);
    html += renderEventLogTile('alert', 'DFS Replication Errors (Last 5)', eventLogs.DFSReplication);
    html += renderEventLogTile('alert', 'DNS Server Errors (Last 5)', eventLogs.DNSServer);

    html += '</div>'; // close dc-tile-grid

    detailEl.innerHTML = html;

    adjustDIHeight();
}

renderDIList('');

if (dcListData.length > 0) {
    selectDIDC(dcListData[0].Name);
}
</script>

</body>
</html>
"@

$html | Out-File -FilePath $ReportPath -Encoding UTF8

Write-Host ""
Write-Host "Dashboard generated successfully:" -ForegroundColor Green
Write-Host $ReportPath -ForegroundColor Yellow
Write-Host ""

Start-Process $ReportPath
