#requires -modules ActiveDirectory

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

$html = @"
<!DOCTYPE html>
<html>
<head>
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
    }
</style>
</head>

<body>

<div class="page-generated">Generated On: $GeneratedOn</div>

<div class="container">

    <div class="header">
        <div class="header-title">Enterprise Active Directory Operational Dashboard</div>
    </div>

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
                    <button class="export-btn" title="Export DC ICMP Status" onclick="downloadCsv('DC_ICMP_Status.csv', dcPingCsv)">⇩</button>
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
                    <button class="export-btn" title="Export AD Sites" onclick="downloadCsv('AD_Sites.csv', adSitesCsv)">⇩</button>
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
                    <button class="export-btn" title="Export Sites with No DCs" onclick="downloadCsv('Sites_With_No_DCs.csv', sitesWithNoDCsCsv)">⇩</button>
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
                    <button class="export-btn" title="Export AD Subnets" onclick="downloadCsv('AD_Subnets.csv', adSubnetsCsv)">⇩</button>
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
                    <button class="export-btn" title="Export Domain Admins" onclick="downloadCsv('Domain_Admins.csv', domainAdminsCsv)">⇩</button>
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
                    <button class="export-btn" title="Export Enterprise Admins" onclick="downloadCsv('Enterprise_Admins.csv', enterpriseAdminsCsv)">⇩</button>
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
                    <button class="export-btn" title="Export Schema Admins" onclick="downloadCsv('Schema_Admins.csv', schemaAdminsCsv)">⇩</button>
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
                    <button class="export-btn" title="Export DNS Zones" onclick="downloadCsv('DNS_Zones.csv', dnsZonesCsv)">⇩</button>
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
                    <button class="export-btn" title="Export GPO Inventory" onclick="downloadCsv('GPO_Inventory.csv', gpoCsv)">⇩</button>
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
                    <button class="export-btn" title="Export Unlinked GPOs" onclick="downloadCsv('Unlinked_GPOs.csv', unlinkedGpoCsv)">⇩</button>
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
                <button class="section-export-btn-inline" title="Export Domain Controllers" onclick="downloadCsv('Domain_Controllers.csv', domainControllersCsv)">⇩</button>
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
                <button class="section-export-btn-inline" title="Export DNS Zones" onclick="downloadCsv('DNS_Zones.csv', dnsZonesCsv)">⇩</button>
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

    <div class="footer">
        Enterprise Active Directory Operational Dashboard | Generated using PowerShell | by Azgar Mohammad
    </div>

</div>

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