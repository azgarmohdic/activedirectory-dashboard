#requires -modules ActiveDirectory
# Last Updated : 16-Jun-2026
# Changes      : GPO report (7 columns incl. Policy Type + WMI Filter), DNS Zones column reorder,
#                Sites & Services Section 2 layout fix, sticky header + fixed footer,
#                footer simplified (no Generated On), tile hover z-index fix

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
    $ADSiteLinksCsvJs = ""
    $ADSiteLinkBridgeFlagJs = "unknown"
}

# ---------------------------------------------------------------------------
# AD Site Links + Bridge All Site Links flag
# ---------------------------------------------------------------------------
try {
    $ADSiteLinks = @(Get-ADReplicationSiteLink -Filter * -Properties *)

    try {
        $ConfigNC = (Get-ADRootDSE).configurationNamingContext
        $IPTransportObj = Get-ADObject -Identity "CN=IP,CN=Inter-Site Transports,CN=Sites,$ConfigNC" -Properties Options -ErrorAction Stop
        # Bit 1 (0x2) of Options = Bridge All Site Links DISABLED
        $ADSiteLinkBridgeFlagJs = if (($IPTransportObj.Options -band 0x2) -eq 0) { 'enabled' } else { 'disabled' }
    }
    catch {
        $ADSiteLinkBridgeFlagJs = 'unknown'
    }

    $ADSiteLinksCsvJs = ConvertTo-JavaScriptString (
        (
            $ADSiteLinks |
            Sort-Object Name |
            Select-Object `
                Name,
                Cost,
                ReplicationFrequencyInMinutes,
                @{Name="SitesIncluded";Expression={
                    ($_.SitesIncluded | ForEach-Object { Get-CleanADSiteName $_ }) -join '|'
                }},
                @{Name="Transport";Expression={
                    if ($_.DistinguishedName -match ',CN=SMTP,') { 'SMTP' } else { 'IP' }
                }} |
            ConvertTo-Csv -NoTypeInformation
        ) -join "`n"
    )
}
catch {
    $ADSiteLinksCsvJs = ""
    $ADSiteLinkBridgeFlagJs = "unknown"
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

# ---------------------------------------------------------------------------
# Identity Risk & Attack Surface tab - deduplicated Privileged Accounts
# (the same 3 groups as the Domain Admins/Enterprise Admins/Schema Admins
# tiles above, unioned by SamAccountName since the same person can be a
# member of more than one).
# ---------------------------------------------------------------------------
$PrivAccountsRaw = [ordered]@{}

# Reuses Get-PrivilegedGroupCsv - the SAME proven function that already
# builds the working Domain Admins/Enterprise Admins/Schema Admins tiles
# above - rather than a fresh AD query, so this can't drift from what those
# tiles already show. Each group is wrapped in its OWN try/catch so one
# group failing to resolve (e.g. Enterprise/Schema Admins only exist in the
# forest root domain) doesn't zero out the other two.
foreach ($pg in @('Domain Admins', 'Enterprise Admins', 'Schema Admins')) {
    try {
        $pgCsvText = Get-PrivilegedGroupCsv -Identity $pg -DomainName $DomainName
        $pgMembers = @($pgCsvText | ConvertFrom-Csv)
        foreach ($m in $pgMembers) {
            if (-not $m.SamAccountName) { continue }
            $key = $m.SamAccountName
            if (-not $PrivAccountsRaw.Contains($key)) {
                $PrivAccountsRaw[$key] = [ordered]@{ Name = $(if ($m.DisplayName) { $m.DisplayName } else { $key }); SamAccountName = $key; Groups = New-Object System.Collections.Generic.List[string] }
            }
            $PrivAccountsRaw[$key].Groups.Add($pg)
        }
    }
    catch {
        Write-Host "Warning: Could not read $pg membership for the Identity Risk tab - $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# Get-PrivilegedGroupCsv doesn't return LastLogonDate, and privileged
# groups are typically small, so a follow-up lookup per account is cheap.
foreach ($key in @($PrivAccountsRaw.Keys)) {
    try {
        $PrivAccountsRaw[$key].LastLogonDate = (Get-ADUser -Identity $key -Properties LastLogonDate -ErrorAction Stop).LastLogonDate
    }
    catch {
        $PrivAccountsRaw[$key].LastLogonDate = $null
    }
}

$PrivAccountsCount = $PrivAccountsRaw.Count
$PrivAccountsTooltip = if ($PrivAccountsCount -eq 0) {
    "<div class='tooltip-line'>No members found</div>"
} else {
    ($PrivAccountsRaw.Values | Sort-Object Name | ForEach-Object {
        $nameSafe = ConvertTo-HtmlSafe $_.Name
        $groupsStr = ConvertTo-HtmlSafe ($_.Groups -join ', ')
        "<div class='tooltip-line'><b>$nameSafe</b><span>$groupsStr</span></div>"
    }) -join ""
}

# ---------------------------------------------------------------------------
# Group & Privilege Architecture tab (Phase F1/F3) - Tier-0 Group Membership,
# member-count threshold, Universal-scope groups nested into a Tier-0
# group, and Pre-Windows 2000 Compatible Access. Deliberately a SEPARATE
# collector from $PrivAccountsRaw above (which only covers Domain Admins/
# Enterprise Admins/Schema Admins and feeds the existing Identity Risk
# tab's "Privileged Accounts" tile) - this extends coverage to the three
# commonly-overlooked operator groups plus the built-in Administrators
# group, on its own new tab, so the existing tile's definition and history
# stay unchanged.
# ---------------------------------------------------------------------------
$Tier0GroupNames               = @('Domain Admins', 'Enterprise Admins', 'Schema Admins', 'Administrators', 'Backup Operators', 'Account Operators', 'Print Operators')
$Tier0MemberThreshold          = 5
$Tier0TokenGroupCountThreshold = 120

$Tier0GroupMembers       = [ordered]@{}
$Tier0MemberFindings     = New-Object System.Collections.Generic.List[object]
$Tier0GroupCountFindings = New-Object System.Collections.Generic.List[object]

foreach ($t0g in $Tier0GroupNames) {
    try {
        $members = @(Get-ADGroupMember -Identity $t0g -Recursive -ErrorAction Stop | Where-Object { $_.objectClass -eq 'user' })
        $Tier0GroupMembers[$t0g] = $members

        if ($members.Count -gt $Tier0MemberThreshold) {
            $Tier0GroupCountFindings.Add([ordered]@{
                Group  = $t0g
                Count  = $members.Count
                Detail = "$($members.Count) members - exceeds the recommended threshold of $Tier0MemberThreshold for $t0g"
            })
        }

        foreach ($m in $members) {
            try {
                $u = Get-ADUser -Identity $m.SamAccountName -Properties Enabled, LastLogonDate, ServicePrincipalName, MemberOf -ErrorAction Stop
            } catch { continue }

            $isStale    = (-not $u.LastLogonDate) -or ($u.LastLogonDate -lt (Get-Date).AddDays(-90))
            $isSpn      = [bool]$u.ServicePrincipalName
            $groupCount = @($u.MemberOf).Count

            # Token-size risk folded in here rather than as a separate
            # collector - it's the same per-member AD round-trip already
            # being made for Enabled/LastLogonDate/SPN, just one more
            # property checked while we're already there.
            # Status/Reason kept as structured pairs (not just a joined
            # string) so the expandable detail panel can render a proper
            # Username | Status | Reason sub-table - one row per issue, so
            # an account with two problems (e.g. disabled AND stale) shows
            # up as two distinct rows there rather than one run-on sentence.
            $reasonPairs = New-Object System.Collections.Generic.List[object]
            if (-not $u.Enabled)                                  { $reasonPairs.Add([ordered]@{ Status = 'Disabled'; Reason = 'Account is disabled' }) }
            if ($isStale)                                          { $reasonPairs.Add([ordered]@{ Status = 'Stale'; Reason = 'No logon in the last 90+ days' }) }
            if ($isSpn)                                            { $reasonPairs.Add([ordered]@{ Status = 'Service Account'; Reason = 'Account has an SPN (Kerberoastable service account)' }) }
            if ($groupCount -gt $Tier0TokenGroupCountThreshold)    { $reasonPairs.Add([ordered]@{ Status = 'Token Size Risk'; Reason = "Carries $groupCount group memberships" }) }

            if ($reasonPairs.Count -gt 0) {
                $shortStatuses = ($reasonPairs | ForEach-Object { $_.Status }) -join ', '
                $Tier0MemberFindings.Add([ordered]@{
                    Account     = $m.SamAccountName
                    Group       = $t0g
                    Detail      = "Member of $t0g - $shortStatuses"
                    ReasonPairs = $reasonPairs
                })
            }
        }
    }
    catch {}
}

$Tier0GroupCountFindingsCount = $Tier0GroupCountFindings.Count
$Tier0MemberFindingsCount     = $Tier0MemberFindings.Count
$Tier0GroupStatus = if ($Tier0MemberFindingsCount -eq 0 -and $Tier0GroupCountFindingsCount -eq 0) { "All Tier-0 group members are enabled, active, and within threshold (Good)" } else { "$Tier0MemberFindingsCount risky member(s), $Tier0GroupCountFindingsCount group(s) over threshold across Tier-0 groups" }
$Tier0GroupHealth = if ($Tier0MemberFindingsCount -gt 0) { "Critical" } elseif ($Tier0GroupCountFindingsCount -gt 0) { "Warning" } else { "Good" }

# --- Universal-scope groups nested into a Tier-0 group (Phase F3) - direct
# group members only (one hop), checking each one's GroupScope. Universal
# groups replicate forest-wide and are usable from any domain in the
# forest, which is unusual and worth a second look as a Tier-0 nesting
# point. ---
$UniversalTier0Findings = New-Object System.Collections.Generic.List[object]
foreach ($t0g in $Tier0GroupNames) {
    try {
        $directGroupMembers = @(Get-ADGroupMember -Identity $t0g -ErrorAction Stop | Where-Object { $_.objectClass -eq 'group' })
        foreach ($gm in $directGroupMembers) {
            try {
                $gmObj = Get-ADGroup -Identity $gm.SamAccountName -Properties GroupScope -ErrorAction Stop
                if ($gmObj.GroupScope -eq 'Universal') {
                    $UniversalTier0Findings.Add([ordered]@{
                        Account = $gmObj.SamAccountName
                        Group   = $t0g
                        Detail  = "Universal-scope group nested into $t0g - replicates forest-wide and is usable from any domain in the forest"
                    })
                }
            } catch {}
        }
    } catch {}
}
$UniversalTier0Count  = $UniversalTier0Findings.Count
$UniversalTier0Status = if ($UniversalTier0Count -eq 0) { "No Universal-scope groups nested into a Tier-0 group (Good)" } else { "$UniversalTier0Count Universal-scope group(s) nested into a Tier-0 group" }
$UniversalTier0Health = if ($UniversalTier0Count -eq 0) { "Good" } else { "Warning" }

# --- Pre-Windows 2000 Compatible Access (Phase F3) - a classic, well-
# documented finding: if this built-in group contains Everyone, Anonymous
# Logon, or even just Authenticated Users, it grants excessive read access
# to security descriptors and enables certain legacy enumeration paths. ---
try {
    $PreWin2000Members    = @(Get-ADGroupMember -Identity 'Pre-Windows 2000 Compatible Access' -ErrorAction Stop)
    $PreWin2000RiskyNames = @($PreWin2000Members | Where-Object { $_.Name -in @('Everyone', 'ANONYMOUS LOGON', 'Authenticated Users') } | ForEach-Object { $_.Name })
    $PreWin2000Count  = $PreWin2000RiskyNames.Count
    $PreWin2000Status = if ($PreWin2000Count -eq 0) { "No Everyone/Anonymous Logon/Authenticated Users in Pre-Windows 2000 Compatible Access (Good)" } else { "$($PreWin2000RiskyNames -join ', ') present in Pre-Windows 2000 Compatible Access - excessive read access" }
    $PreWin2000Health = if ($PreWin2000Count -eq 0) { "Good" } else { "Critical" }
}
catch {
    $PreWin2000RiskyNames = @(); $PreWin2000Count = $null; $PreWin2000Status = "N/A"; $PreWin2000Health = "Unknown"
}

try {
    Import-Module DnsServer -ErrorAction Stop

    $DnsZones = @(Get-DnsServerZone -ErrorAction Stop)

    $DnsZoneCount = $DnsZones.Count
    $ADIntegratedDnsZoneCount = @($DnsZones | Where-Object { $_.IsDsIntegrated -eq $true }).Count
    $StandaloneDnsZoneCount = @($DnsZones | Where-Object { $_.IsDsIntegrated -ne $true }).Count

    $DnsZonesCsvJs = ConvertTo-JavaScriptString (
        (
            $DnsZones |
            Select-Object ZoneName, ZoneType, IsDsIntegrated, ReplicationScope, DynamicUpdate,
                IsReverseLookupZone, SecureSecondaries,
                @{ N = 'MasterServers'; E = { ($_.MasterServers | ForEach-Object { $_.IPAddressToString }) -join ',' } } |
            ConvertTo-Csv -NoTypeInformation
        ) -join "`n"
    )

    # --- DNS hardening (Phase A): zone transfer exposure, aging/scavenging,
    # cache locking. Zone transfer setting is read directly off the zone
    # object (SecureSecondaries) rather than a separate cmdlet call. ---
    $DnsOpenTransferZones = @($DnsZones | Where-Object { $_.SecureSecondaries -eq 'TransferAnyServer' -and -not $_.IsReverseLookupZone })
    $DnsOpenTransferCount = $DnsOpenTransferZones.Count
    $DnsOpenTransferNames = ($DnsOpenTransferZones | Select-Object -First 10 | ForEach-Object { $_.ZoneName }) -join "; "

    $DnsNoAgingZones = @()
    foreach ($z in ($DnsZones | Where-Object { $_.IsDsIntegrated -and $_.ZoneType -eq 'Primary' })) {
        try {
            $aging = Get-DnsServerZoneAging -Name $z.ZoneName -ErrorAction Stop
            if (-not $aging.AgingEnabled) { $DnsNoAgingZones += $z.ZoneName }
        } catch {}
    }
    $DnsNoAgingCount = $DnsNoAgingZones.Count
    $DnsNoAgingNames = ($DnsNoAgingZones | Select-Object -First 10) -join "; "

    try {
        $DnsCacheLockPct = (Get-DnsServerCache -ErrorAction Stop).LockingPercent
    } catch { $DnsCacheLockPct = $null }

    $DnsTransferStatus = if ($DnsOpenTransferCount -eq 0) { "No zones allow transfer to any server (Good)" } else { "$DnsOpenTransferCount zone(s) allow transfer to any server (Critical)" }
    $DnsTransferHealth = if ($DnsOpenTransferCount -eq 0) { "Good" } else { "Critical" }
    $DnsAgingStatus    = if ($DnsNoAgingCount -eq 0) { "Scavenging/aging enabled on all AD-integrated primary zones (Good)" } else { "$DnsNoAgingCount AD-integrated zone(s) have aging disabled - stale records accumulate" }
    $DnsAgingHealth    = if ($DnsNoAgingCount -eq 0) { "Good" } else { "Warning" }

    # --- Dynamic Update risk classification (Phase E1) - upgrades the
    # DynamicUpdate value already collected above from a raw state into an
    # explicit finding. Only "NonsecureAndSecure" is a real risk (any
    # client, authenticated or not depending on config, can create or
    # overwrite records - the classic DNS spoofing/credential-relay setup
    # vector); "Secure" and "None" are both fine and get no penalty.
    # Reverse lookup zones excluded for the same reason the zone-transfer
    # check above excludes them - the realistic attack is a forward
    # record (a fake server/SRV name), not a reverse PTR. ---
    $DnsInsecureUpdateZones = @($DnsZones | Where-Object { $_.DynamicUpdate -eq 'NonsecureAndSecure' -and -not $_.IsReverseLookupZone })
    $DnsInsecureUpdateCount = $DnsInsecureUpdateZones.Count
    $DnsInsecureUpdateNames = ($DnsInsecureUpdateZones | Select-Object -First 10 | ForEach-Object { $_.ZoneName }) -join "; "
    $DnsInsecureUpdateStatus = if ($DnsInsecureUpdateCount -eq 0) { "No zones allow nonsecure dynamic updates (Good)" } else { "$DnsInsecureUpdateCount zone(s) allow nonsecure dynamic updates - any client can register or overwrite records" }
    $DnsInsecureUpdateHealth = if ($DnsInsecureUpdateCount -eq 0) { "Good" } else { "Critical" }

    # --- DNS Zone ACL Audit (Phase E1) - named principals holding a
    # modify-class right (GenericAll/WriteDacl/WriteOwner/GenericWrite/
    # WriteProperty/CreateChild) directly on an AD-integrated zone's own
    # container object, outside the expected DNS-admin set. This is the
    # ACE-level escalation path (edit the zone object itself) - distinct
    # from the Dynamic Update setting above (a protocol-level policy, not
    # an ACL). Same ACL-reading technique already proven for AdminSDHolder
    # and DCSync rights elsewhere in this script, applied to a new target. ---
    if (-not $DomainDN) { $DomainDN = (Get-ADDomain).DistinguishedName }
    $ForestRootDN = $DomainDN
    try { $ForestRootDN = (Get-ADDomain -Identity (Get-ADForest -ErrorAction Stop).RootDomain -ErrorAction Stop).DistinguishedName } catch {}

    function Get-DnsZoneContainerDN {
        param($Zone)
        switch ($Zone.ReplicationScope) {
            'Forest' { "DC=$($Zone.ZoneName),CN=MicrosoftDNS,DC=ForestDnsZones,$ForestRootDN" }
            'Domain' { "DC=$($Zone.ZoneName),CN=MicrosoftDNS,DC=DomainDnsZones,$DomainDN" }
            'Legacy' { "DC=$($Zone.ZoneName),CN=MicrosoftDNS,CN=System,$DomainDN" }
            default  { $null }
        }
    }

    $DnsAclFindings = New-Object System.Collections.Generic.List[object]
    $DnsAclExpectedPrincipals = @('Domain Admins', 'Enterprise Admins', 'Administrators', 'SYSTEM', 'ENTERPRISE DOMAIN CONTROLLERS', 'DnsAdmins', 'CREATOR OWNER')

    foreach ($z in @($DnsZones | Where-Object { $_.IsDsIntegrated -eq $true })) {
        $zoneDN = Get-DnsZoneContainerDN -Zone $z
        if (-not $zoneDN) { continue }
        try {
            $zoneObj = Get-ADObject -Identity $zoneDN -Properties nTSecurityDescriptor -ErrorAction Stop
            $riskyAces = @($zoneObj.nTSecurityDescriptor.Access | Where-Object {
                $_.AccessControlType -eq 'Allow' -and -not $_.IsInherited -and
                ($_.ActiveDirectoryRights.ToString() -match 'GenericAll|WriteDacl|WriteOwner|GenericWrite|WriteProperty|CreateChild')
            })
            foreach ($ace in $riskyAces) {
                $idShort = ($ace.IdentityReference.Value -split '\\')[-1]
                if ($DnsAclExpectedPrincipals -contains $idShort) { continue }
                $DnsAclFindings.Add([ordered]@{
                    Zone    = $z.ZoneName
                    Account = $idShort
                    Detail  = "Holds $($ace.ActiveDirectoryRights) on the '$($z.ZoneName)' DNS zone object - can create or modify DNS records without being a recognized DNS administrator"
                })
            }
        }
        catch {}
    }
    $DnsAclFindingsCount = $DnsAclFindings.Count
    $DnsAclStatus = if ($DnsAclFindingsCount -eq 0) { "No unexpected principals hold modify rights on AD-integrated DNS zones (Good)" } else { "$DnsAclFindingsCount unexpected principal(s) hold modify rights on a DNS zone object" }
    $DnsAclHealth = if ($DnsAclFindingsCount -eq 0) { "Good" } else { "Critical" }

    # --- DnsAdmins Membership Review (Phase E1, not on Copilot's list) -
    # DnsAdmins membership is a well-documented path to SYSTEM on a DC via
    # the DNS server's plugin DLL loading mechanism (ServerLevelPluginDll)
    # - arguably more practically exploited than an over-permissioned zone
    # ACE, and costs almost nothing to check. Listed at the direct-member
    # level only (not expanded through nested groups) to keep this a
    # quick-win check rather than re-running the full attack-path walk. ---
    try {
        $DnsAdminsMembers = @(Get-ADGroupMember -Identity 'DnsAdmins' -ErrorAction Stop)
        $DnsAdminsCount = $DnsAdminsMembers.Count
        $DnsAdminsNames = ($DnsAdminsMembers | Select-Object -First 10 | ForEach-Object { $_.SamAccountName }) -join "; "
        $DnsAdminsStatus = if ($DnsAdminsCount -eq 0) { "DnsAdmins group has no members (Good)" } else { "$DnsAdminsCount named member(s) of DnsAdmins - review for the ServerLevelPluginDll escalation path" }
        $DnsAdminsHealth = if ($DnsAdminsCount -eq 0) { "Good" } else { "Warning" }
    }
    catch {
        $DnsAdminsMembers = @(); $DnsAdminsCount = $null; $DnsAdminsNames = ""; $DnsAdminsStatus = "N/A"; $DnsAdminsHealth = "Unknown"
    }
}
catch {
    $DnsZones = $null
    $DnsZoneCount = "N/A"
    $ADIntegratedDnsZoneCount = "N/A"
    $StandaloneDnsZoneCount = "N/A"
    $DnsZonesCsvJs = ""
    $DnsOpenTransferCount = $null; $DnsOpenTransferNames = ""; $DnsTransferStatus = "N/A"; $DnsTransferHealth = "Unknown"
    $DnsNoAgingCount = $null; $DnsNoAgingNames = ""; $DnsAgingStatus = "N/A"; $DnsAgingHealth = "Unknown"
    $DnsCacheLockPct = $null
    $DnsInsecureUpdateCount = $null; $DnsInsecureUpdateNames = ""; $DnsInsecureUpdateStatus = "N/A"; $DnsInsecureUpdateHealth = "Unknown"
    $DnsAclFindings = New-Object System.Collections.Generic.List[object]
    $DnsAclFindingsCount = $null; $DnsAclStatus = "N/A"; $DnsAclHealth = "Unknown"
    $DnsAdminsMembers = @(); $DnsAdminsCount = $null; $DnsAdminsNames = ""; $DnsAdminsStatus = "N/A"; $DnsAdminsHealth = "Unknown"
}

try {
    Import-Module GroupPolicy -ErrorAction Stop

    $AllGPOs = @(Get-GPO -All)
    $GPOCount = $AllGPOs.Count

    # Build enhanced GPO data: one Get-GPOReport call per GPO captures links + WMI filter
    # GPP cpassword scan target (Phase A hardening check) - the legacy
    # Group Policy Preferences vulnerability (MS14-025): credentials stored
    # in cleartext-reversible cpassword attributes inside SYSVOL XML files.
    $SysvolPoliciesPath = "\\$DomainName\SYSVOL\$DomainName\Policies"
    $GpoRiskyGpoNames   = New-Object System.Collections.Generic.List[string]
    $GpoCpasswordGpoNames = New-Object System.Collections.Generic.List[string]

    $GPODetails = @(
        foreach ($GPO in $AllGPOs) {
            $linkStr = ''
            $wmiName = if ($GPO.WmiFilter) { $GPO.WmiFilter.Name } else { '' }
            $hasSettings = $false
            $restrictedGroupsAdmin = $false
            try {
                $RptXmlText = Get-GPOReport -Guid $GPO.Id -ReportType Xml
                [xml]$Rpt = $RptXmlText
                $linkNodes = @($Rpt.GPO.LinksTo | Where-Object { $_ })
                if ($linkNodes.Count -gt 0) {
                    $linkStr = ($linkNodes | ForEach-Object {
                        $sp = $_.SOMPath
                        if ($sp -match '/$') { 'Domain Root' }
                        else { ($sp -replace '^[^/]+/', '').TrimEnd('/') }
                    }) -join '|'
                }
                $hasSettings = ($null -ne $Rpt.GPO.Computer.ExtensionData) -or ($null -ne $Rpt.GPO.User.ExtensionData)

                # Restricted Groups pushing membership into local Administrators -
                # text-search the raw report rather than parsing namespaced XML
                # nodes, which vary by GPMC version.
                if ($RptXmlText -match 'RestrictedGroups' -and $RptXmlText -match '(?i)Administrators') {
                    $restrictedGroupsAdmin = $true
                    $GpoRiskyGpoNames.Add($GPO.DisplayName)
                }
            }
            catch {}

            # GPP cpassword scan - only the GPOs that actually have a SYSVOL
            # folder containing the legacy *.xml preference files are checked.
            try {
                $gpoSysvolPath = Join-Path -Path $SysvolPoliciesPath -ChildPath "{$($GPO.Id)}"
                if (Test-Path $gpoSysvolPath) {
                    $prefXmlFiles = @(Get-ChildItem -Path $gpoSysvolPath -Recurse -Include 'Groups.xml','Services.xml','ScheduledTasks.xml','Printers.xml','Drives.xml','DataSources.xml' -ErrorAction SilentlyContinue)
                    foreach ($pf in $prefXmlFiles) {
                        $content = Get-Content -Path $pf.FullName -Raw -ErrorAction SilentlyContinue
                        if ($content -match 'cpassword="[^"]+"') {
                            $GpoCpasswordGpoNames.Add($GPO.DisplayName)
                            break
                        }
                    }
                }
            }
            catch {}

            [pscustomobject]@{
                DisplayName      = $GPO.DisplayName
                Id               = $GPO.Id.ToString()
                GpoStatus        = $GPO.GpoStatus.ToString()
                CreationTime     = $GPO.CreationTime.ToString('yyyy-MM-dd')
                ModificationTime = $GPO.ModificationTime.ToString('yyyy-MM-dd')
                WmiFilter        = $wmiName
                LinkedTo         = $linkStr
                IsEmpty          = (-not $hasSettings).ToString()
                RestrictedGroupsAdmin = $restrictedGroupsAdmin.ToString()
            }
        }
    )

    $DisabledGPOCount = @($GPODetails | Where-Object { $_.GpoStatus -ne 'AllSettingsEnabled' }).Count
    $UnlinkedGPOCount = @($GPODetails | Where-Object { -not $_.LinkedTo }).Count
    $GpoRiskyCount      = $GpoRiskyGpoNames.Count
    $GpoCpasswordCount  = $GpoCpasswordGpoNames.Count
    $GpoRiskyStatus     = if ($GpoRiskyCount -eq 0) { "No GPOs push membership into local Administrators (Good)" } else { "$GpoRiskyCount GPO(s) use Restricted Groups to add local Administrators members" }
    $GpoRiskyHealth     = if ($GpoRiskyCount -eq 0) { "Good" } else { "Warning" }
    $GpoCpasswordStatus = if ($GpoCpasswordCount -eq 0) { "No GPP cpassword remnants found (Good)" } else { "$GpoCpasswordCount GPO(s) contain a recoverable GPP cpassword (MS14-025)" }
    $GpoCpasswordHealth = if ($GpoCpasswordCount -eq 0) { "Good" } else { "Critical" }

    $GpoCsvJs = ConvertTo-JavaScriptString (
        ($GPODetails | ConvertTo-Csv -NoTypeInformation) -join "`n"
    )

    $UnlinkedGpoCsvJs = ConvertTo-JavaScriptString (
        ($GPODetails | Where-Object { -not $_.LinkedTo } |
         Select-Object DisplayName, Id, GpoStatus, CreationTime, ModificationTime |
         ConvertTo-Csv -NoTypeInformation) -join "`n"
    )
}
catch {
    $GPOCount         = "N/A"
    $DisabledGPOCount = "N/A"
    $UnlinkedGPOCount = "N/A"
    $GpoCsvJs         = ""
    $UnlinkedGpoCsvJs = ""
    $GpoRiskyGpoNames = New-Object System.Collections.Generic.List[string]
    $GpoCpasswordGpoNames = New-Object System.Collections.Generic.List[string]
    $GpoRiskyCount = $null; $GpoRiskyStatus = "N/A"; $GpoRiskyHealth = "Unknown"
    $GpoCpasswordCount = $null; $GpoCpasswordStatus = "N/A"; $GpoCpasswordHealth = "Unknown"
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
    $Trusts = @(Get-ADTrust -Filter * -Properties Direction, TrustType, TrustAttributes, SIDFilteringQuarantined, SIDFilteringForestAware, SelectiveAuthentication -ErrorAction Stop)
}
catch {
    $Trusts = $null
}

# ---------------------------------------------------------------------------
# Trust Security Deep-Dive (Phase C) - extends the basic trust inventory
# above with risk flags: SID filtering quarantine state, selective
# authentication state, and a forest-wide SID History presence scan. This
# stays read-only and local to AD objects - no cross-forest live probing.
# ---------------------------------------------------------------------------
$TrustRiskFindings = New-Object System.Collections.Generic.List[object]

if ($Trusts) {
    foreach ($t in $Trusts) {
        $dir = $t.Direction.ToString()
        # Inbound-only trusts (we trust them, they don't trust us) carry no
        # SID-filtering exposure risk to OUR domain, so only flag
        # Bidirectional/Outbound where the trusted side could inject SIDs.
        if ($dir -in @('Bidirectional', 'Outbound') -and -not $t.SIDFilteringQuarantined) {
            $TrustRiskFindings.Add([ordered]@{ Account = $t.Name; Detail = "$dir trust without SID filtering quarantine enabled - SID History injection from the trusted side could grant unintended rights in this domain" })
        }
        if ($t.TrustType -eq 'Uplevel' -and -not $t.SelectiveAuthentication -and $dir -in @('Bidirectional', 'Inbound')) {
            $TrustRiskFindings.Add([ordered]@{ Account = $t.Name; Detail = "Trust does not use selective authentication - every account in the trusted domain can attempt to authenticate to resources here" })
        }
    }
}

try {
    $SidHistoryObjects = @(Get-ADObject -LDAPFilter "(sIDHistory=*)" -Properties Name, objectClass -ErrorAction Stop)
    $SidHistoryCount = $SidHistoryObjects.Count
    $SidHistoryNames = ($SidHistoryObjects | Select-Object -First 10 | ForEach-Object { "$($_.Name) ($($_.objectClass))" }) -join "; "
}
catch {
    $SidHistoryObjects = @()
    $SidHistoryCount = $null
    $SidHistoryNames = ""
}

$TrustRiskCount = $TrustRiskFindings.Count
$TrustRiskStatus = if ($TrustRiskCount -eq 0) { "No trust hardening gaps found (Good)" } else { "$TrustRiskCount trust(s) missing SID filtering or selective authentication" }
$TrustRiskHealth = if ($TrustRiskCount -eq 0) { "Good" } else { "Warning" }
$SidHistoryStatus = if ($SidHistoryCount -eq 0) { "No objects with SID History found (Good)" } else { "$SidHistoryCount object(s) carry a SID History entry - confirm each is from a recognized migration" }
$SidHistoryHealth = if ($SidHistoryCount -eq 0) { "Good" } else { "Warning" }

# ---------------------------------------------------------------------------
# Global AD Security Health - one-time forest/domain-level checks
# These run once on the machine executing this script (needs AD module +
# domain admin rights). Results are embedded as JSON into the HTML.
# ---------------------------------------------------------------------------

# --- 1. Schema Version ---
try {
    $SchemaNC      = (Get-ADRootDSE -ErrorAction Stop).schemaNamingContext
    $SchemaObj     = Get-ADObject -Identity $SchemaNC -Properties objectVersion -ErrorAction Stop
    $SchemaVersion = [int]$SchemaObj.objectVersion

    $SchemaVersionLabel = switch ($SchemaVersion) {
        31  { "2000 RTM" }
        30  { "2003 RTM" }
        31  { "2003 R2" }
        44  { "2008 RTM" }
        47  { "2008 R2" }
        56  { "2012 RTM" }
        69  { "2012 R2" }
        87  { "2016" }
        88  { "2019" }
        91  { "2022" }
        default { "Unknown (v$SchemaVersion)" }
    }
}
catch {
    $SchemaVersion      = $null
    $SchemaVersionLabel = "N/A"
}

# --- 2. SYSVOL FRS→DFSR Migration State ---
try {
    $DomainDN          = (Get-ADDomain -ErrorAction Stop).DistinguishedName
    $SysvolMigObj      = Get-ADObject -Identity "CN=DFSR-GlobalSettings,CN=System,$DomainDN" `
                                      -Properties msDFSR-Flags -ErrorAction Stop
    $SysvolMigFlags    = $SysvolMigObj.'msDFSR-Flags'
    $SysvolMigState    = switch ($SysvolMigFlags) {
        0  { "0 - FRS (Deprecated)" }
        1  { "1 - Prepared" }
        2  { "2 - Redirected" }
        3  { "3 - Eliminated (DFSR)" }
        default { "Unknown ($SysvolMigFlags)" }
    }
    $SysvolMigHealth   = switch ($SysvolMigFlags) {
        3       { "Good" }
        0       { "Critical" }
        default { "Warning" }
    }
}
catch {
    # Object may not exist if already fully on DFSR or migration never started
    try {
        $SysvolMigState  = "3 - Eliminated (DFSR)"
        $SysvolMigHealth = "Good"
        $SysvolMigFlags  = 3
    }
    catch {
        $SysvolMigState  = "N/A"
        $SysvolMigHealth = "Unknown"
        $SysvolMigFlags  = $null
    }
}

# --- 3. Fine-Grained Password Policies (PSOs) ---
try {
    $FGPPs     = @(Get-ADFineGrainedPasswordPolicy -Filter * -ErrorAction Stop)
    $FGPPCount = $FGPPs.Count
    $FGPPNames = ($FGPPs | ForEach-Object { $_.Name }) -join ", "
}
catch {
    $FGPPCount = 0
    $FGPPNames = ""
}

# --- 4. Default Administrator Account (RID-500) ---
try {
    $DomainSID   = (Get-ADDomain -ErrorAction Stop).DomainSID.Value
    $AdminRID500 = Get-ADUser -Identity "$DomainSID-500" -Properties Enabled, SamAccountName -ErrorAction Stop

    $AdminIsRenamed  = $AdminRID500.SamAccountName -ne "Administrator"
    $AdminIsDisabled = -not $AdminRID500.Enabled

    $AdminRID500Status = if ($AdminIsRenamed -and $AdminIsDisabled) {
        "Renamed & Disabled (Good)"
    }
    elseif ($AdminIsRenamed) {
        "Renamed but Enabled (Warning)"
    }
    elseif ($AdminIsDisabled) {
        "Disabled but Not Renamed (Warning)"
    }
    else {
        "Default Name, Enabled (Critical)"
    }

    $AdminRID500Health = if ($AdminIsRenamed -and $AdminIsDisabled) { "Good" }
                         elseif ($AdminIsRenamed -or $AdminIsDisabled) { "Warning" }
                         else { "Critical" }

    $AdminRID500Name = $AdminRID500.SamAccountName
}
catch {
    $AdminRID500Status = "N/A"
    $AdminRID500Health = "Unknown"
    $AdminRID500Name   = "N/A"
}

# --- 5. Guest Account ---
try {
    $GuestAccount        = Get-ADUser -Identity "Guest" -Properties Enabled -ErrorAction Stop
    $GuestAccountEnabled = $GuestAccount.Enabled

    $GuestStatus = if ($GuestAccountEnabled) { "Enabled (Critical)" } else { "Disabled (Good)" }
    $GuestHealth = if ($GuestAccountEnabled) { "Critical" } else { "Good" }
}
catch {
    $GuestStatus = "N/A"
    $GuestHealth = "Unknown"
    $GuestAccountEnabled = $null
}

# --- 6. Protected Users Group Coverage ---
# Check whether Domain Admins, Enterprise Admins, Schema Admins members are
# all enrolled in the Protected Users security group.
try {
    $ProtectedUsersMembers = @(Get-ADGroupMember -Identity "Protected Users" -Recursive -ErrorAction Stop | Select-Object -ExpandProperty SamAccountName)

    $PrivGroupNames = @("Domain Admins", "Enterprise Admins", "Schema Admins")
    $ExposedAdmins  = @()

    foreach ($grp in $PrivGroupNames) {
        try {
            $members = @(Get-ADGroupMember -Identity $grp -Recursive -ErrorAction Stop | Where-Object { $_.objectClass -eq 'user' })
            foreach ($m in $members) {
                if ($ProtectedUsersMembers -notcontains $m.SamAccountName) {
                    $ExposedAdmins += "$($m.SamAccountName) ($grp)"
                }
            }
        }
        catch {}
    }

    $ExposedAdminCount  = $ExposedAdmins.Count
    $ProtectedUsersCoverage = if ($ExposedAdminCount -eq 0) { "All Admins Covered (Good)" } else { "$ExposedAdminCount admin(s) NOT in Protected Users" }
    $ProtectedUsersHealth   = if ($ExposedAdminCount -eq 0) { "Good" } else { "Warning" }
    $ExposedAdminsStr       = if ($ExposedAdmins.Count -gt 0) { $ExposedAdmins -join "; " } else { "" }
}
catch {
    $ExposedAdminCount      = $null
    $ProtectedUsersCoverage = "N/A"
    $ProtectedUsersHealth   = "Unknown"
    $ExposedAdminsStr       = ""
}

# --- 7. SPN Duplicates ---
try {
    $SPNOutput     = & setspn -X -F 2>&1
    $DupLines      = @($SPNOutput | Where-Object { $_ -match '^\s+SPN' })
    $SPNDupCount   = $DupLines.Count

    $SPNStatus = if ($SPNDupCount -eq 0) { "No Duplicates Found (Good)" } else { "$SPNDupCount Duplicate SPN(s) Found (Critical)" }
    $SPNHealth = if ($SPNDupCount -eq 0) { "Good" } else { "Critical" }
}
catch {
    $SPNDupCount = $null
    $SPNStatus   = "N/A"
    $SPNHealth   = "Unknown"
}

# --- 8. Accounts with Pre-Authentication Disabled (AS-REP Roastable) ---
try {
    # UAC bit 0x400000 = DONT_REQUIRE_PREAUTH
    $PreAuthDisabled = @(Get-ADUser -LDAPFilter "(userAccountControl:1.2.840.113556.1.4.803:=4194304)" `
                                    -Properties SamAccountName, Enabled, LastLogonDate -ErrorAction Stop |
                          Where-Object { $_.Enabled })

    $PreAuthCount  = $PreAuthDisabled.Count
    $PreAuthStatus = if ($PreAuthCount -eq 0) { "None Found (Good)" } else { "$PreAuthCount Account(s) AS-REP Roastable (Critical)" }
    $PreAuthHealth = if ($PreAuthCount -eq 0) { "Good" } else { "Critical" }
    $PreAuthNames  = if ($PreAuthCount -gt 0) { ($PreAuthDisabled | Select-Object -First 10 | ForEach-Object { $_.SamAccountName }) -join "; " } else { "" }
}
catch {
    $PreAuthCount  = $null
    $PreAuthStatus = "N/A"
    $PreAuthHealth = "Unknown"
    $PreAuthNames  = ""
}

# --- 9. Unconstrained Kerberos Delegation ---
# Computers/users with TrustedForDelegation=True (excluding DCs which require it)
try {
    $DCNames = @($DomainControllers | ForEach-Object { $_.Name })

    $UnconstrainedObjects = @(
        Get-ADObject -LDAPFilter "(userAccountControl:1.2.840.113556.1.4.803:=524288)" `
                     -Properties Name, objectClass, SamAccountName -ErrorAction Stop |
            Where-Object { $DCNames -notcontains $_.Name }
    )

    $UnconstrainedCount  = $UnconstrainedObjects.Count
    $UnconstrainedStatus = if ($UnconstrainedCount -eq 0) { "None Found (Good)" } else { "$UnconstrainedCount Object(s) with Unconstrained Delegation (Critical)" }
    $UnconstrainedHealth = if ($UnconstrainedCount -eq 0) { "Good" } else { "Critical" }
    $UnconstrainedNames  = if ($UnconstrainedCount -gt 0) {
        ($UnconstrainedObjects | Select-Object -First 10 | ForEach-Object {
            "$($_.Name) ($($_.objectClass))"
        }) -join "; "
    } else { "" }
}
catch {
    $UnconstrainedCount  = $null
    $UnconstrainedStatus = "N/A"
    $UnconstrainedHealth = "Unknown"
    $UnconstrainedNames  = ""
}

# --- 10. Password Never Expires ---
# UAC bit 0x10000 (65536) = DONT_EXPIRE_PASSWORD. Enabled accounts with this
# flag set never rotate their password, violating password policy and leaving
# credentials valid indefinitely if compromised.
try {
    $PwdNeverExpires = @(
        Get-ADUser -LDAPFilter "(&(userAccountControl:1.2.840.113556.1.4.803:=65536)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))" `
                   -Properties SamAccountName, Enabled, LastLogonDate -ErrorAction Stop
    )
    $PwdNeverExpiresCount  = $PwdNeverExpires.Count
    $PwdNeverExpiresStatus = if ($PwdNeverExpiresCount -eq 0) { "None Found (Good)" } else { "$PwdNeverExpiresCount Enabled Account(s) - Password Never Expires" }
    $PwdNeverExpiresHealth = if ($PwdNeverExpiresCount -eq 0) { "Good" } elseif ($PwdNeverExpiresCount -le 5) { "Warning" } else { "Critical" }
    $PwdNeverExpiresNames  = if ($PwdNeverExpiresCount -gt 0) {
        ($PwdNeverExpires | Select-Object -First 20 | ForEach-Object { $_.SamAccountName }) -join "; "
    } else { "" }
}
catch {
    $PwdNeverExpiresCount  = $null
    $PwdNeverExpiresStatus = "N/A"
    $PwdNeverExpiresHealth = "Unknown"
    $PwdNeverExpiresNames  = ""
}

# --- 11. Kerberoasting Exposure (accounts with an SPN set) ---
# Lists actual accounts, not just a count, so each can be remediated
# individually. Flags weak Kerberos encryption (RC4 only, no AES) and
# cross-checks privileged group membership - an SPN account that's also
# a Domain Admin is a far higher-value Kerberoasting target.
try {
    $KerberoastUsers = @(
        Get-ADUser -LDAPFilter "(servicePrincipalName=*)" -Properties SamAccountName, ServicePrincipalName, msDS-SupportedEncryptionTypes, Enabled, MemberOf, LastLogonDate -ErrorAction Stop |
            Where-Object { $_.Enabled }
    )

    $PrivGroupNames = @('Domain Admins', 'Enterprise Admins', 'Schema Admins', 'Administrators')

    $KerberoastDetails = @($KerberoastUsers | ForEach-Object {
        $encTypes = $_.'msDS-SupportedEncryptionTypes'
        $hasAes = $encTypes -and (($encTypes -band 8) -or ($encTypes -band 16))
        $isPriv = $false
        foreach ($dn in @($_.MemberOf)) {
            foreach ($pg in $PrivGroupNames) { if ($dn -match "^CN=$pg,") { $isPriv = $true; break } }
            if ($isPriv) { break }
        }
        [ordered]@{ SamAccountName = $_.SamAccountName; WeakEncryption = (-not $hasAes); Privileged = $isPriv; LastLogonDate = $_.LastLogonDate }
    })

    $KerberoastCount     = $KerberoastDetails.Count
    $KerberoastWeakCount = @($KerberoastDetails | Where-Object { $_.WeakEncryption }).Count
    $KerberoastPrivCount = @($KerberoastDetails | Where-Object { $_.Privileged }).Count
    $KerberoastStatus = if ($KerberoastCount -eq 0) { "None Found (Good)" } else { "$KerberoastCount SPN account(s) - $KerberoastWeakCount weak encryption, $KerberoastPrivCount privileged" }
    $KerberoastHealth = if ($KerberoastPrivCount -gt 0) { "Critical" } elseif ($KerberoastCount -gt 0) { "Warning" } else { "Good" }
}
catch {
    $KerberoastDetails   = @()
    $KerberoastCount     = $null
    $KerberoastWeakCount = $null
    $KerberoastPrivCount = $null
    $KerberoastStatus    = "N/A"
    $KerberoastHealth    = "Unknown"
}

# --- 12. DCSync Rights Audit (ACL-based - reads the security descriptor
# directly rather than parsing dsacls text output, which is fragile and
# locale-dependent). Anyone holding BOTH "Replicating Directory Changes"
# and "Replicating Directory Changes All" on the domain root can run
# DCSync and extract every credential in the domain. ---
try {
    $DCSyncGuid1 = [Guid]'1131f6aa-9c07-11d1-f79f-00c04fc2dcd2'
    $DCSyncGuid2 = [Guid]'1131f6ad-9c07-11d1-f79f-00c04fc2dcd2'
    $DomainDN    = (Get-ADDomain).DistinguishedName
    $DomainAcl   = (Get-ADObject -Identity $DomainDN -Properties nTSecurityDescriptor -ErrorAction Stop).nTSecurityDescriptor

    $ExpectedDCSyncPrincipals = @('Domain Admins', 'Enterprise Admins', 'Administrators', 'SYSTEM', 'Domain Controllers', 'Enterprise Domain Controllers')

    $DCSyncHolders = @($DomainAcl.Access | Where-Object {
        ($_.ObjectType -eq $DCSyncGuid1 -or $_.ObjectType -eq $DCSyncGuid2) -and $_.AccessControlType -eq 'Allow'
    } | ForEach-Object {
        $idRef = $_.IdentityReference.Value
        [ordered]@{ Principal = $idRef; Recognized = ($ExpectedDCSyncPrincipals -contains ($idRef -split '\\')[-1]) }
    })

    $DCSyncUnexpected = @($DCSyncHolders | Where-Object { -not $_.Recognized } | ForEach-Object { $_.Principal } | Select-Object -Unique)
    $DCSyncCount  = $DCSyncUnexpected.Count
    $DCSyncStatus = if ($DCSyncCount -eq 0) { "Only expected principals hold replication rights (Good)" } else { "$DCSyncCount unexpected principal(s) hold DCSync rights (Critical)" }
    $DCSyncHealth = if ($DCSyncCount -eq 0) { "Good" } else { "Critical" }
}
catch {
    $DCSyncUnexpected = @()
    $DCSyncCount  = $null
    $DCSyncStatus = "N/A"
    $DCSyncHealth = "Unknown"
}

# --- 13. AdminSDHolder ACL Drift ---
# AdminSDHolder's ACL is the template periodically copied onto every
# protected (Tier 0) object. Unexpected principals here propagate broadly
# and are a common, easy-to-miss persistence technique.
try {
    if (-not $DomainDN) { $DomainDN = (Get-ADDomain).DistinguishedName }
    $AdminSDHolderDN  = "CN=AdminSDHolder,CN=System,$DomainDN"
    $AdminSDHolderAcl = (Get-ADObject -Identity $AdminSDHolderDN -Properties nTSecurityDescriptor -ErrorAction Stop).nTSecurityDescriptor

    $ExpectedSDHolderPrincipals = @('Domain Admins', 'Enterprise Admins', 'Administrators', 'SYSTEM', 'Account Operators', 'Backup Operators', 'Print Operators', 'Server Operators', 'Cert Publishers', 'Authenticated Users', 'CREATOR OWNER')

    $SDHolderAces = @($AdminSDHolderAcl.Access | Where-Object { $_.AccessControlType -eq 'Allow' -and -not $_.IsInherited } | ForEach-Object {
        $idRef = $_.IdentityReference.Value
        [ordered]@{ Principal = $idRef; Recognized = ($ExpectedSDHolderPrincipals -contains ($idRef -split '\\')[-1]) }
    })

    $SDHolderUnexpected = @($SDHolderAces | Where-Object { -not $_.Recognized } | ForEach-Object { $_.Principal } | Select-Object -Unique)
    $SDHolderCount  = $SDHolderUnexpected.Count
    $SDHolderStatus = if ($SDHolderCount -eq 0) { "No unexpected principals on AdminSDHolder (Good)" } else { "$SDHolderCount unexpected principal(s) on AdminSDHolder (Critical)" }
    $SDHolderHealth = if ($SDHolderCount -eq 0) { "Good" } else { "Critical" }
}
catch {
    $SDHolderUnexpected = @()
    $SDHolderCount  = $null
    $SDHolderStatus = "N/A"
    $SDHolderHealth = "Unknown"
}

# ---------------------------------------------------------------------------
# AD-Native Attack Path Summary (Phase D, extended) - deliberately scoped to
# data already inside Active Directory: which GROUPS are nested into Domain
# Admins/Enterprise Admins (an indirect membership path many admins lose
# track of), and which non-standard principals hold a direct ACL right
# (GenericAll/WriteDacl/WriteOwner/GenericWrite) on those group objects
# themselves - a one-hop escalation path ("if I can edit this group's ACL,
# I can add myself to Domain Admins"). This explicitly does NOT enumerate
# local admin groups or sessions on member servers/workstations - that
# would require broad remote access across the whole estate, a different
# risk category than read-only AD queries (see the roadmap, Tier 3).
#
# Extended: rather than stopping at "Group X is nested into Domain Admins,"
# each finding is walked all the way down to the real, named human accounts
# affected, capturing the full hop-by-hop chain (user -> ...intermediate
# group(s)... -> Domain Admins) so the Identity Risk & Attack Surface tab
# can render it as a step-by-step walkthrough instead of one flat sentence.
# Capped at $AttackPathMaxExpand members per starting group - if someone has
# nested an enormous group (e.g. Domain Users) into Domain Admins, that case
# collapses back to one summary finding for the group itself rather than
# flooding the dashboard with thousands of rows.
# ---------------------------------------------------------------------------
$AttackPathFindings       = New-Object System.Collections.Generic.List[object]
$AttackPathStepsByAccount = [ordered]@{}
$AttackPathMaxExpand      = 200

function Expand-GroupToUsers {
    # Walks down a group's membership tree and returns one entry per real
    # USER found beneath it, each carrying the chain of group SamAccountNames
    # from $GroupName (inclusive) down to that user's direct group. Depth-
    # capped and cycle-guarded - legitimate AD nesting never needs more than
    # a handful of hops; this just protects against a misconfigured circular
    # nesting causing runaway recursion.
    param(
        [string]$GroupName,
        [string[]]$VisitedGroups = @()
    )
    $out = New-Object System.Collections.Generic.List[object]

    # A true cycle (revisiting an exact group already in this chain) is
    # reported, not just silently guarded - this is purely additive (the
    # function still returns the same empty list for this branch either
    # way) so it cannot change existing Attack Path behavior, it just also
    # records the finding for the new Nesting Depth & Circular References
    # tile to surface.
    if ($VisitedGroups -contains $GroupName) {
        if (-not $script:DetectedGroupCycles) { $script:DetectedGroupCycles = New-Object System.Collections.Generic.List[string] }
        $cycleDesc = (@($VisitedGroups) + @($GroupName)) -join ' -> '
        if (-not $script:DetectedGroupCycles.Contains($cycleDesc)) { $script:DetectedGroupCycles.Add($cycleDesc) }
        return $out
    }
    if ($VisitedGroups.Count -gt 6) { return $out }
    $newVisited = $VisitedGroups + $GroupName

    try { $members = @(Get-ADGroupMember -Identity $GroupName -ErrorAction Stop) } catch { return $out }

    foreach ($m in $members) {
        if ($m.objectClass -eq 'user') {
            $out.Add([ordered]@{ UserName = $m.SamAccountName; GroupChain = @($GroupName) })
        }
        elseif ($m.objectClass -eq 'group') {
            foreach ($deeper in (Expand-GroupToUsers -GroupName $m.SamAccountName -VisitedGroups $newVisited)) {
                $out.Add([ordered]@{ UserName = $deeper.UserName; GroupChain = @($GroupName) + $deeper.GroupChain })
            }
        }
    }
    return $out
}

function Encode-PathSteps {
    # Encodes an ordered step chain for compact CSV transport:
    # "::Step0|Edge1::Step1|Edge2::Step2" - the JS side splits on "|" then
    # on the first "::" to recover (edge-into-this-step, step label) pairs.
    param([string[]]$Steps, [string[]]$Edges)
    $parts = @()
    for ($i = 0; $i -lt $Steps.Count; $i++) {
        $edge = if ($i -eq 0) { '' } else { $Edges[$i - 1] }
        $parts += "$edge::$($Steps[$i])"
    }
    return ($parts -join '|')
}

function Add-AttackPathEntriesForPrincipal {
    # Shared logic for both the nested-group case and the ACL case: given a
    # starting principal (a group OR a single user/other principal) and the
    # label for the final hop into Tier0Group, expand all the way down to
    # real human accounts and queue one finding per account, each carrying
    # its full step chain.
    param(
        [string]$PrincipalName,
        [bool]$PrincipalIsGroup,
        [string]$Tier0Group,
        [string]$FinalEdgeLabel
    )

    if (-not $PrincipalIsGroup) {
        $steps   = @($PrincipalName, $Tier0Group)
        $edges   = @($FinalEdgeLabel)
        $pathStr = ($steps -join ' -> ') + " ($FinalEdgeLabel)"
        $encoded = Encode-PathSteps -Steps $steps -Edges $edges
        $script:AttackPathFindings.Add([ordered]@{ Account = $PrincipalName; Detail = $pathStr })
        if (-not $script:AttackPathStepsByAccount.Contains($PrincipalName)) { $script:AttackPathStepsByAccount[$PrincipalName] = $encoded }
        return
    }

    $usersUnder = @(Expand-GroupToUsers -GroupName $PrincipalName)

    if ($usersUnder.Count -eq 0 -or $usersUnder.Count -gt $script:AttackPathMaxExpand) {
        # Empty group, enumeration failure, or too large to expand
        # individually - keep one summary finding for the group itself
        # rather than either losing the finding or flooding the dashboard.
        $note    = if ($usersUnder.Count -gt $script:AttackPathMaxExpand) { " - $($usersUnder.Count) members, too many to list individually; review this group's membership directly" } else { "" }
        $steps   = @($PrincipalName, $Tier0Group)
        $edges   = @($FinalEdgeLabel)
        $pathStr = ($steps -join ' -> ') + " ($FinalEdgeLabel$note)"
        $encoded = Encode-PathSteps -Steps $steps -Edges $edges
        $script:AttackPathFindings.Add([ordered]@{ Account = $PrincipalName; Detail = $pathStr })
        if (-not $script:AttackPathStepsByAccount.Contains($PrincipalName)) { $script:AttackPathStepsByAccount[$PrincipalName] = $encoded }
        return
    }

    foreach ($u in $usersUnder) {
        $chainRev = @($u.GroupChain)
        [array]::Reverse($chainRev)
        $steps   = @($u.UserName) + $chainRev + @($Tier0Group)
        $edges   = @(for ($i = 0; $i -lt $chainRev.Count; $i++) { 'member of' }) + @($FinalEdgeLabel)
        $pathStr = ($steps -join ' -> ') + " ($FinalEdgeLabel)"
        $encoded = Encode-PathSteps -Steps $steps -Edges $edges
        $script:AttackPathFindings.Add([ordered]@{ Account = $u.UserName; Detail = $pathStr })
        if (-not $script:AttackPathStepsByAccount.Contains($u.UserName)) { $script:AttackPathStepsByAccount[$u.UserName] = $encoded }
    }
}

try {
    $EscalationRightsGuidNames = @('GenericAll', 'WriteDacl', 'WriteOwner', 'GenericWrite')

    foreach ($tier0Group in @('Domain Admins', 'Enterprise Admins')) {
        try {
            $grpObj = Get-ADGroup -Identity $tier0Group -Properties nTSecurityDescriptor, DistinguishedName -ErrorAction Stop

            # --- Nested group path: groups (not users) that are direct
            # members, each one widening who is effectively a Tier 0 admin
            # without showing up as a direct member of the group itself.
            # Walked all the way down to the real, named human accounts. ---
            $nestedGroups = @(Get-ADGroupMember -Identity $tier0Group -ErrorAction Stop | Where-Object { $_.objectClass -eq 'group' })
            foreach ($ng in $nestedGroups) {
                Add-AttackPathEntriesForPrincipal -PrincipalName $ng.SamAccountName -PrincipalIsGroup $true -Tier0Group $tier0Group -FinalEdgeLabel 'is nested inside'
            }

            # --- One-hop ACL escalation: who can edit this group's own ACL
            # or membership without being a member of it today. If the
            # principal holding the right is itself a group, walked down to
            # the real human accounts the same way as the nesting case. ---
            $grpAcl = $grpObj.nTSecurityDescriptor
            $riskyAces = @($grpAcl.Access | Where-Object {
                $_.AccessControlType -eq 'Allow' -and
                ($EscalationRightsGuidNames -contains $_.ActiveDirectoryRights.ToString().Split(',')[0].Trim()) -or
                ($_.ActiveDirectoryRights.ToString() -match 'GenericAll|WriteDacl|WriteOwner|GenericWrite|WriteProperty')
            })
            $ExpectedGroupAdmins = @('Domain Admins', 'Enterprise Admins', 'Administrators', 'SYSTEM')
            foreach ($ace in $riskyAces) {
                $idShort = ($ace.IdentityReference.Value -split '\\')[-1]
                if ($ExpectedGroupAdmins -contains $idShort) { continue }

                $isGroupPrincipal = $false
                try { $null = Get-ADGroup -Identity $idShort -ErrorAction Stop; $isGroupPrincipal = $true } catch {}

                $rightLabel = "holds $($ace.ActiveDirectoryRights) on"
                Add-AttackPathEntriesForPrincipal -PrincipalName $idShort -PrincipalIsGroup $isGroupPrincipal -Tier0Group $tier0Group -FinalEdgeLabel $rightLabel
            }
        }
        catch {}
    }
}
catch {}

$AttackPathCount  = $AttackPathFindings.Count
$AttackPathStatus = if ($AttackPathCount -eq 0) { "No nested-group or one-hop ACL escalation paths found into Domain Admins/Enterprise Admins (Good)" } else { "$AttackPathCount named account(s) with an escalation path into Domain Admins/Enterprise Admins" }
$AttackPathHealth = if ($AttackPathCount -eq 0) { "Good" } else { "Critical" }

# ---------------------------------------------------------------------------
# Group & Privilege Architecture tab (Phase F1b/F2) - Group-Based Delegation
# (generalized ACL audit + Reset Password extended right), Nesting Depth &
# Circular References, and Foreign Security Principals. Placed here, after
# the Attack Path section, because all three reuse Expand-GroupToUsers
# (defined above) - none of this duplicates Attack Path's own findings for
# Domain Admins/Enterprise Admins specifically; see the comments below for
# exactly how the overlap is avoided.
# ---------------------------------------------------------------------------

# --- Group-Based Delegation - generalizes the ACL-rights detection already
# proven for Domain Admins/Enterprise Admins to the other five Tier-0
# groups (Schema Admins/Administrators/Backup Operators/Account Operators/
# Print Operators), which today have no ACL audit at all. Also closes a
# real gap for ALL SEVEN groups, including Domain Admins/Enterprise
# Admins: a narrowly-scoped Reset Password Control Access Right shows up
# as an ExtendedRight ACE with a specific ObjectType GUID, not one of the
# generic rights flags (GenericAll/WriteDacl/etc.) Attack Path already
# checks for DA/EA - so adding the Reset Password check to DA/EA here does
# NOT duplicate anything, it catches something the existing check
# structurally cannot see. Kept as its own finding category, separate from
# "Attack Path to Domain Admins", to avoid double-counting the generic
# rights Attack Path already reports for DA/EA. ---
$ResetPasswordRightGuid     = [guid]'00299570-246d-11d0-a768-00aa006e0529'
$Tier0AclFindings           = New-Object System.Collections.Generic.List[object]
$Tier0AclExpectedPrincipals = @('Domain Admins', 'Enterprise Admins', 'Administrators', 'SYSTEM', 'ENTERPRISE DOMAIN CONTROLLERS', 'CREATOR OWNER')
$Tier0AclGenericCheckGroups = @('Schema Admins', 'Administrators', 'Backup Operators', 'Account Operators', 'Print Operators')

function Add-Tier0DelegationFinding {
    # Mirrors Add-AttackPathEntriesForPrincipal's expand-to-real-users
    # behavior (reusing the same Expand-GroupToUsers function and the same
    # $AttackPathMaxExpand cap) but writes to $Tier0AclFindings instead of
    # $AttackPathFindings, so this stays a separate finding category.
    param(
        [string]$PrincipalName,
        [bool]$PrincipalIsGroup,
        [string]$Tier0Group,
        [string]$FinalEdgeLabel
    )
    if (-not $PrincipalIsGroup) {
        $script:Tier0AclFindings.Add([ordered]@{ Account = $PrincipalName; Group = $Tier0Group; Detail = "$PrincipalName $FinalEdgeLabel the '$Tier0Group' group object" })
        return
    }
    $usersUnder = @(Expand-GroupToUsers -GroupName $PrincipalName)
    if ($usersUnder.Count -eq 0 -or $usersUnder.Count -gt $script:AttackPathMaxExpand) {
        $script:Tier0AclFindings.Add([ordered]@{ Account = $PrincipalName; Group = $Tier0Group; Detail = "$PrincipalName (group) $FinalEdgeLabel the '$Tier0Group' group object" })
        return
    }
    foreach ($u in $usersUnder) {
        $chainRev = @($u.GroupChain); [array]::Reverse($chainRev)
        $chainStr = ($chainRev -join ' -> ')
        $script:Tier0AclFindings.Add([ordered]@{ Account = $u.UserName; Group = $Tier0Group; Detail = "$($u.UserName) -> $chainStr -> $PrincipalName, which $FinalEdgeLabel the '$Tier0Group' group object" })
    }
}

foreach ($t0g in $Tier0GroupNames) {
    try {
        $t0gObj = Get-ADGroup -Identity $t0g -Properties nTSecurityDescriptor -ErrorAction Stop
        $aces = @($t0gObj.nTSecurityDescriptor.Access | Where-Object { $_.AccessControlType -eq 'Allow' -and -not $_.IsInherited })

        foreach ($ace in $aces) {
            $idShort = ($ace.IdentityReference.Value -split '\\')[-1]
            if ($Tier0AclExpectedPrincipals -contains $idShort) { continue }

            $isResetPwd     = ($ace.ActiveDirectoryRights.ToString() -match 'ExtendedRight') -and ($ace.ObjectType -eq $ResetPasswordRightGuid)
            $isGenericRisky = ($Tier0AclGenericCheckGroups -contains $t0g) -and ($ace.ActiveDirectoryRights.ToString() -match 'GenericAll|WriteDacl|WriteOwner|GenericWrite|WriteProperty')

            if (-not $isResetPwd -and -not $isGenericRisky) { continue }

            $isGroupPrincipal = $false
            try { $null = Get-ADGroup -Identity $idShort -ErrorAction Stop; $isGroupPrincipal = $true } catch {}

            $rightLabel = if ($isResetPwd) { 'holds Reset Password rights on' } else { "holds $($ace.ActiveDirectoryRights) on" }
            Add-Tier0DelegationFinding -PrincipalName $idShort -PrincipalIsGroup $isGroupPrincipal -Tier0Group $t0g -FinalEdgeLabel $rightLabel
        }
    }
    catch {}
}

$Tier0AclFindingsCount = $Tier0AclFindings.Count
$Tier0AclStatus = if ($Tier0AclFindingsCount -eq 0) { "No unexpected delegated rights found on Tier-0 group objects (Good)" } else { "$Tier0AclFindingsCount unexpected delegated right(s) found on Tier-0 group objects" }
$Tier0AclHealth = if ($Tier0AclFindingsCount -eq 0) { "Good" } else { "Critical" }

# --- Nesting Depth & Circular References - generalizes the chain-walker
# to all 7 Tier-0 groups and surfaces two things the Attack Path walker
# doesn't: an explicit hop-count finding when a chain runs deeper than
# $NestingDepthThreshold, and the circular-reference report Expand-
# GroupToUsers now records (see the cycle-recording added to that
# function). Re-walks nested groups already touched by Attack Path for
# DA/EA - a small, accepted amount of redundant querying in exchange for
# not coupling this new check to Attack Path's internal state. ---
$NestingDepthThreshold = 2
$NestingDepthFindings  = New-Object System.Collections.Generic.List[object]

# Records nesting structure for ANY Tier-0 group that has at least one
# nested group beneath it - not just chains that exceed the threshold.
# One level of nesting is normal and worth seeing, not just flagging once
# it gets deep; the threshold only escalates severity, it no longer gates
# whether the finding is captured at all (the earlier version reported "0"
# even when real, shallow nesting existed, since nothing had crossed the
# threshold yet).
foreach ($t0g in $Tier0GroupNames) {
    try {
        $nestedGroupsT0 = @(Get-ADGroupMember -Identity $t0g -ErrorAction Stop | Where-Object { $_.objectClass -eq 'group' })
        if ($nestedGroupsT0.Count -eq 0) { continue }

        $nestedGroupNames = @($nestedGroupsT0 | ForEach-Object { $_.SamAccountName })
        $maxDepth    = 1
        $sampleChain = ''
        foreach ($ng in $nestedGroupsT0) {
            $usersUnderNg = @(Expand-GroupToUsers -GroupName $ng.SamAccountName)
            foreach ($u in $usersUnderNg) {
                if ($u.GroupChain.Count -gt $maxDepth) {
                    $maxDepth = $u.GroupChain.Count
                    $chainRev = @($u.GroupChain); [array]::Reverse($chainRev)
                    $sampleChain = ($chainRev -join ' -> ') + " -> $t0g"
                }
            }
        }

        $exceedsNote = if ($maxDepth -gt $NestingDepthThreshold) { " - exceeds the $NestingDepthThreshold-hop guideline" } else { "" }
        $chainNote   = if ($sampleChain) { " (e.g. $sampleChain)" } else { "" }
        $NestingDepthFindings.Add([ordered]@{
            Group    = $t0g
            MaxDepth = $maxDepth
            Detail   = "Nested via: $($nestedGroupNames -join ', ') - max chain depth $maxDepth hop(s)$chainNote$exceedsNote"
        })
    }
    catch {}
}

$NestingDepthFindingsCount = $NestingDepthFindings.Count
$NestingDepthExceedingCount = @($NestingDepthFindings | Where-Object { $_.MaxDepth -gt $NestingDepthThreshold }).Count
$NestingDepthStatus = if ($NestingDepthFindingsCount -eq 0) { "No nested groups found under any Tier-0 group (Good)" } else { "$NestingDepthFindingsCount Tier-0 group(s) have nested groups, $NestingDepthExceedingCount exceeding $NestingDepthThreshold hops" }
$NestingDepthHealth = if ($NestingDepthExceedingCount -gt 0) { "Warning" } elseif ($NestingDepthFindingsCount -gt 0) { "Good" } else { "Good" }

$GroupCycleCount  = if ($DetectedGroupCycles) { $DetectedGroupCycles.Count } else { 0 }
$GroupCycleStatus = if ($GroupCycleCount -eq 0) { "No circular group membership detected (Good)" } else { "$GroupCycleCount circular group membership chain(s) detected" }
$GroupCycleHealth = if ($GroupCycleCount -eq 0) { "Good" } else { "Critical" }

# --- Foreign Security Principals - external (cross-domain/cross-forest)
# principals represented locally as FSP objects under their own well-known
# container. Flags any FSP found as a direct member of a Tier-0 group as
# Critical, and separately flags an FSP entry that no longer resolves to a
# real object in the trusted domain (the trust partner removed the user/
# group but the local reference remains) as a hygiene finding. SID
# translation requires reaching the trusted domain, so a transient network
# issue could occasionally read as "orphaned" when it isn't - best-effort,
# same caveat as other cross-domain checks in this script. ---
$FspFindings      = New-Object System.Collections.Generic.List[object]
$FspOrphanedCount = 0

try {
    if (-not $DomainDN) { $DomainDN = (Get-ADDomain).DistinguishedName }
    $FspContainer = "CN=ForeignSecurityPrincipals,$DomainDN"
    $FspObjects = @(Get-ADObject -SearchBase $FspContainer -Filter * -ErrorAction Stop | Where-Object { $_.ObjectClass -eq 'foreignSecurityPrincipal' })

    foreach ($fsp in $FspObjects) {
        $resolvedOk = $true
        try { $null = ([System.Security.Principal.SecurityIdentifier]$fsp.Name).Translate([System.Security.Principal.NTAccount]) } catch { $resolvedOk = $false }
        if (-not $resolvedOk) { $FspOrphanedCount++ }
    }

    foreach ($t0g in $Tier0GroupNames) {
        try {
            $t0gMembersAll = @(Get-ADGroupMember -Identity $t0g -ErrorAction Stop)
            foreach ($mm in $t0gMembersAll) {
                if ($FspObjects.DistinguishedName -contains $mm.DistinguishedName) {
                    $FspFindings.Add([ordered]@{
                        Account = $mm.SamAccountName
                        Group   = $t0g
                        Detail  = "Foreign Security Principal is a direct member of $t0g - external (cross-domain/cross-forest) principal holding Tier-0 access"
                    })
                }
            }
        } catch {}
    }
}
catch {}

$FspFindingsCount = $FspFindings.Count
$FspStatus = if ($FspFindingsCount -eq 0 -and $FspOrphanedCount -eq 0) { "No foreign security principals in Tier-0 groups, none orphaned (Good)" } else { "$FspFindingsCount FSP(s) in a Tier-0 group, $FspOrphanedCount orphaned FSP entr(y/ies)" }
$FspHealth = if ($FspFindingsCount -gt 0) { "Critical" } elseif ($FspOrphanedCount -gt 0) { "Warning" } else { "Good" }

# ---------------------------------------------------------------------------
# Group & Privilege Architecture tab - GROUP-LEVEL summary rows. Per
# feedback: per-account detail for Tier-0 group members, FSPs, and
# delegation findings already lives on the Identity Risk & Attack Surface
# tab (they are dual-surfaced there via $IdentityRelevantCategories) - this
# tab should discuss the GROUPS themselves, one row per group, not repeat
# the same individual accounts a second time. Every finding that touches a
# given Tier-0 group (membership risk, size, delegation, Universal
# nesting, FSPs, depth) is folded into ONE row for that group; clicking it
# expands to the full member list and the specific named accounts behind
# each finding. Columns deliberately match the Identity Risk & Attack
# Surface table exactly: Name | Type | Risk | Why | Recommended Action.
# ---------------------------------------------------------------------------
$GpaGroupRows = New-Object System.Collections.Generic.List[object]
$GpaSevRank   = @{ High = 0; Medium = 1; Low = 2 }
$GpaListCap   = 50

function Get-CappedPipeJoin {
    # Caps a list of already-formatted strings to $GpaListCap entries
    # before pipe-joining - this is the part that has to hold up across
    # ANY environment this script runs against, not just the one it was
    # tested in: a Tier-0 group with hundreds of members (or dozens of
    # risky ones) in a large or poorly-tiered enterprise should never
    # produce an unbounded wall of HTML in the detail panel the way a
    # handful of members in a small lab would not have exposed. The true
    # count is tracked separately (see *Total fields below) so the UI can
    # still say "showing 50 of 237" rather than silently truncating.
    param([string[]]$Items, [int]$Cap = $script:GpaListCap)
    if (-not $Items -or $Items.Count -eq 0) { return '' }
    return (($Items | Select-Object -First $Cap) -join '|')
}

foreach ($t0g in $Tier0GroupNames) {
    $allMembers   = @($Tier0GroupMembers[$t0g] | ForEach-Object { $_.SamAccountName })
    $riskyMembers = @($Tier0MemberFindings | Where-Object { $_.Group -eq $t0g })
    $sizeFinding  = $Tier0GroupCountFindings | Where-Object { $_.Group -eq $t0g } | Select-Object -First 1
    $aclFindings  = @($Tier0AclFindings | Where-Object { $_.Group -eq $t0g })
    $univFindings = @($UniversalTier0Findings | Where-Object { $_.Group -eq $t0g })
    $fspInGroup   = @($FspFindings | Where-Object { $_.Group -eq $t0g })
    $nestFinding  = $NestingDepthFindings | Where-Object { $_.Group -eq $t0g } | Select-Object -First 1

    $whyParts = @()
    $actions  = @()
    $sev = 'Low'

    if ($aclFindings.Count -gt 0)  { $whyParts += "$($aclFindings.Count) delegated right(s) incl. possible Reset Password"; $actions += 'Remove the unexpected delegated right from the group object'; $sev = 'High' }
    if ($fspInGroup.Count -gt 0)   { $whyParts += "$($fspInGroup.Count) foreign security principal(s)"; $actions += 'Confirm cross-domain/cross-forest access is still required'; $sev = 'High' }
    if ($riskyMembers.Count -gt 0) { $whyParts += "$($riskyMembers.Count) disabled/stale/SPN/oversized-token member(s)"; $actions += 'Remove disabled/stale members, review service accounts'; if ($sev -ne 'High') { $sev = 'High' } }
    if ($sizeFinding)              { $whyParts += "$($allMembers.Count) total members (exceeds threshold of $Tier0MemberThreshold)"; $actions += 'Review full membership for anyone without an active need'; if ($sev -eq 'Low') { $sev = 'Medium' } }
    if ($univFindings.Count -gt 0) { $whyParts += "$($univFindings.Count) Universal-scope group nested"; $actions += 'Confirm the forest-wide reach is intentional'; if ($sev -eq 'Low') { $sev = 'Medium' } }
    if ($nestFinding -and $nestFinding.MaxDepth -gt $NestingDepthThreshold) { $whyParts += "nesting chain $($nestFinding.MaxDepth) hops deep"; $actions += 'Flatten the nesting or document the chain'; if ($sev -eq 'Low') { $sev = 'Medium' } }
    elseif ($nestFinding) { $whyParts += "$($nestFinding.MaxDepth) level(s) of group nesting (within guideline)" }

    if ($whyParts.Count -eq 0) { continue }  # nothing to report for this group - skip the row entirely (Good)

    # Pipe-delimited so the expandable detail panel can list every named
    # account behind each finding, not just a count - same encoding
    # convention already used elsewhere in this script (e.g. PathSteps).
    # Risky members use a 3-field "Account::Status::Reason" encoding (one
    # entry PER ISSUE, not per account) so an account with two problems
    # renders as two distinct sub-table rows instead of a run-on sentence.
    # Every list is built as an array FIRST (so its true, uncapped count is
    # known) and only joined into the capped pipe-string afterward - this
    # is what lets the UI distinguish "7 risky members, all shown" from
    # "237 members, showing the first 50" in a much larger environment.
    $riskyDetailEntries = @($riskyMembers | ForEach-Object {
        $acct = $_.Account
        $_.ReasonPairs | ForEach-Object { "$acct::$($_.Status)::$($_.Reason)" }
    })
    $aclDetailEntries  = @($aclFindings  | ForEach-Object { "$($_.Account): $($_.Detail)" })
    $univDetailEntries = @($univFindings | ForEach-Object { $_.Account })
    $fspDetailEntries  = @($fspInGroup   | ForEach-Object { $_.Account })

    $GpaGroupRows.Add([PSCustomObject]@{
        Name          = $t0g
        Type          = 'Tier-0 Group'
        Risk          = $sev
        Why           = ($whyParts -join '; ')
        Action        = $(if ($actions.Count -gt 0) { $actions[0] } else { 'Confirm nested members are all meant to be Tier-0' })
        MemberCount   = $allMembers.Count
        AllMembers    = Get-CappedPipeJoin -Items $allMembers
        RiskyDetail   = Get-CappedPipeJoin -Items $riskyDetailEntries
        RiskyTotal    = $riskyDetailEntries.Count
        AclDetail     = Get-CappedPipeJoin -Items $aclDetailEntries
        AclTotal      = $aclDetailEntries.Count
        UnivDetail    = Get-CappedPipeJoin -Items $univDetailEntries
        UnivTotal     = $univDetailEntries.Count
        FspDetail     = Get-CappedPipeJoin -Items $fspDetailEntries
        FspTotal      = $fspDetailEntries.Count
        NestingDetail = $(if ($nestFinding) { $nestFinding.Detail } else { '' })
    })
}

# Circular references and Pre-Windows 2000 aren't tied to one specific
# Tier-0 group, so they get their own standalone rows.
if ($DetectedGroupCycles) {
    foreach ($cyc in $DetectedGroupCycles) {
        $GpaGroupRows.Add([PSCustomObject]@{
            Name = 'Circular group reference'; Type = 'Structural'; Risk = 'High'
            Why = $cyc; Action = 'Identify and remove the circular nesting'
            MemberCount = ''; AllMembers = ''
            RiskyDetail = ''; RiskyTotal = 0; AclDetail = ''; AclTotal = 0
            UnivDetail = ''; UnivTotal = 0; FspDetail = ''; FspTotal = 0
            NestingDetail = ''
        })
    }
}
if ($PreWin2000Count -gt 0) {
    $GpaGroupRows.Add([PSCustomObject]@{
        Name = 'Pre-Windows 2000 Compatible Access'; Type = 'Legacy Group'; Risk = 'High'
        Why = "Contains $($PreWin2000RiskyNames -join ', ')"; Action = 'Remove Everyone/Anonymous Logon/Authenticated Users unless a documented legacy system requires it'
        MemberCount = ''; AllMembers = ''
        RiskyDetail = ''; RiskyTotal = 0; AclDetail = ''; AclTotal = 0
        UnivDetail = ''; UnivTotal = 0; FspDetail = ''; FspTotal = 0
        NestingDetail = ''
    })
}

$GpaGroupRowsSorted = @($GpaGroupRows | Sort-Object { $GpaSevRank[[string]$_.Risk] })
$GpaGroupRowsCsvJs = ConvertTo-JavaScriptString (($GpaGroupRowsSorted | ConvertTo-Csv -NoTypeInformation) -join "`n")

# --- 14. Stale / Inactive Accounts (no logon in 90+ days) ---
try {
    $StaleThreshold  = (Get-Date).AddDays(-90)
    $AllEnabledUsers = @(Get-ADUser -Filter { Enabled -eq $true } -Properties SamAccountName, LastLogonDate, MemberOf -ErrorAction Stop)
    $StaleUsersRaw   = @($AllEnabledUsers | Where-Object { -not $_.LastLogonDate -or $_.LastLogonDate -lt $StaleThreshold })

    $StalePrivCount = 0
    $StaleNames = @($StaleUsersRaw | ForEach-Object {
        $isPriv = $false
        foreach ($dn in @($_.MemberOf)) {
            foreach ($pg in @('Domain Admins', 'Enterprise Admins', 'Schema Admins', 'Administrators')) {
                if ($dn -match "^CN=$pg,") { $isPriv = $true; break }
            }
            if ($isPriv) { break }
        }
        if ($isPriv) { $script:StalePrivCount++ }
        [ordered]@{ SamAccountName = $_.SamAccountName; LastLogonDate = $_.LastLogonDate; Privileged = $isPriv }
    })

    $StaleCount  = $StaleUsersRaw.Count
    $StaleStatus = if ($StaleCount -eq 0) { "No stale accounts found (Good)" } else { "$StaleCount account(s) inactive 90+ days - $StalePrivCount privileged" }
    $StaleHealth = if ($StalePrivCount -gt 0) { "Critical" } elseif ($StaleCount -gt 0) { "Warning" } else { "Good" }
}
catch {
    $StaleNames     = @()
    $StaleCount     = $null
    $StalePrivCount = $null
    $StaleStatus    = "N/A"
    $StaleHealth    = "Unknown"
}

# --- 15. Password Age Distribution (extends the existing PwdNeverExpires
# check with accounts that DO expire but haven't rotated in a long time) ---
try {
    $PwdLastSetUsers = @(Get-ADUser -Filter { Enabled -eq $true } -Properties SamAccountName, PasswordLastSet, PasswordNeverExpires -ErrorAction Stop)
    $NowTs = Get-Date

    $PwdAgeOver1Year = @($PwdLastSetUsers | Where-Object {
        $_.PasswordLastSet -and (-not $_.PasswordNeverExpires) -and (($NowTs - $_.PasswordLastSet).Days -gt 365)
    })
    $PwdAgeOver1YearCount = $PwdAgeOver1Year.Count
    $PwdAgeStatus = if ($PwdAgeOver1YearCount -eq 0) { "No accounts with 365+ day old passwords (Good)" } else { "$PwdAgeOver1YearCount account(s) with passwords unchanged 365+ days" }
    $PwdAgeHealth = if ($PwdAgeOver1YearCount -eq 0) { "Good" } elseif ($PwdAgeOver1YearCount -le 10) { "Warning" } else { "Critical" }
}
catch {
    $PwdAgeOver1YearCount = $null
    $PwdAgeStatus = "N/A"
    $PwdAgeHealth = "Unknown"
}

# --- 16. Constrained Delegation + gMSA Review (extends the existing
# unconstrained delegation check) ---
try {
    $ConstrainedObjects = @(Get-ADObject -LDAPFilter "(msDS-AllowedToDelegateTo=*)" -Properties Name, objectClass, SamAccountName -ErrorAction Stop)
    $ConstrainedCount = $ConstrainedObjects.Count
    $ConstrainedNames = if ($ConstrainedCount -gt 0) {
        ($ConstrainedObjects | Select-Object -First 10 | ForEach-Object { "$($_.Name) ($($_.objectClass))" }) -join "; "
    } else { "" }

    $GmsaAccounts = @(Get-ADServiceAccount -Filter * -ErrorAction SilentlyContinue)
    $GmsaCount = $GmsaAccounts.Count

    $ConstrainedStatus = if ($ConstrainedCount -eq 0) { "No constrained delegation configured" } else { "$ConstrainedCount object(s) with constrained delegation" }
    $ConstrainedHealth = if ($ConstrainedCount -eq 0) { "Good" } else { "Warning" }
}
catch {
    $ConstrainedCount  = $null
    $ConstrainedNames  = ""
    $ConstrainedStatus = "N/A"
    $ConstrainedHealth = "Unknown"
    $GmsaCount         = $null
}

# Serialize Global AD Security data as JSON for embedding in HTML
$GlobalADSecurityData = [ordered]@{
    SchemaVersion           = [ordered]@{ Version = $SchemaVersion; Label = $SchemaVersionLabel }
    SysvolMigration         = [ordered]@{ State = $SysvolMigState; Health = $SysvolMigHealth; Flags = $SysvolMigFlags }
    FGPP                    = [ordered]@{ Count = $FGPPCount; Names = $FGPPNames }
    AdminRID500             = [ordered]@{ Status = $AdminRID500Status; Health = $AdminRID500Health; Name = $AdminRID500Name }
    GuestAccount            = [ordered]@{ Status = $GuestStatus; Health = $GuestHealth; Enabled = $GuestAccountEnabled }
    ProtectedUsers          = [ordered]@{ Status = $ProtectedUsersCoverage; Health = $ProtectedUsersHealth; ExposedCount = $ExposedAdminCount; ExposedAdmins = $ExposedAdminsStr }
    SPNDuplicates           = [ordered]@{ Status = $SPNStatus; Health = $SPNHealth; Count = $SPNDupCount }
    PreAuthDisabled         = [ordered]@{ Status = $PreAuthStatus; Health = $PreAuthHealth; Count = $PreAuthCount; Accounts = $PreAuthNames }
    UnconstrainedDelegation = [ordered]@{ Status = $UnconstrainedStatus; Health = $UnconstrainedHealth; Count = $UnconstrainedCount; Objects = $UnconstrainedNames }
    PwdNeverExpires         = [ordered]@{ Status = $PwdNeverExpiresStatus; Health = $PwdNeverExpiresHealth; Count = $PwdNeverExpiresCount; Accounts = $PwdNeverExpiresNames }
    Kerberoasting           = [ordered]@{ Status = $KerberoastStatus; Health = $KerberoastHealth; Count = $KerberoastCount; WeakCount = $KerberoastWeakCount; PrivCount = $KerberoastPrivCount }
    DCSyncRights            = [ordered]@{ Status = $DCSyncStatus; Health = $DCSyncHealth; Count = $DCSyncCount; Principals = ($DCSyncUnexpected -join "; ") }
    AdminSDHolderDrift      = [ordered]@{ Status = $SDHolderStatus; Health = $SDHolderHealth; Count = $SDHolderCount; Principals = ($SDHolderUnexpected -join "; ") }
    StaleAccounts           = [ordered]@{ Status = $StaleStatus; Health = $StaleHealth; Count = $StaleCount; PrivCount = $StalePrivCount }
    PasswordAge             = [ordered]@{ Status = $PwdAgeStatus; Health = $PwdAgeHealth; Count = $PwdAgeOver1YearCount }
    ConstrainedDelegation   = [ordered]@{ Status = $ConstrainedStatus; Health = $ConstrainedHealth; Count = $ConstrainedCount; Objects = $ConstrainedNames; GmsaCount = $GmsaCount }
    DnsZoneTransfer         = [ordered]@{ Status = $DnsTransferStatus; Health = $DnsTransferHealth; Count = $DnsOpenTransferCount; Zones = $DnsOpenTransferNames }
    DnsAging                = [ordered]@{ Status = $DnsAgingStatus; Health = $DnsAgingHealth; Count = $DnsNoAgingCount; Zones = $DnsNoAgingNames; CacheLockPercent = $DnsCacheLockPct }
    DnsDynamicUpdateRisk    = [ordered]@{ Status = $DnsInsecureUpdateStatus; Health = $DnsInsecureUpdateHealth; Count = $DnsInsecureUpdateCount; Zones = $DnsInsecureUpdateNames }
    DnsZoneAcl              = [ordered]@{ Status = $DnsAclStatus; Health = $DnsAclHealth; Count = $DnsAclFindingsCount }
    DnsAdminsMembership     = [ordered]@{ Status = $DnsAdminsStatus; Health = $DnsAdminsHealth; Count = $DnsAdminsCount; Accounts = $DnsAdminsNames }
    Tier0GroupMembership    = [ordered]@{ Status = $Tier0GroupStatus; Health = $Tier0GroupHealth; Count = $Tier0MemberFindingsCount }
    Tier0GroupSize          = [ordered]@{ Status = "$($Tier0GroupCountFindingsCount) Tier-0 group(s) over the recommended member threshold"; Health = $(if ($Tier0GroupCountFindingsCount -eq 0) { "Good" } else { "Warning" }); Count = $Tier0GroupCountFindingsCount }
    UniversalTier0          = [ordered]@{ Status = $UniversalTier0Status; Health = $UniversalTier0Health; Count = $UniversalTier0Count }
    PreWin2000Access        = [ordered]@{ Status = $PreWin2000Status; Health = $PreWin2000Health; Count = $PreWin2000Count }
    Tier0GroupDelegation    = [ordered]@{ Status = $Tier0AclStatus; Health = $Tier0AclHealth; Count = $Tier0AclFindingsCount }
    NestingDepth            = [ordered]@{ Status = $NestingDepthStatus; Health = $NestingDepthHealth; Count = $NestingDepthExceedingCount; GroupsWithNesting = $NestingDepthFindingsCount }
    GroupCircularReference  = [ordered]@{ Status = $GroupCycleStatus; Health = $GroupCycleHealth; Count = $GroupCycleCount }
    ForeignSecurityPrincipals = [ordered]@{ Status = $FspStatus; Health = $FspHealth; Count = $FspFindingsCount; OrphanedCount = $FspOrphanedCount }
    GpoRestrictedGroups     = [ordered]@{ Status = $GpoRiskyStatus; Health = $GpoRiskyHealth; Count = $GpoRiskyCount }
    GpoCpassword            = [ordered]@{ Status = $GpoCpasswordStatus; Health = $GpoCpasswordHealth; Count = $GpoCpasswordCount }
    TrustHardening          = [ordered]@{ Status = $TrustRiskStatus; Health = $TrustRiskHealth; Count = $TrustRiskCount }
    SidHistory              = [ordered]@{ Status = $SidHistoryStatus; Health = $SidHistoryHealth; Count = $SidHistoryCount; Objects = $SidHistoryNames }
    AttackPathSummary       = [ordered]@{ Status = $AttackPathStatus; Health = $AttackPathHealth; Count = $AttackPathCount }
}

$GlobalADSecurityJson = (ConvertTo-Json -InputObject $GlobalADSecurityData -Depth 4 -Compress) -replace "</script>", "<\/script>"

# ---------------------------------------------------------------------------
# Security Baseline Compliance Score (Phase A) - a focused security-hardening
# percentage, deliberately separate from the operational Health Score above.
# It's the share of security-relevant checks (identity, Kerberos, ACL, DNS,
# GPO hardening - everything in $GlobalADSecurityData) currently passing.
# Checks that failed to collect (Health = Unknown/null) are excluded from
# both the numerator and denominator rather than counted as failures.
# ---------------------------------------------------------------------------
$BaselineCheckHealths = @($GlobalADSecurityData.Values | ForEach-Object { $_.Health } | Where-Object { $_ -and $_ -ne 'Unknown' })
$BaselineTotalChecks  = $BaselineCheckHealths.Count
$BaselinePassChecks   = @($BaselineCheckHealths | Where-Object { $_ -eq 'Good' }).Count
$BaselineCompliancePct = if ($BaselineTotalChecks -gt 0) { [Math]::Round(($BaselinePassChecks / $BaselineTotalChecks) * 100) } else { $null }
$BaselineComplianceColor = if ($null -eq $BaselineCompliancePct) { '#94a3b8' } elseif ($BaselineCompliancePct -ge 90) { '#16a34a' } elseif ($BaselineCompliancePct -ge 70) { '#a16207' } else { '#dc2626' }
$BaselineComplianceText = if ($null -eq $BaselineCompliancePct) { 'Not available' } else { "$BaselineCompliancePct% - $BaselinePassChecks/$BaselineTotalChecks controls passing" }

# ---------------------------------------------------------------------------
# Identity & Kerberos Security - consolidated NAMED findings list (Phase A).
# Unlike $GlobalADSecurityData above (which is counts/status for the Forest
# Overview tile), this is one row per individual account/principal so each
# can be reviewed and remediated specifically. Feeds the new "Identity &
# Kerberos Security" report.
# ---------------------------------------------------------------------------
$IdentityFindingsList = New-Object System.Collections.Generic.List[object]

foreach ($ku in $KerberoastDetails) {
    $sev = if ($ku.Privileged) { 'High' } elseif ($ku.WeakEncryption) { 'Medium' } else { 'Low' }
    $detail = @()
    if ($ku.Privileged)     { $detail += 'privileged account' }
    if ($ku.WeakEncryption) { $detail += 'weak (RC4) encryption supported' }
    if ($detail.Count -eq 0) { $detail += 'AES-capable, non-privileged' }
    $IdentityFindingsList.Add([ordered]@{ Severity = $sev; Category = 'Kerberoasting'; Account = $ku.SamAccountName; Detail = ($detail -join ', ') })
}

foreach ($pa in $PreAuthDisabled) {
    $IdentityFindingsList.Add([ordered]@{ Severity = 'High'; Category = 'AS-REP Roasting'; Account = $pa.SamAccountName; Detail = 'Kerberos pre-authentication disabled' })
}

foreach ($principal in $DCSyncUnexpected) {
    $IdentityFindingsList.Add([ordered]@{ Severity = 'High'; Category = 'DCSync Rights'; Account = $principal; Detail = 'Holds Replicating Directory Changes (All) on the domain - unexpected principal' })
}

foreach ($principal in $SDHolderUnexpected) {
    $IdentityFindingsList.Add([ordered]@{ Severity = 'High'; Category = 'AdminSDHolder Drift'; Account = $principal; Detail = 'Unexpected ACE on AdminSDHolder - propagates to all protected objects' })
}

foreach ($su in $StaleNames) {
    $sev = if ($su.Privileged) { 'High' } else { 'Medium' }
    $lastLogonStr = if ($su.LastLogonDate) { $su.LastLogonDate.ToString('yyyy-MM-dd') } else { 'never' }
    $IdentityFindingsList.Add([ordered]@{ Severity = $sev; Category = 'Stale Account'; Account = $su.SamAccountName; Detail = "No logon since $lastLogonStr$(if ($su.Privileged) { ' - privileged account' } else { '' })" })
}

foreach ($pne in $PwdNeverExpires) {
    $IdentityFindingsList.Add([ordered]@{ Severity = 'Medium'; Category = 'Password Never Expires'; Account = $pne.SamAccountName; Detail = 'Account password is set to never expire' })
}

if ($ConstrainedObjects) {
    foreach ($co in $ConstrainedObjects) {
        $IdentityFindingsList.Add([ordered]@{ Severity = 'Medium'; Category = 'Constrained Delegation'; Account = $co.Name; Detail = "Configured for constrained delegation ($($co.objectClass)) - confirm still required" })
    }
}

if ($DnsOpenTransferZones) {
    foreach ($z in $DnsOpenTransferZones) {
        $IdentityFindingsList.Add([ordered]@{ Severity = 'High'; Category = 'DNS Zone Transfer'; Account = $z.ZoneName; Detail = 'Zone allows transfer to any server - leaks namespace to unauthenticated queriers' })
    }
}

if ($DnsNoAgingZones) {
    foreach ($zn in $DnsNoAgingZones) {
        $IdentityFindingsList.Add([ordered]@{ Severity = 'Medium'; Category = 'DNS Aging/Scavenging'; Account = $zn; Detail = 'AD-integrated zone has aging/scavenging disabled - stale records accumulate' })
    }
}

if ($DnsInsecureUpdateZones) {
    foreach ($zu in $DnsInsecureUpdateZones) {
        $IdentityFindingsList.Add([ordered]@{ Severity = 'High'; Category = 'DNS Dynamic Update Risk'; Account = $zu.ZoneName; Detail = 'Zone allows nonsecure dynamic updates - any client can create or overwrite records (DNS spoofing / credential relay setup)' })
    }
}

foreach ($da in $DnsAclFindings) {
    $IdentityFindingsList.Add([ordered]@{ Severity = 'High'; Category = 'DNS Zone ACL'; Account = $da.Account; Detail = $da.Detail })
}

if ($DnsAdminsMembers) {
    foreach ($dm in $DnsAdminsMembers) {
        $IdentityFindingsList.Add([ordered]@{ Severity = 'Medium'; Category = 'DnsAdmins Membership'; Account = $dm.SamAccountName; Detail = 'Member of DnsAdmins - holds a documented path to SYSTEM on a DC via ServerLevelPluginDll if not tightly justified' })
    }
}

foreach ($rg in $GpoRiskyGpoNames) {
    $IdentityFindingsList.Add([ordered]@{ Severity = 'Medium'; Category = 'GPO Restricted Groups'; Account = $rg; Detail = 'Pushes membership into local Administrators via Restricted Groups - verify intentional' })
}

foreach ($cp in $GpoCpasswordGpoNames) {
    $IdentityFindingsList.Add([ordered]@{ Severity = 'High'; Category = 'GPP cpassword (MS14-025)'; Account = $cp; Detail = 'SYSVOL preference file contains a recoverable cpassword - credential is trivially decryptable' })
}

foreach ($tr in $TrustRiskFindings) {
    $IdentityFindingsList.Add([ordered]@{ Severity = 'Medium'; Category = 'Trust Hardening'; Account = $tr.Account; Detail = $tr.Detail })
}

foreach ($sh in $SidHistoryObjects) {
    $IdentityFindingsList.Add([ordered]@{ Severity = 'Medium'; Category = 'SID History'; Account = $sh.Name; Detail = "Carries a SID History entry ($($sh.objectClass)) - confirm this is from a recognized migration" })
}

foreach ($ap in $AttackPathFindings) {
    $IdentityFindingsList.Add([ordered]@{ Severity = 'High'; Category = 'Attack Path to Domain Admins'; Account = $ap.Account; Detail = $ap.Detail })
}

# --- Group & Privilege Architecture tab findings ---
foreach ($t0m in $Tier0MemberFindings) {
    $IdentityFindingsList.Add([ordered]@{ Severity = 'High'; Category = 'Tier-0 Group Membership'; Account = $t0m.Account; Detail = $t0m.Detail })
}
foreach ($t0c in $Tier0GroupCountFindings) {
    $IdentityFindingsList.Add([ordered]@{ Severity = 'Medium'; Category = 'Tier-0 Group Size'; Account = $t0c.Group; Detail = $t0c.Detail })
}
foreach ($ut0 in $UniversalTier0Findings) {
    $IdentityFindingsList.Add([ordered]@{ Severity = 'Medium'; Category = 'Universal Group Nesting'; Account = $ut0.Account; Detail = $ut0.Detail })
}
if ($PreWin2000Count -gt 0) {
    $IdentityFindingsList.Add([ordered]@{ Severity = 'High'; Category = 'Pre-Windows 2000 Access'; Account = ($PreWin2000RiskyNames -join ', '); Detail = 'Pre-Windows 2000 Compatible Access grants excessive read access to security descriptors and enables certain legacy enumeration paths' })
}
foreach ($t0a in $Tier0AclFindings) {
    $IdentityFindingsList.Add([ordered]@{ Severity = 'High'; Category = 'Group-Based Delegation'; Account = $t0a.Account; Detail = $t0a.Detail })
}
foreach ($nd in $NestingDepthFindings) {
    $IdentityFindingsList.Add([ordered]@{ Severity = 'Medium'; Category = 'Nesting Depth'; Account = $nd.Group; Detail = $nd.Detail })
}
if ($DetectedGroupCycles) {
    foreach ($cyc in $DetectedGroupCycles) {
        $IdentityFindingsList.Add([ordered]@{ Severity = 'High'; Category = 'Circular Group Reference'; Account = '(group chain)'; Detail = "Circular group membership: $cyc" })
    }
}
foreach ($fsp in $FspFindings) {
    $IdentityFindingsList.Add([ordered]@{ Severity = 'High'; Category = 'Foreign Security Principal'; Account = $fsp.Account; Detail = $fsp.Detail })
}

$IdentitySeverityRank = [ordered]@{ 'High' = 0; 'Medium' = 1; 'Low' = 2 }
# NOTE: $IdentityFindingsList is sorted and serialized to CSV further below,
# AFTER $DCInventory is collected - the Phase C Tier-0 contamination check
# (which needs $DCInventory) appends more rows to this same list first.

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

            # Merge the ICMP reachability check (already run from the
            # management host against each DC) into each inventory entry so
            # the Deep Insights "Connectivity Health" tile can show a simple
            # PingStatus alongside LDAP/LDAPS Pass-Fail.
            foreach ($dcEntry in $DCInventory) {
                $pingMatch = $DCPingResults | Where-Object { $_.DomainController -eq $dcEntry.ComputerName } | Select-Object -First 1
                $icmpStatus = if ($pingMatch) { $pingMatch.ICMPStatus } else { "Unknown" }
                $dcEntry | Add-Member -NotePropertyName 'ICMPStatus' -NotePropertyValue $icmpStatus -Force
            }
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

# ---------------------------------------------------------------------------
# Tier-0 Contamination Detection (Phase C) - deliberately scoped to data
# already collected from domain controllers (Recent Interactive Logons),
# NOT new remote scanning across the server/workstation estate. Flags any
# account that interactively logged onto a DC and is NOT a recognized
# Tier-0 administrator or a machine account - on a properly tiered domain,
# nobody else should ever be logging onto a DC directly.
#
# Caveat (by design, not a bug): Recent Interactive Logons only keeps the
# last 5 logon events per DC, so this is a sampling signal, not a
# guarantee every contamination event will be caught.
# ---------------------------------------------------------------------------
try {
    $Tier0RecognizedNames = New-Object System.Collections.Generic.HashSet[string]
    foreach ($grp in @('Domain Admins', 'Enterprise Admins', 'Schema Admins', 'Administrators')) {
        try {
            Get-ADGroupMember -Identity $grp -Recursive -ErrorAction Stop | Where-Object { $_.objectClass -eq 'user' } | ForEach-Object {
                [void]$Tier0RecognizedNames.Add($_.SamAccountName.ToLower())
            }
        } catch {}
    }

    $Tier0ContaminationFindings = New-Object System.Collections.Generic.List[object]
    foreach ($dcEntry in $DCInventory) {
        $recentLogons = @($dcEntry.DeepInsights.RecentLogons)
        foreach ($logon in $recentLogons) {
            if (-not $logon.Account) { continue }
            $bareName = ($logon.Account -split '\\')[-1]
            if ($bareName -match '\$$') { continue }  # machine account, not a human admin
            if ($Tier0RecognizedNames.Contains($bareName.ToLower())) { continue }

            $Tier0ContaminationFindings.Add([ordered]@{
                Account = $bareName
                DC      = $dcEntry.ComputerName
                Detail  = "$($logon.LogonType) logon at $($logon.Time) - account is not a recognized Tier 0 administrator"
            })
        }
    }

    $Tier0ContaminationCount = $Tier0ContaminationFindings.Count
}
catch {
    $Tier0ContaminationFindings = New-Object System.Collections.Generic.List[object]
    $Tier0ContaminationCount = $null
}

if ($Tier0ContaminationCount) {
    foreach ($t0 in $Tier0ContaminationFindings) {
        $IdentityFindingsList.Add([ordered]@{ Severity = 'High'; Category = 'Tier-0 Contamination'; Account = "$($t0.Account) on $($t0.DC)"; Detail = $t0.Detail })
    }
}
# NOTE: the $HealthScore deduction and $TopRisksList/$SecurityRisksList
# entries for this finding are added LATER in the script, right after
# those lists are actually declared (search for "Tier0Contamination score
# entry" below) - they cannot be added here because $TopRisksList and
# $SecurityRisksList do not exist yet at this point in the script (they
# are declared further down, alongside $ExecSecChecks). Only the
# $IdentityFindingsList contribution above stays here, since that list
# already exists this early and needs every contributor before
# $IdentityFindingsSorted is built just below.

# Sort and serialize $IdentityFindingsList NOW - after every Phase A/C
# contributor (including the Tier-0 check above, which needs $DCInventory)
# has had a chance to add its rows.
$IdentityFindingsSorted = @($IdentityFindingsList | Sort-Object { $IdentitySeverityRank[[string]$_.Severity] })

$IdentityFindingsCsvJs = ConvertTo-JavaScriptString (
    (
        $IdentityFindingsSorted |
        ForEach-Object { [PSCustomObject]@{ Severity = $_.Severity; Category = $_.Category; Account = $_.Account; Detail = $_.Detail } } |
        ConvertTo-Csv -NoTypeInformation
    ) -join "`n"
)

# NOTE: the Group & Privilege Architecture tab's table no longer reads from
# $IdentityFindingsSorted directly - it has its own per-group aggregated
# rows ($GpaGroupRows / $GpaGroupRowsCsvJs, built earlier alongside the
# Phase F1-F3 collectors), since per-account detail for these same
# categories already lives on the Identity Risk & Attack Surface tab below
# and showing it twice was redundant.

# ---------------------------------------------------------------------------
# Identity Risk & Attack Surface tab - remaining tile tooltips. Built here
# (after $IdentityFindingsSorted) so the "High Risk Identities" tile can
# include Tier-0 Contamination and Attack Path findings, which are only
# added to the list after $DCInventory is collected, further up.
# ---------------------------------------------------------------------------
$StaleAccountsTooltip = if (-not $StaleNames -or $StaleNames.Count -eq 0) {
    "<div class='tooltip-line'>No stale accounts found</div>"
} else {
    ($StaleNames | Sort-Object SamAccountName | ForEach-Object {
        $nameSafe = ConvertTo-HtmlSafe $_.SamAccountName
        $lastLogonStr = if ($_.LastLogonDate) { $_.LastLogonDate.ToString('dd-MMM-yyyy') } else { 'never' }
        $detail = ConvertTo-HtmlSafe "Last logon: $lastLogonStr$(if ($_.Privileged) { ' - privileged' } else { '' })"
        "<div class='tooltip-line'><b>$nameSafe</b><span>$detail</span></div>"
    }) -join ""
}

$RiskyServiceAccountsTooltip = if (-not $KerberoastDetails -or $KerberoastDetails.Count -eq 0) {
    "<div class='tooltip-line'>No SPN accounts found</div>"
} else {
    ($KerberoastDetails | Sort-Object SamAccountName | ForEach-Object {
        $nameSafe = ConvertTo-HtmlSafe $_.SamAccountName
        $bits = @('SPN')
        if ($_.WeakEncryption) { $bits += 'weak encryption' }
        if ($_.Privileged)     { $bits += 'privileged' }
        $detail = ConvertTo-HtmlSafe ($bits -join ' + ')
        "<div class='tooltip-line'><b>$nameSafe</b><span>$detail</span></div>"
    }) -join ""
}

# The Identity Risk & Attack Surface tab is specifically about WHO is risky
# (a person or service account/principal), not WHAT configuration is risky.
# $IdentityFindingsSorted also carries infrastructure-level findings (DNS
# zones, GPOs, trusts) that belong to their own reports (DNS Zones, Group
# Policies, the trust inventory) - those are excluded here so this tab
# doesn't show a DNS zone name or a GPO name in a "Name" column meant for
# accounts.
$IdentityRelevantCategories = @(
    'Kerberoasting', 'AS-REP Roasting', 'DCSync Rights', 'AdminSDHolder Drift',
    'Stale Account', 'Password Never Expires', 'Constrained Delegation',
    'Tier-0 Contamination', 'Attack Path to Domain Admins', 'SID History',
    'DNS Zone ACL', 'DnsAdmins Membership',
    'Tier-0 Group Membership', 'Group-Based Delegation', 'Foreign Security Principal'
)
$IdentityOnlyFindings = @($IdentityFindingsSorted | Where-Object { $IdentityRelevantCategories -contains $_.Category })

# High Risk Identities: distinct accounts (not rows) carrying a High
# severity finding anywhere in the identity-only findings above,
# deduplicated and combining every reason that account showed up for.
$HighRiskAccountMap = [ordered]@{}
foreach ($f in $IdentityOnlyFindings) {
    if ($f.Severity -ne 'High') { continue }
    $acct = $f.Account
    if (-not $HighRiskAccountMap.Contains($acct)) {
        $HighRiskAccountMap[$acct] = New-Object System.Collections.Generic.List[string]
    }
    $HighRiskAccountMap[$acct].Add($f.Category)
}
$HighRiskIdentitiesCount = $HighRiskAccountMap.Count
$HighRiskIdentitiesTooltip = if ($HighRiskIdentitiesCount -eq 0) {
    "<div class='tooltip-line'>No High severity identities found</div>"
} else {
    ($HighRiskAccountMap.GetEnumerator() | Sort-Object Key | ForEach-Object {
        $nameSafe = ConvertTo-HtmlSafe $_.Key
        $detail = ConvertTo-HtmlSafe ($_.Value -join ', ')
        "<div class='tooltip-line'><b>$nameSafe</b><span>$detail</span></div>"
    }) -join ""
}

# ---------------------------------------------------------------------------
# Critical & High risk users table (Identity Risk & Attack Surface tab) -
# one row per ACCOUNT (not per finding), combining every finding that
# account has into a single short "Why" phrase. Risk tier here is a 4-level
# display refinement of the 3-tier (High/Medium/Low) severity already used
# elsewhere: a High-severity account that is ALSO a privileged group member
# displays as "Critical"; a High-severity account that is not privileged
# displays as "High"; Medium stays "Medium". Privileged-only accounts with
# no other finding still display as "Critical". Type is "Service" if the
# account ever appears under the Kerberoasting category (it has an SPN),
# otherwise "User".
#
# Each account is ALSO tagged with which of the 4 tiles it belongs to
# (stale / privileged / service / highrisk), tracked independently of the
# display Risk tier above - this is what makes the tiles clickable: e.g. an
# SPN account that's AES-capable and non-privileged is "Low" risk and
# wouldn't normally appear in the default table, but it still needs to show
# up when the "Risky Service Accounts" tile is clicked, since it's counted
# in that tile's number.
# ---------------------------------------------------------------------------
$PrivAccountNamesSet = @{}
if ($PrivAccountsRaw) { foreach ($k in $PrivAccountsRaw.Keys) { $PrivAccountNamesSet[$k] = $true } }

# LastLogonDate lookup, gathered from whichever collector already happened
# to fetch it for that account (Stale, AS-REP, Kerberoasting, Privileged) -
# no extra AD round-trips beyond what those collectors already did.
$AccountLastLogonMap = @{}
foreach ($su in $StaleNames)        { if ($su.SamAccountName -and -not $AccountLastLogonMap.ContainsKey($su.SamAccountName))        { $AccountLastLogonMap[$su.SamAccountName]        = $su.LastLogonDate } }
foreach ($pa2 in $PreAuthDisabled)  { if ($pa2.SamAccountName -and -not $AccountLastLogonMap.ContainsKey($pa2.SamAccountName))      { $AccountLastLogonMap[$pa2.SamAccountName]       = $pa2.LastLogonDate } }
foreach ($ku2 in $KerberoastDetails){ if ($ku2.SamAccountName -and -not $AccountLastLogonMap.ContainsKey($ku2.SamAccountName))      { $AccountLastLogonMap[$ku2.SamAccountName]       = $ku2.LastLogonDate } }
if ($PrivAccountsRaw) {
    foreach ($pv in $PrivAccountsRaw.Values) { if ($pv.SamAccountName -and -not $AccountLastLogonMap.ContainsKey($pv.SamAccountName)) { $AccountLastLogonMap[$pv.SamAccountName] = $pv.LastLogonDate } }
}

$AllAccountsMap = [ordered]@{}
$DisplayRank = @{ 'Critical' = 0; 'High' = 1; 'Medium' = 2; 'Low' = 3 }
$SevRawRank  = @{ 'High' = 0; 'Medium' = 1; 'Low' = 2 }

function Ensure-IdentityAccountEntry($acct) {
    if (-not $AllAccountsMap.Contains($acct)) {
        $AllAccountsMap[$acct] = [ordered]@{
            Type      = 'User'
            Why       = New-Object System.Collections.Generic.List[string]
            Tags      = New-Object System.Collections.Generic.List[string]
            Sev       = 'Low'
            PathSteps = ''
        }
    }
}

foreach ($f in $IdentityOnlyFindings) {
    $acct = $f.Account
    Ensure-IdentityAccountEntry $acct
    $entry = $AllAccountsMap[$acct]

    $shortWhy = switch -Wildcard ($f.Category) {
        'Kerberoasting'                  { $entry.Type = 'Service'; if (-not $entry.Tags.Contains('service')) { $entry.Tags.Add('service') }; if ($f.Detail -match 'privileged') { 'SPN + Privileged' } else { 'SPN' } }
        'AS-REP Roasting'                { 'AS-REP roastable' }
        'DCSync Rights'                  { 'DCSync rights' }
        'AdminSDHolder Drift'            { 'AdminSDHolder ACE' }
        'Stale Account'                  { if (-not $entry.Tags.Contains('stale')) { $entry.Tags.Add('stale') }; 'Stale' }
        'Tier-0 Contamination'           { 'Logged onto a DC' }
        'Attack Path to Domain Admins'   { 'Path to Domain Admins' }
        'GPP cpassword*'                 { 'Recoverable GPP password' }
        'Password Never Expires'        { 'Password never expires' }
        'Constrained Delegation'        { 'Constrained delegation' }
        'DNS Zone ACL'                   { 'DNS zone ACL right' }
        'DnsAdmins Membership'           { 'Member of DnsAdmins' }
        'Tier-0 Group Membership'        { 'Tier-0 group risk' }
        'Group-Based Delegation'         { 'Delegated Tier-0 right' }
        'Foreign Security Principal'     { 'Foreign security principal' }
        default                          { $f.Category }
    }
    if (-not $entry.Why.Contains($shortWhy)) { $entry.Why.Add($shortWhy) }
    if ($SevRawRank[$f.Severity] -lt $SevRawRank[$entry.Sev]) { $entry.Sev = $f.Severity }

    # Carry the hop-by-hop chain through to the expandable detail panel so
    # it can render the walkthrough view instead of just the flat Detail
    # sentence. First matching path wins if an account somehow has more than
    # one (mirrors how Why already collapses duplicates to one label).
    if ($f.Category -eq 'Attack Path to Domain Admins' -and -not $entry.PathSteps -and $AttackPathStepsByAccount.Contains($acct)) {
        $entry.PathSteps = $AttackPathStepsByAccount[$acct]
    }
}

# Fold in Privileged Accounts even when they have no other finding attached
# - group membership itself is the reason they belong in this table.
foreach ($pa in @($PrivAccountsRaw.Values)) {
    $acct = $pa.SamAccountName
    Ensure-IdentityAccountEntry $acct
    $entry = $AllAccountsMap[$acct]
    if (-not $entry.Tags.Contains('privileged')) { $entry.Tags.Add('privileged') }
    # Names the SPECIFIC group(s) (e.g. "Member of Domain Admins") rather
    # than a generic "Privileged" label, so the expanded panel can show
    # exactly which group to review.
    $privGroupsWhy = "Member of $($pa.Groups -join ', ')"
    if (-not $entry.Why.Contains($privGroupsWhy)) { $entry.Why.Add($privGroupsWhy) }
}

# Final display Risk tier, the "highrisk" tag (Critical or High), and the
# LastLogon string for the expandable detail panel.
foreach ($acct in @($AllAccountsMap.Keys)) {
    $entry  = $AllAccountsMap[$acct]
    $isPriv = $entry.Tags.Contains('privileged')
    $entry.Risk = if ($entry.Sev -eq 'High') { if ($isPriv) { 'Critical' } else { 'High' } }
                  elseif ($entry.Sev -eq 'Medium') { 'Medium' }
                  else { if ($isPriv) { 'Critical' } else { 'Low' } }
    if ($entry.Risk -in @('Critical', 'High') -and -not $entry.Tags.Contains('highrisk')) { $entry.Tags.Add('highrisk') }

    $llDate = $AccountLastLogonMap[$acct]
    $entry.LastLogon = if ($AccountLastLogonMap.ContainsKey($acct)) {
        if ($null -eq $llDate) { 'Never' } else { $llDate.ToString('dd-MMM-yyyy') }
    } else { 'Unknown' }
}

$AllAccountsRows = @($AllAccountsMap.Keys | ForEach-Object {
    $entry = $AllAccountsMap[$_]
    [PSCustomObject]@{
        Name      = $_
        Type      = $entry.Type
        Risk      = $entry.Risk
        Why       = ($entry.Why -join ' + ')
        Tags      = ($entry.Tags -join ',')
        LastLogon = $entry.LastLogon
        PathSteps = $entry.PathSteps
    }
})

# Default view (matches what was already shown): Critical/High/Medium only.
$CriticalUsersRows = @($AllAccountsRows | Where-Object { $_.Risk -ne 'Low' } | Sort-Object { $DisplayRank[$_.Risk] }, Name)

# Full dataset (including Low) embedded for client-side tile filtering.
$IdentityUsersCsvJs = ConvertTo-JavaScriptString (
    (
        $AllAccountsRows | Sort-Object { $DisplayRank[$_.Risk] }, Name |
        ConvertTo-Csv -NoTypeInformation
    ) -join "`n"
)

# Uses the exact same .rpt-b badge classes as the Reports tab (rpt-crit,
# rpt-medium), scoped under #identityUsersTable instead of #rptPanel since
# this lives on a different tab. "Critical" gets a new id-critical modifier
# (a deeper red than High) - Reports itself has no Critical tier today, so
# this is additive, not a re-theme of existing reports.
$RiskBadgeClass = @{ 'Critical' = 'rpt-b id-critical'; 'High' = 'rpt-b rpt-crit'; 'Medium' = 'rpt-b rpt-medium' }

$CriticalUsersTableRowsHtml = if ($CriticalUsersRows.Count -eq 0) {
    '<tr><td colspan="4" style="text-align:center;padding:20px;color:#94a3b8;font-style:italic;">No critical or high risk identities detected</td></tr>'
} else {
    ($CriticalUsersRows | ForEach-Object {
        $badgeCls = $RiskBadgeClass[$_.Risk]
        $nameSafe = ConvertTo-HtmlSafe $_.Name
        $whySafe  = ConvertTo-HtmlSafe $_.Why
        "<tr>" +
        "<td class=`"rpt-dc`">$nameSafe</td>" +
        "<td>$($_.Type)</td>" +
        "<td><span class=`"$badgeCls`">$($_.Risk)</span></td>" +
        "<td>$whySafe</td>" +
        "</tr>"
    }) -join ""
}

# Use -InputObject (not the pipeline) so ConvertTo-Json always emits a JSON
# array, even when there is exactly one Domain Controller. Piping a
# single-element array to ConvertTo-Json unwraps it and produces a bare
# JSON object ("{...}") instead of "[{...}]", which breaks the dcListData.
# filter/find/forEach calls in the DC Health, Security Posture, and Deep
# Insights tabs (all blank) on single-DC environments.
$DCListJson = (ConvertTo-Json -InputObject $DCListData -Depth 6 -Compress) -replace "</script>", "<\/script>"

if ($DCInventory.Count -gt 0) {
    $DCInventoryJson = (ConvertTo-Json -InputObject $DCInventory -Depth 8 -Compress) -replace "</script>", "<\/script>"
}
else {
    $DCInventoryJson = "[]"
}

# ---------------------------------------------------------------------------
# Executive Summary - overall health score, top risks, and trend history
#
# The score is computed entirely server-side from data already collected
# above (Global AD Security health checks, DC reachability, GPO health).
# This keeps it consistent and avoids duplicating the client-side JS
# classifiers used by the Security Posture / Deep Insights tabs. The gauge
# and trend chart are rendered as plain inline SVG (no charting library),
# so the report has zero external dependencies and works fully offline -
# important since most AD environments have no internet access.
# ---------------------------------------------------------------------------

$HealthScore = 100
$TopRisksList      = New-Object System.Collections.Generic.List[object]
# Flat list of every INDIVIDUAL finding (one row per DC per risk, or one row
# with DC="Forest-wide" for checks that aren't tied to a specific DC). Feeds
# the "Security Risks Need Attention" report - scales to any DC count,
# unlike the Top Risks tile which only shows a top-6 summary.
$SecurityRisksList = New-Object System.Collections.Generic.List[object]

# Builds the Top Risks tile text for a DC-scoped finding: names up to 3
# affected DCs, then "+N more" with a pointer to the full report - so the
# tile stays readable whether 1 DC or 200 DCs are affected.
function Get-ExecDcNamesText {
    param([string[]]$Names)
    if (-not $Names -or $Names.Count -eq 0) { return "" }
    $shown = ($Names | Select-Object -First 3) -join ", "
    $extra = $Names.Count - 3
    if ($extra -gt 0) {
        return "$shown +$extra more - see the Security Risks Need Attention report"
    }
    return $shown
}

$ExecSecChecks = @(
    [ordered]@{ Health = $SysvolMigHealth;        Detail = "SYSVOL migration is not in the fully-migrated, robust state. Logon scripts and Group Policy file replication may be inconsistent." }
    [ordered]@{ Health = $AdminRID500Health;      Detail = "The built-in Administrator account (RID 500) is not both renamed and disabled, leaving a well-known high-value target enabled." }
    [ordered]@{ Health = $GuestHealth;            Detail = "The built-in Guest account is enabled, allowing a well-known low-privilege account to remain usable." }
    [ordered]@{ Health = $ProtectedUsersHealth;   Detail = "$ExposedAdminCount privileged admin account(s) are not members of the Protected Users group, leaving them exposed to credential theft techniques." }
    [ordered]@{ Health = $SPNHealth;              Detail = "$SPNDupCount duplicate Service Principal Name(s) found, which can cause Kerberos authentication failures." }
    [ordered]@{ Health = $PreAuthHealth;          Detail = "$PreAuthCount account(s) have Kerberos pre-authentication disabled - a known AS-REP Roasting attack vector." }
    [ordered]@{ Health = $UnconstrainedHealth;    Detail = "$UnconstrainedCount object(s) are configured for unconstrained Kerberos delegation - a high-value lateral movement risk." }
    [ordered]@{ Health = $PwdNeverExpiresHealth;  Detail = "$PwdNeverExpiresCount account(s) have passwords set to never expire." }
    [ordered]@{ Health = $KerberoastHealth;       Detail = "$KerberoastCount account(s) have an SPN set (Kerberoastable) - $KerberoastWeakCount use weak encryption, $KerberoastPrivCount are privileged. See the Identity & Kerberos Security report for named accounts." }
    [ordered]@{ Health = $DCSyncHealth;           Detail = "$DCSyncCount unexpected principal(s) hold DCSync rights (Replicating Directory Changes) on the domain - this allows full credential extraction." }
    [ordered]@{ Health = $SDHolderHealth;         Detail = "$SDHolderCount unexpected principal(s) found on the AdminSDHolder ACL - this propagates to every protected (Tier 0) object." }
    [ordered]@{ Health = $StaleHealth;            Detail = "$StaleCount account(s) have not logged on in 90+ days - $StalePrivCount are privileged. See the Identity & Kerberos Security report for named accounts." }
    [ordered]@{ Health = $PwdAgeHealth;           Detail = "$PwdAgeOver1YearCount account(s) have not rotated their password in 365+ days." }
    [ordered]@{ Health = $ConstrainedHealth;      Detail = "$ConstrainedCount object(s) are configured for constrained Kerberos delegation - review to confirm each is still required." }
    [ordered]@{ Health = $DnsTransferHealth;      Detail = "$DnsOpenTransferCount DNS zone(s) allow zone transfer to any server, which can leak the entire DNS namespace to an attacker." }
    [ordered]@{ Health = $DnsAgingHealth;         Detail = "$DnsNoAgingCount AD-integrated DNS zone(s) have aging/scavenging disabled, allowing stale records to accumulate." }
    [ordered]@{ Health = $GpoRiskyHealth;         Detail = "$GpoRiskyCount GPO(s) use Restricted Groups to push membership into the local Administrators group - verify each is intentional." }
    [ordered]@{ Health = $GpoCpasswordHealth;     Detail = "$GpoCpasswordCount GPO(s) contain a recoverable Group Policy Preferences cpassword (MS14-025) - credentials are trivially decryptable." }
    [ordered]@{ Health = $TrustRiskHealth;        Detail = "$TrustRiskCount trust(s) are missing SID filtering quarantine or selective authentication hardening." }
    [ordered]@{ Health = $SidHistoryHealth;       Detail = "$SidHistoryCount object(s) carry a SID History entry - confirm each is from a recognized domain migration, not an injection attempt." }
    [ordered]@{ Health = $AttackPathHealth;       Detail = "$AttackPathCount escalation path finding(s) into Domain Admins/Enterprise Admins (nested groups or one-hop ACL rights) - see the Identity & Kerberos Security report." }
    [ordered]@{ Health = $DnsInsecureUpdateHealth; Detail = "$DnsInsecureUpdateCount DNS zone(s) allow nonsecure dynamic updates - any client can create or overwrite records." }
    [ordered]@{ Health = $DnsAclHealth;           Detail = "$DnsAclFindingsCount unexpected principal(s) hold modify rights on a DNS zone object - see the Identity & Kerberos Security report for named accounts." }
    [ordered]@{ Health = $DnsAdminsHealth;        Detail = "$DnsAdminsCount named member(s) of DnsAdmins - a documented path to SYSTEM on a DC via ServerLevelPluginDll if not tightly justified." }
    [ordered]@{ Health = $Tier0GroupHealth;       Detail = "$Tier0MemberFindingsCount risky Tier-0 group member(s) (disabled, stale, SPN, or token-size risk) across Domain Admins/Enterprise Admins/Schema Admins/Administrators/Backup Operators/Account Operators/Print Operators." }
    [ordered]@{ Health = $Tier0AclHealth;         Detail = "$Tier0AclFindingsCount unexpected delegated right(s) on a Tier-0 group object, including any Reset Password right - see the Identity & Kerberos Security report for named accounts." }
    [ordered]@{ Health = $GroupCycleHealth;       Detail = "$GroupCycleCount circular group membership chain(s) detected - a group nests back into itself through one or more intermediate groups." }
    [ordered]@{ Health = $FspHealth;              Detail = "$FspFindingsCount foreign security principal(s) found inside a Tier-0 group, $FspOrphanedCount orphaned FSP entr(y/ies) overall." }
    [ordered]@{ Health = $PreWin2000Health;       Detail = "Pre-Windows 2000 Compatible Access contains: $($PreWin2000RiskyNames -join ', ')." }
)

foreach ($ExecChk in $ExecSecChecks) {
    switch ($ExecChk.Health) {
        'Critical' {
            $HealthScore -= 8
            $TopRisksList.Add([ordered]@{ Severity = 'High'; Text = $ExecChk.Detail })
            $SecurityRisksList.Add([ordered]@{ Severity = 'High'; Risk = $ExecChk.Detail; DC = 'Forest-wide'; Source = 'Global AD Security' })
        }
        'Warning'  {
            $HealthScore -= 3
            $TopRisksList.Add([ordered]@{ Severity = 'Medium'; Text = $ExecChk.Detail })
            $SecurityRisksList.Add([ordered]@{ Severity = 'Medium'; Risk = $ExecChk.Detail; DC = 'Forest-wide'; Source = 'Global AD Security' })
        }
        'Unknown'  { $HealthScore -= 1 }
        default    { }
    }
}

# --- Tier0Contamination score entry - the actual finding/IdentityFindingsList
# row for this check was already added much earlier (right after
# $Tier0ContaminationFindings was computed, before $IdentityFindingsSorted
# was built) - this is just the score deduction and Top Risks/Security
# Risks entries, which had to wait until $TopRisksList/$SecurityRisksList
# existed. $Tier0ContaminationCount/$Tier0ContaminationFindings are
# script-scoped variables already set earlier and are still valid here. ---
if ($Tier0ContaminationCount) {
    $HealthScore -= [Math]::Min($Tier0ContaminationCount * 6, 18)
    # NOTE: $Tier0ContaminationFindings.Add() stores [ordered]@{} hashtables,
    # not PSCustomObjects - Select-Object's -ExpandProperty cannot expand a
    # property from a Hashtable (it only works via .NET/PSObject reflection,
    # which a Hashtable's keys are not exposed through), so this uses
    # ForEach-Object instead, which DOES support a hashtable's dot-property
    # access correctly.
    $Tier0AccountNames = (($Tier0ContaminationFindings | Select-Object -First 3) | ForEach-Object { $_.Account }) -join ', '
    $Tier0ExtraCount = $Tier0ContaminationCount - 3
    $Tier0NamesText = if ($Tier0ExtraCount -gt 0) { "$Tier0AccountNames +$Tier0ExtraCount more" } else { $Tier0AccountNames }
    $TopRisksList.Add([ordered]@{ Severity = 'High'; Text = "Non-privileged account logon(s) detected on a domain controller: $Tier0NamesText - see the Identity & Kerberos Security report." })
    $SecurityRisksList.Add([ordered]@{ Severity = 'High'; Risk = 'Non-privileged account logon(s) detected on a domain controller'; DC = 'Forest-wide'; Source = 'Tier-0 Contamination' })
}

# DC ICMP reachability - named per failed DC instead of "X of Y", so the
# tile stays actionable at any DC count.
$DCPingFailedNames = @($DCPingResults | Where-Object { $_.ICMPStatus -ne "Reachable" } | ForEach-Object { $_.DomainController })

if ($DCPingFailedNames.Count -gt 0) {
    $HealthScore -= [Math]::Min($DCPingFailedNames.Count * 10, 30)
    $TopRisksList.Add([ordered]@{ Severity = 'High'; Text = "Not responding to ICMP ping: $(Get-ExecDcNamesText $DCPingFailedNames)" })
    foreach ($pingDc in $DCPingFailedNames) {
        $SecurityRisksList.Add([ordered]@{ Severity = 'High'; Risk = 'Domain controller is not responding to ICMP ping'; DC = $pingDc; Source = 'Connectivity' })
    }
}

if (($UnlinkedGPOCount -is [int]) -and ($UnlinkedGPOCount -gt 0)) {
    $HealthScore -= [Math]::Min($UnlinkedGPOCount, 10)
    $TopRisksList.Add([ordered]@{ Severity = 'Medium'; Text = "$UnlinkedGPOCount Group Policy Object(s) are not linked anywhere and are not being applied." })
    $SecurityRisksList.Add([ordered]@{ Severity = 'Medium'; Risk = "$UnlinkedGPOCount Group Policy Object(s) are not linked anywhere and are not being applied"; DC = 'Forest-wide'; Source = 'Group Policy' })
}

# AD replication health - reuses the same per-DC repadmin data already
# collected by Get-DCInventory.ps1 for the "AD Replication Partners" report,
# so this stays consistent with that report instead of re-querying.
$ReplTotalLinks  = 0
$ReplFailedLinks = 0
$ReplFailedDcNames = New-Object System.Collections.Generic.List[string]

foreach ($ExecDcEntry in $DCInventory) {
    $ExecReplArr = @($ExecDcEntry.Replication)
    $ExecDcHasReplFail = $false

    foreach ($ExecRepl in $ExecReplArr) {
        if (-not $ExecRepl) { continue }
        $ReplTotalLinks++

        $ExecConsecFail = 0
        $ExecLastResult = 0
        try { if ($ExecRepl.ConsecutiveFailures) { $ExecConsecFail = [int]$ExecRepl.ConsecutiveFailures } } catch {}
        try { if ($ExecRepl.LastResult)          { $ExecLastResult = [int]$ExecRepl.LastResult } }          catch {}

        if ($ExecConsecFail -gt 0 -or $ExecLastResult -ne 0) { $ReplFailedLinks++; $ExecDcHasReplFail = $true }
    }

    if ($ExecDcHasReplFail) { $ReplFailedDcNames.Add($ExecDcEntry.ComputerName) }
}

$ReplHealthKnown = $ReplTotalLinks -gt 0

if ($ReplHealthKnown -and $ReplFailedLinks -gt 0) {
    $HealthScore -= [Math]::Min($ReplFailedLinks * 6, 20)
    $TopRisksList.Add([ordered]@{ Severity = 'High'; Text = "AD replication failing on: $(Get-ExecDcNamesText $ReplFailedDcNames) - see the AD Replication Partners report for details." })
    foreach ($replDc in $ReplFailedDcNames) {
        $SecurityRisksList.Add([ordered]@{ Severity = 'High'; Risk = 'AD replication link(s) are failing'; DC = $replDc; Source = 'AD Replication' })
    }
}

# DNS health - reuses the same per-DC SRV record check already collected
# by Get-DCInventory.ps1 for the Deep Insights "DNS Health" tile.
$DnsDcsChecked    = 0
$DnsDcsSrvMissingNames = New-Object System.Collections.Generic.List[string]

foreach ($ExecDnsEntry in $DCInventory) {
    if (-not $ExecDnsEntry.DeepInsights) { continue }
    if (-not $ExecDnsEntry.DeepInsights.DNSHealth) { continue }
    $DnsDcsChecked++

    $ExecSrvOk = $ExecDnsEntry.DeepInsights.DNSHealth.SRVRecordRegistered
    if ($ExecSrvOk -eq $false) { $DnsDcsSrvMissingNames.Add($ExecDnsEntry.ComputerName) }
}

$DnsHealthKnown = $DnsDcsChecked -gt 0
# Kept as an integer count too - the Executive Summary DNS gauge/tile
# calculations further down reference $DnsDcsSrvMissing as a count.
$DnsDcsSrvMissing = $DnsDcsSrvMissingNames.Count

if ($DnsHealthKnown -and $DnsDcsSrvMissingNames.Count -gt 0) {
    $HealthScore -= [Math]::Min($DnsDcsSrvMissingNames.Count * 10, 20)
    $TopRisksList.Add([ordered]@{ Severity = 'High'; Text = "Missing LDAP SRV record in DNS on: $(Get-ExecDcNamesText $DnsDcsSrvMissingNames) - clients and other DCs may be unable to locate them." })
    foreach ($dnsDc in $DnsDcsSrvMissingNames) {
        $SecurityRisksList.Add([ordered]@{ Severity = 'High'; Risk = 'Missing LDAP SRV record in DNS'; DC = $dnsDc; Source = 'DNS' })
    }
}

# ---------------------------------------------------------------------------
# Consolidated CRITICAL findings from Security Posture, Deep Insights, and
# DC Health Check - per user request, only critical/high-severity checks
# are folded into the Executive Summary score (not every warning-level
# item from those tabs, to keep the risk list readable). Reads the same
# $DCInventory data and applies the same thresholds as the client-side JS
# classifiers in those tabs, aggregated by DC COUNT rather than one line
# per DC, so the score reflects dashboard-wide health, not just the 8
# original identity/security checks.
# ---------------------------------------------------------------------------
$CrSmb1Dcs          = New-Object System.Collections.Generic.List[string]
$CrTlsLegacyDcs     = New-Object System.Collections.Generic.List[string]
$CrLdapIntegrityDcs = New-Object System.Collections.Generic.List[string]
$CrLdapChannelDcs   = New-Object System.Collections.Generic.List[string]
$CrFwDomainDcs      = New-Object System.Collections.Generic.List[string]
$CrNetbiosDcs       = New-Object System.Collections.Generic.List[string]
$CrBackupDcs        = New-Object System.Collections.Generic.List[string]
$CrLingeringDcs     = New-Object System.Collections.Generic.List[string]
$CrUsnRollbackDcs   = New-Object System.Collections.Generic.List[string]
$CrDfsrDcs          = New-Object System.Collections.Generic.List[string]
$CrSvcDownDcs       = New-Object System.Collections.Generic.List[string]
$CrDcdiagDcs        = New-Object System.Collections.Generic.List[string]
$CrDcCheckedCount   = 0

foreach ($ExecCrDc in $DCInventory) {
    $CrDcCheckedCount++
    $crDcName = $ExecCrDc.ComputerName
    $crSp = $ExecCrDc.SecurityPosture

    if ($crSp) {
        if ($crSp.SMB -and $crSp.SMB.SMB1Enabled -eq $true) { $CrSmb1Dcs.Add($crDcName) }

        if ($crSp.TLS) {
            $crLegacyTls = @($crSp.TLS | Where-Object {
                $crProto = ([string]$_.Protocol).ToLower()
                ($crProto -match 'ssl' -or $crProto -eq 'tls 1.0' -or $crProto -eq 'tls1.0') -and ($_.Enabled -eq $true -or $_.Enabled -eq 1)
            })
            if ($crLegacyTls.Count -gt 0) { $CrTlsLegacyDcs.Add($crDcName) }
        }

        if ($crSp.LDAPSigning) {
            if ($crSp.LDAPSigning.LDAPServerIntegrity -ne 2)        { $CrLdapIntegrityDcs.Add($crDcName) }
            if ($crSp.LDAPSigning.LdapEnforceChannelBinding -ne 2)  { $CrLdapChannelDcs.Add($crDcName) }
        }

        if ($crSp.Firewall -and $crSp.Firewall.Profiles) {
            $crDomainProfile = $crSp.Firewall.Profiles | Where-Object { $_.Name -eq 'Domain' } | Select-Object -First 1
            if ($crDomainProfile -and -not $crDomainProfile.Enabled) { $CrFwDomainDcs.Add($crDcName) }
        }

        if ($crSp.NetBIOS -and $crSp.NetBIOS.Adapters) {
            $crNbEnabled = @($crSp.NetBIOS.Adapters | Where-Object { ([string]$_.Setting).ToLower() -eq 'enabled' })
            if ($crNbEnabled.Count -gt 0) { $CrNetbiosDcs.Add($crDcName) }
        }
    }

    if ($ExecCrDc.ADBackup -and ([string]$ExecCrDc.ADBackup.Status).ToLower() -eq 'critical') { $CrBackupDcs.Add($crDcName) }

    $crDi = $ExecCrDc.DeepInsights
    if ($crDi) {
        if ($crDi.LingeringObjectEvents -and @($crDi.LingeringObjectEvents).Count -gt 0) { $CrLingeringDcs.Add($crDcName) }
        if ($crDi.USNRollbackEvents -and @($crDi.USNRollbackEvents).Count -gt 0)         { $CrUsnRollbackDcs.Add($crDcName) }

        $crDfsrCritical = $false
        if ($crDi.DFSRBacklog) {
            $crErrState = @($crDi.DFSRBacklog | Where-Object { $_.State -eq 'In Error' })
            if ($crErrState.Count -gt 0) { $crDfsrCritical = $true }
        }
        if ($crDi.SharesHealth -and $crDi.SharesHealth.NETLOGON -and $crDi.SharesHealth.NETLOGON.Status -ne 'Shared') { $crDfsrCritical = $true }
        if ($null -ne $crDi.MaxOfflineTimeInDays -and $crDi.MaxOfflineTimeInDays -lt 30) { $crDfsrCritical = $true }
        if ($crDfsrCritical) { $CrDfsrDcs.Add($crDcName) }

        if ($crDi.DCDiag) {
            $crDcdiagFail = $false
            foreach ($crTestKey in @('NetLogons', 'Replications', 'Advertising', 'FSMOCheck')) {
                $crVal = $crDi.DCDiag.$crTestKey
                if ($crVal -and ([string]$crVal).ToLower() -ne 'passed') { $crDcdiagFail = $true }
            }
            if ($crDcdiagFail) { $CrDcdiagDcs.Add($crDcName) }
        }
    }

    if ($ExecCrDc.Services) {
        $crSvcDown = $false
        foreach ($crSvcName in @('DNS', 'Netlogon', 'NTDS', 'Kdc', 'DFSR', 'W32Time')) {
            $crSvc = $ExecCrDc.Services | Where-Object { $_.Name -eq $crSvcName -or $_.ServiceName -eq $crSvcName } | Select-Object -First 1
            $crSvcStatus = if ($crSvc) { if ($crSvc.Status) { $crSvc.Status } else { $crSvc.State } } else { $null }
            if ($crSvcStatus -and ([string]$crSvcStatus).ToLower() -ne 'running') { $crSvcDown = $true }
        }
        if ($crSvcDown) { $CrSvcDownDcs.Add($crDcName) }
    }
}

$ExecConsolidatedChecks = @(
    @{ Dcs = $CrSmb1Dcs;          Risk = "SMBv1 enabled - a legacy, insecure protocol";                                Source = "Security Posture" }
    @{ Dcs = $CrTlsLegacyDcs;     Risk = "Legacy SSL/TLS 1.0 enabled";                                                 Source = "Security Posture" }
    @{ Dcs = $CrLdapIntegrityDcs; Risk = "LDAP server signing not required";                                           Source = "Security Posture" }
    @{ Dcs = $CrLdapChannelDcs;   Risk = "LDAP channel binding not enforced";                                          Source = "Security Posture" }
    @{ Dcs = $CrFwDomainDcs;      Risk = "Windows Firewall Domain profile turned off";                                 Source = "Security Posture" }
    @{ Dcs = $CrNetbiosDcs;       Risk = "NetBIOS over TCP/IP enabled";                                                Source = "Security Posture" }
    @{ Dcs = $CrBackupDcs;        Risk = "AD system state backup critically overdue or missing";                      Source = "Deep Insights" }
    @{ Dcs = $CrLingeringDcs;     Risk = "Lingering object events detected (1388/1988)";                               Source = "Deep Insights" }
    @{ Dcs = $CrUsnRollbackDcs;   Risk = "USN rollback event detected (2095) - possible AD split-brain replication";   Source = "Deep Insights" }
    @{ Dcs = $CrDfsrDcs;          Risk = "SYSVOL/DFSR replication in an error state";                                  Source = "Deep Insights" }
    @{ Dcs = $CrSvcDownDcs;       Risk = "Core AD service (DNS/Netlogon/NTDS/KDC/DFSR/W32Time) not running";           Source = "DC Health Check" }
    @{ Dcs = $CrDcdiagDcs;        Risk = "Failing one or more dcdiag tests (NetLogons/Replications/Advertising/FSMO)"; Source = "DC Health Check" }
)

foreach ($ExecCrChk in $ExecConsolidatedChecks) {
    if ($ExecCrChk.Dcs.Count -gt 0) {
        $HealthScore -= [Math]::Min($ExecCrChk.Dcs.Count * 5, 15)
        $TopRisksList.Add([ordered]@{ Severity = 'High'; Text = "$($ExecCrChk.Risk) on: $(Get-ExecDcNamesText $ExecCrChk.Dcs)" })
        foreach ($crDcName2 in $ExecCrChk.Dcs) {
            $SecurityRisksList.Add([ordered]@{ Severity = 'High'; Risk = $ExecCrChk.Risk; DC = $crDcName2; Source = $ExecCrChk.Source })
        }
    }
}

if ($HealthScore -lt 0)   { $HealthScore = 0 }
if ($HealthScore -gt 100) { $HealthScore = 100 }
$HealthScore = [int][Math]::Round($HealthScore)

$HealthScoreLabel = if ($HealthScore -ge 90) { "Excellent" }
                    elseif ($HealthScore -ge 75) { "Good" }
                    elseif ($HealthScore -ge 50) { "Fair" }
                    else { "Needs Attention" }

$HealthScoreColor = if ($HealthScore -ge 90) { "#16a34a" }
                     elseif ($HealthScore -ge 75) { "#65a30d" }
                     elseif ($HealthScore -ge 50) { "#854d0e" }
                     else { "#dc2626" }

$ExecSeverityRank = [ordered]@{ 'High' = 0; 'Medium' = 1; 'Low' = 2 }
$TopRisksSorted = @($TopRisksList | Sort-Object { $ExecSeverityRank[[string]$_.Severity] } | Select-Object -First 6)

$TopRisksHtml = if ($TopRisksSorted.Count -eq 0) {
    '<div class="exec-risk-empty">&#10003; No significant risks detected</div>'
}
else {
    ($TopRisksSorted | ForEach-Object {
        $ExecBadgeClass = if ($_.Severity -eq 'High') { 'exec-risk-high' } else { 'exec-risk-medium' }
        $ExecBadgeLabel = if ($_.Severity -eq 'High') { 'HIGH' } else { 'MEDIUM' }
        "<div class=`"exec-risk-row`"><span class=`"exec-risk-badge $ExecBadgeClass`">$ExecBadgeLabel</span><span class=`"exec-risk-text`">$($_.Text)</span></div>"
    }) -join "`n"
}

# Full (unsorted-by-count) risk list, used only for the CSV export - the
# tile itself still shows just the top 6. The export button only appears
# when there are MORE risks than fit in the tile, so it doesn't clutter
# the card on a healthy environment with few risks.
$TopRisksAllSorted = @($TopRisksList | Sort-Object { $ExecSeverityRank[[string]$_.Severity] })

$TopRisksCsvJs = ConvertTo-JavaScriptString (
    (
        $TopRisksAllSorted |
        ForEach-Object { [PSCustomObject]@{ Severity = $_.Severity; Text = $_.Text } } |
        ConvertTo-Csv -NoTypeInformation
    ) -join "`n"
)

$TopRisksExportBtn = if ($TopRisksAllSorted.Count -gt $TopRisksSorted.Count) {
    "<button class=`"export-btn`" title=`"Export Full Risk List`" onclick=`"downloadCsv('Top_Risks_Full_List.csv', topRisksCsv)`">$([char]0x21E9)</button>"
} else { "" }

# Full per-DC risk list for the "Security Risks Need Attention" report -
# sorted High first, then Medium, then Low.
$SecurityRisksSorted = @($SecurityRisksList | Sort-Object { $ExecSeverityRank[[string]$_.Severity] })

$SecurityRisksCsvJs = ConvertTo-JavaScriptString (
    (
        $SecurityRisksSorted |
        ForEach-Object { [PSCustomObject]@{ Severity = $_.Severity; Risk = $_.Risk; DC = $_.DC; Source = $_.Source } } |
        ConvertTo-Csv -NoTypeInformation
    ) -join "`n"
)

# Gauge geometry (r=50 circle, circumference = 2*pi*r)
$ExecGaugeCircumference = [Math]::Round((2 * [Math]::PI * 50), 1)
$ExecGaugeDash = [Math]::Round(($HealthScore / 100) * $ExecGaugeCircumference, 1)
$ExecGaugeDashArray = "$ExecGaugeDash $ExecGaugeCircumference"

# Health score history (CSV stored alongside the script, grows over each
# run). Tracked per-RUN (not per-day) - every execution appends a new point
# with a full date+time stamp, so testing fixes back-to-back shows up
# immediately on the trend instead of waiting for a new calendar day.
$HealthHistoryPath = Join-Path -Path $PSScriptRoot -ChildPath "ADHealthHistory.csv"
$RunTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm"

$HealthHistoryRows = @()
if (Test-Path $HealthHistoryPath) {
    try { $HealthHistoryRows = @(Import-Csv -Path $HealthHistoryPath) } catch { $HealthHistoryRows = @() }
}

$HealthHistoryRows += [PSCustomObject]@{ Date = $RunTimestamp; Score = $HealthScore }

if ($HealthHistoryRows.Count -gt 200) {
    $HealthHistoryRows = @($HealthHistoryRows | Select-Object -Last 200)
}

try {
    $HealthHistoryRows | Export-Csv -Path $HealthHistoryPath -NoTypeInformation -Force
}
catch {
    Write-Host "Warning: Could not write health history file - $($_.Exception.Message)" -ForegroundColor Yellow
}

# Trend SVG polyline (precomputed server-side - no charting library needed)
$ExecTrendPointCount = $HealthHistoryRows.Count
$ExecTrendSvgPoints = ""
$ExecTrendSvgW = 460
$ExecTrendSvgH = 90
$ExecTrendPadX = 8
$ExecTrendPadY = 10

if ($ExecTrendPointCount -ge 2) {
    $ExecUsableW = $ExecTrendSvgW - (2 * $ExecTrendPadX)
    $ExecUsableH = $ExecTrendSvgH - (2 * $ExecTrendPadY)
    $ExecStepX = $ExecUsableW / ($ExecTrendPointCount - 1)
    $ExecPts = @()

    for ($ExecI = 0; $ExecI -lt $ExecTrendPointCount; $ExecI++) {
        $ExecScoreVal = [double]$HealthHistoryRows[$ExecI].Score
        $ExecX = $ExecTrendPadX + ($ExecI * $ExecStepX)
        $ExecY = $ExecTrendPadY + $ExecUsableH - (($ExecScoreVal / 100) * $ExecUsableH)
        $ExecPts += "$([Math]::Round($ExecX,1)),$([Math]::Round($ExecY,1))"
    }

    $ExecTrendSvgPoints = $ExecPts -join " "
}

$ExecTrendStartLabel = if ($HealthHistoryRows.Count -gt 0) { $HealthHistoryRows[0].Date } else { "" }
$ExecTrendEndLabel   = if ($HealthHistoryRows.Count -gt 0) { $HealthHistoryRows[-1].Date } else { "" }

# Pre-build the trend chart HTML as a plain string (avoids nesting a
# here-string inside the larger $html here-string below, which PowerShell's
# tokenizer does not handle reliably).
if ($ExecTrendPointCount -ge 2) {
    $ExecTrendHtml =
        '<svg viewBox="0 0 ' + $ExecTrendSvgW + ' ' + $ExecTrendSvgH + '" width="100%" height="90" preserveAspectRatio="none">' +
        '<polyline points="' + $ExecTrendSvgPoints + '" fill="none" stroke="#1e3a8a" stroke-width="2.5" stroke-linejoin="round" stroke-linecap="round"/>' +
        '</svg>' +
        '<div class="exec-trend-footer"><span>' + $ExecTrendStartLabel + '</span><span>' + $ExecTrendEndLabel + '</span></div>'
}
else {
    $ExecTrendHtml = '<div class="exec-trend-empty">Trend builds up as this report runs over time.<br/>This is your first recorded health score.</div>'
}

# Precompute the small conditional bits used in the metric cards below, so
# the $html here-string only ever does plain $variable interpolation.
$ExecDcOnlineCount = $DomainControllerCount - $DCPingFailedCount
$ExecRiskCount      = $TopRisksSorted.Count

if ($DCPingFailedCount -eq 0) {
    $ExecDcSubColor = '#16a34a'
    $ExecDcSubText  = 'All online'
}
else {
    $ExecDcSubColor = '#dc2626'
    $ExecDcSubText  = "$DCPingFailedCount unreachable"
}

if ($TopRisksSorted.Count -eq 0) {
    $ExecRiskSubColor = '#16a34a'
    $ExecRiskSubText  = 'None detected'
}
else {
    $ExecRiskSubColor = '#854d0e'
    $ExecRiskSubText  = 'Needs review'
}

$ExecGpoIsUnlinked = ($UnlinkedGPOCount -is [int]) -and ($UnlinkedGPOCount -gt 0)
if ($ExecGpoIsUnlinked) {
    $ExecGpoSubColor = '#854d0e'
    $ExecGpoSubText  = "$UnlinkedGPOCount unlinked"
}
else {
    $ExecGpoSubColor = '#16a34a'
    $ExecGpoSubText  = 'All linked'
}

if (-not $ReplHealthKnown) {
    $ExecReplValue    = "N/A"
    $ExecReplSubColor = '#64748b'
    $ExecReplSubText  = 'Not collected'
}
elseif ($ReplFailedLinks -eq 0) {
    $ExecReplValue    = "$ReplTotalLinks/$ReplTotalLinks"
    $ExecReplSubColor = '#16a34a'
    $ExecReplSubText  = 'All in sync'
}
else {
    $ExecReplValue    = "$($ReplTotalLinks - $ReplFailedLinks)/$ReplTotalLinks"
    $ExecReplSubColor = '#dc2626'
    $ExecReplSubText  = "$ReplFailedLinks failing"
}

if (-not $DnsHealthKnown) {
    $ExecDnsValue     = "N/A"
    $ExecDnsSubColor  = '#64748b'
    $ExecDnsSubText   = 'Not collected'
}
elseif ($DnsDcsSrvMissing -eq 0) {
    $ExecDnsValue     = "$DnsDcsChecked/$DnsDcsChecked"
    $ExecDnsSubColor  = '#16a34a'
    $ExecDnsSubText   = 'SRV records OK'
}
else {
    $ExecDnsValue     = "$($DnsDcsChecked - $DnsDcsSrvMissing)/$DnsDcsChecked"
    $ExecDnsSubColor  = '#dc2626'
    $ExecDnsSubText   = "$DnsDcsSrvMissing missing SRV"
}

# ---- Mini donut rings for DC / DNS / Replication (r=28 circle) ----
$ExecRingCircumference = [Math]::Round((2 * [Math]::PI * 28), 1)

if ($DomainControllerCount -gt 0) {
    $ExecDcRingPct = $ExecDcOnlineCount / $DomainControllerCount
}
else {
    $ExecDcRingPct = 0
}
$ExecDcRingColor = if ($DCPingFailedCount -eq 0) { '#16a34a' } else { '#dc2626' }
$ExecDcRingDash  = "$([Math]::Round($ExecDcRingPct * $ExecRingCircumference, 1)) $ExecRingCircumference"

if (-not $DnsHealthKnown) {
    $ExecDnsRingDash  = "0 $ExecRingCircumference"
    $ExecDnsRingColor = '#94a3b8'
}
else {
    $ExecDnsRingPct   = ($DnsDcsChecked - $DnsDcsSrvMissing) / $DnsDcsChecked
    $ExecDnsRingColor = if ($DnsDcsSrvMissing -eq 0) { '#16a34a' } else { '#dc2626' }
    $ExecDnsRingDash  = "$([Math]::Round($ExecDnsRingPct * $ExecRingCircumference, 1)) $ExecRingCircumference"
}

if (-not $ReplHealthKnown) {
    $ExecReplRingDash  = "0 $ExecRingCircumference"
    $ExecReplRingColor = '#94a3b8'
}
else {
    $ExecReplRingPct   = ($ReplTotalLinks - $ReplFailedLinks) / $ReplTotalLinks
    $ExecReplRingColor = if ($ReplFailedLinks -eq 0) { '#16a34a' } else { '#dc2626' }
    $ExecReplRingDash  = "$([Math]::Round($ExecReplRingPct * $ExecRingCircumference, 1)) $ExecRingCircumference"
}

# ---- Bar-fill percentages for the KPI cards (always defined, even when a
# metric isn't collected, so the progress bar markup never interpolates a
# blank width) ----
$ExecDcBarPct   = [Math]::Round($ExecDcRingPct * 100)
$ExecDnsBarPct  = if ($DnsHealthKnown)  { [Math]::Round($ExecDnsRingPct * 100) }  else { 0 }
$ExecReplBarPct = if ($ReplHealthKnown) { [Math]::Round($ExecReplRingPct * 100) } else { 0 }

# ---- Bar chart geometry shared by the GPO and Security Risks charts ----
# Plot area: x=30..230, y=15 (top) to y=105 (bottom, the zero line).
$ExecChartTop     = 15
$ExecChartBottom  = 105
$ExecChartUsableH = $ExecChartBottom - $ExecChartTop

# ---- Group Policies chart data: Total / Linked / Unlinked ----
$ExecGpoKnown = ($GPOCount -is [int])

if ($ExecGpoKnown) {
    $ExecGpoUnlinkedNum = if ($UnlinkedGPOCount -is [int]) { $UnlinkedGPOCount } else { 0 }
    $ExecGpoLinkedNum   = $GPOCount - $ExecGpoUnlinkedNum
    $ExecGpoTotalNum    = $GPOCount
}
else {
    $ExecGpoUnlinkedNum = 0
    $ExecGpoLinkedNum   = 0
    $ExecGpoTotalNum    = 0
}

$ExecGpoChartMax = [Math]::Max($ExecGpoTotalNum, 1)

$ExecGpoTotalBarH    = [Math]::Round(($ExecGpoTotalNum    / $ExecGpoChartMax) * $ExecChartUsableH, 1)
$ExecGpoLinkedBarH   = [Math]::Round(($ExecGpoLinkedNum   / $ExecGpoChartMax) * $ExecChartUsableH, 1)
$ExecGpoUnlinkedBarH = [Math]::Round(($ExecGpoUnlinkedNum / $ExecGpoChartMax) * $ExecChartUsableH, 1)

$ExecGpoTotalBarY    = [Math]::Round($ExecChartBottom - $ExecGpoTotalBarH, 1)
$ExecGpoLinkedBarY   = [Math]::Round($ExecChartBottom - $ExecGpoLinkedBarH, 1)
$ExecGpoUnlinkedBarY = [Math]::Round($ExecChartBottom - $ExecGpoUnlinkedBarH, 1)

$ExecGpoAxisMaxLabel = $ExecGpoTotalNum
$ExecGpoAxisMidLabel = [Math]::Round($ExecGpoTotalNum / 2, 0)

$ExecGpoTotalLabelY    = [Math]::Max($ExecGpoTotalBarY - 5, 10)
$ExecGpoLinkedLabelY   = [Math]::Max($ExecGpoLinkedBarY - 5, 10)
$ExecGpoUnlinkedLabelY = [Math]::Max($ExecGpoUnlinkedBarY - 5, 10)

# ---- Security Risks chart data: Total / High / Medium ----
# Sourced from $SecurityRisksList (the same per-DC finding list that feeds
# the "Security Risks Need Attention" report), not $TopRisksList (which
# collapses all DCs for a given check into a single summary row) - so the
# numbers shown here always match what the report shows.
$ExecHighCount   = @($SecurityRisksList | Where-Object { $_.Severity -eq 'High' }).Count
$ExecMediumCount = @($SecurityRisksList | Where-Object { $_.Severity -eq 'Medium' }).Count
$ExecSecTotalNum = $ExecHighCount + $ExecMediumCount

$ExecSecChartMax = [Math]::Max($ExecSecTotalNum, 1)

$ExecSecTotalBarH = [Math]::Round(($ExecSecTotalNum / $ExecSecChartMax) * $ExecChartUsableH, 1)
$ExecHighBarH     = [Math]::Round(($ExecHighCount   / $ExecSecChartMax) * $ExecChartUsableH, 1)
$ExecMediumBarH   = [Math]::Round(($ExecMediumCount / $ExecSecChartMax) * $ExecChartUsableH, 1)

$ExecSecTotalBarY = [Math]::Round($ExecChartBottom - $ExecSecTotalBarH, 1)
$ExecHighBarY     = [Math]::Round($ExecChartBottom - $ExecHighBarH, 1)
$ExecMediumBarY   = [Math]::Round($ExecChartBottom - $ExecMediumBarH, 1)

$ExecSecAxisMaxLabel = $ExecSecTotalNum
$ExecSecAxisMidLabel = [Math]::Round($ExecSecTotalNum / 2, 0)

$ExecSecTotalLabelY  = [Math]::Max($ExecSecTotalBarY - 5, 10)
$ExecHighLabelY      = [Math]::Max($ExecHighBarY - 5, 10)
$ExecMediumLabelY    = [Math]::Max($ExecMediumBarY - 5, 10)

$html = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Enterprise Active Directory Operational Dashboard</title>

<style>
    html, body {
        height: 100%;
    }

    body {
        margin: 0;
        padding: 0;
        font-family: "Segoe UI", Arial, sans-serif;
        background: #f0f4f8;
        color: #1f2937;
        min-height: 100vh;
    }

    .sticky-top {
        position: fixed;
        top: 0;
        left: 0;
        right: 0;
        z-index: 300;
        background: #f0f4f8;
        border-bottom: 1px solid #e2e8f0;
        box-shadow: 0 2px 10px rgba(15, 23, 42, 0.07);
        padding: 8px 16px 10px 16px;
    }

    .page-generated {
        width: 92%;
        max-width: 1400px;
        margin: 10px auto 0 auto;
        text-align: right;
        font-size: 12px;
        color: #64748b;
    }

    .container {
        width: 92%;
        margin: 0 auto;
        padding-bottom: 34px;
    }

    .header {
        margin: 0;
        background: linear-gradient(90deg, #1e3a8a, #1e40af, #2563eb);
        padding: 14px 24px;
        text-align: center;
        border-radius: 12px 12px 0 0;
    }

    .header-title {
        font-size: 28px;
        font-weight: 800;
        color: #ffffff;
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
        z-index: 10;
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
        z-index: 10;
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

    /* ---- Executive Summary tab ---- */
    .exec-grid {
        max-width: 950px;
        margin: 0 auto;
    }

    /* The toggle is taken OUT of normal flow (position:absolute) so it
       never reserves horizontal space inside the row - reserving space
       here would shrink the 4-tile grid and break its column alignment
       with the rows below. It floats just outside the row's right edge
       instead. */
    .exec-row1-wrap {
        position: relative;
        margin-bottom: 16px;
    }

    .exec-row1-wrap .exec-row1 {
        margin-bottom: 0;
    }

    .exec-kpi-toggle {
        position: absolute;
        top: 0;
        left: 100%;
        margin-left: 10px;
        display: flex;
        flex-direction: column;
        gap: 6px;
    }

    .exec-kpi-toggle-btn {
        width: 26px;
        height: 26px;
        border-radius: 8px;
        border: 1px solid #e2e8f0;
        background: #ffffff;
        color: #94a3b8;
        display: flex;
        align-items: center;
        justify-content: center;
        cursor: pointer;
    }

    .exec-kpi-toggle-btn:hover {
        background: #f8fafc;
    }

    .exec-kpi-toggle-btn.active {
        background: #eff6ff;
        border-color: #93c5fd;
        color: #2563eb;
    }

    .exec-row1 {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 16px;
        margin-bottom: 16px;
    }

    .exec-row2 {
        display: grid;
        grid-template-columns: repeat(2, 1fr);
        gap: 16px;
        margin-bottom: 22px;
    }

    .exec-tile {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 16px;
        box-shadow: 0 8px 18px rgba(15, 23, 42, 0.07);
        padding: 16px;
        text-align: center;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        min-height: 150px;
    }

    .exec-card {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 18px;
        box-shadow: 0 8px 20px rgba(15, 23, 42, 0.08);
        padding: 18px;
    }

    .exec-gauge-card {
        border: 2px solid #2563eb;
    }

    .exec-gauge-label {
        font-size: 13px;
        font-weight: 400;
        color: #64748b;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        margin-top: 6px;
    }

    .exec-gauge-sublabel {
        font-size: 11px;
        font-weight: 600;
        margin-top: 4px;
    }

    .exec-ring-label {
        font-size: 13px;
        font-weight: 400;
        color: #64748b;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        margin-bottom: 8px;
    }

    .exec-kpi-label-row {
        display: flex;
        align-items: center;
        gap: 6px;
        margin-bottom: 8px;
    }

    .exec-kpi-label-row svg {
        flex-shrink: 0;
    }

    .exec-ring-sub {
        font-size: 11px;
        font-weight: 600;
        margin-top: 6px;
    }

    /* KPI card style (Concept 2: number + slim horizontal bar, no ring) */
    .exec-kpi-card-left {
        text-align: left;
        align-items: flex-start;
    }

    .exec-kpi-value {
        font-size: 22px;
        font-weight: 500;
        margin-top: 6px;
        font-family: 'Segoe UI', Arial, sans-serif;
        line-height: 1;
    }

    .exec-kpi-value-lg {
        font-size: 28px;
    }

    .exec-kpi-suffix {
        font-size: 12px;
        font-weight: 500;
        color: #94a3b8;
    }

    .exec-kpi-bar-track {
        height: 6px;
        width: 100%;
        background: #e2e8f0;
        border-radius: 4px;
        margin-top: 10px;
        overflow: hidden;
    }

    .exec-kpi-bar-fill {
        height: 100%;
        border-radius: 4px;
    }

    .exec-kpi-status {
        font-size: 11px;
        font-weight: 500;
        margin-top: 8px;
    }

    .exec-chart-card {
        flex-direction: column;
        align-items: stretch;
        text-align: left;
        padding: 16px 18px;
    }

    .exec-chart-header {
        display: flex;
        justify-content: space-between;
        align-items: baseline;
        margin-bottom: 4px;
    }

    .exec-chart-title {
        font-size: 13px;
        font-weight: 700;
        color: #475569;
    }

    .exec-chart-total {
        font-size: 13px;
        font-weight: 800;
        color: #0f172a;
    }

    .exec-bottom-row {
        display: grid;
        grid-template-columns: repeat(2, 1fr);
        gap: 16px;
    }

    .exec-card-title {
        font-size: 14px;
        font-weight: 600;
        color: #475569;
        margin-bottom: 12px;
    }

    .exec-card-title-row {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 12px;
        margin-bottom: 12px;
    }

    .exec-risk-row {
        display: flex;
        align-items: center;
        gap: 10px;
        padding: 9px 0;
        border-bottom: 1px solid #f1f5f9;
    }

    .exec-risk-row:last-child {
        border-bottom: none;
    }

    .exec-risk-badge {
        flex: 0 0 auto;
        font-size: 10px;
        font-weight: 700;
        padding: 3px 9px;
        border-radius: 6px;
        letter-spacing: 0.03em;
    }

    .exec-risk-high {
        background: #fef2f2;
        color: #b91c1c;
    }

    .exec-risk-medium {
        background: #ffedd5;
        color: #c2410c;
    }

    .exec-risk-text {
        font-size: 12.5px;
        color: #1e293b;
        line-height: 1.4;
    }

    .exec-risk-empty {
        text-align: center;
        padding: 24px 10px;
        color: #16a34a;
        font-size: 13px;
        font-weight: 700;
    }

    .exec-trend-footer {
        display: flex;
        justify-content: space-between;
        font-size: 10px;
        color: #94a3b8;
        margin-top: 4px;
    }

    .exec-trend-empty {
        text-align: center;
        padding: 24px 10px;
        color: #94a3b8;
        font-size: 12px;
        font-style: italic;
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
        background: #fef9c3;
        color: #a16207;
    }

    .health-critical {
        background: #fef2f2;
        color: #b91c1c;
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

    tr:nth-child(even) {
        background: #eff6ff;
    }

    tr:hover {
        background: #dbeafe;
    }

    .footer {
        position: fixed;
        bottom: 0;
        left: 0;
        right: 0;
        z-index: 200;
        background: #f0f4f8;
        border-top: 1px solid #e2e8f0;
        padding: 6px 24px;
        font-size: 11px;
        color: #64748b;
        text-align: center;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        box-shadow: 0 -2px 10px rgba(15, 23, 42, 0.07);
    }

    /* =====================================================================
       Header subtitle / footer copyright
       ===================================================================== */

    .header-subtitle {
        margin-top: 6px;
        font-size: 13px;
        font-weight: 500;
        color: #bfdbfe;
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

    .tab-indicator-strip {
        background: #6b7280;
        height: 7px;
        position: relative;
        z-index: 5;
    }

    .tab-indicator-arrow {
        position: absolute;
        top: 7px;
        left: 0;
        width: 0;
        height: 0;
        border-left: 7px solid transparent;
        border-right: 7px solid transparent;
        border-top: 8px solid #6b7280;
        transform: translateX(-50%);
        transition: left 0.25s ease;
    }

    .tab-bar {
        max-width: 100%;
        margin: 0;
        display: flex;
        gap: 0;
        background: linear-gradient(180deg, #ffffff, #f7f9fb);
        border-bottom: 3px solid #29ABE2;
        padding: 0;
        position: relative;
        z-index: 1;
        height: 38px;
        box-sizing: border-box;
    }

    .tab-button {
        appearance: none;
        border: none;
        border-right: 1px solid #e5e7eb;
        background: transparent;
        cursor: pointer;
        padding: 0 10px;
        font-size: 13px;
        font-weight: 400;
        font-family: inherit;
        color: #3a3a3a;
        flex: 1;
        white-space: nowrap;
        letter-spacing: 0.02em;
        transition: color 0.15s ease;
        display: flex;
        align-items: center;
        justify-content: center;
    }

    .tab-button:last-child {
        border-right: none;
    }

    .tab-button:hover:not(.active) {
        color: #1d4ed8;
    }

    .tab-button.active {
        color: #29ABE2;
        background: transparent;
    }

    .page-ribbon {
        background: linear-gradient(90deg, #29ABE2, #4FC3F7);
        position: relative;
        height: 38px;
    }

    .ribbon-title {
        font-size: 20px;
        font-weight: 400;
        color: #ffffff;
        letter-spacing: 0.02em;
        position: absolute;
        top: 50%;
        transform: translateY(-50%);
        white-space: nowrap;
    }

    .ribbon-crumb {
        font-size: 11px;
        color: #e3f6fd;
        position: absolute;
        top: 50%;
        transform: translateY(-50%);
        white-space: nowrap;
    }

    .ribbon-crumb b {
        color: #ffffff;
        font-weight: 700;
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
        z-index: 10;
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
    .dc-pill-warning { background: #fef9c3; color: #a16207; }
    .dc-pill-critical { background: #fef2f2; color: #b91c1c; }
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

    .dc-conn-grid {
        display: grid;
        grid-template-columns: repeat(5, 1fr);
        gap: 8px;
    }

    .dc-conn-cell {
        display: flex;
        flex-direction: column;
        align-items: center;
        gap: 4px;
        padding: 7px 6px;
        border-radius: 8px;
        background: #f8fafc;
        text-align: center;
    }

    .dc-conn-label {
        font-size: 12px;
        font-weight: 600;
        color: #64748b;
        min-height: 20px;
        display: flex;
        align-items: center;
        justify-content: center;
    }

    .dc-conn-cell .health {
        font-size: 13px;
        margin-top: 0;
    }

    .dc-role-section-label {
        font-size: 11px;
        font-weight: 600;
        color: #94a3b8;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        margin: 10px 0 6px;
    }

    .dc-role-section-label:first-of-type {
        margin-top: 0;
    }

    .dc-role-list {
        display: flex;
        flex-wrap: wrap;
        gap: 6px;
    }

    .dc-role-badge {
        font-size: 12px;
        font-weight: 600;
        padding: 4px 10px;
        border-radius: 6px;
    }

    .dc-role-badge-good {
        background: #dcfce7;
        color: #166534;
    }

    .dc-role-badge-warning {
        background: #fef9c3;
        color: #854d0e;
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

    /* Caps the Replication Partners tile at ~5 visible rows (header + 5
       data rows) regardless of how many partners a DC has - the export
       button + note below covers seeing the full list. */
    .dc-repl-scroll {
        max-height: 190px;
        overflow-y: auto;
        border-radius: 10px;
    }

    .dc-repl-note {
        font-size: 10.5px;
        color: #64748b;
        margin-top: 8px;
        padding: 0 2px;
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

        .dc-conn-grid {
            grid-template-columns: repeat(2, 1fr);
        }
    }

    /* =====================================================================
       Reports tab
       ===================================================================== */

    .rpt-tab-wrap {
        position: relative;
        flex: 1;
    }

    .rpt-tab-wrap .tab-button {
        width: 100%;
        height: 100%;
        border-right: none;
    }

    .rpt-dropdown {
        display: none;
        position: absolute;
        top: 100%;
        left: 0;
        right: 0;
        background: #ffffff;
        border: 1px solid #cbd5e1;
        border-radius: 8px;
        box-shadow: 0 6px 20px rgba(15,23,42,0.12);
        z-index: 200;
        overflow: hidden;
        min-width: 210px;
        padding-top: 4px;
    }

    .rpt-tab-wrap:hover .rpt-dropdown,
    .rpt-tab-wrap.open .rpt-dropdown {
        display: block;
    }

    .rpt-dd-header {
        padding: 6px 12px;
        font-size: 10px;
        font-weight: 700;
        color: #94a3b8;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        background: #f8fafc;
        border-bottom: 1px solid #e2e8f0;
    }

    .rpt-dd-item {
        padding: 9px 14px;
        font-size: 13px;
        font-weight: 500;
        color: #334155;
        cursor: pointer;
        transition: background 0.1s;
    }

    .rpt-dd-item:hover {
        background: #eff6ff;
        color: #1e40af;
    }

    .rpt-dd-item.active {
        background: #eff6ff;
        color: #1d4ed8;
        font-weight: 700;
    }

    .rpt-report-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 14px;
    }

    .rpt-report-title {
        font-size: 16px;
        font-weight: 700;
        color: #1e293b;
    }

    .rpt-export-btn {
        appearance: none;
        border: 1px solid #cbd5e1;
        background: #ffffff;
        color: #374151;
        font-size: 12px;
        font-weight: 600;
        font-family: inherit;
        padding: 6px 14px;
        border-radius: 6px;
        cursor: pointer;
        transition: background 0.15s, border-color 0.15s, color 0.15s;
    }

    .rpt-export-btn:hover {
        background: #eff6ff;
        border-color: #93c5fd;
        color: #1e40af;
    }

    /* ---- Reports tab: global theme toggle (applies to every report) ----
       Note: the grouped header row (.rpt-group-row) is disabled dashboard-
       wide via buildGroupRow() returning '', so it never renders and needs
       no theme rules here. */
    #rptPanel[data-theme="A"] .rpt-header-row th {
        background: #fafbfc;
        color: #24292f;
        border-bottom: 1px solid #e1e4e8;
        border-right: 1px solid #e1e4e8;
    }
    #rptPanel[data-theme="A"] .rpt-header-row th.rpt-gd { border-left: 1px solid #d0d7de; }
    #rptPanel[data-theme="A"] table tbody tr:nth-child(even) td { background: #fcfcfd; }
    #rptPanel[data-theme="A"] table tbody tr:hover td { background: #f6f8fa; }
    #rptPanel[data-theme="A"] td.rpt-dc { color: #0969da; border-right: 1px solid #eef0f2; }
    #rptPanel[data-theme="A"] td.rpt-gd { border-left: 1px solid #eef0f2; }
    /* Azure Portal style: no filled pill - plain text with a small color
       dot in front (::before), using currentColor so the dot always
       matches whichever status color is set below. */
    #rptPanel[data-theme="A"] .rpt-b {
        background: transparent;
        border-radius: 0;
        padding: 0 0 0 13px;
        position: relative;
        font-weight: 400;
    }
    #rptPanel[data-theme="A"] .rpt-b::before {
        content: '';
        position: absolute;
        left: 0;
        top: 50%;
        transform: translateY(-50%);
        width: 6px;
        height: 6px;
        border-radius: 50%;
        background: currentColor;
    }
    #rptPanel[data-theme="A"] .rpt-good   { color: #1a7f37; }
    #rptPanel[data-theme="A"] .rpt-warn   { color: #a16207; }
    #rptPanel[data-theme="A"] .rpt-medium { color: #c2410c; }
    #rptPanel[data-theme="A"] .rpt-crit   { color: #b91c1c; }
    #rptPanel[data-theme="A"] .rpt-na     { color: #656d76; }
    #rptPanel[data-theme="A"] .rpt-info   { color: #0969da; }
    #rptPanel[data-theme="A"] .rpt-purple { color: #6639ba; }
    #rptPanel[data-theme="A"] .rpt-pink   { color: #bf3989; }

    /* Identity Risk & Attack Surface's table, and the Group & Privilege
       Architecture table, both mirror the EXACT same Theme A/Theme B rules
       as #rptPanel (header colors, badge shape), scoped to their own table
       IDs instead - they live on different tabs so the #rptPanel-scoped
       rules above don't reach them. toggleReportTheme() keeps both
       elements' data-theme attribute in sync with #rptPanel's, so switching
       style in Reports also restyles both tables. */
    #identityUsersTable[data-theme="A"] .rpt-header-row th,
    #groupPrivTable[data-theme="A"] .rpt-header-row th {
        background: #fafbfc;
        color: #24292f;
        border-bottom: 1px solid #e1e4e8;
        border-right: 1px solid #e1e4e8;
    }
    #identityUsersTable[data-theme="A"] table tbody tr:nth-child(even) td,
    #groupPrivTable[data-theme="A"] table tbody tr:nth-child(even) td { background: #fcfcfd; }
    #identityUsersTable[data-theme="A"] table tbody tr:hover td,
    #groupPrivTable[data-theme="A"] table tbody tr:hover td { background: #f6f8fa; }
    #identityUsersTable[data-theme="A"] .rpt-b,
    #groupPrivTable[data-theme="A"] .rpt-b {
        background: transparent;
        border-radius: 0;
        padding: 0 0 0 13px;
        position: relative;
        font-weight: 400;
    }
    #identityUsersTable[data-theme="A"] .rpt-b::before,
    #groupPrivTable[data-theme="A"] .rpt-b::before {
        content: '';
        position: absolute;
        left: 0;
        top: 50%;
        transform: translateY(-50%);
        width: 6px;
        height: 6px;
        border-radius: 50%;
        background: currentColor;
    }
    #identityUsersTable[data-theme="A"] .id-critical,
    #groupPrivTable[data-theme="A"] .id-critical { color: #7f1d1d; }
    #identityUsersTable[data-theme="A"] .rpt-crit,
    #groupPrivTable[data-theme="A"] .rpt-crit     { color: #b91c1c; }
    #identityUsersTable[data-theme="A"] .rpt-medium,
    #groupPrivTable[data-theme="A"] .rpt-medium   { color: #c2410c; }

    #identityUsersTable[data-theme="B"] .rpt-header-row th,
    #groupPrivTable[data-theme="B"] .rpt-header-row th {
        background: #3a4150;
        color: #ffffff;
    }
    #identityUsersTable[data-theme="B"] table tbody tr:nth-child(even) td,
    #groupPrivTable[data-theme="B"] table tbody tr:nth-child(even) td { background: #f7f8fa; }
    #identityUsersTable[data-theme="B"] table tbody tr:hover td,
    #groupPrivTable[data-theme="B"] table tbody tr:hover td { background: #eef0f3; }
    #identityUsersTable[data-theme="B"] .rpt-b,
    #groupPrivTable[data-theme="B"] .rpt-b {
        border-radius: 10px;
        border: 1px solid;
        background: #ffffff;
        font-weight: 600;
    }
    #identityUsersTable[data-theme="B"] .id-critical,
    #groupPrivTable[data-theme="B"] .id-critical { border-color: #7f1d1d; color: #7f1d1d; }
    #identityUsersTable[data-theme="B"] .rpt-crit,
    #groupPrivTable[data-theme="B"] .rpt-crit     { border-color: #b91c1c; color: #b91c1c; }
    #identityUsersTable[data-theme="B"] .rpt-medium,
    #groupPrivTable[data-theme="B"] .rpt-medium   { border-color: #c2410c; color: #c2410c; }

    /* Both tables use table-layout:fixed with explicit column widths (see
       each <colgroup>) so expanding a row's detail panel never resizes the
       columns - the shared .rpt-styled-table td rule forces nowrap, which
       both other reports rely on AND would make long text (Why, Recommended
       Action, the detail panel) overflow horizontally instead of wrapping,
       so it's overridden here, scoped to these two tables only. */
    #identityUsersTable td,
    #groupPrivTable td {
        white-space: normal;
        word-break: break-word;
        vertical-align: top;
    }
    #identityUsersTable .id-detail-row td,
    #groupPrivTable .id-detail-row td {
        max-width: 0;
        overflow: hidden;
    }

    #rptPanel[data-theme="B"] .rpt-header-row th {
        background: #3a4150;
        color: #ffffff;
    }
    #rptPanel[data-theme="B"] .rpt-header-row th.rpt-gd { border-left: 2px solid #5a6373; }
    #rptPanel[data-theme="B"] table tbody tr:nth-child(even) td { background: #f7f8fa; }
    #rptPanel[data-theme="B"] table tbody tr:hover td { background: #eef0f3; }
    #rptPanel[data-theme="B"] td.rpt-dc { color: #3a4150; border-right: 1px solid #e8eaed; }
    #rptPanel[data-theme="B"] td.rpt-gd { border-left: 2px solid #e8eaed; }
    #rptPanel[data-theme="B"] .rpt-b {
        border-radius: 10px;
        border: 1px solid;
        background: #ffffff;
        font-weight: 600;
    }
    #rptPanel[data-theme="B"] .rpt-good   { border-color: #1a7f5a; color: #1a7f5a; }
    #rptPanel[data-theme="B"] .rpt-warn   { border-color: #a16207; color: #a16207; }
    #rptPanel[data-theme="B"] .rpt-medium { border-color: #c2410c; color: #c2410c; }
    #rptPanel[data-theme="B"] .rpt-crit   { border-color: #b91c1c; color: #b91c1c; }
    #rptPanel[data-theme="B"] .rpt-na     { border-color: #9aa3ad; color: #9aa3ad; background: #ffffff; }
    #rptPanel[data-theme="B"] .rpt-info   { border-color: #1a5fa0; color: #1a5fa0; }
    #rptPanel[data-theme="B"] .rpt-purple { border-color: #6639ba; color: #6639ba; }
    #rptPanel[data-theme="B"] .rpt-pink   { border-color: #bf3989; color: #bf3989; }

    .rpt-tbl-wrap {
        overflow-x: auto;
        border-radius: 8px;
        border: 1px solid #e2e8f0;
    }

    .rpt-styled-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 12px;
    }

    .rpt-styled-table .rpt-header-row th {
        background: #1e40af;
        color: #ffffff;
        padding: 7px 10px;
        text-align: left;
        font-weight: 600;
        white-space: nowrap;
        border-right: 1px solid #2563eb;
    }

    /* Sortable column headers - works regardless of the active theme's
       header background color, so the hover affordance is always visible. */
    .rpt-header-row th[onclick]:hover {
        filter: brightness(0.92);
    }

    .rpt-styled-table .rpt-header-row th:last-child { border-right: none; }
    .rpt-styled-table .rpt-header-row th.rpt-gd { border-left: 2px solid #60a5fa; }

    .rpt-styled-table tbody tr:nth-child(even) td { background: #eff6ff; }
    .rpt-styled-table tbody tr:hover td { background: #dbeafe; }

    .rpt-styled-table td {
        padding: 6px 10px;
        border-bottom: 1px solid #e2e8f0;
        border-right: 1px solid #f1f5f9;
        white-space: nowrap;
        color: #1e293b;
    }

    .rpt-styled-table td:last-child { border-right: none; }
    .rpt-styled-table td.rpt-dc { font-weight: 700; color: #1e40af; border-right: 1px solid #cbd5e1; }
    .rpt-styled-table td.rpt-gd { border-left: 2px solid #cbd5e1; }

    /* The Identity & Kerberos Security and Security Risks Need Attention
       reports both expand a detail row on click now (same pattern as the
       Identity Risk and Group & Privilege Architecture tabs) - the shared
       .rpt-styled-table td rule above forces nowrap, which is fine for
       normal rows but would make the detail panel's wrapped paragraphs
       and sub-tables overflow horizontally instead of wrapping, so it's
       overridden specifically for the detail row's cell. */
    .rpt-styled-table .id-detail-row td {
        white-space: normal;
        word-break: break-word;
        vertical-align: top;
    }

    .rpt-b { display: inline-block; font-size: 10px; font-weight: 600; padding: 2px 8px; border-radius: 10px; }
    .rpt-good { background: #dcfce7; color: #166534; }
    .rpt-warn { background: #fef9c3; color: #854d0e; }
    .rpt-crit { background: #fee2e2; color: #991b1b; }
    .rpt-na     { background: #f1f5f9; color: #64748b; }
    .rpt-info   { background: #e0f2fe; color: #075985; }
    .rpt-purple { background: #ede9fe; color: #4c1d95; }
    .rpt-pink   { background: #fbeaf0; color: #72243e; }
    .rpt-disk-warn { color: #854d0e; font-weight: 600; }
    .rpt-disk-crit { color: #991b1b; font-weight: 600; }
    .rpt-disk-ok   { color: #166534; }

    .rpt-note {
        font-size: 11px;
        color: #64748b;
        margin-top: 10px;
        display: flex;
        align-items: center;
        gap: 6px;
        flex-wrap: wrap;
    }
</style>
</head>

<body>

<div class="sticky-top" id="stickyTop">
    <div class="header">
        <div class="header-title">Enterprise Active Directory Health Dashboard</div>
        <div class="header-subtitle">Forest: $ForestRootDomain &nbsp;|&nbsp; Domain: $DomainName &nbsp;|&nbsp; $DomainControllerCount Domain Controllers &nbsp;|&nbsp; Generated: $GeneratedOn</div>
    </div>
    <div class="tab-indicator-strip" id="tabIndicatorStrip">
        <div class="tab-indicator-arrow" id="tabIndicatorArrow"></div>
    </div>
    <div class="tab-bar">
        <button class="tab-button active" id="tab-btn-execsummary" onclick="showTab('execsummary')"><span class="tab-label">Executive Summary</span></button>
        <button class="tab-button" id="tab-btn-forest" onclick="showTab('forest')"><span class="tab-label">Forest Overview</span></button>
        <button class="tab-button" id="tab-btn-dchealth" onclick="showTab('dchealth')"><span class="tab-label">Domain Controller Health</span></button>
        <button class="tab-button" id="tab-btn-deepinsights" onclick="showTab('deepinsights')"><span class="tab-label">Deep Insights</span></button>
        <button class="tab-button" id="tab-btn-secposture" onclick="showTab('secposture')"><span class="tab-label">Security Posture</span></button>
        <button class="tab-button" id="tab-btn-identityrisk" onclick="showTab('identityrisk')"><span class="tab-label">Identity Risk &amp; Attack Surface</span></button>
        <button class="tab-button" id="tab-btn-groupsprivilege" onclick="showTab('groupsprivilege')"><span class="tab-label">Group &amp; Privilege Architecture</span></button>
        <div class="rpt-tab-wrap" id="rptTabWrap">
            <button class="tab-button" id="tab-btn-reports" onclick="showTab('reports')">
                &#9783; <span class="tab-label">Reports</span> &#9663;
            </button>
            <div class="rpt-dropdown" id="rptDropdown">
                <div class="rpt-dd-header">Export reports</div>
                <div class="rpt-dd-item" id="rdd-health" onclick="switchReport('health')">&#9783; DC Health Check</div>
                <div class="rpt-dd-item" id="rdd-inventory" onclick="switchReport('inventory')">&#9635; DC Server Inventory</div>
                <div class="rpt-dd-item" id="rdd-dcuptime" onclick="switchReport('dcuptime')">&#9201; DC Uptime &amp; Reboot Status</div>
                <div class="rpt-dd-item" id="rdd-dcbysite" onclick="switchReport('dcbysite')">&#9671; DC List by Site</div>
                <div class="rpt-dd-item" id="rdd-replpartners" onclick="switchReport('replpartners')">&#8652; AD Replication Partners</div>
                <div class="rpt-dd-item" id="rdd-sites" onclick="switchReport('sites')">&#9671; Sites &amp; Services</div>
                <div class="rpt-dd-item" id="rdd-gpo" onclick="switchReport('gpo')">&#9783; Group Policies</div>
                <div class="rpt-dd-item" id="rdd-dnszones" onclick="switchReport('dnszones')">&#9670; DNS Zones</div>
                <div class="rpt-dd-item" id="rdd-forwarders" onclick="switchReport('forwarders')">&#8594; Conditional Forwarders</div>
                <div class="rpt-dd-item" id="rdd-security" onclick="switchReport('security')">&#9673; Security Posture Matrix</div>
                <div class="rpt-dd-item" id="rdd-connectivity" onclick="switchReport('connectivity')">&#9741; Connectivity Matrix</div>
                <div class="rpt-dd-item" id="rdd-secrisks" onclick="switchReport('secrisks')">&#9888; Security Risks Need Attention</div>
                <div class="rpt-dd-item" id="rdd-identitysec" onclick="switchReport('identitysec')">&#128272; Identity &amp; Kerberos Security</div>
            </div>
        </div>
    </div>
    <div class="page-ribbon" id="pageRibbon">
        <div class="ribbon-title" id="ribbonTitle">Executive Summary</div>
        <div class="ribbon-crumb" id="ribbonCrumb">Dashboard&nbsp;/&nbsp;<b>Executive Summary</b></div>
    </div>
</div>

<div class="container" id="mainContainer">

    <div id="tab-execsummary" class="tab-content active">

    <div class="exec-grid">

        <div class="exec-row1-wrap">
        <div class="exec-row1" id="execRow1Bars">
            <div class="exec-tile exec-gauge-card exec-kpi-card-left">
                <div class="exec-gauge-label">Overall Health</div>
                <div class="exec-kpi-value exec-kpi-value-lg" style="color:$HealthScoreColor;">$HealthScore<span class="exec-kpi-suffix">/100</span></div>
                <div class="exec-kpi-bar-track">
                    <div class="exec-kpi-bar-fill" style="width:$HealthScore%;background:$HealthScoreColor;"></div>
                </div>
                <div class="exec-kpi-status" style="color:$HealthScoreColor;">$HealthScoreLabel</div>
            </div>

            <div class="exec-tile exec-kpi-card-left">
                <div class="exec-kpi-label-row">
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="$ExecDcRingColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="6" rx="1.5"/><rect x="3" y="14" width="18" height="6" rx="1.5"/><circle cx="7" cy="7" r="0.6" fill="$ExecDcRingColor" stroke="none"/><circle cx="7" cy="17" r="0.6" fill="$ExecDcRingColor" stroke="none"/></svg>
                    <div class="exec-ring-label" style="margin-bottom:0;">Domain Controllers</div>
                </div>
                <div class="exec-kpi-value" style="color:$ExecDcRingColor;">$ExecDcOnlineCount/$DomainControllerCount</div>
                <div class="exec-kpi-bar-track">
                    <div class="exec-kpi-bar-fill" style="width:$ExecDcBarPct%;background:$ExecDcRingColor;"></div>
                </div>
                <div class="exec-kpi-status" style="color:$ExecDcSubColor;">$ExecDcSubText</div>
            </div>

            <div class="exec-tile exec-kpi-card-left">
                <div class="exec-kpi-label-row">
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="$ExecDnsRingColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M3 12h18"/><path d="M12 3a14 14 0 0 1 3.5 9 14 14 0 0 1-3.5 9 14 14 0 0 1-3.5-9 14 14 0 0 1 3.5-9z"/></svg>
                    <div class="exec-ring-label" style="margin-bottom:0;">DNS Status</div>
                </div>
                <div class="exec-kpi-value" style="color:$ExecDnsRingColor;">$ExecDnsValue</div>
                <div class="exec-kpi-bar-track">
                    <div class="exec-kpi-bar-fill" style="width:$ExecDnsBarPct%;background:$ExecDnsRingColor;"></div>
                </div>
                <div class="exec-kpi-status" style="color:$ExecDnsSubColor;">$ExecDnsSubText</div>
            </div>

            <div class="exec-tile exec-kpi-card-left">
                <div class="exec-kpi-label-row">
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="$ExecReplRingColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12a9 9 0 1 1-2.64-6.36"/><polyline points="21 3 21 9 15 9"/></svg>
                    <div class="exec-ring-label" style="margin-bottom:0;">Replication Health</div>
                </div>
                <div class="exec-kpi-value" style="color:$ExecReplRingColor;">$ExecReplValue</div>
                <div class="exec-kpi-bar-track">
                    <div class="exec-kpi-bar-fill" style="width:$ExecReplBarPct%;background:$ExecReplRingColor;"></div>
                </div>
                <div class="exec-kpi-status" style="color:$ExecReplSubColor;">$ExecReplSubText</div>
            </div>
        </div>

        <div class="exec-row1" id="execRow1Donuts" style="display:none;">
            <div class="exec-tile exec-gauge-card">
                <div class="exec-gauge-label">Overall Health</div>
                <svg viewBox="0 0 120 120" width="86" height="86">
                    <circle cx="60" cy="60" r="50" fill="none" stroke="#f1f5f9" stroke-width="7"/>
                    <circle cx="60" cy="60" r="50" fill="none" stroke="$HealthScoreColor" stroke-width="7" stroke-dasharray="$ExecGaugeDashArray" stroke-linecap="round" transform="rotate(-90 60 60)"/>
                    <text x="60" y="57" text-anchor="middle" font-size="22" font-weight="500" fill="#0f172a" font-family="Segoe UI, Arial, sans-serif">$HealthScore</text>
                    <text x="60" y="72" text-anchor="middle" font-size="9" fill="#94a3b8" font-family="Segoe UI, Arial, sans-serif">/ 100</text>
                </svg>
                <div class="exec-kpi-status" style="color:$HealthScoreColor;margin-top:6px;">$HealthScoreLabel</div>
            </div>

            <div class="exec-tile">
                <div class="exec-kpi-label-row" style="justify-content:center;">
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="$ExecDcRingColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="6" rx="1.5"/><rect x="3" y="14" width="18" height="6" rx="1.5"/><circle cx="7" cy="7" r="0.6" fill="$ExecDcRingColor" stroke="none"/><circle cx="7" cy="17" r="0.6" fill="$ExecDcRingColor" stroke="none"/></svg>
                    <div class="exec-ring-label" style="margin-bottom:0;">Domain Controllers</div>
                </div>
                <svg viewBox="0 0 70 70" width="60" height="60">
                    <circle cx="35" cy="35" r="28" fill="none" stroke="#f1f5f9" stroke-width="5"/>
                    <circle cx="35" cy="35" r="28" fill="none" stroke="$ExecDcRingColor" stroke-width="5" stroke-dasharray="$ExecDcRingDash" stroke-linecap="round" transform="rotate(-90 35 35)"/>
                    <text x="35" y="39" text-anchor="middle" font-size="13" font-weight="500" fill="#0f172a" font-family="Segoe UI, Arial, sans-serif">$ExecDcOnlineCount/$DomainControllerCount</text>
                </svg>
                <div class="exec-kpi-status" style="color:$ExecDcSubColor;margin-top:6px;">$ExecDcSubText</div>
            </div>

            <div class="exec-tile">
                <div class="exec-kpi-label-row" style="justify-content:center;">
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="$ExecDnsRingColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M3 12h18"/><path d="M12 3a14 14 0 0 1 3.5 9 14 14 0 0 1-3.5 9 14 14 0 0 1-3.5-9 14 14 0 0 1 3.5-9z"/></svg>
                    <div class="exec-ring-label" style="margin-bottom:0;">DNS Status</div>
                </div>
                <svg viewBox="0 0 70 70" width="60" height="60">
                    <circle cx="35" cy="35" r="28" fill="none" stroke="#f1f5f9" stroke-width="5"/>
                    <circle cx="35" cy="35" r="28" fill="none" stroke="$ExecDnsRingColor" stroke-width="5" stroke-dasharray="$ExecDnsRingDash" stroke-linecap="round" transform="rotate(-90 35 35)"/>
                    <text x="35" y="39" text-anchor="middle" font-size="13" font-weight="500" fill="#0f172a" font-family="Segoe UI, Arial, sans-serif">$ExecDnsValue</text>
                </svg>
                <div class="exec-kpi-status" style="color:$ExecDnsSubColor;margin-top:6px;">$ExecDnsSubText</div>
            </div>

            <div class="exec-tile">
                <div class="exec-kpi-label-row" style="justify-content:center;">
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="$ExecReplRingColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12a9 9 0 1 1-2.64-6.36"/><polyline points="21 3 21 9 15 9"/></svg>
                    <div class="exec-ring-label" style="margin-bottom:0;">Replication Health</div>
                </div>
                <svg viewBox="0 0 70 70" width="60" height="60">
                    <circle cx="35" cy="35" r="28" fill="none" stroke="#f1f5f9" stroke-width="5"/>
                    <circle cx="35" cy="35" r="28" fill="none" stroke="$ExecReplRingColor" stroke-width="5" stroke-dasharray="$ExecReplRingDash" stroke-linecap="round" transform="rotate(-90 35 35)"/>
                    <text x="35" y="39" text-anchor="middle" font-size="12" font-weight="500" fill="#0f172a" font-family="Segoe UI, Arial, sans-serif">$ExecReplValue</text>
                </svg>
                <div class="exec-kpi-status" style="color:$ExecReplSubColor;margin-top:6px;">$ExecReplSubText</div>
            </div>
        </div>

        <div class="exec-kpi-toggle">
            <button class="exec-kpi-toggle-btn active" id="execKpiBtnBars" onclick="toggleExecKpiStyle('bars')" title="Bar style view">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><line x1="4" y1="6" x2="20" y2="6"/><line x1="4" y1="12" x2="14" y2="12"/><line x1="4" y1="18" x2="17" y2="18"/></svg>
            </button>
            <button class="exec-kpi-toggle-btn" id="execKpiBtnDonuts" onclick="toggleExecKpiStyle('donuts')" title="Circle chart view">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="9"/></svg>
            </button>
        </div>
        </div>

        <div class="exec-row2">
            <div class="exec-tile exec-chart-card">
                <div class="exec-chart-header">
                    <div class="exec-chart-title">Group Policies</div>
                    <div class="exec-chart-total" style="color:$ExecGpoSubColor;">$ExecGpoSubText</div>
                </div>
                <svg viewBox="0 0 240 130" width="100%" height="130" preserveAspectRatio="xMidYMid meet">
                    <defs>
                        <linearGradient id="gpoBlueGrad" x1="0" y1="0" x2="0" y2="1">
                            <stop offset="0%" stop-color="#60a5fa"/>
                            <stop offset="100%" stop-color="#2563eb"/>
                        </linearGradient>
                        <linearGradient id="gpoGreenGrad" x1="0" y1="0" x2="0" y2="1">
                            <stop offset="0%" stop-color="#4ade80"/>
                            <stop offset="100%" stop-color="#16a34a"/>
                        </linearGradient>
                        <linearGradient id="gpoAmberGrad" x1="0" y1="0" x2="0" y2="1">
                            <stop offset="0%" stop-color="#fbbf24"/>
                            <stop offset="100%" stop-color="#d97706"/>
                        </linearGradient>
                    </defs>
                    <line x1="30" y1="$ExecChartTop" x2="30" y2="$ExecChartBottom" stroke="#e2e8f0" stroke-width="1"/>
                    <line x1="30" y1="$ExecChartBottom" x2="230" y2="$ExecChartBottom" stroke="#e2e8f0" stroke-width="1"/>
                    <text x="26" y="19" text-anchor="end" font-size="8" fill="#94a3b8">$ExecGpoAxisMaxLabel</text>
                    <text x="26" y="62" text-anchor="end" font-size="8" fill="#94a3b8">$ExecGpoAxisMidLabel</text>
                    <text x="26" y="108" text-anchor="end" font-size="8" fill="#94a3b8">0</text>
                    <rect x="55" y="$ExecGpoTotalBarY" width="34" height="$ExecGpoTotalBarH" fill="url(#gpoBlueGrad)" rx="3"/>
                    <text x="72" y="$ExecGpoTotalLabelY" text-anchor="middle" font-size="12" font-weight="800" fill="#0C447C">$ExecGpoTotalNum</text>
                    <rect x="115" y="$ExecGpoLinkedBarY" width="34" height="$ExecGpoLinkedBarH" fill="url(#gpoGreenGrad)" rx="3"/>
                    <text x="132" y="$ExecGpoLinkedLabelY" text-anchor="middle" font-size="12" font-weight="800" fill="#166534">$ExecGpoLinkedNum</text>
                    <rect x="175" y="$ExecGpoUnlinkedBarY" width="34" height="$ExecGpoUnlinkedBarH" fill="url(#gpoAmberGrad)" rx="3"/>
                    <text x="192" y="$ExecGpoUnlinkedLabelY" text-anchor="middle" font-size="12" font-weight="800" fill="#854d0e">$ExecGpoUnlinkedNum</text>
                    <text x="72" y="120" text-anchor="middle" font-size="9" fill="#64748b">Total</text>
                    <text x="132" y="120" text-anchor="middle" font-size="9" fill="#64748b">Linked</text>
                    <text x="192" y="120" text-anchor="middle" font-size="9" fill="#64748b">Unlinked</text>
                </svg>
            </div>

            <div class="exec-tile exec-chart-card">
                <div class="exec-chart-header">
                    <div class="exec-chart-title">Security Risks</div>
                    <div class="exec-chart-total" style="color:$ExecRiskSubColor;">$ExecRiskSubText</div>
                </div>
                <div style="font-size:11px;color:#64748b;margin:-4px 0 6px;">Baseline compliance: <span style="font-weight:700;color:$BaselineComplianceColor;">$BaselineComplianceText</span></div>
                <svg viewBox="0 0 240 130" width="100%" height="130" preserveAspectRatio="xMidYMid meet">
                    <defs>
                        <linearGradient id="secBlueGrad" x1="0" y1="0" x2="0" y2="1">
                            <stop offset="0%" stop-color="#60a5fa"/>
                            <stop offset="100%" stop-color="#2563eb"/>
                        </linearGradient>
                        <linearGradient id="secRedGrad" x1="0" y1="0" x2="0" y2="1">
                            <stop offset="0%" stop-color="#f87171"/>
                            <stop offset="100%" stop-color="#dc2626"/>
                        </linearGradient>
                        <linearGradient id="secAmberGrad" x1="0" y1="0" x2="0" y2="1">
                            <stop offset="0%" stop-color="#fbbf24"/>
                            <stop offset="100%" stop-color="#d97706"/>
                        </linearGradient>
                    </defs>
                    <line x1="30" y1="$ExecChartTop" x2="30" y2="$ExecChartBottom" stroke="#e2e8f0" stroke-width="1"/>
                    <line x1="30" y1="$ExecChartBottom" x2="230" y2="$ExecChartBottom" stroke="#e2e8f0" stroke-width="1"/>
                    <text x="26" y="19" text-anchor="end" font-size="8" fill="#94a3b8">$ExecSecAxisMaxLabel</text>
                    <text x="26" y="62" text-anchor="end" font-size="8" fill="#94a3b8">$ExecSecAxisMidLabel</text>
                    <text x="26" y="108" text-anchor="end" font-size="8" fill="#94a3b8">0</text>
                    <rect x="55" y="$ExecSecTotalBarY" width="34" height="$ExecSecTotalBarH" fill="url(#secBlueGrad)" rx="3"/>
                    <text x="72" y="$ExecSecTotalLabelY" text-anchor="middle" font-size="12" font-weight="800" fill="#0C447C">$ExecSecTotalNum</text>
                    <rect x="115" y="$ExecHighBarY" width="34" height="$ExecHighBarH" fill="url(#secRedGrad)" rx="3"/>
                    <text x="132" y="$ExecHighLabelY" text-anchor="middle" font-size="12" font-weight="800" fill="#991b1b">$ExecHighCount</text>
                    <rect x="175" y="$ExecMediumBarY" width="34" height="$ExecMediumBarH" fill="url(#secAmberGrad)" rx="3"/>
                    <text x="192" y="$ExecMediumLabelY" text-anchor="middle" font-size="12" font-weight="800" fill="#854d0e">$ExecMediumCount</text>
                    <text x="72" y="120" text-anchor="middle" font-size="9" fill="#64748b">Total</text>
                    <text x="132" y="120" text-anchor="middle" font-size="9" fill="#64748b">High</text>
                    <text x="192" y="120" text-anchor="middle" font-size="9" fill="#64748b">Medium</text>
                </svg>
            </div>
        </div>

        <div class="exec-bottom-row">
            <div class="exec-card">
                <div class="exec-card-title-row">
                    <div class="exec-card-title" style="margin-bottom:0">Top Risks Requiring Attention</div>
                    $TopRisksExportBtn
                </div>
                $TopRisksHtml
            </div>

            <div class="exec-card">
                <div class="exec-card-title">Health Score Trend</div>
                $ExecTrendHtml
            </div>
        </div>

        <div style="height:60px;"></div>

    </div>

    </div>

    <div id="tab-forest" class="tab-content">

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

            <div class="config-tile">
                <div class="config-label">Schema Version</div>
                <div class="config-value">$SchemaVersion</div>
                <div class="config-footer">$SchemaVersionLabel (objectVersion on CN=Schema)</div>
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

    <div class="section-title">Global AD Security Health</div>

    <div id="globalSecHealthGrid" class="dashboard"></div>

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

    <div id="tab-identityrisk" class="tab-content">

        <div class="section-title">Identity Risk &amp; Attack Surface</div>

        <div class="dashboard">

            <div class="tile accent-orange" id="idTileStale" onclick="filterIdentityUsersTable('stale', this)" style="cursor:pointer;">
                <div class="tile-top">
                    <div class="tile-title">Stale Accounts</div>
                    <div class="tile-icon">90d</div>
                </div>
                <div>
                    <div class="tile-value">$StaleCount</div>
                    <div class="tile-footer">Click to filter table &middot; hover for accounts</div>
                </div>
                <div class="tooltip-box">$StaleAccountsTooltip</div>
            </div>

            <div class="tile accent-red" id="idTilePrivileged" onclick="filterIdentityUsersTable('privileged', this)" style="cursor:pointer;">
                <div class="tile-top">
                    <div class="tile-title">Privileged Accounts</div>
                    <div class="tile-icon">T0</div>
                </div>
                <div>
                    <div class="tile-value">$PrivAccountsCount</div>
                    <div class="tile-footer">Click to filter table &middot; hover for accounts</div>
                </div>
                <div class="tooltip-box">$PrivAccountsTooltip</div>
            </div>

            <div class="tile accent-purple" id="idTileService" onclick="filterIdentityUsersTable('service', this)" style="cursor:pointer;">
                <div class="tile-top">
                    <div class="tile-title">Risky Service Accounts</div>
                    <div class="tile-icon">SPN</div>
                </div>
                <div>
                    <div class="tile-value">$KerberoastCount</div>
                    <div class="tile-footer">Click to filter table &middot; hover for accounts</div>
                </div>
                <div class="tooltip-box">$RiskyServiceAccountsTooltip</div>
            </div>

            <div class="tile accent-red" id="idTileHighRisk" onclick="filterIdentityUsersTable('highrisk', this)" style="cursor:pointer;">
                <div class="tile-top">
                    <div class="tile-title">High Risk Identities</div>
                    <div class="tile-icon">!</div>
                </div>
                <div>
                    <div class="tile-value">$HighRiskIdentitiesCount</div>
                    <div class="tile-footer">Click to filter table &middot; hover for accounts</div>
                </div>
                <div class="tooltip-box">$HighRiskIdentitiesTooltip</div>
            </div>

        </div>

        <div style="max-width:950px;margin:28px auto 0;display:flex;align-items:baseline;justify-content:space-between;">
            <div class="section-title" id="idTableSectionTitle" style="margin:0;">Critical &amp; High Risk Users</div>
            <a href="javascript:void(0)" id="idShowAllLink" onclick="filterIdentityUsersTable(null, null)" style="font-size:12px;color:#2563eb;text-decoration:none;display:none;">&#8634; Show all (consolidated view)</a>
        </div>

        <div id="identityUsersTable" data-theme="A" style="max-width:950px;margin:0 auto;">
            <div class="rpt-tbl-wrap">
                <table class="rpt-styled-table" style="width:100%;table-layout:fixed;">
                    <colgroup>
                        <col style="width:16%;">
                        <col style="width:8%;">
                        <col style="width:9%;">
                        <col style="width:30%;">
                        <col style="width:37%;">
                    </colgroup>
                    <thead>
                        <tr class="rpt-header-row">
                            <th>Name <span style="font-weight:400;opacity:0.6;font-size:10px;">(click to expand)</span></th>
                            <th>Type</th>
                            <th>Risk</th>
                            <th>Why</th>
                            <th>Recommended Action</th>
                        </tr>
                    </thead>
                    <tbody id="idUsersTbody">
                        <tr><td colspan="5" style="text-align:center;padding:20px;color:#94a3b8;font-style:italic;">Loading...</td></tr>
                    </tbody>
                </table>
            </div>
        </div>
        <div style="max-width:950px;margin:8px auto 0;font-size:11px;color:#94a3b8;" id="idTableNote">&#8505; Sorted Critical &#8594; High &#8594; Medium. Every Critical/High/Medium identity finding is shown - this list is not capped. "Type" is Service if the account has an SPN, otherwise User. Matches whichever Reports style (A/B) is currently selected.</div>

    </div>

    <div id="tab-groupsprivilege" class="tab-content">

        <div class="section-title">Group &amp; Privilege Architecture</div>

        <div class="dashboard">

            <div class="tile accent-red" id="gpaTileMembership" onclick="filterGroupPrivTable('membership', this)" style="cursor:pointer;">
                <div class="tile-top">
                    <div class="tile-title">Tier-0 Group Membership</div>
                    <div class="tile-icon">T0</div>
                </div>
                <div>
                    <div class="tile-value">$Tier0MemberFindingsCount</div>
                    <div class="tile-footer">Across 7 privileged groups &middot; click to filter</div>
                </div>
            </div>

            <div class="tile accent-orange" id="gpaTileNesting" onclick="filterGroupPrivTable('nesting', this)" style="cursor:pointer;">
                <div class="tile-top">
                    <div class="tile-title">Nesting Depth &amp; Cycles</div>
                    <div class="tile-icon">&#8595;&#8595;</div>
                </div>
                <div>
                    <div class="tile-value">$($NestingDepthFindingsCount + $GroupCycleCount)</div>
                    <div class="tile-footer">Deep chains + circular references &middot; click to filter</div>
                </div>
            </div>

            <div class="tile accent-purple" id="gpaTileFsp" onclick="filterGroupPrivTable('fsp', this)" style="cursor:pointer;">
                <div class="tile-top">
                    <div class="tile-title">Foreign Security Principals</div>
                    <div class="tile-icon">FSP</div>
                </div>
                <div>
                    <div class="tile-value">$FspFindingsCount</div>
                    <div class="tile-footer">External principals in Tier-0 &middot; click to filter</div>
                </div>
            </div>

            <div class="tile accent-red" id="gpaTileDelegation" onclick="filterGroupPrivTable('delegation', this)" style="cursor:pointer;">
                <div class="tile-top">
                    <div class="tile-title">Group-Based Delegation</div>
                    <div class="tile-icon">KEY</div>
                </div>
                <div>
                    <div class="tile-value">$Tier0AclFindingsCount</div>
                    <div class="tile-footer">Incl. Reset Password rights &middot; click to filter</div>
                </div>
            </div>

            <div class="tile accent-cyan" id="gpaTileLegacy" onclick="filterGroupPrivTable('legacy', this)" style="cursor:pointer;">
                <div class="tile-top">
                    <div class="tile-title">Legacy &amp; Universal Risk</div>
                    <div class="tile-icon">OLD</div>
                </div>
                <div>
                    <div class="tile-value">$($UniversalTier0Count + $PreWin2000Count)</div>
                    <div class="tile-footer">Pre-Win2000 + Universal nesting &middot; click to filter</div>
                </div>
            </div>

        </div>

        <div style="max-width:950px;margin:28px auto 0;display:flex;align-items:baseline;justify-content:space-between;">
            <div class="section-title" id="gpaTableSectionTitle" style="margin:0;">All Group &amp; Privilege Findings</div>
            <a href="javascript:void(0)" id="gpaShowAllLink" onclick="filterGroupPrivTable(null, null)" style="font-size:12px;color:#2563eb;text-decoration:none;display:none;">&#8634; Show all</a>
        </div>

        <div id="groupPrivTable" data-theme="A" style="max-width:950px;margin:0 auto;">
            <div class="rpt-tbl-wrap">
                <table class="rpt-styled-table" style="width:100%;table-layout:fixed;">
                    <colgroup>
                        <col style="width:16%;">
                        <col style="width:8%;">
                        <col style="width:9%;">
                        <col style="width:30%;">
                        <col style="width:37%;">
                    </colgroup>
                    <thead>
                        <tr class="rpt-header-row">
                            <th>Name <span style="font-weight:400;opacity:0.6;font-size:10px;">(click to expand)</span></th>
                            <th>Type</th>
                            <th>Risk</th>
                            <th>Why</th>
                            <th>Recommended Action</th>
                        </tr>
                    </thead>
                    <tbody id="gpaTbody">
                        <tr><td colspan="5" style="text-align:center;padding:20px;color:#94a3b8;font-style:italic;">Loading...</td></tr>
                    </tbody>
                </table>
            </div>
        </div>
        <div style="max-width:950px;margin:8px auto 0;font-size:11px;color:#94a3b8;" id="gpaTableNote">&#8505; One row per group - covers Domain Admins/Enterprise Admins/Schema Admins/Administrators/Backup Operators/Account Operators/Print Operators, plus nesting, delegation, foreign security principals, and legacy group hygiene. Click a row to see the full member/finding list. Per-account detail for these same accounts also lives on the Identity Risk &amp; Attack Surface tab.</div>

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

    <div id="tab-reports" class="tab-content">
        <div id="rptPlaceholder" style="text-align:center;padding:60px 20px;color:#94a3b8;">
            <div style="font-size:42px;margin-bottom:14px;">&#9783;</div>
            <div style="font-size:18px;font-weight:700;color:#475569;margin-bottom:8px;">Select a Report</div>
            <div style="font-size:14px;">Hover over the <strong>Reports</strong> tab above and choose a report from the menu.</div>
        </div>
        <div id="rptPanel" style="display:none;" data-theme="A">
            <div class="rpt-report-header">
                <div class="rpt-report-title" id="rptTitle"></div>
                <div style="display:flex;gap:8px;">
                    <button class="rpt-export-btn" id="rptThemeToggleBtn" onclick="toggleReportTheme()">&#127912; Switch Style (B)</button>
                    <button class="rpt-export-btn" onclick="exportCurrentReport()">&#8659; Export to Excel</button>
                </div>
            </div>
            <div class="rpt-tbl-wrap">
                <table id="rptTable" class="rpt-styled-table" style="width:100%;margin:0 0 25px 0;">
                    <thead id="rptThead"></thead>
                    <tbody id="rptTbody"></tbody>
                </table>
            </div>
            <div class="rpt-note" id="rptNote"></div>
        </div>
    </div>


</div>

<div class="footer">
    <span class="footer-line">Enterprise Active Directory Health Dashboard | Generated using PowerShell</span>
    <span class="footer-line footer-copyright">&nbsp;&nbsp;|&nbsp;&nbsp;&copy; $(Get-Date -Format yyyy) Azgar Mohammad. All rights reserved. | For internal IT operations use only.</span>
</div>

<script id="dc-list-data" type="application/json">$DCListJson</script>
<script id="dc-inventory-data" type="application/json">$DCInventoryJson</script>
<script id="global-ad-security-data" type="application/json">$GlobalADSecurityJson</script>

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

// Builds a CSV on the fly from the currently-displayed DC's live
// Replication data (dcInventoryByName is populated further down once the
// dc-inventory-data JSON is parsed - this function isn't called until a
// user clicks the export button, well after that has happened).
function exportReplicationPartnersCsv(dcName) {
    const inv = dcInventoryByName[dcName];
    if (!inv || !inv.Replication || inv.Replication.length === 0) { return; }

    function csvCell(v) {
        const s = String(v === null || v === undefined ? '' : v);
        return '"' + s.replace(/"/g, '""') + '"';
    }

    const rows = ['Partner,Partition,Last Success,Failures,Last Result'];
    inv.Replication.forEach(function (r) {
        rows.push([csvCell(r.Partner), csvCell(r.Partition), csvCell(r.LastSuccess), csvCell(r.ConsecutiveFailures), csvCell(r.LastResult)].join(','));
    });

    downloadCsv('Replication_Partners_' + dcName + '.csv', rows.join('\r\n'));
}

const domainControllersCsv = '$DomainControllersCsvJs';
const dcPingCsv = '$DCPingCsvJs';
const topRisksCsv = '$TopRisksCsvJs';
const securityRisksCsv = '$SecurityRisksCsvJs';
const identityFindingsCsv = '$IdentityFindingsCsvJs';
const identityUsersCsv = '$IdentityUsersCsvJs';
const gpaGroupRowsCsv = '$GpaGroupRowsCsvJs';
const adSitesCsv = '$ADSitesCsvJs';
const sitesWithNoDCsCsv = '$SitesWithNoDCsCsvJs';
const adSubnetsCsv = '$ADSubnetsCsvJs';
const adSiteLinksCsv = '$ADSiteLinksCsvJs';
const adSiteLinkBridgeFlag = '$ADSiteLinkBridgeFlagJs';
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
const globalADSecurityData = JSON.parse(document.getElementById('global-ad-security-data').textContent || '{}');

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
        ? '<span class="health health-critical">Reboot Pending</span>'
        : '<span class="health health-good">No Reboot Pending</span>';

    html += '<div class="dc-tile">' + tileTitle('monitor', 'System & Uptime') +
        kvRow('Operating System', sys.Caption) +
        kvRow('OS Build', sys.BuildNumber) +
        kvRow('Last Boot Time', sys.LastBootTime) +
        kvRow('Uptime', sys.UptimeReadable) +
        '<div class="dc-kv-row"><span>Pending Reboot Status</span><span>' + rebootBadge + '</span></div>' +
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

    // Replication partners - capped to ~5 visible rows via scroll, with an
    // export button (same pattern as the Forest Overview tiles) for the
    // complete list.
    const hasReplData = inv.Replication && inv.Replication.length > 0;
    const replExportBtn = hasReplData
        ? '<button class="export-btn" title="Export Replication Partners" onclick="exportReplicationPartnersCsv(\'' + meta.Name.replace(/'/g, "\\'") + '\')">&#8659;</button>'
        : '';

    html += '<div class="dc-tile dc-tile-wide">' + tileTitle('refresh', 'Replication Partners', replExportBtn);

    if (hasReplData) {
        html += '<div class="dc-repl-scroll"><table class="dc-mini-table"><tr><th>Partner</th><th>Partition</th><th>Last Success</th><th>Failures</th><th>Last Result</th></tr>';

        inv.Replication.forEach(function (r) {
            const failCls = (r.ConsecutiveFailures && r.ConsecutiveFailures !== '0') ? 'health-warning' : 'health-good';

            html += '<tr><td>' + escapeHtml(r.Partner) + '</td><td>' + escapeHtml(r.Partition) + '</td><td>' + escapeHtml(r.LastSuccess) + '</td>' +
                '<td><span class="health ' + failCls + '">' + escapeHtml(r.ConsecutiveFailures) + '</span></td><td>' + escapeHtml(r.LastResult) + '</td></tr>';
        });

        html += '</table></div>';
        html += '<div class="dc-repl-note">&#8505; For complete list of replication partners status download the report</div>';
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

// =====================================================================
// Global AD Security Health section - Forest Overview tab
// Renders 7 one-time domain/forest-level security check tiles into
// the #globalSecHealthGrid div.
// =====================================================================

function renderGlobalADSecurity() {
    const grid = document.getElementById('globalSecHealthGrid');
    if (!grid) { return; }

    const d = globalADSecurityData || {};

    // Build a tile matching the Forest Overview static tile structure exactly:
    // tile-top with text title + tile-icon abbreviation, health badge, tooltip-box
    function secTile(accentClass, abbr, title, statusText, healthClass, footerText, tooltipHtml) {
        const badge = '<span class="health health-' + healthClass + '">'
                    + healthClass.charAt(0).toUpperCase() + healthClass.slice(1)
                    + '</span>';
        const tip = tooltipHtml ? '<div class="tooltip-box">' + tooltipHtml + '</div>' : '';
        return '<div class="tile ' + accentClass + '">'
             + '<div class="tile-top">'
             +   '<div class="tile-title">' + escapeHtml(title) + '</div>'
             +   '<div class="tile-icon">' + escapeHtml(abbr) + '</div>'
             + '</div>'
             + '<div>'
             +   '<div class="tile-value" style="font-size:13px;line-height:1.4;margin-top:4px">' + escapeHtml(statusText) + '</div>'
             +   badge
             +   '<div class="tile-footer">' + escapeHtml(footerText) + '</div>'
             + '</div>'
             + tip
             + '</div>';
    }

    // Split a semicolon-delimited account/object list into individual tooltip-line rows
    function listToTip(label, raw) {
        if (!raw) { return ''; }
        var items = raw.split(';').map(function(s) { return s.trim(); }).filter(Boolean);
        if (items.length === 0) { return ''; }
        return items.map(function(item) {
            return '<div class="tooltip-line"><b>' + escapeHtml(label) + '</b><span>' + escapeHtml(item) + '</span></div>';
        }).join('');
    }

    function health(h) {
        if (!h) { return 'unknown'; }
        return h.toLowerCase();
    }

    var tiles = '';

    // 1. FGPP
    var fgpp = d.FGPP || {};
    var fgppFooter = (fgpp.Count > 0 && fgpp.Names) ? fgpp.Names : 'No fine-grained password policies defined';
    var fgppTip = fgpp.Names ? listToTip('PSO', fgpp.Names) : '';
    tiles += secTile('accent-blue', 'FG', 'Fine-Grained Password Policies',
        fgpp.Count != null ? fgpp.Count + ' PSO(s) Configured' : 'N/A',
        fgpp.Count > 0 ? 'good' : 'unknown',
        fgppFooter, fgppTip);

    // 2. Admin RID-500
    var adm = d.AdminRID500 || {};
    var admTip = adm.Name
        ? '<div class="tooltip-line"><b>SamAccountName</b><span>' + escapeHtml(adm.Name) + '</span></div>'
        : '';
    tiles += secTile('accent-purple', 'AD', 'Default Admin Account (RID-500)',
        adm.Status || 'N/A', health(adm.Health),
        'Rename + disable to reduce target surface', admTip);

    // 3. Guest Account
    var gst = d.GuestAccount || {};
    tiles += secTile(gst.Enabled ? 'accent-red' : 'accent-green', 'GU', 'Guest Account',
        gst.Status || 'N/A', health(gst.Health),
        'Should be disabled on all domain members', '');

    // 4. Protected Users Coverage
    var pu = d.ProtectedUsers || {};
    var puTip = listToTip('Not in Protected Users', pu.ExposedAdmins || '');
    tiles += secTile(pu.ExposedCount > 0 ? 'accent-red' : 'accent-green', 'PU', 'Protected Users Coverage',
        pu.Status || 'N/A', health(pu.Health),
        'DA/EA/Schema Admins in Protected Users group', puTip);

    // 5. SPN Duplicates
    var spn = d.SPNDuplicates || {};
    tiles += secTile(spn.Count > 0 ? 'accent-red' : 'accent-green', 'SPN', 'SPN Duplicates',
        spn.Status || 'N/A', health(spn.Health),
        'Duplicate SPNs cause Kerberos service failures', '');

    // 6. Pre-Auth Disabled (AS-REP Roastable)
    var pa = d.PreAuthDisabled || {};
    var paTip = listToTip('AS-REP Roastable', pa.Accounts || '');
    tiles += secTile(pa.Count > 0 ? 'accent-red' : 'accent-green', 'PA', 'Pre-Auth Disabled Accounts',
        pa.Status || 'N/A', health(pa.Health),
        'AS-REP roastable - no Kerberos pre-authentication', paTip);

    // 7. Unconstrained Delegation
    var ud = d.UnconstrainedDelegation || {};
    var udTip = listToTip('Object', ud.Objects || '');
    tiles += secTile(ud.Count > 0 ? 'accent-red' : 'accent-green', 'UD', 'Unconstrained Delegation',
        ud.Status || 'N/A', health(ud.Health),
        'Non-DC accounts with TrustedForDelegation=True', udTip);

    // 8. Password Never Expires
    var pne = d.PwdNeverExpires || {};
    var pneTip = listToTip('Account', pne.Accounts || '');
    tiles += secTile(pne.Count > 0 ? (pne.Count > 5 ? 'accent-red' : 'accent-orange') : 'accent-green', 'PW', 'Password Never Expires',
        pne.Status || 'N/A', health(pne.Health),
        'Enabled accounts with DONT_EXPIRE_PASSWORD flag set', pneTip);

    grid.innerHTML = tiles;
}

renderGlobalADSecurity();

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

function installedRolesStatus(roles) {
    if (!roles || !roles.Core) { return 'unknown'; }
    if (roles.Extra && roles.Extra.length > 0) { return 'warning'; }

    return 'good';
}

function firewallStatus(fw) {
    const profiles = (fw && fw.Profiles) || [];

    if (profiles.length === 0) { return 'unknown'; }

    const domain = profiles.find(function (p) { return p.Name === 'Domain'; });
    if (domain && !domain.Enabled) { return 'critical'; }
    if (profiles.some(function (p) { return !p.Enabled; })) { return 'warning'; }

    return 'good';
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
    consider(installedRolesStatus(sp.InstalledRoles));
    consider(firewallStatus(sp.Firewall));

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

    // 1. Installed server roles (AD DS best-practice check) &#8212; full-width, first
    const installedRoles = sp.InstalledRoles || {};
    const rolesSt = installedRolesStatus(installedRoles);
    const coreRoles = installedRoles.Core || [];
    const extraRoles = installedRoles.Extra || [];
    let rolesTip = '';

    if (rolesSt === 'warning') {
        rolesTip = secTip(
            'This domain controller has Windows Server roles installed beyond AD DS and DNS. Running additional roles on a DC increases its attack surface and can make patching and maintenance riskier, since the DC cannot easily be taken offline.',
            'Where possible, move non-AD DS roles (file services, certificate services, etc.) to dedicated member servers and keep domain controllers running only AD DS, DNS, and their management tools.',
            'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/reducing-the-active-directory-attack-surface',
            'Reducing the Active Directory attack surface &ndash; Microsoft Learn'
        );
    }

    html += '<div class="dc-tile dc-tile-wide">' + tileTitle('package', 'Installed Server Roles', dcPill(rolesSt));

    if (coreRoles.length > 0 || extraRoles.length > 0) {
        html += '<div class="dc-role-section-label">Core AD DS roles</div><div class="dc-role-list">';
        coreRoles.forEach(function (r) {
            html += '<span class="dc-role-badge dc-role-badge-good">' + escapeHtml(r) + '</span>';
        });
        html += '</div>';

        if (extraRoles.length > 0) {
            html += '<div class="dc-role-section-label">Additional roles</div><div class="dc-role-list">';
            extraRoles.forEach(function (r) {
                html += '<span class="dc-role-badge dc-role-badge-warning">' + escapeHtml(r) + '</span>';
            });
            html += '</div>';
        }
    }
    else {
        html += kvRow('No installed-roles data collected', '-');
    }

    html += rolesTip + '</div>';

    // 2a. Windows Firewall profile status (paired with NetBIOS in row 2)
    const fw = sp.Firewall || {};
    const fwSt = firewallStatus(fw);
    const fwProfiles = fw.Profiles || [];
    let fwTip = '';

    if (fwSt === 'critical') {
        fwTip = secTip(
            'The Windows Firewall Domain profile is disabled on this domain controller. Without the Domain profile active, the DC is exposed to lateral movement and unauthorized access on the internal network.',
            'Enable the Windows Firewall Domain profile immediately. Use Group Policy (Computer Configuration > Windows Settings > Security Settings > Windows Firewall with Advanced Security) to enforce firewall profiles on all DCs.',
            'https://learn.microsoft.com/en-us/windows/security/threat-protection/windows-firewall/best-practices-configuring',
            'Windows Firewall best practices &ndash; Microsoft Learn'
        );
    }
    else if (fwSt === 'warning') {
        fwTip = secTip(
            'One or more Windows Firewall profiles (Private or Public) are disabled on this domain controller. While the Domain profile may be active, disabled profiles reduce defence-in-depth.',
            'Enable all three firewall profiles (Domain, Private, Public) on domain controllers as a defence-in-depth measure.',
            'https://learn.microsoft.com/en-us/windows/security/threat-protection/windows-firewall/best-practices-configuring',
            'Windows Firewall best practices &ndash; Microsoft Learn'
        );
    }

    html += '<div class="dc-tile">' + tileTitle('shield', 'Windows Firewall Profiles', dcPill(fwSt));

    if (fwProfiles.length > 0) {
        fwProfiles.forEach(function (p) {
            // "Disabled" always renders the same color regardless of which
            // profile it is - severity (Domain off = more severe than
            // Private/Public off) is already conveyed once via the tile's
            // overall Critical/Warning badge above, not per-row.
            const cls = p.Enabled ? 'health-good' : 'health-critical';
            const val = p.Enabled ? 'Enabled' : 'Disabled';
            html += '<div class="dc-kv-row"><span>' + escapeHtml(p.Name) + ' Profile</span><span class="health ' + cls + '">' + val + '</span></div>';
        });
    }
    else {
        html += kvRow('No firewall profile data collected', '-');
    }

    html += fwTip + '</div>';

    // 2b. NetBIOS over TCP/IP (NetBT) &#8212; sits beside Firewall in row 2
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

    // 3. SMB Protocol Security | 4. TLS / Schannel (Row 3 — paired)
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

    // 6. TLS / Schannel
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

    // 5. LDAPS certificate
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

    // 6. LDAP signing
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

    // 7. NTLM hardening
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

    // 8. Audit policy coverage
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

    // 9. Recent hotfix installed updates (full-width, last)
    const hotfixes = sp.Hotfixes || [];
    const rebootPending = !!(sys && sys.PendingReboot);

    html += '<div class="dc-tile dc-tile-wide">' + tileTitle('refresh', 'Recent Hotfix Installed Updates',
        rebootPending ? '<span class="dc-pill dc-pill-critical">Reboot Pending</span>' : '');

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
            '<div style="margin-top:8px; padding:10px 12px; border-radius:6px; background:#fef2f2; color:#b91c1c; font-size:12.5px;">' +
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

function dfsrStatus(arr, shares, maxOffline) {
    if (!arr || arr.length === 0) { return 'unknown'; }

    let result = 'good';

    if (arr.some(function (f) { return f.State === 'In Error'; })) { result = 'critical'; }
    else if (!arr.every(function (f) { return f.State === 'Normal'; })) { result = 'warning'; }

    if (shares && shares.NETLOGON && shares.NETLOGON.Status !== 'Shared') { result = 'critical'; }

    if (maxOffline !== null && !isNaN(maxOffline)) {
        const rank = { good: 0, warning: 1, critical: 2 };
        const offlineRank = (maxOffline < 30) ? 'critical' : (maxOffline < 60) ? 'warning' : 'good';
        if (rank[offlineRank] > rank[result]) { result = offlineRank; }
    }

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
    consider(dfsrStatus(di.DFSRBacklog, di.SharesHealth, di.MaxOfflineTimeInDays));
    consider(backupStatus(inv.ADBackup));
    consider(dnsHealthStatus(di.DNSHealth));
    consider(eventLogStatus(di.EventLogs && di.EventLogs.DirectoryService));
    consider(eventLogStatus(di.EventLogs && di.EventLogs.DFSReplication));
    consider(eventLogStatus(di.EventLogs && di.EventLogs.DNSServer));

    // MaxTokenSize and StrictReplicationConsistency (moved from Security Posture)
    const sp2 = inv.SecurityPosture || {};
    const mts2 = sp2.MaxTokenSize;
    consider(mts2 === null || mts2 === undefined ? 'unknown' : (mts2 >= 65535 ? 'good' : 'critical'));
    const src2 = sp2.StrictReplicationConsistency;
    consider(src2 === null || src2 === undefined ? 'unknown' : (src2 === 1 ? 'good' : 'critical'));

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

    // 1. Connectivity Health Check - simplified Ping/LDAP/LDAPS pass-fail
    // summary intended for non-technical audiences. Deliberately separate
    // from the technical "LDAP Signing" (Security Posture) and "LDAPS
    // Certificate" tiles, which require interpretation. Shown first as the
    // "vitals" check before the deeper DCDIAG diagnostics.
    const pingOk = inv.ICMPStatus === 'Reachable';
    const pingKnown = inv.ICMPStatus !== undefined && inv.ICMPStatus !== null;

    const dnsInfo = inv.DNS || {};
    const dnsKnown = !!dnsInfo.Status;
    const dnsOk = dnsInfo.Status === 'Pass';

    const kerberosInfo = inv.Kerberos || {};
    const kerberosKnown = !!kerberosInfo.Status;
    const kerberosOk = kerberosInfo.Status === 'Pass';

    const ldapInfo = inv.LDAP || {};
    const ldapKnown = !!ldapInfo.Status;
    const ldapOk = ldapInfo.Status === 'Pass';

    const ldapsInfo = inv.LDAPS || {};
    const ldapsKnown = !!ldapsInfo.Status && ldapsInfo.Status !== 'Unknown';
    const ldapsOk = !!ldapsInfo.PortOpen && (ldapsInfo.Status === 'OK' || ldapsInfo.Status === 'Warning');

    const isGC = !!meta.IsGlobalCatalog;
    const gcInfo = inv.GlobalCatalog || {};
    const gcKnown = isGC && !!gcInfo.Status;
    const gcOk = gcInfo.Status === 'Pass';

    const gcSslInfo = inv.GlobalCatalogSSL || {};
    const gcSslKnown = isGC && !!gcSslInfo.Status;
    const gcSslOk = gcSslInfo.Status === 'Pass';

    const adwsInfo = inv.ADWS || {};
    const adwsKnown = !!adwsInfo.Status;
    const adwsOk = adwsInfo.Status === 'Pass';

    const timeSyncInfo = inv.TimeSync || {};
    const timeKnown = !!timeSyncInfo.Status && timeSyncInfo.Status !== 'Unknown';
    const timeOk = timeSyncInfo.Status === 'OK';

    const smbInfo = inv.SMB || {};
    const smbKnown = !!smbInfo.Status;
    const smbOk = smbInfo.Status === 'Pass';

    const connRows = [
        { label: 'Ping (ICMP)', known: pingKnown, ok: pingOk, pass: 'Passed', fail: 'Failed' },
        { label: 'DNS (53)', known: dnsKnown, ok: dnsOk, pass: 'Passed', fail: 'Failed' },
        { label: 'Kerberos (88)', known: kerberosKnown, ok: kerberosOk, pass: 'Passed', fail: 'Failed' },
        { label: 'LDAP (389)', known: ldapKnown, ok: ldapOk, pass: 'Passed', fail: 'Failed' },
        { label: 'LDAPS (636)', known: ldapsKnown, ok: ldapsOk, pass: 'Passed', fail: 'Failed' },
        { label: 'GC (3268)', known: gcKnown, ok: gcOk, pass: 'Passed', fail: 'Failed', na: !isGC },
        { label: 'GC SSL (3269)', known: gcSslKnown, ok: gcSslOk, pass: 'Passed', fail: 'Failed', na: !isGC },
        { label: 'ADWS (9389)', known: adwsKnown, ok: adwsOk, pass: 'Passed', fail: 'Failed' },
        { label: 'Time (123)', known: timeKnown, ok: timeOk, pass: 'Passed', fail: 'Failed' },
        { label: 'SMB (445)', known: smbKnown, ok: smbOk, pass: 'Passed', fail: 'Failed' }
    ];

    let connSt = 'good';
    if (connRows.every(function (r) { return r.na || !r.known; })) {
        connSt = 'unknown';
    }
    else if (connRows.some(function (r) { return !r.na && r.known && !r.ok; })) {
        connSt = 'critical';
    }

    html += '<div class="dc-tile dc-tile-wide">' + tileTitle('network', 'Connectivity Health Check', dcPill(connSt)) + '<div class="dc-conn-grid">';

    connRows.forEach(function (r) {
        let cls, val;
        if (r.na) {
            cls = 'health-unknown';
            val = 'N/A';
        }
        else if (!r.known) {
            cls = 'health-unknown';
            val = 'Unknown';
        }
        else {
            cls = r.ok ? 'health-good' : 'health-critical';
            val = r.ok ? r.pass : r.fail;
        }
        html += '<div class="dc-conn-cell"><span class="dc-conn-label">' + escapeHtml(r.label) + '</span><span class="health ' + cls + '">' + escapeHtml(val) + '</span></div>';
    });

    html += '</div></div>';

    // 2. DCDIAG summary
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

    // 3. SYSVOL / DFSR replication health
    const dfsrArr = di.DFSRBacklog || [];
    const shares = di.SharesHealth || {};
    const maxOfflineDays = (di.MaxOfflineTimeInDays !== undefined && di.MaxOfflineTimeInDays !== null) ? parseInt(di.MaxOfflineTimeInDays, 10) : null;
    const dfsrSt = dfsrStatus(dfsrArr, shares, maxOfflineDays);
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
            const label = (f.ReplicatedFolderName === 'SYSVOL Share') ? 'SYSVOL Status' : f.ReplicatedFolderName;
            html += '<div class="dc-kv-row"><span>' + escapeHtml(label) + '</span><span class="health ' + cls + '">' + escapeHtml(f.State) + '</span></div>';
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

    if (maxOfflineDays !== null && !isNaN(maxOfflineDays)) {
        const offlineCls = (maxOfflineDays >= 60) ? 'health-good' : (maxOfflineDays >= 30) ? 'health-warning' : 'health-critical';
        html += '<div class="dc-kv-row"><span>Max Offline Time</span><span class="health ' + offlineCls + '">' + maxOfflineDays + ' days</span></div>';
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

    // 5. MaxTokenSize | 6. Strict Replication Consistency (moved from Security Posture)
    const spDI = inv.SecurityPosture || {};
    const mts = spDI.MaxTokenSize;
    const mtsGood = (mts !== null && mts !== undefined && mts >= 65535);
    const mtsStatus = mts === null || mts === undefined ? 'unknown' : (mtsGood ? 'good' : 'critical');

    const mtsTip = secTip(
        'MaxTokenSize controls the maximum Kerberos token buffer. The default value of 12,000 bytes is too small for accounts in 150+ groups, causing silent authentication failures (HTTP 400, IIS 401 loops).',
        'Set MaxTokenSize = 65535 on all DCs and member servers per KB2732618. A system reboot is required after the change.',
        'https://learn.microsoft.com/en-us/troubleshoot/windows-server/windows-security/maxtokensize-registry-entry-kb2732618',
        'KB2732618 &ndash; MaxTokenSize &ndash; Microsoft Learn'
    );

    html += '<div class="dc-tile">' + tileTitle('lock', 'MaxTokenSize', dcPill(mtsStatus));
    html += kvRow('Current Value', mts !== null && mts !== undefined ? String(mts) : 'Not Set (OS Default 12000)');
    html += kvRow('Recommended', '65535 (KB2732618)');
    html += kvRow('Status', mtsGood ? 'Optimal' : (mts === null || mts === undefined ? 'Not configured - OS default applies' : 'Below recommended value - users in 150+ groups at risk'));
    html += mtsTip;
    html += '</div>';

    const src = spDI.StrictReplicationConsistency;
    const srcGood = (src === 1);
    const srcStatus = src === null || src === undefined ? 'unknown' : (srcGood ? 'good' : 'critical');

    const srcTip = secTip(
        'Strict Replication Consistency (value 1) prevents a DC from accepting objects from a source DC that has been offline long enough to accumulate lingering objects. When set to 0 (loose mode), these stale objects silently propagate to the rest of the forest.',
        'Set HKLM\\SYSTEM\\CurrentControlSet\\Services\\NTDS\\Parameters\\"Strict Replication Consistency" = 1 on all DCs. This is the default on Windows Server 2008 R2 and later, but may be disabled on upgraded environments.',
        'https://learn.microsoft.com/en-us/troubleshoot/windows-server/active-directory/information-lingering-objects',
        'Lingering objects &ndash; Microsoft Learn'
    );

    html += '<div class="dc-tile">' + tileTitle('shield', 'Strict Replication Consistency', dcPill(srcStatus));
    html += kvRow('Registry Value', src !== null && src !== undefined ? String(src) : 'Not Set');
    html += kvRow('Required Value', '1 (Strict Mode)');
    html += kvRow('Status', srcGood ? 'Strict mode enabled - lingering objects blocked' : (src === 0 ? 'Loose mode - lingering objects can propagate' : 'Not configured - check registry manually'));
    html += srcTip;
    html += '</div>';

    // 7. Lingering Object Events (1388 / 1988) - targeted sweep separate from generic error tile
    const lingeringEvents = di.LingeringObjectEvents;
    const lingeringStatus = !Array.isArray(lingeringEvents) ? 'unknown'
        : lingeringEvents.length === 0 ? 'good' : 'critical';

    html += '<div class="dc-tile dc-tile-wide">' + tileTitle('alert', 'Lingering Object Events (1388 / 1988)', dcPill(lingeringStatus));

    if (!Array.isArray(lingeringEvents)) {
        html += kvRow('Status', 'Data not yet collected - re-run Get-DCInventory.ps1');
    } else if (lingeringEvents.length === 0) {
        html += '<div class="dc-kv-row"><span style="color:var(--color-success,#16a34a)">&#10003; No lingering object events found in Directory Service log</span></div>';
        html += '<div style="margin-top:6px;font-size:12px;color:var(--muted,#6b7280)">Event 1388 = object accepted in loose mode &bull; Event 1988 = object blocked in strict mode</div>';
    } else {
        html += '<div style="margin-bottom:6px;font-size:12px;color:#b91c1c">Lingering objects detected - run <code>repadmin /removelingeringobjects</code> to remediate.</div>';
        html += '<table class="dc-mini-table"><tr><th>Time</th><th>Event ID</th><th>Message</th></tr>';
        lingeringEvents.forEach(function (e) {
            html += '<tr><td>' + escapeHtml(e.Time) + '</td><td>' + escapeHtml(String(e.EventID)) + '</td><td>' + escapeHtml(e.Message) + '</td></tr>';
        });
        html += '</table>';
    }
    html += '</div>';

    // 6. USN Rollback Events (2095) - snapshot restore / split-brain detection
    const usnEvents = di.USNRollbackEvents;
    const usnStatus = !Array.isArray(usnEvents) ? 'unknown'
        : usnEvents.length === 0 ? 'good' : 'critical';

    html += '<div class="dc-tile dc-tile-wide">' + tileTitle('alert', 'USN Rollback Events (2095)', dcPill(usnStatus));

    if (!Array.isArray(usnEvents)) {
        html += kvRow('Status', 'Data not yet collected - re-run Get-DCInventory.ps1');
    } else if (usnEvents.length === 0) {
        html += '<div class="dc-kv-row"><span style="color:var(--color-success,#16a34a)">&#10003; No USN rollback events found in Directory Service log</span></div>';
        html += '<div style="margin-top:6px;font-size:12px;color:var(--muted,#6b7280)">Event 2095 indicates an improper VM snapshot restore caused AD split-brain replication</div>';
    } else {
        html += '<div style="margin-bottom:6px;font-size:12px;color:#b91c1c">USN rollback detected - this DC may have a split-brain replication issue. Immediate investigation required.</div>';
        html += '<table class="dc-mini-table"><tr><th>Time</th><th>Event ID</th><th>Message</th></tr>';
        usnEvents.forEach(function (e) {
            html += '<tr><td>' + escapeHtml(e.Time) + '</td><td>' + escapeHtml(String(e.EventID)) + '</td><td>' + escapeHtml(e.Message) + '</td></tr>';
        });
        html += '</table>';
    }
    html += '</div>';

    // 7. Recent interactive logons (full-width)
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

    // 8-10. Top 5 recent error-level events per log (full-width, last)
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

// =====================================================================
// Reports tab
// =====================================================================

var currentReport = null;

// Sort state for the Security Risks Need Attention report - persists across
// re-renders so clicking a header twice flips direction. Defaults to
// severity order (High, then Medium, then Low) until the user picks a
// column to sort by instead.
var secRisksSortCol = null;
var secRisksSortDir = 1;

function toggleSecRisksSort(col) {
    if (secRisksSortCol === col) {
        secRisksSortDir = -secRisksSortDir;
    } else {
        secRisksSortCol = col;
        secRisksSortDir = 1;
    }
    renderSecurityRisksReport();
}

// ---- dropdown hover wiring (delayed close to bridge the hover gap) ----
(function () {
    var wrap  = document.getElementById('rptTabWrap');
    var dd    = document.getElementById('rptDropdown');
    var timer = null;

    function openMenu()  { if (timer) { clearTimeout(timer); timer = null; } dd.style.display = 'block'; }
    function startClose(){ timer = setTimeout(function () { dd.style.display = 'none'; }, 250); }

    wrap.addEventListener('mouseenter', openMenu);
    wrap.addEventListener('mouseleave', startClose);
    dd.addEventListener('mouseenter',   openMenu);
    dd.addEventListener('mouseleave',   startClose);
})();

var rptLabels = {
    health:        'DC Health Check',
    inventory:     'DC Server Inventory',
    dcuptime:      'DC Uptime & Reboot Status',
    dcbysite:      'DC List by Site',
    replpartners:  'AD Replication Partners',
    sites:        'Sites & Services',
    gpo:          'Group Policies',
    dnszones:     'DNS Zones',
    forwarders:   'Conditional Forwarders',
    security:     'Security Posture Matrix',
    connectivity: 'Connectivity Matrix',
    secrisks:     'Security Risks Need Attention',
    identitysec:  'Identity & Kerberos Security'
};

// Top-level tab display names for the page ribbon. This is intentionally
// separate from rptLabels (the Reports sub-menu names) - the ribbon always
// reflects which top-level TAB is active, never which report is loaded.
var tabDisplayNames = {
    execsummary:  'Executive Summary',
    forest:       'Forest Overview',
    dchealth:     'Domain Controller Health',
    deepinsights: 'Deep Insights',
    secposture:   'Security Posture',
    identityrisk: 'Identity Risk & Attack Surface',
    groupsprivilege: 'Group & Privilege Architecture',
    reports:      'Reports'
};

function switchReport(type) {
    currentReport = type;
    document.querySelectorAll('.rpt-dd-item').forEach(function (el) { el.classList.remove('active'); });
    document.getElementById('rdd-' + type).classList.add('active');
    // Note: the Reports tab button label and the page ribbon intentionally
    // stay "Reports" no matter which report is selected from the dropdown -
    // only the report panel content below changes. Do not overwrite the
    // tab-btn-reports label or the ribbon here.
    document.getElementById('rptPlaceholder').style.display = 'none';
    document.getElementById('rptPanel').style.display = 'block';
    renderReport(type);
    document.getElementById('rptDropdown').style.display = 'none';
    // Make sure the Reports tab is active
    showTab('reports');
}

// The ribbon TITLE is pinned to a fixed x-position - directly under the
// label text of the FIRST tab (Executive Summary) - regardless of which
// tab is actually active. This is intentional per the agreed design.
function positionRibbonTitle() {
    var ribbon     = document.getElementById('pageRibbon');
    var firstLabel = document.querySelector('#tab-btn-execsummary .tab-label');
    var title      = document.getElementById('ribbonTitle');
    if (!ribbon || !firstLabel || !title) { return; }
    var ribbonRect = ribbon.getBoundingClientRect();
    var labelRect  = firstLabel.getBoundingClientRect();
    title.style.left = (labelRect.left - ribbonRect.left) + 'px';
}

// The ribbon CRUMB's trailing letter is pinned to a fixed x-position too -
// directly under the trailing letter of the LAST tab's label (Reports),
// regardless of which tab is actually active. Same fixed-anchor concept as
// positionRibbonTitle above, just anchored to the last tab instead of the
// first, and aligning the right edge instead of the left edge.
function positionRibbonCrumb() {
    var ribbon    = document.getElementById('pageRibbon');
    var lastLabel = document.querySelector('#tab-btn-reports .tab-label');
    var crumb     = document.getElementById('ribbonCrumb');
    if (!ribbon || !lastLabel || !crumb) { return; }
    var ribbonRect = ribbon.getBoundingClientRect();
    var labelRect  = lastLabel.getBoundingClientRect();
    var rightEdge  = labelRect.right - ribbonRect.left;
    var crumbWidth = crumb.getBoundingClientRect().width;
    crumb.style.left  = (rightEdge - crumbWidth) + 'px';
    crumb.style.right = 'auto';
}

function showTab(tabName) {
    document.querySelectorAll('.tab-content').forEach(function (el) { el.classList.remove('active'); });
    document.querySelectorAll('.tab-button').forEach(function (el) { el.classList.remove('active'); });
    document.getElementById('tab-' + tabName).classList.add('active');
    var activeBtn = document.getElementById('tab-btn-' + tabName);
    if (activeBtn) { activeBtn.classList.add('active'); }

    // Move the pointer arrow above the strip to sit centered over the active tab
    var strip = document.getElementById('tabIndicatorStrip');
    var arrow = document.getElementById('tabIndicatorArrow');
    if (strip && arrow && activeBtn) {
        var stripRect = strip.getBoundingClientRect();
        var btnRect   = activeBtn.getBoundingClientRect();
        var centerX   = (btnRect.left + btnRect.width / 2) - stripRect.left;
        arrow.style.left = centerX + 'px';
    }

    // Update the page ribbon - always reflects the top-level tab, never a
    // sub-report selected from the Reports dropdown.
    var ribbonTitle = document.getElementById('ribbonTitle');
    var ribbonCrumb = document.getElementById('ribbonCrumb');
    var displayName = tabDisplayNames[tabName] || tabName;
    if (ribbonTitle) { ribbonTitle.textContent = displayName; }
    if (ribbonCrumb) { ribbonCrumb.innerHTML = 'Dashboard&nbsp;/&nbsp;<b>' + displayName + '</b>'; }

    // Re-measure both fixed anchors: title under the first tab's label,
    // crumb under the last tab's (Reports) label - neither depends on
    // which tab is actually active.
    positionRibbonTitle();
    positionRibbonCrumb();

    // Reports: do NOT auto-render &#8212; user must pick from the dropdown menu
}

// ---- helpers ----
function rptBadge(text, cls) { return '<span class="rpt-b ' + cls + '">' + String(text) + '</span>'; }
function rptGood(t)   { return rptBadge(t, 'rpt-good'); }
function rptWarn(t)   { return rptBadge(t, 'rpt-warn'); }
function rptMedium(t) { return rptBadge(t, 'rpt-medium'); }
function rptCrit(t)   { return rptBadge(t, 'rpt-crit'); }
function rptNa(t)     { return rptBadge(t, 'rpt-na'); }

function passBadge(status, passText, failText) {
    if (!status) { return rptNa('Unknown'); }
    return status === 'Pass' ? rptGood(passText || 'Passed') : rptCrit(failText || 'Failed');
}

function svcBadge(status) {
    if (!status) { return rptNa('Unknown'); }
    var s = String(status).toLowerCase();
    if (s === 'running') { return rptGood('Running'); }
    if (s === 'stopped') { return rptCrit('Stopped'); }
    return rptWarn(status);
}

function diskPct(free, total) {
    if (!total || total === 0) { return null; }
    return Math.round((free / total) * 100);
}

function diskPctCell(pct) {
    if (pct === null || pct === undefined) { return ''; }
    var cls = pct < 15 ? 'rpt-disk-crit' : pct < 30 ? 'rpt-disk-warn' : 'rpt-disk-ok';
    return '<span class="' + cls + '">' + pct + '%</span>';
}

function diskGBCell(gb, pct) {
    if (gb === null || gb === undefined || gb === '') { return ''; }
    var cls = (pct !== null && pct < 15) ? 'rpt-disk-crit' : (pct !== null && pct < 30) ? 'rpt-disk-warn' : 'rpt-disk-ok';
    return '<span class="' + cls + '">' + gb + '</span>';
}

function machineTypeBadge(hw) {
    if (!hw) { return rptNa('Unknown'); }
    var model = String(hw.Model || hw.ComputerModel || '').toLowerCase();
    var mfr   = String(hw.Manufacturer || '').toLowerCase();
    if (model.indexOf('vmware') !== -1 || mfr.indexOf('vmware') !== -1) { return '<span class="rpt-b rpt-info">VMware</span>'; }
    if (model.indexOf('virtual') !== -1 || mfr.indexOf('microsoft') !== -1) { return '<span class="rpt-b rpt-info">Hyper-V</span>'; }
    return '<span class="rpt-b rpt-good">Physical</span>';
}

function dcdiagResult(di, testName) {
    if (!di || !di.DCDiag) { return rptNa('N/A'); }
    var dd = di.DCDiag;
    // DCDiag is stored as a plain object: { NetLogons: "Passed", Replications: "Passed", ... }
    var keys = Object.keys(dd);
    var result = null;
    for (var i = 0; i < keys.length; i++) {
        if (keys[i].toLowerCase() === testName.toLowerCase()) { result = dd[keys[i]]; break; }
    }
    if (result === null || result === undefined) { return rptNa('N/A'); }
    var r = String(result).toLowerCase();
    if (r === 'passed') { return rptGood('Passed'); }
    if (r === 'failed') { return rptCrit('Failed'); }
    return rptWarn(String(result));
}

function getServiceStatus(inv, svcName) {
    var svcs = inv.Services || [];
    for (var i = 0; i < svcs.length; i++) {
        var n = String(svcs[i].Name || svcs[i].ServiceName || '').toLowerCase();
        if (n === svcName.toLowerCase()) { return svcs[i].Status || svcs[i].State || ''; }
    }
    return '';
}

function getDisk(inv, driveLetter) {
    var disks = inv.Disks || [];
    var letter = driveLetter.replace(':', '').toUpperCase();
    for (var i = 0; i < disks.length; i++) {
        var dl = String(disks[i].DeviceID || disks[i].Drive || disks[i].DriveLetter || '').replace(':', '').toUpperCase();
        if (dl === letter) { return disks[i]; }
    }
    return null;
}

// The grouped header row (column-category banner above the column names) is
// disabled dashboard-wide by returning an empty string here. All report
// renderers still call buildGroupRow([...]) + buildHeaderRow([...]) - this
// single change removes the group row from every report at once, so each
// renderer didn't need to be edited individually.
function buildGroupRow(groups) {
    return '';
}

function buildHeaderRow(cols) {
    var html = '<tr class="rpt-header-row">';
    cols.forEach(function (c) {
        var cls = c.divider ? ' rpt-gd' : '';
        html += '<th class="' + cls + '">' + c.label + '</th>';
    });
    return html + '</tr>';
}

function buildDataRow(cells) {
    var html = '<tr>';
    cells.forEach(function (c) {
        var cls = (c.dc ? 'rpt-dc' : '') + (c.divider ? ' rpt-gd' : '');
        html += '<td class="' + cls.trim() + '">' + (c.html !== undefined ? c.html : escapeHtml(String(c.val !== undefined ? c.val : ''))) + '</td>';
    });
    return html + '</tr>';
}

// ---- report renderers ----
function renderHealthReport() {
    document.getElementById('rptTitle').textContent = 'DC Health Check';
    document.getElementById('rptNote').innerHTML = '&#8505; DCDiag results sourced from Deep Insights collection &#8212; no additional overhead.';

    var thead = buildGroupRow([
        { span: 1, empty: true },
        { span: 3, label: 'Connectivity' },
        { span: 6, label: 'Services', divider: true },
        { span: 4, label: 'DCDiag Tests', divider: true }
    ]) + buildHeaderRow([
        { label: 'DC Name' },
        { label: 'Ping' }, { label: 'LDAP (389)' }, { label: 'LDAPS (636)' },
        { label: 'DNS Service',      divider: true },
        { label: 'Netlogon Service' },
        { label: 'NTDS (AD DS)' },
        { label: 'KDC Service' },
        { label: 'DFSR Service' },
        { label: 'W32TM Service' },
        { label: 'Netlogon Test',    divider: true },
        { label: 'Replication Test' },
        { label: 'Advertising Test' },
        { label: 'FSMO Check' }
    ]);

    var tbody = '';
    dcInventoryData.forEach(function (inv) {
        var ldapInfo  = inv.LDAP  || {};
        var ldapsInfo = inv.LDAPS || {};
        var pingOk    = inv.ICMPStatus === 'Reachable';
        var di        = inv.DeepInsights || {};

        tbody += buildDataRow([
            { html: escapeHtml(inv.ComputerName), dc: true },
            { html: pingOk ? rptGood('Passed') : rptCrit('Failed') },
            { html: (ldapInfo.Status === 'Pass') ? rptGood('Passed') : (ldapInfo.Status === 'Fail' ? rptCrit('Failed') : rptNa('Unknown')) },
            { html: (ldapsInfo.PortOpen) ? rptGood('Passed') : rptCrit('Failed') },
            { html: svcBadge(getServiceStatus(inv, 'DNS')),      divider: true },
            { html: svcBadge(getServiceStatus(inv, 'Netlogon')) },
            { html: svcBadge(getServiceStatus(inv, 'NTDS')) },
            { html: svcBadge(getServiceStatus(inv, 'Kdc')) },
            { html: svcBadge(getServiceStatus(inv, 'DFSR')) },
            { html: svcBadge(getServiceStatus(inv, 'W32Time')) },
            { html: dcdiagResult(di, 'NetLogons'),     divider: true },
            { html: dcdiagResult(di, 'Replications') },
            { html: dcdiagResult(di, 'Advertising') },
            { html: dcdiagResult(di, 'FSMOCheck') }
        ]);
    });

    document.getElementById('rptThead').innerHTML = thead;
    document.getElementById('rptTbody').innerHTML = tbody;
}

function renderInventoryReport() {
    document.getElementById('rptTitle').textContent = 'DC Server Inventory';
    document.getElementById('rptNote').innerHTML =
        '<span style="color:#854d0e;font-weight:600;">Amber</span> = C: or D: free space 15&#8211;30% &nbsp;&middot;&nbsp; ' +
        '<span style="color:#991b1b;font-weight:600;">Red</span> = free space below 15% &nbsp;&middot;&nbsp; ' +
        'Blank D: columns = drive not present on this DC';

    var thead = buildGroupRow([
        { span: 3, empty: true },
        { span: 2, label: 'Hardware' },
        { span: 3, label: 'Drive C:', divider: true },
        { span: 3, label: 'Drive D:', divider: true },
        { span: 1, label: 'Machine',  divider: true }
    ]) + buildHeaderRow([
        { label: 'DC Name' },
        { label: 'IPv4 Address' },
        { label: 'Operating System' },
        { label: 'CPU Cores' },
        { label: 'RAM (GB)' },
        { label: 'C: Total (GB)', divider: true },
        { label: 'C: Free (GB)' },
        { label: 'C: Free (%)' },
        { label: 'D: Total (GB)', divider: true },
        { label: 'D: Free (GB)' },
        { label: 'D: Free (%)' },
        { label: 'Machine Type',  divider: true }
    ]);

    var tbody = '';
    dcInventoryData.forEach(function (inv) {
        var sys    = inv.System   || {};
        var hw     = inv.Hardware || {};
        var netArr = Array.isArray(inv.Network) ? inv.Network : [];
        var netPri = netArr[0] || {};

        // IP: prefer EthernetSettings.IPv4Address (primary adapter), fall back to Network[0].IPAddress
        var ip = (inv.EthernetSettings && inv.EthernetSettings.IPv4Address)
               || netPri.IPAddress || '';

        // OS: collector stores Caption = full name, e.g. "Microsoft Windows
        // Server 2019 Standard Edition" - shown as-is, not abbreviated.
        var osVer = String(sys.Caption || 'Unknown');

        var cDisk = getDisk(inv, 'C');
        var dDisk = getDisk(inv, 'D');

        var cTotal = cDisk ? Math.round((cDisk.SizeGB || cDisk.TotalGB || cDisk.Size || 0)) : null;
        var cFree  = cDisk ? Math.round((cDisk.FreeGB || cDisk.FreeSpaceGB || 0)) : null;
        var cPct   = (cTotal && cFree !== null) ? diskPct(cFree, cTotal) : null;

        var dTotal = dDisk ? Math.round((dDisk.SizeGB || dDisk.TotalGB || dDisk.Size || 0)) : null;
        var dFree  = dDisk ? Math.round((dDisk.FreeGB || dDisk.FreeSpaceGB || 0)) : null;
        var dPct   = (dTotal && dFree !== null) ? diskPct(dFree, dTotal) : null;

        tbody += buildDataRow([
            { html: escapeHtml(inv.ComputerName), dc: true },
            { val: ip },
            { val: osVer },
            { val: hw.NumberOfCores || '' },
            { val: hw.TotalMemoryGB || '' },
            { html: cTotal !== null ? String(cTotal) : '',          divider: true },
            { html: cFree  !== null ? diskGBCell(cFree, cPct) : '' },
            { html: cPct   !== null ? diskPctCell(cPct) : '' },
            { html: dTotal !== null ? String(dTotal) : '',          divider: true },
            { html: dFree  !== null ? diskGBCell(dFree, dPct) : '' },
            { html: dPct   !== null ? diskPctCell(dPct) : '' },
            { html: machineTypeBadge(hw),                           divider: true }
        ]);
    });

    document.getElementById('rptThead').innerHTML = thead;
    document.getElementById('rptTbody').innerHTML = tbody;
}

function renderDcBySiteReport() {
    document.getElementById('rptTitle').textContent = 'DC List by Site';
    document.getElementById('rptNote').innerHTML =
        '&#8505; Domain Controllers sorted by AD Site. ' +
        'Total: <strong>' + dcListData.length + '</strong> DC' + (dcListData.length !== 1 ? 's' : '') +
        ' across <strong>' + (function() {
            var s = {}; dcListData.forEach(function(d) { s[d.Site || ''] = 1; }); return Object.keys(s).length;
        })() + '</strong> site(s).';

    var thead = buildGroupRow([
        { span: 3, label: 'Domain Controller Details' },
        { span: 1, label: 'AD Site', divider: true }
    ]) + buildHeaderRow([
        { label: 'DC Name' },
        { label: 'IPv4 Address' },
        { label: 'Operating System' },
        { label: 'AD Site', divider: true }
    ]);

    // Sort by Site, then DC Name within each site
    var sorted = dcListData.slice().sort(function(a, b) {
        var sA = (a.Site || '').toLowerCase();
        var sB = (b.Site || '').toLowerCase();
        if (sA < sB) { return -1; }
        if (sA > sB) { return  1; }
        return (a.Name || '').toLowerCase().localeCompare((b.Name || '').toLowerCase());
    });

    var tbody = '';
    sorted.forEach(function(dc) {
        // OS: shorten to "Windows Server YYYY"
        var os = String(dc.OperatingSystem || '');
        var osMatch = os.match(/\b(2008|2012|2016|2019|2022|2025)\b/);
        if (osMatch) { os = 'Windows Server ' + osMatch[0]; }

        tbody += buildDataRow([
            { html: escapeHtml(dc.Name || ''), dc: true },
            { val: dc.IPv4Address || '' },
            { val: os },
            { html: '<span class="rpt-b rpt-info">' + escapeHtml(dc.Site || '&#8212;') + '</span>', divider: true }
        ]);
    });

    if (sorted.length === 0) {
        tbody = '<tr><td colspan="4" style="text-align:center;padding:20px;color:#94a3b8;font-style:italic;">No DC data available</td></tr>';
    }

    document.getElementById('rptThead').innerHTML = thead;
    document.getElementById('rptTbody').innerHTML = tbody;
}

function renderReplPartnersReport() {
    document.getElementById('rptTitle').textContent = 'AD Replication Partners';
    document.getElementById('rptNote').innerHTML =
        '&#8505; One row per DC &#8594; partner pair. Result and failure count reflect the <strong>worst state across all naming contexts</strong> ' +
        '(Domain, Schema, Configuration). Last Attempt = most recent of last success or last failure time. Data sourced from repadmin /showrepl per DC.';

    var thead = buildGroupRow([
        { span: 2, label: 'Domain Controller' },
        { span: 2, label: 'Replication Partner', divider: true },
        { span: 2, label: 'Last Replication', divider: true },
        { span: 3, label: 'Status', divider: true }
    ]) + buildHeaderRow([
        { label: 'DC Name' },
        { label: 'DC Site' },
        { label: 'Partner DC',    divider: true },
        { label: 'Partner Site' },
        { label: 'Last Attempt',  divider: true },
        { label: 'Last Success' },
        { label: 'Result',        divider: true },
        { label: 'Consec. Failures' },
        { label: 'Transport' }
    ]);

    var tbody = '';
    var totalLinks = 0;
    var failLinks  = 0;

    dcInventoryData.forEach(function(inv) {
        var replArr = inv.Replication || [];
        var dcMeta  = dcListData.find(function(d) { return d.Name === inv.ComputerName; }) || {};
        var dcSite  = dcMeta.Site || '';

        if (replArr.length === 0) {
            tbody += buildDataRow([
                { html: escapeHtml(inv.ComputerName), dc: true },
                { val: dcSite },
                { html: '<span style="color:#94a3b8;font-style:italic;">No replication data collected</span>', divider: true },
                { val: '' }, { val: '', divider: true }, { val: '' },
                { html: rptNa('N/A'), divider: true }, { val: '' }, { val: '' }
            ]);
            return;
        }

        // Group by partner - deduplicate across naming contexts
        var byPartner = {};
        replArr.forEach(function(r) {
            var pKey = (r.Partner || '').toLowerCase();
            if (!byPartner[pKey]) {
                byPartner[pKey] = {
                    partner:    r.Partner     || '',
                    site:       r.PartnerSite || '',
                    transport:  r.Transport   || 'IP',
                    lastSuccess:    r.LastSuccess    || '',
                    lastFailTime:   r.LastFailureTime || '',
                    failures:   parseInt(r.ConsecutiveFailures, 10) || 0,
                    lastResult: parseInt(r.LastResult, 10) || 0
                };
            } else {
                // Worst result wins
                var res = parseInt(r.LastResult, 10) || 0;
                if (res !== 0 && byPartner[pKey].lastResult === 0) { byPartner[pKey].lastResult = res; }
                // Max failures
                var f = parseInt(r.ConsecutiveFailures, 10) || 0;
                if (f > byPartner[pKey].failures) { byPartner[pKey].failures = f; }
                // Most recent success
                if (r.LastSuccess > byPartner[pKey].lastSuccess) { byPartner[pKey].lastSuccess = r.LastSuccess; }
                // Most recent failure time
                if (r.LastFailureTime > byPartner[pKey].lastFailTime) { byPartner[pKey].lastFailTime = r.LastFailureTime; }
            }
        });

        Object.keys(byPartner).sort().forEach(function(pKey) {
            var p = byPartner[pKey];
            totalLinks++;

            // Last attempt = most recent of success or failure timestamp
            var lastAttempt = p.lastSuccess;
            if (p.lastFailTime && p.lastFailTime > lastAttempt) { lastAttempt = p.lastFailTime; }

            // Shorten partner name - strip domain suffix for display
            var partnerShort = p.partner.replace(/\.[^.]+\.[^.]+$/, '');

            var resultHtml;
            if (p.lastResult === 0 && p.failures === 0) {
                resultHtml = rptGood('&#10004; Success');
            } else if (p.lastResult !== 0) {
                resultHtml = rptCrit('&#10006; Error ' + p.lastResult);
                failLinks++;
            } else {
                resultHtml = rptWarn('Warning');
                failLinks++;
            }

            var failHtml = (p.failures === 0)
                ? '<span style="color:#6b7280;">0</span>'
                : (p.failures <= 2 ? rptWarn(p.failures) : rptCrit(p.failures));

            tbody += buildDataRow([
                { html: escapeHtml(inv.ComputerName), dc: true },
                { val: dcSite },
                { html: escapeHtml(partnerShort), divider: true },
                { val: p.site },
                { val: lastAttempt,    divider: true },
                { val: p.lastSuccess },
                { html: resultHtml,    divider: true },
                { html: failHtml },
                { val: p.transport || 'IP' }
            ]);
        });
    });

    if (!tbody) {
        tbody = '<tr><td colspan="9" style="text-align:center;padding:20px;color:#94a3b8;font-style:italic;">No replication data available. Ensure DC Health collection is enabled.</td></tr>';
    }

    document.getElementById('rptThead').innerHTML = thead;
    document.getElementById('rptTbody').innerHTML = tbody;

    // Append summary to note bar
    if (totalLinks > 0) {
        var summaryColor = failLinks > 0 ? '#991b1b' : '#166534';
        var summaryText  = failLinks > 0
            ? failLinks + ' link' + (failLinks !== 1 ? 's' : '') + ' failing'
            : 'All links healthy';
        document.getElementById('rptNote').innerHTML +=
            ' &nbsp;&middot;&nbsp; <strong>' + totalLinks + '</strong> replication link' + (totalLinks !== 1 ? 's' : '') +
            ' &nbsp;&middot;&nbsp; <span style="font-weight:600;color:' + summaryColor + ';">' + summaryText + '</span>';
    }
}

function renderSecurityReport() {
    document.getElementById('rptTitle').textContent = 'Security Posture Matrix';
    document.getElementById('rptNote').innerHTML = '&#8505; Data sourced from Security Posture collectors. Red = critical finding, Amber = warning.';

    var thead = buildGroupRow([
        { span: 1, empty: true },
        { span: 2, label: 'SMB' },
        { span: 1, label: 'TLS',            divider: true },
        { span: 2, label: 'LDAP Hardening', divider: true },
        { span: 3, label: 'Firewall',        divider: true },
        { span: 2, label: 'Other',           divider: true }
    ]) + buildHeaderRow([
        { label: 'DC Name' },
        { label: 'SMBv1' }, { label: 'SMB Signing' },
        { label: 'TLS Legacy',        divider: true },
        { label: 'LDAP Integrity',    divider: true },
        { label: 'Channel Binding' },
        { label: 'FW Domain',         divider: true },
        { label: 'FW Private' },
        { label: 'FW Public' },
        { label: 'NetBIOS',           divider: true },
        { label: 'Extra Roles' }
    ]);

    var tbody = '';
    dcInventoryData.forEach(function (inv) {
        var sp   = inv.SecurityPosture || {};
        var smb  = sp.SMB   || {};
        var tls  = sp.TLS   || [];
        var sig  = sp.LDAPSigning || {};
        var fw   = sp.Firewall    || {};
        var nb   = sp.NetBIOS     || {};
        var roles = sp.InstalledRoles || {};

        var smb1Html   = (smb.SMB1Enabled === true || smb.SMB1Enabled === 'True')  ? rptCrit('Enabled')  : rptGood('Disabled');
        var smbSignHtml = (smb.SigningRequired === true || smb.SigningRequired === 'True') ? rptGood('Required') : rptWarn('Not Required');

        var legacyTls = tls.filter(function (p) {
            var n = String(p.Protocol || '').toLowerCase();
            return (n.indexOf('ssl') !== -1 || n === 'tls 1.0' || n === 'tls1.0') && (p.Enabled === 1 || p.Enabled === true);
        });
        var tlsHtml = legacyTls.length > 0 ? rptCrit(legacyTls[0].Protocol) : rptGood('None');

        var intVal = sig.LDAPServerIntegrity;
        var intHtml = intVal === 2 ? rptGood('Require') : (intVal === 1 ? rptWarn('Negotiate') : rptCrit('None'));

        var cbVal = sig.LdapEnforceChannelBinding;
        var cbHtml = cbVal === 2 ? rptGood('Always') : (cbVal === 1 ? rptWarn('When supp.') : rptCrit('Never'));

        function fwPro(name) {
            var profs = fw.Profiles || [];
            for (var i = 0; i < profs.length; i++) {
                if (String(profs[i].Name || '').toLowerCase() === name.toLowerCase()) {
                    return profs[i].Enabled ? rptGood('On') : (name === 'Domain' ? rptCrit('Off') : rptWarn('Off'));
                }
            }
            return rptNa('N/A');
        }

        var nbAdapters = nb.Adapters || [];
        var nbEnabled  = nbAdapters.some(function (a) { return String(a.Setting || '').toLowerCase() === 'enabled'; });
        var nbHtml     = nbEnabled ? rptCrit('Enabled') : rptGood('Disabled');

        var extraCount = (roles.Extra || []).length;
        var rolesHtml  = extraCount > 0 ? rptWarn(extraCount) : rptGood('0');

        tbody += buildDataRow([
            { html: escapeHtml(inv.ComputerName), dc: true },
            { html: smb1Html }, { html: smbSignHtml },
            { html: tlsHtml,   divider: true },
            { html: intHtml,   divider: true },
            { html: cbHtml },
            { html: fwPro('Domain'),  divider: true },
            { html: fwPro('Private') },
            { html: fwPro('Public') },
            { html: nbHtml,    divider: true },
            { html: rolesHtml }
        ]);
    });

    document.getElementById('rptThead').innerHTML = thead;
    document.getElementById('rptTbody').innerHTML = tbody;
}

function renderConnectivityReport() {
    document.getElementById('rptTitle').textContent = 'Connectivity Matrix';
    document.getElementById('rptNote').innerHTML = '&#8505; N/A = check not applicable for this DC (e.g., GC ports on a non-GC DC).';

    var thead = buildGroupRow([
        { span: 1, empty: true },
        { span: 5, label: 'Core AD Ports' },
        { span: 3, label: 'GC / Extended', divider: true },
        { span: 2, label: 'Other',          divider: true },
        { span: 1, empty: true }
    ]) + buildHeaderRow([
        { label: 'DC Name' },
        { label: 'Ping' }, { label: 'DNS (53)' }, { label: 'Kerberos (88)' }, { label: 'LDAP (389)' }, { label: 'LDAPS (636)' },
        { label: 'GC (3268)',     divider: true },
        { label: 'GC SSL (3269)' },
        { label: 'ADWS (9389)' },
        { label: 'Time (123)',    divider: true },
        { label: 'SMB (445)' },
        { label: 'Score' }
    ]);

    var tbody = '';
    dcInventoryData.forEach(function (inv) {
        var isGC = inv.IsGlobalCatalog === true || inv.IsGlobalCatalog === 'True' ||
                   (inv.System && (inv.System.IsGlobalCatalog === true || inv.System.IsGlobalCatalog === 'True'));

        function portBadge(sectionName, na) {
            if (na) { return rptNa('N/A'); }
            var s = inv[sectionName] || {};
            if (!s.Status) { return rptNa('Unknown'); }
            return s.Status === 'Pass' ? rptGood('Passed') : rptCrit('Failed');
        }

        var pingOk  = inv.ICMPStatus === 'Reachable';
        var ldapInfo  = inv.LDAP            || {};
        var ldapsInfo = inv.LDAPS           || {};
        var gcInfo    = inv.GlobalCatalog   || {};
        var gcSslInfo = inv.GlobalCatalogSSL || {};
        var adwsInfo  = inv.ADWS            || {};
        // TimeSync.Status from w32tm: "OK" → Pass, anything else → Fail
        var timeInfo  = { Status: (inv.TimeSync && inv.TimeSync.Status === 'OK') ? 'Pass'
                                : (inv.TimeSync && inv.TimeSync.Status)           ? 'Fail' : null };
        var smbInfo   = inv.SMB             || {};

        var ldapHtml  = ldapInfo.Status === 'Pass' ? rptGood('Passed') : (ldapInfo.Status ? rptCrit('Failed') : rptNa('Unknown'));
        var ldapsHtml = ldapsInfo.PortOpen ? rptGood('Passed') : rptCrit('Failed');
        var gcHtml    = !isGC ? rptNa('N/A') : (gcInfo.Status === 'Pass' ? rptGood('Passed') : rptCrit('Failed'));
        var gcSslHtml = !isGC ? rptNa('N/A') : (gcSslInfo.Status === 'Pass' ? rptGood('Passed') : rptCrit('Failed'));
        var adwsHtml  = adwsInfo.Status === 'Pass' ? rptGood('Passed') : (adwsInfo.Status ? rptCrit('Failed') : rptNa('Unknown'));
        var timeHtml  = timeInfo.Status === 'Pass' ? rptGood('Passed') : (timeInfo.Status ? rptCrit('Failed') : rptNa('Unknown'));
        var smbHtml   = smbInfo.Status === 'Pass' ? rptGood('Passed') : (smbInfo.Status ? rptCrit('Failed') : rptNa('Unknown'));

        var coreChecks = [pingOk, ldapInfo.Status === 'Pass', !!ldapsInfo.PortOpen,
                          adwsInfo.Status === 'Pass', timeInfo.Status === 'Pass', smbInfo.Status === 'Pass'];
        var gcChecks   = isGC ? [gcInfo.Status === 'Pass', gcSslInfo.Status === 'Pass'] : [];
        var passed = coreChecks.concat(gcChecks).filter(Boolean).length;
        var total  = coreChecks.length + gcChecks.length;

        tbody += buildDataRow([
            { html: escapeHtml(inv.ComputerName), dc: true },
            { html: pingOk ? rptGood('Passed') : rptCrit('Failed') },
            { html: portBadge('DNS') },
            { html: portBadge('Kerberos') },
            { html: ldapHtml },
            { html: ldapsHtml },
            { html: gcHtml,    divider: true },
            { html: gcSslHtml },
            { html: adwsHtml },
            { html: timeHtml,  divider: true },
            { html: smbHtml },
            { val: passed + '/' + total }
        ]);
    });

    document.getElementById('rptThead').innerHTML = thead;
    document.getElementById('rptTbody').innerHTML = tbody;
}

// ---- DNS zone CSV parser ----
function parseDnsZonesCsv() {
    var raw = (typeof dnsZonesCsv !== 'undefined') ? dnsZonesCsv : '';
    if (!raw) { return []; }
    var lines = raw.split('\n').filter(function (l) { return l.trim(); });
    if (lines.length < 2) { return []; }
    var headers = lines[0].replace(/"/g, '').split(',');
    var rows = [];
    for (var i = 1; i < lines.length; i++) {
        var vals = lines[i].match(/("([^"]*)"|([^,]*))(,|$)/g) || [];
        var obj = {};
        vals.forEach(function (v, idx) {
            var clean = v.replace(/,$/,'').replace(/^"|"$/g,'');
            if (headers[idx]) { obj[headers[idx].trim()] = clean; }
        });
        rows.push(obj);
    }
    return rows;
}

function renderDnsZonesReport() {
    document.getElementById('rptTitle').textContent = 'DNS Zones';
    document.getElementById('rptNote').innerHTML =
        '&#8505; Forward lookup zones are listed first, followed by reverse lookup zones. AD-Integrated zones replicate automatically &#8212; no zone transfers needed.';

    var thead = buildGroupRow([
        { span: 2, empty: true },
        { span: 2, label: 'Zone Classification', divider: true },
        { span: 1, label: 'Replication',         divider: true },
        { span: 1, label: 'Updates',             divider: true }
    ]) + buildHeaderRow([
        { label: 'Zone Name' },
        { label: 'Lookup Type' },
        { label: 'Zone Type',         divider: true },
        { label: 'AD Integrated' },
        { label: 'Replication Scope', divider: true },
        { label: 'Dynamic Updates',   divider: true }
    ]);

    var zones = parseDnsZonesCsv().filter(function (z) { return (z.ZoneType || '') !== 'Forwarder'; });

    // Compute the reverse-lookup flag once, then sort Forward zones before
    // Reverse zones. Array.sort is stable, so relative order within each
    // group is preserved exactly as collected.
    zones.forEach(function (z) {
        z._isRev = (z.IsReverseLookupZone || '').toLowerCase() === 'true'
            || (z.ZoneName || '').indexOf('.in-addr.arpa') !== -1
            || (z.ZoneName || '').indexOf('.ip6.arpa') !== -1;
    });
    zones.sort(function (a, b) { return (a._isRev ? 1 : 0) - (b._isRev ? 1 : 0); });

    var tbody = '';

    zones.forEach(function (z) {
        var zType = z.ZoneType || '';

        var typePill = '';
        if (zType === 'Primary')   { typePill = '<span class="rpt-b rpt-purple">Primary</span>'; }
        else if (zType === 'Secondary') { typePill = '<span class="rpt-b rpt-na">Secondary</span>'; }
        else if (zType === 'Stub') { typePill = '<span class="rpt-b rpt-pink">Stub</span>'; }
        else { typePill = '<span class="rpt-b rpt-na">' + escapeHtml(zType) + '</span>'; }

        var isAD = (z.IsDsIntegrated || '').toLowerCase() === 'true';
        var adPill = isAD
            ? '<span class="rpt-b rpt-info">AD-Integrated</span>'
            : '<span class="rpt-b rpt-na">Standard</span>';

        var lookupPill = z._isRev
            ? '<span class="rpt-b rpt-warn">Reverse</span>'
            : '<span class="rpt-b rpt-good">Forward</span>';

        var scope = isAD ? (z.ReplicationScope || 'Domain') : 'N/A';
        var scopeCls = (scope === 'Forest') ? 'rpt-info' : (scope === 'Domain') ? 'rpt-purple' : '';
        var scopeHtml = scope === 'N/A'
            ? '<span class="rpt-b rpt-na">N/A</span>'
            : '<span class="rpt-b ' + scopeCls + '">' + escapeHtml(scope) + '</span>';

        var dynUp = z.DynamicUpdate || 'None';
        var dynHtml = '';
        if (dynUp === 'Secure') { dynHtml = '<span class="rpt-b rpt-good">Secure Only</span>'; }
        else if (dynUp === 'NonsecureAndSecure') { dynHtml = '<span class="rpt-b rpt-warn">NonSecure+Secure</span>'; }
        else { dynHtml = '<span class="rpt-b rpt-na">None</span>'; }

        tbody += buildDataRow([
            { html: '<span style="font-family:monospace;font-size:11px;font-weight:600;">' + escapeHtml(z.ZoneName || '') + '</span>', dc: true },
            { html: lookupPill },
            { html: typePill,    divider: true },
            { html: adPill },
            { html: scopeHtml,   divider: true },
            { html: dynHtml,     divider: true }
        ]);
    });

    if (!tbody) { tbody = '<tr><td colspan="6" style="text-align:center;color:#94a3b8;padding:20px;">No DNS zone data collected</td></tr>'; }

    document.getElementById('rptThead').innerHTML = thead;
    document.getElementById('rptTbody').innerHTML = tbody;
}

function renderForwardersReport() {
    document.getElementById('rptTitle').textContent = 'Conditional Forwarders';
    document.getElementById('rptNote').innerHTML =
        '&#8505; AD-Integrated forwarders replicate automatically to all DCs in the specified scope. If more than 2 forwarder IPs exist, additional IPs appear comma-separated in the Secondary column.';

    var thead = buildGroupRow([
        { span: 1, empty: true },
        { span: 2, label: 'Forwarder IPs' },
        { span: 2, label: 'AD Replication', divider: true }
    ]) + buildHeaderRow([
        { label: 'Zone / Domain Name' },
        { label: 'Primary Forwarder IP' },
        { label: 'Secondary Forwarder IP(s)' },
        { label: 'AD Integrated',    divider: true },
        { label: 'Replication Scope' }
    ]);

    var zones = parseDnsZonesCsv();
    var forwarders = zones.filter(function (z) { return z.ZoneType === 'Forwarder'; });
    var tbody = '';

    forwarders.forEach(function (z) {
        var ips = (z.MasterServers || '').split(',').map(function (ip) { return ip.trim(); }).filter(Boolean);
        var ip1Html = ips[0] ? '<span style="font-family:monospace;font-size:10px;font-weight:600;padding:2px 7px;border-radius:5px;background:#f8fafc;border:0.5px solid #cbd5e1;">' + escapeHtml(ips[0]) + '</span>' : '&#8212;';
        var ip2Html = ips.length > 1
            ? ips.slice(1).map(function (ip) {
                return '<span style="font-family:monospace;font-size:10px;font-weight:600;padding:2px 7px;border-radius:5px;background:#f8fafc;border:0.5px solid #cbd5e1;margin-right:3px;">' + escapeHtml(ip) + '</span>';
              }).join('')
            : '&#8212;';

        var isAD = (z.IsDsIntegrated || '').toLowerCase() === 'true';
        var adPill = isAD
            ? '<span class="rpt-b rpt-info">AD-Integrated</span>'
            : '<span class="rpt-b rpt-na">Standard</span>';

        var scope = isAD ? (z.ReplicationScope || 'Domain') : 'N/A';
        var scopeCls = (scope === 'Forest') ? 'rpt-info' : (scope === 'Domain') ? 'rpt-purple' : '';
        var scopeHtml = scope === 'N/A'
            ? '<span class="rpt-b rpt-na">N/A</span>'
            : '<span class="rpt-b ' + scopeCls + '">' + escapeHtml(scope) + '</span>';

        tbody += buildDataRow([
            { html: '<span style="font-family:monospace;font-size:11px;font-weight:600;">' + escapeHtml(z.ZoneName || '') + '</span>', dc: true },
            { html: ip1Html },
            { html: ip2Html },
            { html: adPill,    divider: true },
            { html: scopeHtml }
        ]);
    });

    if (!tbody) { tbody = '<tr><td colspan="5" style="text-align:center;color:#94a3b8;padding:20px;">No conditional forwarders found</td></tr>'; }

    document.getElementById('rptThead').innerHTML = thead;
    document.getElementById('rptTbody').innerHTML = tbody;
}

// =====================================================================
// Generic CSV parser (shared by Sites/GPO renderers)
// =====================================================================
function parseSimpleCsv(raw) {
    if (!raw) { return []; }
    var lines = raw.split('\n').filter(function (l) { return l.trim(); });
    if (lines.length < 2) { return []; }
    var headers = lines[0].replace(/"/g, '').split(',');
    var rows = [];
    for (var i = 1; i < lines.length; i++) {
        var vals = lines[i].match(/("([^"]*)"|([^,]*))(,|$)/g) || [];
        var obj  = {};
        vals.forEach(function (v, idx) {
            var c = v.replace(/,$/, '').replace(/^"|"$/g, '');
            if (headers[idx]) { obj[headers[idx].trim()] = c; }
        });
        rows.push(obj);
    }
    return rows;
}

// =====================================================================
// Sites & Services report renderer
// =====================================================================
function renderSitesReport() {
    // -- parse data sources --
    var sites      = parseSimpleCsv(typeof adSitesCsv      !== 'undefined' ? adSitesCsv      : '');
    var subnets    = parseSimpleCsv(typeof adSubnetsCsv     !== 'undefined' ? adSubnetsCsv    : '');
    var dcs        = parseSimpleCsv(typeof domainControllersCsv !== 'undefined' ? domainControllersCsv : '');
    var links      = parseSimpleCsv(typeof adSiteLinksCsv   !== 'undefined' ? adSiteLinksCsv  : '');
    var bridgeFlag = (typeof adSiteLinkBridgeFlag !== 'undefined') ? adSiteLinkBridgeFlag : 'unknown';

    // -- DC count per site --
    var dcCountBySite = {};
    dcs.forEach(function (dc) {
        var s = dc.Site || '';
        dcCountBySite[s] = (dcCountBySite[s] || 0) + 1;
    });

    // -- subnets grouped by site --
    var subnetsBySite = {};
    subnets.forEach(function (sub) {
        var s = sub.Site || '';
        if (!subnetsBySite[s]) { subnetsBySite[s] = []; }
        subnetsBySite[s].push(sub);
    });

    // -- alert counts --
    var noDCCount  = 0;
    var noSubCount = 0;
    sites.forEach(function (s) {
        if (!dcCountBySite[s.Name])                                     { noDCCount++;  }
        if (!subnetsBySite[s.Name] || subnetsBySite[s.Name].length === 0) { noSubCount++; }
    });

    // ─── SECTION 1: Sites & Subnets ───────────────────────────────────
    var thead1 = buildGroupRow([
        { span: 2, label: 'Site' },
        { span: 3, label: 'Associated Subnets', divider: true }
    ]) + buildHeaderRow([
        { label: 'Site Name' },
        { label: 'Description' },
        { label: 'Subnet (CIDR)',       divider: true },
        { label: 'Location' },
        { label: 'Subnet Description' }
    ]);

    var tbody1 = '';
    // Alert banners as first row
    if (noDCCount > 0 || noSubCount > 0) {
        tbody1 += '<tr><td colspan="5" style="padding:6px 0 4px 0;border:none;background:transparent;">';
        if (noDCCount  > 0) { tbody1 += '<div style="background:#fef2f2;border:1px solid #fca5a5;border-radius:6px;padding:6px 12px;margin-bottom:5px;font-size:11px;color:#7f1d1d;font-weight:600;">&#9888; ' + noDCCount  + ' site(s) have no Domain Controllers &#8212; clients cannot authenticate locally.</div>'; }
        if (noSubCount > 0) { tbody1 += '<div style="background:#fff7ed;border:1px solid #fdba74;border-radius:6px;padding:6px 12px;font-size:11px;color:#7c2d12;font-weight:600;">&#9888; ' + noSubCount + ' site(s) have no subnets &#8212; AD site cost routing may be incorrect.</div>'; }
        tbody1 += '</td></tr>';
    }

    if (sites.length === 0) {
        tbody1 += '<tr><td colspan="5" style="text-align:center;padding:20px;color:#94a3b8;font-style:italic;">No AD site data available</td></tr>';
    } else {
        sites.forEach(function (site) {
            var nm       = site.Name || '';
            var desc     = site.Description || '';
            var siteSubs = subnetsBySite[nm] || [];
            var dcCnt    = dcCountBySite[nm] || 0;
            var noDC     = dcCnt === 0;
            var noSub    = siteSubs.length === 0;

            var rowBg  = noDC ? '#fef2f2' : (noSub ? '#fff7ed' : '#eff6ff');
            var bdrClr = noDC ? '#fca5a5' : (noSub ? '#fdba74' : '#bfdbfe');
            var nmClr  = noDC ? '#991b1b' : (noSub ? '#92400e' : '#1e40af');

            var dcBadge  = noDC ? rptCrit('&#10006; No DCs')
                                 : rptGood('&#10004; ' + dcCnt + ' DC' + (dcCnt > 1 ? 's' : ''));
            var subBadge = noSub ? rptWarn('&#9888; No subnets')
                                  : '<span class="rpt-b rpt-info">' + siteSubs.length + ' subnet' + (siteSubs.length !== 1 ? 's' : '') + '</span>';

            tbody1 += '<tr><td colspan="5" style="background:' + rowBg + ';border-top:2px solid ' + bdrClr + ';border-bottom:1px solid ' + bdrClr + ';padding:7px 10px;">';
            tbody1 += '<span style="font-weight:800;font-size:12px;color:' + nmClr + ';margin-right:6px;">' + escapeHtml(nm) + '</span>';
            tbody1 += dcBadge + ' ' + subBadge;
            if (desc) { tbody1 += '<div style="font-size:10px;color:#475569;margin-top:3px;">' + escapeHtml(desc) + '</div>'; }
            tbody1 += '</td></tr>';

            if (noSub) {
                tbody1 += '<tr><td colspan="5" style="font-size:10px;color:#94a3b8;font-style:italic;padding:7px 10px;border-bottom:1px solid #f1f5f9;">No subnets associated with this site</td></tr>';
            } else {
                siteSubs.forEach(function (sub) {
                    var loc = sub.Location || '';
                    var locCell = loc ? '<span class="rpt-b rpt-purple">' + escapeHtml(loc) + '</span>' : '';
                    tbody1 += '<tr style="border-bottom:1px solid #f1f5f9;">';
                    tbody1 += '<td style="padding:6px 10px;"></td>';
                    tbody1 += '<td style="font-size:10px;color:#94a3b8;font-style:italic;padding:6px 10px;">' + escapeHtml(nm) + '</td>';
                    tbody1 += '<td style="border-left:2px solid #dbeafe;padding:6px 10px;"><span style="font-family:monospace;font-size:11px;font-weight:700;background:#f1f5f9;color:#0f172a;padding:2px 7px;border-radius:5px;border:0.5px solid #cbd5e1;">' + escapeHtml(sub.Name) + '</span></td>';
                    tbody1 += '<td style="padding:6px 10px;">' + locCell + '</td>';
                    tbody1 += '<td style="padding:6px 10px;font-size:11px;color:#374151;">' + escapeHtml(sub.Description || '') + '</td>';
                    tbody1 += '</tr>';
                });
            }
        });
    }

    // ─── SECTION 2: Site Links ─────────────────────────────────────────
    // Uses the same .rpt-header-row / .rpt-gd / .rpt-b classes as every
    // other report so the global theme toggle covers this section too -
    // no group header row, matching the rest of the dashboard.
    var thead2 = buildHeaderRow([
        { label: 'Link Name' },
        { label: 'Connected Sites' },
        { label: 'Cost',            divider: true },
        { label: 'Interval (min)' },
        { label: 'Transport' }
    ]);

    var tbody2 = '';
    if (links.length === 0) {
        tbody2 = '<tr><td colspan="5" style="text-align:center;padding:18px;color:#94a3b8;font-style:italic;">No site link data available</td></tr>';
    } else {
        links.forEach(function (lnk) {
            var cost     = parseInt(lnk.Cost || '0', 10);
            var interval = lnk.ReplicationFrequencyInMinutes || '';
            var transport = lnk.Transport || 'IP';
            var siteList  = (lnk.SitesIncluded || '').split('|').filter(Boolean);
            var costClr   = cost >= 400 ? '#cf222e' : '#1a7f37';
            var transBadge = transport === 'SMTP'
                ? '<span class="rpt-b rpt-warn">SMTP</span>'
                : '<span class="rpt-b rpt-info">IP</span>';
            var siteBadges = siteList.map(function (s) {
                return '<span class="rpt-b rpt-info" style="margin-right:3px;">' + escapeHtml(s) + '</span>';
            }).join('');
            tbody2 += buildDataRow([
                { html: '<span style="font-weight:700;">' + escapeHtml(lnk.Name) + '</span>' },
                { html: siteBadges },
                { html: '<span style="font-weight:800;font-size:14px;color:' + costClr + ';">' + cost + '</span>', divider: true },
                { val: interval + ' min' },
                { html: transBadge }
            ]);
        });
    }

    // Bridge All Site Links tile
    var bColor = bridgeFlag === 'enabled'  ? '#16a34a' : bridgeFlag === 'disabled' ? '#dc2626' : '#94a3b8';
    var bLabel = bridgeFlag === 'enabled'  ? 'Enabled &#10004;'
               : bridgeFlag === 'disabled' ? 'Disabled &#10006;' : 'Unknown';
    var bDesc  = bridgeFlag === 'enabled'
               ? 'AD automatically bridges all IP site links &#8212; replication paths are calculated transitively across all sites (recommended default).'
               : bridgeFlag === 'disabled'
               ? 'Bridge All Site Links is DISABLED. You must create site link bridge objects manually for transitive replication between sites.'
               : 'Unable to determine the Bridge All Site Links setting.';
    var bridgeTile = '<tr><td colspan="5" style="padding:10px 0 4px 0;border:none;background:transparent;">'
        + '<div style="background:#fff;border-radius:9px;box-shadow:0 2px 8px rgba(15,23,42,0.07);padding:13px 16px;display:flex;align-items:flex-start;gap:14px;">'
        + '<span style="font-size:24px;line-height:1.2;">&#128279;</span>'
        + '<div style="flex:1;">'
        + '<div style="font-size:10px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:4px;">Bridge All Site Links &mdash; IP Transport</div>'
        + '<div style="font-size:14px;font-weight:800;color:' + bColor + ';margin-bottom:4px;">' + bLabel + '</div>'
        + '<div style="font-size:10.5px;color:#64748b;">' + bDesc + '</div>'
        + '</div></div></td></tr>';

    // ─── Render into DOM ───────────────────────────────────────────────
    document.getElementById('rptTitle').textContent = 'Sites and Services';
    document.getElementById('rptThead').innerHTML   = thead1;
    document.getElementById('rptTbody').innerHTML   = tbody1;

    // Section 2 injected into rptNote
    var sec2 = '<div style="display:flex;align-items:center;justify-content:space-between;margin:20px 0 10px 0;">';
    sec2 += '<div style="font-size:12px;font-weight:800;color:#1e3a5f;text-transform:uppercase;letter-spacing:0.06em;border-left:4px solid #2563eb;padding-left:10px;">Site Links &amp; Replication</div>';
    sec2 += '<button onclick="exportSiteLinksSec2()" style="display:flex;align-items:center;gap:6px;padding:6px 14px;background:#1e40af;color:#fff;border:none;border-radius:7px;font-size:11px;font-weight:600;cursor:pointer;letter-spacing:0.03em;">&#8659; Export Site Links</button>';
    sec2 += '</div>';
    sec2 += '<div style="overflow-x:auto;"><table id="siteLinksSec2Table" class="rpt-styled-table" style="width:100%;table-layout:fixed;border-radius:10px;overflow:hidden;box-shadow:0 3px 12px rgba(15,23,42,0.08);background:#fff;margin:0;">' +
            '<colgroup><col style="width:30%;"><col style="width:38%;"><col style="width:10%;"><col style="width:12%;"><col style="width:10%;"></colgroup>';
    sec2 += '<thead>' + thead2 + '</thead>';
    sec2 += '<tbody>' + tbody2 + bridgeTile + '</tbody></table></div>';
    sec2 += '<p style="font-size:10.5px;color:#64748b;margin-top:10px;">&#8505; Cost: lower = preferred path. Default cost is 100, default interval is 180 min. Cost &ge; 400 shown in red. Bridge All Site Links is enabled by default in AD.</p>';
    var noteEl = document.getElementById('rptNote');
    noteEl.style.display = 'block';
    noteEl.innerHTML = sec2;
}

// =====================================================================
// Group Policies report renderer
// =====================================================================
function renderGpoReport() {
    var gpos = parseSimpleCsv(typeof gpoCsv !== 'undefined' ? gpoCsv : '');

    var thead = buildHeaderRow([
        { label: '#' },
        { label: 'GPO Name' },
        { label: 'GPO GUID' },
        { label: 'Status' },
        { label: 'Policy Type' },
        { label: 'Links' },
        { label: 'WMI Filter' }
    ]);

    var tbody = '';
    if (gpos.length === 0) {
        tbody = '<tr><td colspan="7" style="text-align:center;padding:20px;color:#94a3b8;font-style:italic;">No GPO data available</td></tr>';
    } else {
        var emptyCount = 0;
        gpos.forEach(function (g, idx) {
            var nm      = g.DisplayName || '';
            var guid    = g.Id          || '';
            var status  = g.GpoStatus   || '';
            var linked  = g.LinkedTo    || '';
            var isEmpty = (g.IsEmpty || '').toLowerCase() === 'true';
            var wmi     = g.WmiFilter   || '';

            var isUnlinked = !linked;
            var isDisabled = status === 'AllSettingsDisabled';
            if (isEmpty) { emptyCount++; }

            // Status badge
            var stBadge = '';
            if (status === 'AllSettingsEnabled')           { stBadge = rptGood('Enabled'); }
            else if (status === 'AllSettingsDisabled')     { stBadge = rptNa('Disabled'); }
            else if (status === 'UserSettingsDisabled')    { stBadge = rptWarn('User Off'); }
            else if (status === 'ComputerSettingsDisabled'){ stBadge = rptWarn('Comp Off'); }
            else                                           { stBadge = rptNa(status || 'Unknown'); }

            // Policy Type badge derived from GpoStatus + IsEmpty - uses the
            // same shared rpt-b classes as every other badge in the
            // dashboard, so the global theme toggle covers this too.
            var ptBadge = '';
            if (isEmpty) {
                ptBadge = '<span class="rpt-b rpt-crit">&#9888; Empty</span>';
            } else if (status === 'AllSettingsEnabled') {
                ptBadge = '<span class="rpt-b rpt-good">Both</span>';
            } else if (status === 'ComputerSettingsDisabled') {
                ptBadge = '<span class="rpt-b rpt-info">User</span>';
            } else if (status === 'UserSettingsDisabled') {
                ptBadge = '<span class="rpt-b rpt-purple">Computer</span>';
            } else if (status === 'AllSettingsDisabled') {
                ptBadge = '<span class="rpt-b rpt-na">None</span>';
            } else {
                ptBadge = '<span class="rpt-b rpt-na">&#8212;</span>';
            }

            // Links badge
            var linkBadge = isUnlinked ? rptWarn('No Links') : rptGood('Linked');

            // WMI filter badge - color comes from the shared rpt-info class
            // (theme-able); max-width/ellipsis/tooltip stay inline since
            // those are layout behavior, not color, and don't conflict.
            var wmiBadge = wmi
                ? '<span class="rpt-b rpt-info" style="max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="' + escapeHtml(wmi) + '">' + escapeHtml(wmi) + '</span>'
                : '<span class="rpt-b rpt-na">&#8212;</span>';

            var rowStyle = 'border-bottom:1px solid #f1f5f9;' + (isEmpty ? 'background:#fef2f2;' : (isUnlinked ? 'background:#fffbeb;' : ''));
            tbody += '<tr style="' + rowStyle + '">';
            tbody += '<td style="padding:7px 10px;color:#94a3b8;font-size:11px;text-align:center;">' + (idx + 1) + '</td>';
            tbody += '<td style="padding:7px 10px;font-weight:700;font-size:12px;color:' + (isDisabled ? '#94a3b8' : '#0f172a') + ';">' + escapeHtml(nm) + '</td>';
            tbody += '<td style="padding:7px 10px;font-family:monospace;font-size:10px;color:#64748b;">' + escapeHtml(guid) + '</td>';
            tbody += '<td style="padding:7px 10px;">' + stBadge + '</td>';
            tbody += '<td style="padding:7px 10px;">' + ptBadge + '</td>';
            tbody += '<td style="padding:7px 10px;">' + linkBadge + '</td>';
            tbody += '<td style="padding:7px 10px;">' + wmiBadge + '</td>';
            tbody += '</tr>';
        });
    }

    var noteUnlinked = gpos.filter(function (g) { return !g.LinkedTo; }).length;
    var emptyFinal   = gpos.filter(function (g) { return (g.IsEmpty || '').toLowerCase() === 'true'; }).length;
    var wmiCount     = gpos.filter(function (g) { return g.WmiFilter; }).length;

    document.getElementById('rptTitle').textContent = 'Group Policy Objects (GPO) Inventory';
    document.getElementById('rptThead').innerHTML   = thead;
    document.getElementById('rptTbody').innerHTML   = tbody;
    document.getElementById('rptNote').innerHTML    =
        '&#8505; Total: ' + gpos.length + ' GPO(s). ' +
        (emptyFinal   > 0 ? emptyFinal   + ' empty (no settings defined). ' : '') +
        (noteUnlinked > 0 ? noteUnlinked + ' unlinked (not applied anywhere). ' : '') +
        (wmiCount     > 0 ? wmiCount     + ' use a WMI filter (conditionally applied). ' : '') +
        'Policy Type reflects which configuration class is active.';
}

// =====================================================================
// Security Risks Need Attention report renderer - one row per individual
// finding (DC + risk pair), so it scales whether there are 4 DCs or 400,
// unlike the Top Risks tile on the Executive Summary which only shows a
// top-6 summary with truncated DC name lists.
// =====================================================================
function renderSecurityRisksReport() {
    var risks = parseSimpleCsv(typeof securityRisksCsv !== 'undefined' ? securityRisksCsv : '');

    var sevRank = { High: 0, Medium: 1, Low: 2 };

    if (secRisksSortCol === 'Severity') {
        risks.sort(function (a, b) {
            var ra = sevRank.hasOwnProperty(a.Severity) ? sevRank[a.Severity] : 3;
            var rb = sevRank.hasOwnProperty(b.Severity) ? sevRank[b.Severity] : 3;
            return (ra - rb) * secRisksSortDir;
        });
    } else if (secRisksSortCol === 'DC') {
        risks.sort(function (a, b) {
            return String(a.DC || '').localeCompare(String(b.DC || '')) * secRisksSortDir;
        });
    } else if (secRisksSortCol === 'Source') {
        risks.sort(function (a, b) {
            return String(a.Source || '').localeCompare(String(b.Source || '')) * secRisksSortDir;
        });
    } else {
        // Default order - no column picked yet: High, then Medium, then Low.
        risks.sort(function (a, b) {
            var ra = sevRank.hasOwnProperty(a.Severity) ? sevRank[a.Severity] : 3;
            var rb = sevRank.hasOwnProperty(b.Severity) ? sevRank[b.Severity] : 3;
            return ra - rb;
        });
    }

    function sortArrow(col) {
        // Always show a sort icon, even before the user has clicked anything,
        // so it's obvious at a glance that the column is sortable - not just
        // discoverable by accident. Active column gets a bold single arrow
        // showing direction; inactive sortable columns get a faint two-way
        // icon as a persistent "click to sort" affordance.
        if (secRisksSortCol !== col) {
            return ' <span style="opacity:0.35;font-size:10px;">&#8693;</span>';
        }
        return secRisksSortDir === 1
            ? ' <span style="font-size:11px;">&#9650;</span>'
            : ' <span style="font-size:11px;">&#9660;</span>';
    }

    var thead = '<tr class="rpt-header-row">' +
        '<th style="cursor:pointer;user-select:none;" title="Click to sort by severity" onclick="toggleSecRisksSort(\'Severity\')">Severity' + sortArrow('Severity') + '</th>' +
        '<th>Risk <span style="font-weight:400;opacity:0.6;font-size:10px;">(click row to expand)</span></th>' +
        '<th class="rpt-gd" style="cursor:pointer;user-select:none;" title="Click to sort by domain controller" onclick="toggleSecRisksSort(\'DC\')">Domain Controller' + sortArrow('DC') + '</th>' +
        '<th style="cursor:pointer;user-select:none;" title="Click to sort by source" onclick="toggleSecRisksSort(\'Source\')">Source' + sortArrow('Source') + '</th>' +
        '</tr>';

    var tbody = '';
    if (risks.length === 0) {
        tbody = '<tr><td colspan="4" style="text-align:center;padding:20px;color:#94a3b8;font-style:italic;">No significant risks detected</td></tr>';
    } else {
        tbody = risks.map(function (r, idx) {
            var sev = r.Severity || '';
            var sevBadge = sev === 'High' ? rptCrit('High') : (sev === 'Medium' ? rptMedium('Medium') : rptNa(sev || 'Low'));
            var dcCell = (r.DC === 'Forest-wide')
                ? '<span class="rpt-b rpt-na">Forest-wide</span>'
                : escapeHtml(r.DC || '');

            return '<tr style="border-bottom:1px solid #f1f5f9;cursor:pointer;" onclick="toggleRptDetailRow(' + idx + ', this)">' +
                '<td style="padding:7px 10px;">' + sevBadge + '</td>' +
                '<td class="rpt-dc" style="padding:7px 10px;font-size:12px;color:#0f172a;"><span class="id-expand-chevron" style="display:inline-block;width:14px;color:#94a3b8;">&#9656;</span>' + escapeHtml(r.Risk || '') + '</td>' +
                '<td style="padding:7px 10px;font-weight:700;font-size:12px;color:#0f172a;">' + dcCell + '</td>' +
                '<td style="padding:7px 10px;font-size:11.5px;color:#64748b;">' + escapeHtml(r.Source || '') + remediationBadge(r.Source) + '</td>' +
                '</tr>';
        }).join('');
    }

    var highCount   = risks.filter(function (r) { return r.Severity === 'High'; }).length;
    var mediumCount = risks.filter(function (r) { return r.Severity === 'Medium'; }).length;

    rptCurrentRows   = risks;
    rptDetailBuilder = buildSecurityRiskDetailHtml;
    rptDetailColSpan = 4;

    document.getElementById('rptTitle').textContent = 'Security Risks Need Attention';
    document.getElementById('rptThead').innerHTML   = thead;
    document.getElementById('rptTbody').innerHTML   = tbody;
    document.getElementById('rptNote').innerHTML    =
        '&#8505; Total: ' + risks.length + ' finding(s) - ' + highCount + ' High, ' + mediumCount + ' Medium. ' +
        'Each row is one risk on one domain controller (or "Forest-wide" for findings not tied to a specific DC). ' +
        'Click a column header to sort, click a row to expand why it matters and how to fix it.';
}

// =====================================================================
// DC Uptime & Reboot Status report renderer - one row per DC, pulled
// directly from dcInventoryData (System tile data) joined with dcListData
// for the IPv4 address, same way renderDCDetail() already does it.
// =====================================================================
function renderDcUptimeReport() {
    var thead = buildHeaderRow([
        { label: 'DC Name' },
        { label: 'IPv4 Address', divider: true },
        { label: 'Operating System' },
        { label: 'Last Boot Time' },
        { label: 'Uptime' },
        { label: 'Pending Reboot Status' }
    ]);

    var tbody = '';
    if (!dcInventoryData || dcInventoryData.length === 0) {
        tbody = '<tr><td colspan="6" style="text-align:center;padding:20px;color:#94a3b8;font-style:italic;">No domain controller data available</td></tr>';
    } else {
        dcInventoryData.forEach(function (inv) {
            var sys    = inv.System || {};
            var dcMeta = dcListData.find(function (d) { return d.Name === inv.ComputerName; }) || {};
            var ip     = dcMeta.IPv4Address || '';

            var rebootBadge = sys.PendingReboot ? rptCrit('Reboot Pending') : rptGood('No Reboot Pending');

            tbody += buildDataRow([
                { html: escapeHtml(inv.ComputerName), dc: true },
                { html: escapeHtml(ip), divider: true },
                { html: escapeHtml(sys.Caption || 'Unknown') },
                { html: escapeHtml(sys.LastBootTime || 'Unknown') },
                { html: escapeHtml(sys.UptimeReadable || 'Unknown') },
                { html: rebootBadge }
            ]);
        });
    }

    var rebootCount = dcInventoryData.filter(function (inv) {
        return inv.System && inv.System.PendingReboot;
    }).length;

    document.getElementById('rptTitle').textContent = 'DC Uptime & Reboot Status';
    document.getElementById('rptThead').innerHTML   = thead;
    document.getElementById('rptTbody').innerHTML   = tbody;
    document.getElementById('rptNote').innerHTML    =
        '&#8505; Total: ' + dcInventoryData.length + ' domain controller(s). ' +
        (rebootCount > 0
            ? rebootCount + ' DC(s) have a reboot pending - see the Domain Controller Health tab for the reason.'
            : 'No domain controllers have a reboot pending.');
}

// =====================================================================
// Remediation + MITRE ATT&CK knowledge base (Phase B) - keyed by the same
// Category/Source strings already used in the Identity & Kerberos Security
// and Security Risks Need Attention reports, so both can show a
// what/why/how-to-fix panel and a MITRE technique tag per finding without
// any new per-finding data collection.
// =====================================================================
var remediationGuide = {
    'Kerberoasting': {
        why: 'Any account with an SPN can have its Kerberos service ticket requested and cracked offline. Privileged or RC4-only accounts are cracked fastest.',
        fix: 'Use a gMSA where possible, enforce AES (set msDS-SupportedEncryptionTypes to AES128/256 only), and use a long random password for any account that must remain a standard user.',
        ref: 'https://learn.microsoft.com/en-us/windows-server/security/kerberos/kerberos-authentication-overview',
        mitre: 'T1558.003'
    },
    'AS-REP Roasting': {
        why: 'Accounts with Kerberos pre-authentication disabled let an attacker request an AS-REP and crack it offline without any credentials.',
        fix: 'Re-enable "Do not require Kerberos preauthentication" (clear the UAC flag) unless there is a specific, documented reason it must stay disabled.',
        ref: 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/appendix-d--securing-built-in-administrator-accounts-in-active-directory',
        mitre: 'T1558.004'
    },
    'DCSync Rights': {
        why: 'Replicating Directory Changes + All lets a principal request password hashes for any account in the domain, equivalent to a full domain compromise.',
        fix: 'Remove the right from any principal outside Domain Admins/Enterprise Admins/Domain Controllers. Audit how it was granted (often an over-scoped delegation).',
        ref: 'https://learn.microsoft.com/en-us/defender-for-identity/understand-lateral-movement-paths',
        mitre: 'T1003.006'
    },
    'AdminSDHolder Drift': {
        why: 'AdminSDHolder&#8217;s ACL is copied onto every protected (Tier 0) object on a timer (SDProp). An unexpected ACE here silently propagates broad rights everywhere.',
        fix: 'Remove the unexpected ACE from AdminSDHolder directly, then confirm it does not reappear on the next SDProp cycle (60 minutes by default).',
        ref: 'https://learn.microsoft.com/en-us/troubleshoot/windows-server/identity/adminsdholder-permissions-not-propagated',
        mitre: 'T1098'
    },
    'Stale Account': {
        why: 'Inactive accounts are rarely monitored and are a quiet path for persistence if credentials are ever compromised, especially if privileged.',
        fix: 'Disable accounts inactive 90+ days after confirming with the owner/business; remove privileged group membership immediately regardless of disable timing.',
        ref: 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/appendix-c--protected-accounts-and-groups-in-active-directory',
        mitre: 'T1078'
    },
    'Password Never Expires': {
        why: 'A password that never rotates remains valid indefinitely if it is ever leaked, with no automatic remediation window.',
        fix: 'Remove the never-expires flag and move the account to a managed service account (gMSA) if it is a service account, or enforce normal rotation if it is a user.',
        ref: 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/manage/understand-default-user-rights-and-permissions',
        mitre: 'T1078'
    },
    'Constrained Delegation': {
        why: 'Constrained delegation lets a service impersonate users to specific targets - a misconfigured or stale target list is a lateral movement path.',
        fix: 'Confirm every delegation target is still required; migrate to resource-based constrained delegation where possible; remove unused entries.',
        ref: 'https://learn.microsoft.com/en-us/windows-server/security/kerberos/kerberos-constrained-delegation-overview',
        mitre: 'T1187'
    },
    'DNS Zone Transfer': {
        why: 'A zone open to transfer-to-any-server hands your entire DNS namespace (hostnames, roles, sometimes internal naming conventions) to anyone who asks.',
        fix: 'Set the zone&#8217;s transfer setting to "Only to servers listed in the Name Servers tab" or disable transfers entirely if no secondaries exist.',
        ref: 'https://learn.microsoft.com/en-us/windows-server/networking/dns/deploy/security',
        mitre: 'T1590.002'
    },
    'DNS Aging/Scavenging': {
        why: 'Without aging/scavenging, stale DNS records accumulate indefinitely, increasing the chance of dangling/hijackable records and incorrect lookups.',
        fix: 'Enable aging on the zone and configure scavenging on at least one DNS server per zone with sensible no-refresh/refresh intervals (commonly 7 days each).',
        ref: 'https://learn.microsoft.com/en-us/troubleshoot/windows-server/networking/dns-scavenging-best-practices',
        mitre: ''
    },
    'GPO Restricted Groups': {
        why: 'Restricted Groups that add members to local Administrators silently grant admin rights everywhere the GPO is linked - easy to lose track of over time.',
        fix: 'Review every Restricted Groups setting against current need; prefer a documented, narrowly-scoped tiering model over broad Restricted Groups pushes.',
        ref: 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/implementing-least-privilege-administrative-models',
        mitre: 'T1098'
    },
    'GPP cpassword (MS14-025)': {
        why: 'Group Policy Preferences cpassword values use a published, static AES key - any authenticated user can decrypt the credential in seconds.',
        fix: 'Remove the preference item and rotate the exposed credential immediately; this is a 2014 vulnerability with no legitimate reason to still be present.',
        ref: 'https://support.microsoft.com/en-us/topic/ms14-025-vulnerability-in-group-policy-preferences-could-allow-elevation-of-privilege-18870e57-5d11-1190-7f1e-8f51953de4ed',
        mitre: 'T1552.006'
    },
    'Global AD Security': {
        why: 'Foundational forest/domain-wide identity hygiene checks (SYSVOL, RID 500, Guest, Protected Users, SPN duplicates, delegation, password policy).',
        fix: 'See the specific finding text for the exact control and recommended remediation.',
        ref: 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory-installations',
        mitre: ''
    },
    'Security Posture': {
        why: 'Per-DC configuration hardening (SMB, LDAP signing, firewall, NetBIOS, NTLM) - a weak setting on even one DC widens the attack surface for the whole domain.',
        fix: 'Apply the hardening setting via GPO at the OU containing your domain controllers so it stays consistent and survives a DC rebuild.',
        ref: 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory-installations',
        mitre: ''
    },
    'Deep Insights': {
        why: 'Operational integrity signals (backup recency, lingering objects, USN rollback, DFSR health) - these indicate the directory itself may be in a fragile or inconsistent state.',
        fix: 'See the specific finding text; most require a targeted repadmin/dfsrdiag investigation rather than a simple setting change.',
        ref: 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/manage/troubleshoot/troubleshooting-active-directory-replication-problems',
        mitre: ''
    },
    'DC Health Check': {
        why: 'A failing core service or dcdiag test on a domain controller can degrade authentication and replication for every client that uses it.',
        fix: 'Investigate the specific service/test on the named DC; check the System and Directory Service event logs for the underlying cause.',
        ref: 'https://learn.microsoft.com/en-us/troubleshoot/windows-server/identity/useful-tools-for-troubleshooting-ad',
        mitre: ''
    },
    'Connectivity': {
        why: 'A DC that does not respond to ICMP may also be failing other services - ping is a coarse but fast first signal of an unreachable or overloaded DC.',
        fix: 'Confirm the DC is powered on and on the network; check Windows Firewall ICMP rules; investigate further via the DC Health Check report.',
        ref: 'https://learn.microsoft.com/en-us/troubleshoot/windows-server/identity/useful-tools-for-troubleshooting-ad',
        mitre: ''
    },
    'AD Replication': {
        why: 'A failing replication link means changes (including security-relevant ones, like a disabled account) may not reach every DC promptly.',
        fix: 'Run repadmin /showrepl and repadmin /replsummary on the affected DC to identify the specific partner and error.',
        ref: 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/manage/troubleshoot/troubleshooting-active-directory-replication-problems',
        mitre: ''
    },
    'DNS': {
        why: 'A missing LDAP SRV record means clients and other DCs may be unable to locate this domain controller at all.',
        fix: 'Restart the Netlogon service on the affected DC to re-register its SRV records, then verify with nslookup/dcdiag /test:dns.',
        ref: 'https://learn.microsoft.com/en-us/troubleshoot/windows-server/identity/useful-tools-for-troubleshooting-ad',
        mitre: ''
    },
    'Group Policy': {
        why: 'Unlinked GPOs are not actively harmful, but they indicate policy sprawl and make it harder to know what is actually enforced.',
        fix: 'Confirm the GPO is genuinely unused, then either link it where intended or remove it to keep the policy set auditable.',
        ref: 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory-installations',
        mitre: ''
    },
    'Tier-0 Contamination': {
        why: 'Domain controllers are Tier 0 - only Tier 0 administrators should ever log onto one directly. A non-privileged account logging on (even via RDP or a cached/unlock session) suggests the tiering boundary is not being respected.',
        fix: 'Confirm why the account accessed the DC; if it was a one-off troubleshooting session, document it; if it is routine, move that workflow off the DC entirely (e.g. use RSAT from an admin workstation instead).',
        ref: 'https://learn.microsoft.com/en-us/windows-server/identity/securing-privileged-access/securing-privileged-access',
        mitre: 'T1078.002'
    },
    'Trust Hardening': {
        why: 'Without SID filtering quarantine, a compromised trusted domain can inject SID History values that grant rights in this domain. Without selective authentication, every account in the trusted domain can attempt to authenticate here.',
        fix: 'Enable SID filtering quarantine on outbound/bidirectional trusts unless a specific migration scenario requires it disabled; enable selective authentication on forest/external trusts and explicitly grant only the access actually needed.',
        ref: 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory-installations',
        mitre: 'T1199'
    },
    'SID History': {
        why: 'A legitimate SID History entry from a domain migration is normal, but it is also exactly what a SID History injection attack (e.g. via DCShadow) would produce - presence alone cannot distinguish the two from this dashboard.',
        fix: 'Confirm every flagged object against your migration records. Anything unexplained should be investigated immediately and the SID History attribute cleared if illegitimate.',
        ref: 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/manage/understand-security-identifiers',
        mitre: 'T1134.005'
    },
    'Attack Path to Domain Admins': {
        why: 'Either a group is nested into Domain Admins/Enterprise Admins (every member of that group is an effective Tier 0 admin, often unintentionally), or a principal holds a direct ACL right on the group object itself, letting it add members without ever being one today.',
        fix: 'For nested groups: confirm every member of the nested group is meant to be a Tier 0 administrator, or remove the nesting. For ACL findings: remove the unexpected right from the group object - only Domain Admins/Enterprise Admins/SYSTEM should be able to modify it.',
        ref: 'https://learn.microsoft.com/en-us/windows-server/identity/securing-privileged-access/securing-privileged-access',
        mitre: 'T1098'
    },
    'DNS Zone ACL': {
        why: 'A non-standard principal holding Write/GenericAll/CreateChild on a DNS zone object can create or modify records directly - including a fake _ldap._tcp/_kerberos._tcp SRV record or a spoofed server name, which clients and other DCs will trust and use for authentication.',
        fix: 'Remove the unexpected right from the zone object - only Domain Admins/Enterprise Admins/SYSTEM/DnsAdmins should be able to modify it. Audit how the right was granted, since this is rarely set intentionally at the per-zone level.',
        ref: 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory-installations',
        mitre: 'T1556'
    },
    'DnsAdmins Membership': {
        why: 'DnsAdmins membership is a well-documented path to SYSTEM on a domain controller: a member can set the DNS Server service&#8217;s ServerLevelPluginDll registry value to an attacker-controlled DLL path, which loads with SYSTEM privileges the next time the DNS service starts.',
        fix: 'Remove any member who does not have a specific, documented operational need for DNS administration. Treat this group with the same care as Domain Admins, not as a routine IT-support group.',
        ref: 'https://learn.microsoft.com/en-us/windows-server/identity/securing-privileged-access/securing-privileged-access',
        mitre: 'T1574.002'
    },
    'DNS Dynamic Update Risk': {
        why: 'A zone set to allow nonsecure dynamic updates accepts record creation/overwrite from any client, not just Kerberos-authenticated AD-integrated updates - an attacker on the network can register a fake server or service record and have legitimate clients authenticate to it.',
        fix: 'Change the zone&#8217;s dynamic update setting to "Secure only" unless a specific, documented legacy system requires nonsecure updates - if so, isolate that system&#8217;s updates to its own zone rather than leaving the whole zone open.',
        ref: 'https://learn.microsoft.com/en-us/windows-server/networking/dns/deploy/security',
        mitre: 'T1584'
    },
    'Tier-0 Group Membership': {
        why: 'Domain Admins/Enterprise Admins/Schema Admins/Administrators/Backup Operators/Account Operators/Print Operators all carry dangerous built-in rights - a disabled, stale, or service account inside any of them is a quiet, often-forgotten foothold. Backup Operators can read any file via the backup privilege, bypassing normal ACLs; Print Operators can log on locally to a domain controller and load drivers.',
        fix: 'Remove disabled or stale accounts from Tier-0 groups immediately. Confirm any service account membership is still required, and prefer a narrowly-scoped delegation over standing Tier-0 membership where possible.',
        ref: 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/appendix-b--privileged-accounts-and-groups-in-active-directory',
        mitre: 'T1078.002'
    },
    'Tier-0 Group Size': {
        why: 'A Tier-0 group with an unusually large membership is harder to audit and more likely to contain access nobody actively reviews - every additional member is one more potential point of compromise with full administrative impact.',
        fix: 'Review the full membership list and remove anyone without an active, documented need for this specific group. Consider time-bound (PAM/PIM-style) elevation instead of standing membership for anyone who only occasionally needs it.',
        ref: 'https://learn.microsoft.com/en-us/windows-server/identity/securing-privileged-access/securing-privileged-access',
        mitre: 'T1078.002'
    },
    'Universal Group Nesting': {
        why: 'Universal groups replicate to the global catalog and are usable from any domain in the forest - nesting one into a Tier-0 group is unusual and means that group&#8217;s effective Tier-0 reach extends further than a typical Domain Local or Global group would.',
        fix: 'Confirm this is intentional. If the broader, forest-wide reach is not actually needed, convert to a Domain Local or Global group, or remove the nesting.',
        ref: 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/manage/understand-group-scope',
        mitre: 'T1069.002'
    },
    'Pre-Windows 2000 Access': {
        why: 'The Pre-Windows 2000 Compatible Access group exists for backward compatibility with very old systems. If it contains Everyone or Anonymous Logon, it grants excessive read access to security descriptors on user and group objects and can enable certain legacy enumeration techniques.',
        fix: 'Remove Everyone/Anonymous Logon/Authenticated Users from this group unless a specific, documented legacy system still requires it - this compatibility mode has not been needed for any supported Windows version in a very long time.',
        ref: 'https://learn.microsoft.com/en-us/troubleshoot/windows-server/identity/pre-windows-2000-compatible-access-group',
        mitre: 'T1087.002'
    },
    'Group-Based Delegation': {
        why: 'A principal holding a delegated right (a generic ACL right, or specifically Reset Password) on a Tier-0 group object can escalate without ever being a member of that group today - the classic "helpdesk group can reset a Domain Admin&#8217;s password" scenario.',
        fix: 'Remove the unexpected right from the Tier-0 group object. Only Domain Admins/Enterprise Admins/SYSTEM should be able to modify these groups or reset passwords for their members.',
        ref: 'https://learn.microsoft.com/en-us/windows-server/identity/securing-privileged-access/securing-privileged-access',
        mitre: 'T1098'
    },
    'Nesting Depth': {
        why: 'A membership chain more than two or three hops deep into a Tier-0 group is hard to audit - the people with effective Tier-0 access are not visible from the group&#8217;s own membership list, only by walking the full chain.',
        fix: 'Flatten the nesting where practical - add the real intended administrators closer to the Tier-0 group directly, or document the chain explicitly if the structure must stay as-is.',
        ref: 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory-installations',
        mitre: 'T1069.002'
    },
    'Circular Group Reference': {
        why: 'A group that nests back into itself through one or more intermediate groups (Group A contains Group B which contains Group A) is a structural misconfiguration - it does not directly grant access, but it signals the group structure has not been properly maintained and can hide other issues.',
        fix: 'Identify and remove the circular nesting. Document the intended hierarchy so it does not recur.',
        ref: 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/manage/understand-group-scope',
        mitre: 'T1069.002'
    },
    'Foreign Security Principal': {
        why: 'A Foreign Security Principal inside a Tier-0 group means an account or group from another domain or forest holds full administrative access here - the security of this domain now depends partly on the security practices of the trusted domain.',
        fix: 'Confirm this cross-domain/cross-forest access is intentional and still required. Remove it if not, and prefer a narrower, resource-specific grant over Tier-0 membership for any legitimate cross-domain need.',
        ref: 'https://learn.microsoft.com/en-us/windows-server/identity/ad-ds/plan/security-best-practices/best-practices-for-securing-active-directory-installations',
        mitre: 'T1078.002'
    }
};

// =====================================================================
// Calibrated Risk Scoring Engine (Phase C) - Impact x Likelihood x
// Exposure, each 1-10, normalized to 0-100. This is DISPLAY-ONLY: it does
// not change $HealthScore's point-deduction math, so the existing
// ADHealthHistory.csv trend stays comparable across versions, per the
// roadmap's explicit calibration requirement. Shown as a supplementary
// "Risk score" next to the existing Severity badge.
// =====================================================================
var riskScoreModel = {
    'Kerberoasting':              { impact: 8, likelihood: 6, exposure: 7 },
    'AS-REP Roasting':            { impact: 8, likelihood: 7, exposure: 7 },
    'DCSync Rights':               { impact: 10, likelihood: 9, exposure: 8 },
    'AdminSDHolder Drift':         { impact: 9, likelihood: 7, exposure: 7 },
    'Stale Account':               { impact: 5, likelihood: 4, exposure: 5 },
    'Password Never Expires':      { impact: 4, likelihood: 4, exposure: 5 },
    'Constrained Delegation':      { impact: 6, likelihood: 5, exposure: 5 },
    'DNS Zone Transfer':           { impact: 6, likelihood: 6, exposure: 8 },
    'DNS Aging/Scavenging':        { impact: 3, likelihood: 3, exposure: 3 },
    'GPO Restricted Groups':       { impact: 7, likelihood: 5, exposure: 6 },
    'GPP cpassword (MS14-025)':    { impact: 9, likelihood: 9, exposure: 6 },
    'Tier-0 Contamination':        { impact: 7, likelihood: 5, exposure: 6 },
    'Trust Hardening':             { impact: 6, likelihood: 4, exposure: 5 },
    'SID History':                 { impact: 7, likelihood: 4, exposure: 5 },
    'Attack Path to Domain Admins': { impact: 10, likelihood: 6, exposure: 7 },
    'DNS Zone ACL':                 { impact: 8, likelihood: 5, exposure: 6 },
    'DnsAdmins Membership':         { impact: 9, likelihood: 4, exposure: 5 },
    'DNS Dynamic Update Risk':      { impact: 7, likelihood: 6, exposure: 7 },
    'Tier-0 Group Membership':      { impact: 8, likelihood: 5, exposure: 6 },
    'Tier-0 Group Size':            { impact: 5, likelihood: 4, exposure: 4 },
    'Universal Group Nesting':      { impact: 5, likelihood: 3, exposure: 5 },
    'Pre-Windows 2000 Access':      { impact: 6, likelihood: 4, exposure: 6 },
    'Group-Based Delegation':       { impact: 9, likelihood: 6, exposure: 6 },
    'Nesting Depth':                { impact: 5, likelihood: 4, exposure: 4 },
    'Circular Group Reference':     { impact: 4, likelihood: 2, exposure: 3 },
    'Foreign Security Principal':   { impact: 8, likelihood: 4, exposure: 5 }
};

function riskScoreFor(category) {
    var m = riskScoreModel[category];
    if (!m) { return null; }
    var raw = m.impact * m.likelihood * m.exposure;
    var normalized = Math.round((raw / 1000) * 100);
    return { impact: m.impact, likelihood: m.likelihood, exposure: m.exposure, score: normalized };
}

function riskScoreBadge(category) {
    var rs = riskScoreFor(category);
    if (!rs) { return ''; }
    var color = rs.score >= 80 ? '#b91c1c' : (rs.score >= 60 ? '#a16207' : (rs.score >= 40 ? '#92400e' : '#475569'));
    var tip = 'Impact ' + rs.impact + ' x Likelihood ' + rs.likelihood + ' x Exposure ' + rs.exposure + ' (each 1-10), normalized to 0-100. Supplementary to Severity - does not change the Health Score.';
    return '<span title="' + escapeHtml(tip) + '" style="font-size:10px;font-weight:700;color:' + color + ';border:1px solid ' + color + ';border-radius:6px;padding:1px 5px;cursor:help;white-space:nowrap;">Risk ' + rs.score + '</span>';
}

// Plain-text tooltip (native title attribute) rather than a positioned
// hover panel - reliable inside a horizontally-scrollable table, where an
// absolutely-positioned panel risks being clipped by the scroll wrapper.
function remediationTipText(category) {
    var g = remediationGuide[category];
    if (!g) { return ''; }
    var mitreStr = g.mitre ? (' | MITRE ATT&CK: ' + g.mitre) : '';
    return 'Why it matters: ' + g.why + ' | How to fix: ' + g.fix + mitreStr + ' | See: ' + g.ref;
}

// Small inline "i" badge that carries the remediation tooltip - appended
// next to a Category/Source cell's text so it's visually clear there's
// more detail on hover, without changing the cell's existing text/color.
function remediationBadge(category) {
    var tip = remediationTipText(category);
    if (!tip) { return ''; }
    return ' <span title="' + escapeHtml(tip) + '" style="display:inline-flex;align-items:center;justify-content:center;width:14px;height:14px;border-radius:50%;background:#e2e8f0;color:#475569;font-size:10px;font-weight:700;cursor:help;vertical-align:middle;">i</span>';
}

// =====================================================================
// Identity & Kerberos Security report renderer (Phase A) - one row per
// named account/principal across Kerberoasting, AS-REP Roasting, DCSync
// rights, AdminSDHolder drift, stale accounts, password-never-expires,
// and constrained delegation. Sortable the same way as Security Risks
// Need Attention, including a persistent sort-icon affordance.
// =====================================================================
var identitySortCol = null;
var identitySortDir = 1;

function toggleIdentitySort(col) {
    if (identitySortCol === col) {
        identitySortDir = -identitySortDir;
    } else {
        identitySortCol = col;
        identitySortDir = 1;
    }
    renderIdentitySecurityReport();
}

// =====================================================================
// Identity Risk & Attack Surface tile filtering - clicking a tile narrows
// the "Critical & High Risk Users" table below to just that tile's
// accounts (including Low-risk ones that wouldn't show in the default
// view, e.g. an AES-capable non-privileged SPN account, so the table
// always matches the tile's own count). Clicking "Show all" or the same
// tile data isn't toggled off automatically - the explicit link is more
// discoverable than an implicit toggle.
// =====================================================================
var idTileLabels = {
    stale:      'Stale Accounts',
    privileged: 'Privileged Accounts',
    service:    'Risky Service Accounts',
    highrisk:   'High Risk Identities'
};

// Playbook keyed by the exact short "Why" phrases built server-side. Each
// account's Why is a " + "-joined combination of these (e.g. "Stale +
// Privileged") - the expandable panel splits on " + " and looks up each
// piece, then merges the attack risks and actions across all of them so an
// account with multiple reasons gets one combined view, not duplicates.
var idPlaybook = {
    'SPN': {
        attackRisk: 'Kerberoastable - its service ticket can be requested and cracked offline to recover the password.',
        actions: ['Rotate the password to a long random value (25+ chars) or migrate to a gMSA', 'Enforce AES encryption (msDS-SupportedEncryptionTypes)', 'Remove the SPN if the service no longer needs it']
    },
    'SPN + Privileged': {
        attackRisk: 'Highest-value Kerberoasting target - a cracked password grants privileged access directly.',
        actions: ['Treat as urgent - rotate the password immediately or migrate to a gMSA', 'Enforce AES encryption (msDS-SupportedEncryptionTypes)', 'Reconsider whether this account needs both an SPN and privileged membership']
    },
    'AS-REP roastable': {
        attackRisk: 'Pre-authentication is disabled - an attacker can request an AS-REP and crack it offline with no credentials at all.',
        actions: ['Re-enable Kerberos pre-authentication unless there is a specific, documented reason it must stay off', 'If it must stay off, ensure the password is long and random']
    },
    'DCSync rights': {
        attackRisk: 'Can request password hashes for any account in the domain - equivalent to a full domain compromise.',
        actions: ['Remove the Replicating Directory Changes (All) right from this principal immediately', 'Audit how the right was granted - often an over-scoped delegation', 'Rotate krbtgt and other high-value credentials if compromise is suspected']
    },
    'AdminSDHolder ACE': {
        attackRisk: 'Propagates to every protected (Tier 0) object on the next SDProp cycle (60 minutes by default) - a quiet persistence technique.',
        actions: ['Remove the unexpected ACE from AdminSDHolder directly', 'Confirm it does not reappear after the next SDProp cycle', 'Investigate how/when it was added']
    },
    'Stale': {
        attackRisk: 'Inactive accounts are rarely monitored - a quiet path for persistence if credentials are ever compromised.',
        actions: ['Disable the account immediately if confirmed unused', 'If still required, reset the password and document an owner', 'Remove from any privileged group if not actively needed', 'Monitor for any authentication attempts after disabling']
    },
    'Logged onto a DC': {
        attackRisk: 'Domain controllers are Tier 0 - any non-admin logon here suggests the tiering boundary is not being respected.',
        actions: ['Confirm why this account accessed the DC', 'If it was one-off troubleshooting, document it', 'If routine, move the workflow off the DC (e.g. RSAT from an admin workstation)']
    },
    'Path to Domain Admins': {
        attackRisk: 'A nested group membership or ACL right gives this principal an escalation path into Domain Admins/Enterprise Admins.',
        actions: ['Confirm every member of a nested group is meant to be Tier 0 - remove the nesting if not', 'Remove any unexpected ACL right on the Domain Admins/Enterprise Admins group object']
    },
    'Recoverable GPP password': {
        attackRisk: 'Group Policy Preferences cpassword uses a published, static key - any authenticated user can decrypt it in seconds.',
        actions: ['Remove the preference item from the GPO', 'Rotate the exposed credential immediately', 'Search SYSVOL for other legacy preference files with the same issue']
    },
    'Password never expires': {
        attackRisk: 'A password that never rotates stays valid indefinitely if it is ever leaked.',
        actions: ['Remove the never-expires flag', 'Migrate service accounts to a gMSA where possible', 'Enforce normal rotation for user accounts']
    },
    'Constrained delegation': {
        attackRisk: 'Lets this account impersonate users to specific targets - a stale or over-broad target list is a lateral movement path.',
        actions: ['Confirm every delegation target is still required', 'Migrate to resource-based constrained delegation where possible', 'Remove unused delegation entries']
    },
    'Privileged': {
        attackRisk: 'Membership in a Tier 0 group (Domain Admins/Enterprise Admins/Schema Admins) makes this one of the highest-value targets in the domain.',
        actions: ['Confirm this account still needs Tier 0 membership', 'Ensure it is used only from a Privileged Access Workstation', 'Enroll in Protected Users if not already covered', 'Review sign-in activity periodically']
    },
    'DNS zone ACL right': {
        attackRisk: 'Can create or modify records on a DNS zone directly - including a fake SRV/A record that clients and other DCs would trust and authenticate to.',
        actions: ['Remove the unexpected right from the zone object', 'Confirm only Domain Admins/Enterprise Admins/SYSTEM/DnsAdmins can modify it', 'Audit how the right was granted']
    },
    'Member of DnsAdmins': {
        attackRisk: 'DnsAdmins membership is a documented path to SYSTEM on a domain controller via the DNS service&#8217;s ServerLevelPluginDll setting.',
        actions: ['Confirm this account has a specific, documented need for DNS administration', 'Remove from DnsAdmins if not actively required', 'Treat with the same scrutiny as Domain Admins membership']
    },
    'Tier-0 group risk': {
        attackRisk: 'Member of a Tier-0 group (Domain Admins/Enterprise Admins/Schema Admins/Administrators/Backup Operators/Account Operators/Print Operators) while disabled, stale, a service account, or carrying an unusually large number of group memberships.',
        actions: ['Remove if disabled or stale', 'Confirm service account membership is still required', 'Investigate why this account carries this many group memberships if flagged for token size']
    },
    'Delegated Tier-0 right': {
        attackRisk: 'Holds a delegated right (a generic ACL right, or specifically Reset Password) on a Tier-0 group object - can escalate without ever being a member of that group today.',
        actions: ['Remove the unexpected right from the Tier-0 group object', 'Confirm only Domain Admins/Enterprise Admins/SYSTEM can modify it or reset passwords for its members', 'Audit how the right was granted']
    },
    'Foreign security principal': {
        attackRisk: 'An account or group from another domain or forest holds direct membership in a Tier-0 group - this domain&#8217;s security now depends partly on the trusted domain&#8217;s security practices.',
        actions: ['Confirm this cross-domain/cross-forest access is intentional and still required', 'Remove if not', 'Prefer a narrower, resource-specific grant over Tier-0 membership for legitimate cross-domain needs']
    }
};

// "Member of <Group1, Group2>" reasons (built server-side from the
// account's actual privileged group membership) match the generic
// 'Privileged' playbook entry, but with the specific group name(s)
// substituted into the action text, e.g. "Confirm this account still
// needs membership in Domain Admins" instead of a generic "Tier 0".
function idPlaybookFor(reason) {
    if (idPlaybook[reason]) { return idPlaybook[reason]; }
    if (reason.indexOf('Member of ') === 0) {
        var groups = reason.substring('Member of '.length);
        var base = idPlaybook['Privileged'];
        return {
            attackRisk: 'Member of ' + groups + ' (a Tier 0 group) - one of the highest-value targets in the domain.',
            actions: base.actions.map(function (a) {
                return a.replace('Tier 0 membership', 'membership in ' + groups);
            })
        };
    }
    return null;
}

// Decodes the "::Step0|Edge1::Step1|Edge2::Step2" chain produced server-side
// (Encode-PathSteps in the PowerShell collector) into an ordered array of
// {edge, label} pairs - edge is the relationship INTO that step (empty for
// the first/starting account).
function parsePathSteps(encoded) {
    if (!encoded) { return []; }
    return encoded.split('|').map(function (part) {
        var idx = part.indexOf('::');
        if (idx === -1) { return { edge: '', label: part }; }
        return { edge: part.substring(0, idx), label: part.substring(idx + 2) };
    });
}

// Renders the attack path as the literal sequence an attacker (or an
// over-permissioned account) would walk - same underlying finding as the
// Attack Risk text below, shown step-by-step with arrows instead of one
// flat sentence. The final node (the Tier 0 target) is highlighted in red.
// This is a structured DISPLAY of a static finding already collected -
// not a live simulation of attacker behavior.
function buildAttackPathStepsHtml(encoded) {
    var steps = parsePathSteps(encoded);
    if (steps.length < 2) { return ''; }

    var chain = '<div style="display:flex;flex-wrap:wrap;align-items:center;gap:6px;">';
    steps.forEach(function (s, i) {
        if (i > 0) {
            chain += '<span style="color:#94a3b8;font-size:11px;white-space:nowrap;">&#8594; <i>' + escapeHtml(s.edge) + '</i> &#8594;</span>';
        }
        var isLast = (i === steps.length - 1);
        var style = isLast
            ? 'background:#fee2e2;border:1px solid #dc2626;color:#7f1d1d;font-weight:700;'
            : 'background:#eef2ff;border:1px solid #c7d2fe;color:#1e3a8a;font-weight:600;';
        chain += '<span style="padding:3px 9px;border-radius:6px;font-size:12px;' + style + '">' + escapeHtml(s.label) + '</span>';
    });
    chain += '</div>';

    var lastEdge = steps[steps.length - 1].edge || '';
    var target   = steps[steps.length - 1].label;
    var impact = (lastEdge.indexOf('nested') !== -1 || lastEdge.indexOf('member of') !== -1)
        ? 'No further action required - every account in this chain is already an effective member of ' + target + ' today.'
        : 'One step from full membership: the account at the end of this chain can add itself directly to ' + target + ' using the right it already holds.';

    return chain + '<div style="margin-top:8px;font-size:11.5px;color:#7f1d1d;"><b>Impact:</b> ' + escapeHtml(impact) + '</div>';
}

function buildIdentityDetailHtml(row) {
    var reasons = (row.Why || '').split(' + ').map(function (s) { return s.trim(); }).filter(Boolean);
    var attackRisks = [];
    var actionsSeen = {};
    var actions = [];

    reasons.forEach(function (r) {
        var pb = idPlaybookFor(r);
        if (!pb) { return; }
        if (attackRisks.indexOf(pb.attackRisk) === -1) { attackRisks.push(pb.attackRisk); }
        pb.actions.forEach(function (a) {
            if (!actionsSeen[a]) { actionsSeen[a] = true; actions.push(a); }
        });
    });

    var sevBadge = '<span class="' + (idBadgeClassFor(row.Risk)) + '">' + escapeHtml(row.Risk) + '</span>';
    var whyHtml = reasons.length
        ? '<ul style="margin:4px 0 0;padding-left:18px;">' + reasons.map(function (r) { return '<li>' + escapeHtml(r) + '</li>'; }).join('') + '</ul>'
        : '<span style="color:#94a3b8;">&mdash;</span>';
    var attackRiskHtml = attackRisks.length ? escapeHtml(attackRisks.join(' ')) : 'No specific attack path documented for this combination.';
    var actionsHtml = actions.length
        ? '<ol style="margin:4px 0 0;padding-left:18px;">' + actions.map(function (a) { return '<li>' + escapeHtml(a) + '</li>'; }).join('') + '</ol>'
        : '<span style="color:#94a3b8;">No specific actions documented.</span>';
    var pathHtml = buildAttackPathStepsHtml(row.PathSteps);
    var pathSectionHtml = pathHtml
        ? '<div style="margin-top:10px;"><b>Attack Path Walkthrough:</b><div style="margin-top:6px;">' + pathHtml + '</div></div>'
        : '';

    return '<div style="padding:14px 18px;background:#f8fafc;border-radius:8px;font-size:12.5px;line-height:1.7;">' +
        '<div><b>Risk:</b> ' + sevBadge + '</div>' +
        '<div style="margin-top:10px;"><b>Why:</b>' + whyHtml + '</div>' +
        pathSectionHtml +
        '<div style="margin-top:10px;"><b>Attack Risk:</b><br>' + attackRiskHtml + '</div>' +
        '<div style="margin-top:10px;"><b>Recommended Actions:</b>' + actionsHtml + '</div>' +
        '<div style="margin-top:10px;"><b>Owner:</b> <span style="color:#94a3b8;">Not tracked by this tool</span></div>' +
        '<div style="margin-top:4px;"><b>Last Logon:</b> ' + escapeHtml(row.LastLogon || 'Unknown') + '</div>' +
        '</div>';
}

function idBadgeClassFor(risk) {
    var badgeClass = { Critical: 'rpt-b id-critical', High: 'rpt-b rpt-crit', Medium: 'rpt-b rpt-medium', Low: 'rpt-b rpt-na' };
    return badgeClass[risk] || 'rpt-b rpt-na';
}

// Short, one-line version of the FIRST applicable playbook action, for the
// "Recommended Action" table column (the full numbered list only shows in
// the expanded detail panel).
function idShortAction(why) {
    var reasons = (why || '').split(' + ').map(function (s) { return s.trim(); }).filter(Boolean);
    for (var i = 0; i < reasons.length; i++) {
        var pb = idPlaybookFor(reasons[i]);
        if (pb && pb.actions && pb.actions.length) { return pb.actions[0]; }
    }
    return 'Review account';
}

var idCurrentRows = [];

function toggleIdentityDetail(rowIndex, trEl) {
    var existing = trEl.nextElementSibling;
    if (existing && existing.classList && existing.classList.contains('id-detail-row')) {
        existing.parentNode.removeChild(existing);
        trEl.querySelector('.id-expand-chevron').innerHTML = '&#9656;';
        return;
    }
    // Close any other open detail row first - one at a time keeps the table readable.
    document.querySelectorAll('.id-detail-row').forEach(function (el) { el.parentNode.removeChild(el); });
    document.querySelectorAll('.id-expand-chevron').forEach(function (el) { el.innerHTML = '&#9656;'; });

    var row = idCurrentRows[rowIndex];
    if (!row) { return; }

    var detailTr = document.createElement('tr');
    detailTr.className = 'id-detail-row';
    var detailTd = document.createElement('td');
    detailTd.colSpan = 5;
    detailTd.style.padding = '0';
    detailTd.innerHTML = buildIdentityDetailHtml(row);
    detailTr.appendChild(detailTd);
    trEl.parentNode.insertBefore(detailTr, trEl.nextSibling);
    trEl.querySelector('.id-expand-chevron').innerHTML = '&#9662;';
}

function filterIdentityUsersTable(tag, tileEl) {
    var rows = parseSimpleCsv(typeof identityUsersCsv !== 'undefined' ? identityUsersCsv : '');
    var displayRank = { Critical: 0, High: 1, Medium: 2, Low: 3 };

    var filtered = !tag
        ? rows.filter(function (r) { return r.Risk !== 'Low'; })
        : rows.filter(function (r) { return (r.Tags || '').split(',').indexOf(tag) !== -1; });

    filtered.sort(function (a, b) { return (displayRank[a.Risk] || 9) - (displayRank[b.Risk] || 9) || a.Name.localeCompare(b.Name); });
    idCurrentRows = filtered;

    var tbody = document.getElementById('idUsersTbody');
    if (!tbody) { return; }

    if (filtered.length === 0) {
        tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;padding:20px;color:#94a3b8;font-style:italic;">No accounts found for this filter</td></tr>';
    } else {
        tbody.innerHTML = filtered.map(function (r, idx) {
            return '<tr onclick="toggleIdentityDetail(' + idx + ', this)" style="cursor:pointer;">' +
                '<td class="rpt-dc"><span class="id-expand-chevron" style="display:inline-block;width:14px;color:#94a3b8;">&#9656;</span>' + escapeHtml(r.Name) + '</td>' +
                '<td>' + escapeHtml(r.Type) + '</td>' +
                '<td><span class="' + idBadgeClassFor(r.Risk) + '">' + escapeHtml(r.Risk) + '</span></td>' +
                '<td>' + escapeHtml(r.Why) + '</td>' +
                '<td>' + escapeHtml(idShortAction(r.Why)) + '</td>' +
                '</tr>';
        }).join('');
    }

    // Highlight the active tile (if any) and reset the others.
    document.querySelectorAll('#idTileStale, #idTilePrivileged, #idTileService, #idTileHighRisk').forEach(function (el) {
        el.style.boxShadow = '';
    });
    if (tileEl) { tileEl.style.boxShadow = '0 0 0 2px #2563eb'; }

    var titleEl = document.getElementById('idTableSectionTitle');
    var noteEl  = document.getElementById('idTableNote');
    var linkEl  = document.getElementById('idShowAllLink');

    if (!tag) {
        if (titleEl) { titleEl.textContent = 'Critical & High Risk Users'; }
        if (noteEl)  { noteEl.innerHTML = '&#8505; Sorted Critical → High → Medium. Every Critical/High/Medium identity finding is shown - this list is not capped. Click a name to see the full recommended actions.'; }
        if (linkEl)  { linkEl.style.display = 'none'; }
    } else {
        if (titleEl) { titleEl.textContent = idTileLabels[tag] || 'Filtered Users'; }
        if (noteEl)  { noteEl.innerHTML = '&#8505; Showing every account counted in the "' + (idTileLabels[tag] || tag) + '" tile (' + filtered.length + '), including any not otherwise shown by default.'; }
        if (linkEl)  { linkEl.style.display = 'inline'; }
    }
}

// =====================================================================
// Group & Privilege Architecture tab - ONE ROW PER GROUP (not per
// account - per-account detail for the same Tier-0 members, FSPs, and
// delegation findings already lives on the Identity Risk & Attack
// Surface tab above, so this tab discusses the groups themselves instead
// of repeating the same names a second time). Columns deliberately match
// that tab exactly: Name | Type | Risk | Why | Recommended Action.
// Clicking a row expands the full member list and the specific named
// accounts behind whichever findings touch that group.
// =====================================================================
var gpaTileLabels = {
    membership: 'Tier-0 Group Membership',
    nesting:    'Nesting Depth & Cycles',
    fsp:        'Foreign Security Principals',
    delegation: 'Group-Based Delegation',
    legacy:     'Legacy & Universal Group Risk'
};
// Tile filters check which DETAIL FIELD is populated on a row, rather than
// a category string - each row can carry several findings at once now
// (e.g. one group can have both a delegation finding AND an oversized
// member count), so a simple category list no longer fits.
var gpaTileFilters = {
    membership: function (r) { return !!(r.RiskyDetail || r.MemberCount && r.Why.indexOf('exceeds threshold') !== -1); },
    nesting:    function (r) { return r.Type === 'Structural' || !!r.NestingDetail; },
    fsp:        function (r) { return !!r.FspDetail; },
    delegation: function (r) { return !!r.AclDetail; },
    legacy:     function (r) { return r.Type === 'Legacy Group' || !!r.UnivDetail; }
};
var gpaCurrentRows = [];

// Shared plain-bordered-grid builder for every sub-table in this detail
// panel (Risky members, Delegated rights, Universal groups nested,
// Foreign security principals). Every cell gets an explicit border,
// background, and text color so nothing ambient bleeds in (that's what
// caused the earlier stray blue header). Sized by content but capped at
// max-width:80% rather than width:100% - and explicitly margin-left:0 /
// align-self:flex-start so it sits flush at the left edge under the
// panel's existing padding, never centered, regardless of what an
// ancestor's flex/grid alignment might otherwise do. word-break is
// reset to normal (with overflow-wrap as the fallback for a genuinely
// unbreakable token) and each cell gets a sane min-width - without both
// of those, the inherited aggressive word-break rule from the surrounding
// report table let the column collapse so narrow that ordinary words got
// chopped mid-token, which is what went wrong the first time.
function gpaExcelTable(headers, rowsOfCells) {
    var gridBorder = '1px solid #cbd5e1';
    var cellBase = 'border:' + gridBorder + ';word-break:normal;overflow-wrap:break-word;min-width:110px;';
    var thStyle = 'text-align:left;padding:6px 10px;background:#f1f5f9;color:#0f172a;font-weight:600;font-size:12px;white-space:nowrap;' + cellBase;
    var tdStyle = 'padding:6px 10px;background:#ffffff;font-size:12px;color:#1e293b;' + cellBase;

    var headHtml = headers.map(function (h) { return '<th style="' + thStyle + '">' + escapeHtml(h) + '</th>'; }).join('');
    var bodyHtml = rowsOfCells.map(function (cells) {
        return '<tr>' + cells.map(function (c) { return '<td style="' + tdStyle + '">' + escapeHtml(c) + '</td>'; }).join('') + '</tr>';
    }).join('');

    return '<table style="display:table;margin:6px 0 0 0;margin-left:0;align-self:flex-start;max-width:80%;border-collapse:collapse;border:' + gridBorder + ';">' +
        '<thead><tr>' + headHtml + '</tr></thead><tbody>' + bodyHtml + '</tbody></table>';
}

// Pipe-delimited detail fields ("name: reason|name: reason") -> the shared
// grid above. Entries with no ": reason" suffix (e.g. a bare account
// name) just leave the second column blank rather than showing nothing.
function gpaDetailTable(raw, col1Label, col2Label) {
    if (!raw) { return '<span style="color:#94a3b8;">&mdash;</span>'; }
    var rowsOfCells = raw.split('|').filter(Boolean).map(function (s) {
        var idx = s.indexOf(': ');
        return idx === -1 ? [s, ''] : [s.substring(0, idx), s.substring(idx + 2)];
    });
    return gpaExcelTable([col1Label || 'Account', col2Label || 'Detail'], rowsOfCells);
}

// Risky-member entries are encoded "Account::Status::Reason", one entry
// PER ISSUE (an account with two problems gets two entries).
function gpaRiskyMembersTable(raw) {
    if (!raw) { return '<span style="color:#94a3b8;">&mdash;</span>'; }
    var rowsOfCells = raw.split('|').filter(Boolean).map(function (s) { return s.split('::'); });
    return gpaExcelTable(['Username', 'Status', 'Reason'], rowsOfCells);
}

// Plain list of names (e.g. AllMembers) -> a single-column grid, same
// styling as everything else here rather than a comma-separated blob -
// this is the part that has to hold up in a much larger AD environment,
// where a Tier-0 group could have far more members than whatever was
// used to test this.
function gpaNameListTable(raw, label) {
    if (!raw) { return '<span style="color:#94a3b8;">No members</span>'; }
    var rowsOfCells = raw.split('|').filter(Boolean).map(function (s) { return [s]; });
    return gpaExcelTable([label || 'Name'], rowsOfCells);
}

// PowerShell caps every list at 50 entries before it ever reaches the
// page (Get-CappedPipeJoin) - shownCount is always what's actually in the
// pipe-delimited string, trueTotal is the real, uncapped count from the
// matching *Total/MemberCount field. This renders the note ONLY when they
// differ, so a small environment with 6 members never sees a pointless
// "showing 6 of 6".
function gpaCapNote(trueTotal, shownCount) {
    var total = parseInt(trueTotal, 10) || 0;
    if (total <= shownCount) { return ''; }
    return '<div style="margin-top:4px;font-size:11px;color:#94a3b8;">Showing ' + shownCount + ' of ' + total + ' - the rest are omitted here to keep this panel readable.</div>';
}

function buildGroupPrivDetailHtml(row) {
    var sevBadge = '<span class="' + idBadgeClassFor(row.Risk) + '">' + escapeHtml(row.Risk) + '</span>';

    var memberSection = '';
    if (row.AllMembers) {
        var shownMembers = row.AllMembers.split('|').filter(Boolean).length;
        memberSection = '<div style="margin-top:10px;"><b>All members (' + (row.MemberCount || shownMembers) + '):</b>' +
            gpaNameListTable(row.AllMembers, 'Member') + gpaCapNote(row.MemberCount, shownMembers) + '</div>';
    }

    var riskySection = row.RiskyDetail
        ? '<div style="margin-top:10px;"><b>Risky members:</b>' + gpaRiskyMembersTable(row.RiskyDetail) +
          gpaCapNote(row.RiskyTotal, row.RiskyDetail.split('|').filter(Boolean).length) + '</div>'
        : '';

    var sections = riskySection + [
        ['Delegated rights', row.AclDetail, row.AclTotal, 'Account', 'Right'],
        ['Universal groups nested', row.UnivDetail, row.UnivTotal, 'Group', ''],
        ['Foreign security principals', row.FspDetail, row.FspTotal, 'Account', '']
    ].filter(function (s) { return s[1]; }).map(function (s) {
        var shown = s[1].split('|').filter(Boolean).length;
        return '<div style="margin-top:10px;"><b>' + s[0] + ':</b>' + gpaDetailTable(s[1], s[3], s[4]) + gpaCapNote(s[2], shown) + '</div>';
    }).join('');

    var nestingSection = row.NestingDetail
        ? '<div style="margin-top:10px;"><b>Nesting:</b> ' + escapeHtml(row.NestingDetail) + '</div>'
        : '';

    return '<div style="padding:14px 18px;background:#f8fafc;border-radius:8px;font-size:12.5px;line-height:1.7;">' +
        '<div><b>Risk:</b> ' + sevBadge + '</div>' +
        '<div style="margin-top:10px;"><b>Why:</b> ' + escapeHtml(row.Why || '') + '</div>' +
        '<div style="margin-top:10px;"><b>Recommended Action:</b> ' + escapeHtml(row.Action || '') + '</div>' +
        memberSection + sections + nestingSection +
        '</div>';
}

function toggleGroupPrivDetail(rowIndex, trEl) {
    var existing = trEl.nextElementSibling;
    if (existing && existing.classList && existing.classList.contains('id-detail-row')) {
        existing.parentNode.removeChild(existing);
        trEl.querySelector('.id-expand-chevron').innerHTML = '&#9656;';
        return;
    }
    document.querySelectorAll('#gpaTbody .id-detail-row').forEach(function (el) { el.parentNode.removeChild(el); });
    document.querySelectorAll('#gpaTbody .id-expand-chevron').forEach(function (el) { el.innerHTML = '&#9656;'; });

    var row = gpaCurrentRows[rowIndex];
    if (!row) { return; }

    var detailTr = document.createElement('tr');
    detailTr.className = 'id-detail-row';
    var detailTd = document.createElement('td');
    detailTd.colSpan = 5;
    detailTd.style.padding = '0';
    detailTd.innerHTML = buildGroupPrivDetailHtml(row);
    detailTr.appendChild(detailTd);
    trEl.parentNode.insertBefore(detailTr, trEl.nextSibling);
    trEl.querySelector('.id-expand-chevron').innerHTML = '&#9662;';
}

function filterGroupPrivTable(tag, tileEl) {
    var rows = parseSimpleCsv(typeof gpaGroupRowsCsv !== 'undefined' ? gpaGroupRowsCsv : '');
    var sevRank = { High: 0, Medium: 1, Low: 2 };

    var filtered = !tag
        ? rows
        : rows.filter(gpaTileFilters[tag] || function () { return true; });

    filtered.sort(function (a, b) { return (sevRank[a.Risk] || 9) - (sevRank[b.Risk] || 9); });
    gpaCurrentRows = filtered;

    var tbody = document.getElementById('gpaTbody');
    if (!tbody) { return; }

    if (filtered.length === 0) {
        tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;padding:20px;color:#94a3b8;font-style:italic;">No findings for this filter (Good)</td></tr>';
    } else {
        tbody.innerHTML = filtered.map(function (r, idx) {
            return '<tr onclick="toggleGroupPrivDetail(' + idx + ', this)" style="cursor:pointer;">' +
                '<td class="rpt-dc"><span class="id-expand-chevron" style="display:inline-block;width:14px;color:#94a3b8;">&#9656;</span>' + escapeHtml(r.Name) + '</td>' +
                '<td>' + escapeHtml(r.Type) + '</td>' +
                '<td><span class="' + idBadgeClassFor(r.Risk) + '">' + escapeHtml(r.Risk) + '</span></td>' +
                '<td>' + escapeHtml(r.Why) + '</td>' +
                '<td>' + escapeHtml(r.Action) + '</td>' +
                '</tr>';
        }).join('');
    }

    document.querySelectorAll('#gpaTileMembership, #gpaTileNesting, #gpaTileFsp, #gpaTileDelegation, #gpaTileLegacy').forEach(function (el) {
        el.style.boxShadow = '';
    });
    if (tileEl) { tileEl.style.boxShadow = '0 0 0 2px #2563eb'; }

    var titleEl = document.getElementById('gpaTableSectionTitle');
    var noteEl  = document.getElementById('gpaTableNote');
    var linkEl  = document.getElementById('gpaShowAllLink');

    if (!tag) {
        if (titleEl) { titleEl.textContent = 'All Group & Privilege Findings'; }
        if (noteEl)  { noteEl.innerHTML = '&#8505; One row per group. Click a row to see the full member list and named accounts behind each finding.'; }
        if (linkEl)  { linkEl.style.display = 'none'; }
    } else {
        if (titleEl) { titleEl.textContent = gpaTileLabels[tag] || 'Filtered Findings'; }
        if (noteEl)  { noteEl.innerHTML = '&#8505; Showing every group counted in the "' + (gpaTileLabels[tag] || tag) + '" tile (' + filtered.length + ').'; }
        if (linkEl)  { linkEl.style.display = 'inline'; }
    }
}

// =====================================================================
// Shared expand-on-click detail panel for the flat report tables (Identity
// & Kerberos Security, Security Risks Need Attention) - #rptTbody is
// shared across all 10 Reports dropdown items, so rather than one toggle
// function per report, a single shared one reads whichever row array and
// detail-builder the CURRENTLY rendered report last set. Re-rendering a
// report replaces #rptTbody.innerHTML wholesale, which already destroys
// any previously open detail row, so no extra cleanup is needed when
// switching reports.
// =====================================================================
var rptCurrentRows   = [];
var rptDetailBuilder = null;
var rptDetailColSpan = 4;

function toggleRptDetailRow(rowIndex, trEl) {
    var existing = trEl.nextElementSibling;
    if (existing && existing.classList && existing.classList.contains('id-detail-row')) {
        existing.parentNode.removeChild(existing);
        trEl.querySelector('.id-expand-chevron').innerHTML = '&#9656;';
        return;
    }
    document.querySelectorAll('#rptTbody .id-detail-row').forEach(function (el) { el.parentNode.removeChild(el); });
    document.querySelectorAll('#rptTbody .id-expand-chevron').forEach(function (el) { el.innerHTML = '&#9656;'; });

    var row = rptCurrentRows[rowIndex];
    if (!row || !rptDetailBuilder) { return; }

    var detailTr = document.createElement('tr');
    detailTr.className = 'id-detail-row';
    var detailTd = document.createElement('td');
    detailTd.colSpan = rptDetailColSpan;
    detailTd.style.padding = '0';
    detailTd.innerHTML = rptDetailBuilder(row);
    detailTr.appendChild(detailTd);
    trEl.parentNode.insertBefore(detailTr, trEl.nextSibling);
    trEl.querySelector('.id-expand-chevron').innerHTML = '&#9662;';
}

// Shared core for both reports below - they differ only in which field
// holds the remediationGuide category key and which field/label
// identifies the row's subject (an account for Identity findings, a DC
// for Security Risks).
function buildFindingDetailHtmlGeneric(row, categoryKey, subjectKey, subjectLabel) {
    var category = row[categoryKey];
    var g = remediationGuide[category];
    var sevBadge   = '<span class="' + idBadgeClassFor(row.Severity) + '">' + escapeHtml(row.Severity || '') + '</span>';
    var whyHtml    = g ? escapeHtml(g.why) : 'No additional guidance documented for this category.';
    var fixHtml    = g ? escapeHtml(g.fix) : '<span style="color:#94a3b8;">No specific actions documented.</span>';
    var mitreHtml  = (g && g.mitre) ? ('<div style="margin-top:6px;"><b>MITRE ATT&amp;CK:</b> ' + escapeHtml(g.mitre) + '</div>') : '';
    var refHtml    = (g && g.ref) ? ('<div style="margin-top:6px;"><b>Reference:</b> <a href="' + escapeHtml(g.ref) + '" target="_blank" rel="noopener">' + escapeHtml(g.ref) + '</a></div>') : '';
    var subjectVal = row[subjectKey];
    var subjectHtml = subjectVal ? ('<div style="margin-top:10px;"><b>' + subjectLabel + ':</b> ' + escapeHtml(subjectVal) + '</div>') : '';
    var detailText = row.Detail || row.Risk || '';

    return '<div style="padding:14px 18px;background:#f8fafc;border-radius:8px;font-size:12.5px;line-height:1.7;">' +
        '<div><b>Severity:</b> ' + sevBadge + '</div>' +
        subjectHtml +
        '<div style="margin-top:10px;"><b>Finding:</b> ' + escapeHtml(detailText) + '</div>' +
        '<div style="margin-top:10px;"><b>Why it matters:</b> ' + whyHtml + '</div>' +
        '<div style="margin-top:10px;"><b>How to fix:</b> ' + fixHtml + '</div>' +
        mitreHtml + refHtml +
        '</div>';
}
function buildIdentityFindingDetailHtml(row) { return buildFindingDetailHtmlGeneric(row, 'Category', 'Account', 'Account / Principal'); }
function buildSecurityRiskDetailHtml(row)    { return buildFindingDetailHtmlGeneric(row, 'Source', 'DC', 'Domain Controller'); }

function renderIdentitySecurityReport() {
    var findings = parseSimpleCsv(typeof identityFindingsCsv !== 'undefined' ? identityFindingsCsv : '');

    var sevRank = { High: 0, Medium: 1, Low: 2 };

    if (identitySortCol === 'Category') {
        findings.sort(function (a, b) { return String(a.Category || '').localeCompare(String(b.Category || '')) * identitySortDir; });
    } else if (identitySortCol === 'Account') {
        findings.sort(function (a, b) { return String(a.Account || '').localeCompare(String(b.Account || '')) * identitySortDir; });
    } else {
        findings.sort(function (a, b) {
            var ra = sevRank.hasOwnProperty(a.Severity) ? sevRank[a.Severity] : 3;
            var rb = sevRank.hasOwnProperty(b.Severity) ? sevRank[b.Severity] : 3;
            return ra - rb;
        });
    }

    function sortIcon(col) {
        if (identitySortCol !== col) { return ' <span style="opacity:0.35;font-size:10px;">&#8693;</span>'; }
        return identitySortDir === 1 ? ' <span style="font-size:11px;">&#9650;</span>' : ' <span style="font-size:11px;">&#9660;</span>';
    }

    var thead = '<tr class="rpt-header-row">' +
        '<th>Severity</th>' +
        '<th class="rpt-gd" style="cursor:pointer;user-select:none;" title="Click to sort by category" onclick="toggleIdentitySort(\'Category\')">Category' + sortIcon('Category') + '</th>' +
        '<th style="cursor:pointer;user-select:none;" title="Click to sort by account" onclick="toggleIdentitySort(\'Account\')">Account / Principal' + sortIcon('Account') + '</th>' +
        '<th>Detail <span style="font-weight:400;opacity:0.6;font-size:10px;">(click row to expand)</span></th>' +
        '</tr>';

    var tbody = '';
    if (findings.length === 0) {
        tbody = '<tr><td colspan="4" style="text-align:center;padding:20px;color:#94a3b8;font-style:italic;">No identity or Kerberos security findings detected</td></tr>';
    } else {
        tbody = findings.map(function (f, idx) {
            var sev = f.Severity || '';
            var sevBadge = sev === 'High' ? rptCrit('High') : (sev === 'Medium' ? rptMedium('Medium') : rptNa(sev || 'Low'));
            return '<tr style="border-bottom:1px solid #f1f5f9;cursor:pointer;" onclick="toggleRptDetailRow(' + idx + ', this)">' +
                '<td style="padding:7px 10px;">' + sevBadge + ' ' + riskScoreBadge(f.Category) + '</td>' +
                '<td class="rpt-dc" style="padding:7px 10px;font-size:12px;color:#0f172a;font-weight:600;"><span class="id-expand-chevron" style="display:inline-block;width:14px;color:#94a3b8;">&#9656;</span>' + escapeHtml(f.Category || '') + remediationBadge(f.Category) + '</td>' +
                '<td style="padding:7px 10px;font-weight:700;font-size:12px;color:#0f172a;">' + escapeHtml(f.Account || '') + '</td>' +
                '<td style="padding:7px 10px;font-size:11.5px;color:#64748b;">' + escapeHtml(f.Detail || '') + '</td>' +
                '</tr>';
        }).join('');
    }

    var highCount   = findings.filter(function (f) { return f.Severity === 'High'; }).length;
    var mediumCount = findings.filter(function (f) { return f.Severity === 'Medium'; }).length;

    rptCurrentRows   = findings;
    rptDetailBuilder = buildIdentityFindingDetailHtml;
    rptDetailColSpan = 4;

    document.getElementById('rptTitle').textContent = 'Identity & Kerberos Security';
    document.getElementById('rptThead').innerHTML   = thead;
    document.getElementById('rptTbody').innerHTML   = tbody;
    document.getElementById('rptNote').innerHTML    =
        '&#8505; Total: ' + findings.length + ' finding(s) - ' + highCount + ' High, ' + mediumCount + ' Medium. ' +
        'Covers Kerberoasting, AS-REP Roasting, DCSync rights, AdminSDHolder drift, stale accounts, password-never-expires, and constrained delegation. ' +
        'Click Category or Account to sort, click a row to expand why it matters and how to fix it.';
}

// Global Reports theme toggle - applies to every report at once, since they
// all share the same .rpt-header-row / .rpt-group-row / .rpt-b (+good/warn/
// crit/na) classes. Pure CSS attribute-selector theming - no re-render of
// the currently displayed report's content is needed, just the data-theme
// attribute flips and the existing CSS rules re-apply instantly.
// Executive Summary KPI row style toggle - lets leadership switch between
// the default slim-bar cards and a circular-donut view for the same 4
// metrics (Overall Health, Domain Controllers, DNS Status, Replication
// Health). Both rows are rendered server-side already with the right
// values - this just shows/hides the two pre-built rows, no re-render.
function toggleExecKpiStyle(style) {
    var barsRow    = document.getElementById('execRow1Bars');
    var donutsRow  = document.getElementById('execRow1Donuts');
    var btnBars    = document.getElementById('execKpiBtnBars');
    var btnDonuts  = document.getElementById('execKpiBtnDonuts');
    if (!barsRow || !donutsRow) { return; }

    if (style === 'donuts') {
        barsRow.style.display   = 'none';
        donutsRow.style.display = 'grid';
        if (btnBars)   { btnBars.classList.remove('active'); }
        if (btnDonuts) { btnDonuts.classList.add('active'); }
    } else {
        barsRow.style.display   = 'grid';
        donutsRow.style.display = 'none';
        if (btnDonuts) { btnDonuts.classList.remove('active'); }
        if (btnBars)   { btnBars.classList.add('active'); }
    }
}

function toggleReportTheme() {
    var panel = document.getElementById('rptPanel');
    var btn   = document.getElementById('rptThemeToggleBtn');
    if (!panel) { return; }
    var next = (panel.getAttribute('data-theme') === 'A') ? 'B' : 'A';
    panel.setAttribute('data-theme', next);
    if (btn) { btn.innerHTML = (next === 'A') ? '&#127912; Switch Style (B)' : '&#127912; Switch Style (A)'; }

    // Keep the Identity Risk & Attack Surface table in sync with the SAME
    // toggle - it lives outside #rptPanel (a different tab entirely), so it
    // needs its own data-theme attribute mirrored here rather than relying
    // on the #rptPanel-scoped CSS.
    var idTable = document.getElementById('identityUsersTable');
    if (idTable) { idTable.setAttribute('data-theme', next); }

    var gpaTable = document.getElementById('groupPrivTable');
    if (gpaTable) { gpaTable.setAttribute('data-theme', next); }
}

function renderReport(type) {
    // Reset rptNote display to CSS default (flex) before each render.
    // renderSitesReport() overrides to 'block' for its multi-element Section 2.
    document.getElementById('rptNote').style.display = '';

    if (type === 'health')            { renderHealthReport(); }
    else if (type === 'inventory')    { renderInventoryReport(); }
    else if (type === 'dcbysite')     { renderDcBySiteReport(); }
    else if (type === 'replpartners') { renderReplPartnersReport(); }
    else if (type === 'security')     { renderSecurityReport(); }
    else if (type === 'connectivity') { renderConnectivityReport(); }
    else if (type === 'dnszones')     { renderDnsZonesReport(); }
    else if (type === 'forwarders')   { renderForwardersReport(); }
    else if (type === 'sites')        { renderSitesReport(); }
    else if (type === 'gpo')          { renderGpoReport(); }
    else if (type === 'secrisks')     { renderSecurityRisksReport(); }
    else if (type === 'dcuptime')     { renderDcUptimeReport(); }
    else if (type === 'identitysec')  { renderIdentitySecurityReport(); }

    document.querySelectorAll('.rpt-dd-item').forEach(function (el) { el.classList.remove('active'); });
    var el = document.getElementById('rdd-' + type);
    if (el) { el.classList.add('active'); }
}

// ---- export to Excel via SheetJS ----
function stripHtmlFromSheet(ws) {
    var range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    for (var R = range.s.r; R <= range.e.r; R++) {
        for (var C = range.s.c; C <= range.e.c; C++) {
            var addr = XLSX.utils.encode_cell({ r: R, c: C });
            if (ws[addr] && ws[addr].v) {
                ws[addr].v = String(ws[addr].v).replace(/<[^>]+>/g, '').trim();
                ws[addr].t = 's';
            }
        }
    }
}

function exportCurrentReport() {
    if (typeof XLSX === 'undefined') {
        alert('SheetJS library not loaded. Please check your internet connection.');
        return;
    }
    var tbl = document.getElementById('rptTable');
    var wb  = XLSX.utils.book_new();
    var ws  = XLSX.utils.table_to_sheet(tbl, { raw: false });
    stripHtmlFromSheet(ws);

    var names = {
        health: 'DC_Health_Check', inventory: 'DC_Server_Inventory',
        dcbysite: 'DC_List_by_Site',
        replpartners: 'AD_Replication_Partners',
        security: 'Security_Posture', connectivity: 'Connectivity_Matrix',
        dnszones: 'DNS_Zones', forwarders: 'Conditional_Forwarders',
        sites: 'Sites_and_Services_Section1', gpo: 'Group_Policies',
        secrisks: 'Security_Risks_Need_Attention',
        dcuptime: 'DC_Uptime_and_Reboot_Status',
        identitysec: 'Identity_and_Kerberos_Security'
    };
    var sheetName = (names[currentReport] || 'Report').replace(/_/g, ' ');
    var fileName  = (names[currentReport] || 'Report') + '_' + new Date().toISOString().slice(0, 10) + '.xlsx';

    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    XLSX.writeFile(wb, fileName);
}

function exportSiteLinksSec2() {
    if (typeof XLSX === 'undefined') {
        alert('SheetJS library not loaded. Please check your internet connection.');
        return;
    }
    var tbl = document.getElementById('siteLinksSec2Table');
    if (!tbl) { alert('Site Links table not found. Please open the Sites & Services report first.'); return; }
    var wb  = XLSX.utils.book_new();
    var ws  = XLSX.utils.table_to_sheet(tbl, { raw: false });
    stripHtmlFromSheet(ws);
    XLSX.utils.book_append_sheet(wb, ws, 'Site Links');
    XLSX.writeFile(wb, 'Site_Links_' + new Date().toISOString().slice(0, 10) + '.xlsx');
}

// initialise dropdown active state
document.getElementById('rdd-health').classList.add('active');

// Dynamic top padding - keeps content clear of the fixed sticky header
(function() {
    function adjustTopPadding() {
        var hdr = document.getElementById('stickyTop');
        var cnt = document.getElementById('mainContainer');
        if (hdr && cnt) { cnt.style.paddingTop = (hdr.offsetHeight + 16) + 'px'; }
    }
    adjustTopPadding();
    window.addEventListener('resize', adjustTopPadding);
})();

// Position the tab pointer arrow and ribbon correctly on first load, since
// showTab() only runs when a tab is clicked - the default active tab
// (Executive Summary) needs the same sync to happen once up front.
(function() {
    function positionArrowOnLoad() {
        var strip = document.getElementById('tabIndicatorStrip');
        var arrow = document.getElementById('tabIndicatorArrow');
        var activeBtn = document.querySelector('.tab-button.active');
        if (strip && arrow && activeBtn) {
            var stripRect = strip.getBoundingClientRect();
            var btnRect   = activeBtn.getBoundingClientRect();
            var centerX   = (btnRect.left + btnRect.width / 2) - stripRect.left;
            arrow.style.left = centerX + 'px';
        }
        positionRibbonTitle();
        positionRibbonCrumb();
    }
    positionArrowOnLoad();
    window.addEventListener('resize', positionArrowOnLoad);
})();

// Render the Identity Risk & Attack Surface table's default (consolidated)
// view once up front - the tab's tbody starts as a "Loading..." placeholder
// since the table is fully JS-rendered (so the tile-click filtering and the
// default view share one code path instead of two that could drift apart).
filterIdentityUsersTable(null, null);
filterGroupPrivTable(null, null);
</script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

</body>
</html>
"@

$html | Out-File -FilePath $ReportPath -Encoding UTF8

Write-Host ""
Write-Host "Dashboard generated successfully:" -ForegroundColor Green
Write-Host $ReportPath -ForegroundColor Yellow
Write-Host ""

Start-Process $ReportPath
