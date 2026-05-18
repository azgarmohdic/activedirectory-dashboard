#requires -modules ActiveDirectory

Import-Module ActiveDirectory -ErrorAction Stop

$ReportPath = "C:\Temp\AD_Dashboard_Report.html"

if (!(Test-Path "C:\Temp")) {
    New-Item -Path "C:\Temp" -ItemType Directory | Out-Null
}

Write-Host "Collecting Active Directory information..." -ForegroundColor Cyan

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

    if ($null -eq $Value) { return "Not Configured" }

    if ($Value -is [TimeSpan]) {
        return "$([math]::Abs($Value.Days)) Days"
    }

    return $Value.ToString()
}

function Get-ADGroupMemberDetailsSafe {
    param(
        [Parameter(Mandatory)]
        [string]$Identity
    )

    try {
        $Members = Get-ADGroupMember -Identity $Identity -Recursive -ErrorAction Stop |
            Sort-Object Name |
            Select-Object Name, SamAccountName, objectClass

        if (!$Members) {
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
                "<div class='tooltip-line'><b>$Name</b> <span>($Sam - $Type)</span></div>"
            }
            else {
                "<div class='tooltip-line'><b>$Name</b> <span>($Type)</span></div>"
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
            Tooltip = "<div class='tooltip-line'>Unable to read group members</div>"
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

    $ExportData = Get-ADGroupMember -Identity $Identity -Recursive -ErrorAction SilentlyContinue |
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
        }

    return $ExportData | ConvertTo-Csv -NoTypeInformation
}

$Domain = Get-ADDomain
$Forest = Get-ADForest

$DomainName = $Domain.DNSRoot
$ForestFunctionalLevel = $Forest.ForestMode
$DomainFunctionalLevel = $Domain.DomainMode
$GeneratedOn = Get-Date -Format "dd-MMM-yyyy hh:mm:ss tt"

try {
    $RecycleBinFeature = Get-ADOptionalFeature -Identity "Recycle Bin Feature" -ErrorAction Stop
    if ($RecycleBinFeature.EnabledScopes.Count -gt 0) {
        $ADRecycleBinStatus = "Enabled"
    }
    else {
        $ADRecycleBinStatus = "Disabled"
    }
}
catch {
    $ADRecycleBinStatus = "N/A"
}

try {
    $ConfigNC = (Get-ADRootDSE).configurationNamingContext
    $DirectoryServicePath = "CN=Directory Service,CN=Windows NT,CN=Services,$ConfigNC"
    $DirectoryService = Get-ADObject -Identity $DirectoryServicePath -Properties tombstoneLifetime

    if ($DirectoryService.tombstoneLifetime) {
        $TombstoneLifetime = "$($DirectoryService.tombstoneLifetime) Days"
    }
    else {
        $TombstoneLifetime = "Default / Not Explicitly Set"
    }
}
catch {
    $TombstoneLifetime = "N/A"
}

$DomainControllers = Get-ADDomainController -Filter *
$DomainControllerCount = $DomainControllers.Count

$DomainControllersCsv = $DomainControllers |
    Select-Object `
        @{Name="DomainController";Expression={$_.HostName}},
        IPv4Address,
        @{Name="IsGC";Expression={$_.IsGlobalCatalog}},
        OperatingSystem,
        Site |
    ConvertTo-Csv -NoTypeInformation

$DomainControllersCsvJs = ConvertTo-JavaScriptString ($DomainControllersCsv -join "`n")

$FSMORoles = [ordered]@{
    "Schema Master"         = $Forest.SchemaMaster
    "Domain Naming Master"  = $Forest.DomainNamingMaster
    "PDC Emulator"          = $Domain.PDCEmulator
    "RID Master"            = $Domain.RIDMaster
    "Infrastructure Master" = $Domain.InfrastructureMaster
}

$FSMORoleHolders = $FSMORoles.Values | Select-Object -Unique
$FSMORoleHolderCount = $FSMORoleHolders.Count

$FSMOTooltip = ($FSMORoles.GetEnumerator() | ForEach-Object {
    $RoleName = ConvertTo-HtmlSafe $_.Key
    $RoleDC   = ConvertTo-HtmlSafe $_.Value
    "<div class='tooltip-line'><b>$RoleName</b><span>$RoleDC</span></div>"
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
    $PasswordPolicyTooltip = "<div class='tooltip-line'>Unable to read default domain password policy</div>"
}

try {
    Import-Module GroupPolicy -ErrorAction Stop
    $AllGPOs = Get-GPO -All
    $GPOCount = $AllGPOs.Count

    $DisabledGPOs = $AllGPOs | Where-Object {
        $_.GpoStatus -ne "AllSettingsEnabled"
    }

    $DisabledGPOCount = $DisabledGPOs.Count

    $UnlinkedGPOs = foreach ($GPO in $AllGPOs) {
        try {
            [xml]$GpoReport = Get-GPOReport -Guid $GPO.Id -ReportType Xml
            if ($GpoReport.GPO.LinksTo.Count -eq 0) {
                $GPO
            }
        }
        catch {}
    }

    $UnlinkedGPOCount = ($UnlinkedGPOs | Measure-Object).Count

    $GpoCsv = $AllGPOs |
        Select-Object DisplayName, Id, Owner, GpoStatus, CreationTime, ModificationTime |
        ConvertTo-Csv -NoTypeInformation

    $GpoCsvJs = ConvertTo-JavaScriptString ($GpoCsv -join "`n")
}
catch {
    $GPOCount = "N/A"
    $DisabledGPOCount = "N/A"
    $UnlinkedGPOCount = "N/A"
    $GpoCsvJs = ""
}

$TotalUsers = (Get-ADUser -LDAPFilter "(objectCategory=person)" -ResultSetSize $null).Count

$ActiveUsers = (Get-ADUser `
    -LDAPFilter "(&(objectCategory=person)(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))" `
    -ResultSetSize $null).Count

$DisabledUsers = (Get-ADUser `
    -LDAPFilter "(&(objectCategory=person)(objectClass=user)(userAccountControl:1.2.840.113556.1.4.803:=2))" `
    -ResultSetSize $null).Count

$TotalGroups = (Get-ADGroup -LDAPFilter "(objectClass=group)" -ResultSetSize $null).Count

$SecurityGroups = (Get-ADGroup `
    -LDAPFilter "(&(objectClass=group)(groupType:1.2.840.113556.1.4.803:=2147483648))" `
    -ResultSetSize $null).Count

$DistributionGroups = (Get-ADGroup `
    -LDAPFilter "(&(objectClass=group)(!(groupType:1.2.840.113556.1.4.803:=2147483648)))" `
    -ResultSetSize $null).Count

$TotalComputers = (Get-ADComputer -LDAPFilter "(objectCategory=computer)" -ResultSetSize $null).Count

$ActiveComputers = (Get-ADComputer `
    -LDAPFilter "(&(objectCategory=computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))" `
    -ResultSetSize $null).Count

$DisabledComputers = (Get-ADComputer `
    -LDAPFilter "(&(objectCategory=computer)(userAccountControl:1.2.840.113556.1.4.803:=2))" `
    -ResultSetSize $null).Count

$DomainAdmins     = Get-ADGroupMemberDetailsSafe -Identity "Domain Admins"
$EnterpriseAdmins = Get-ADGroupMemberDetailsSafe -Identity "Enterprise Admins"
$SchemaAdmins     = Get-ADGroupMemberDetailsSafe -Identity "Schema Admins"

$DomainAdminsCsv = Get-PrivilegedGroupCsv -Identity "Domain Admins" -DomainName $DomainName
$EnterpriseAdminsCsv = Get-PrivilegedGroupCsv -Identity "Enterprise Admins" -DomainName $DomainName
$SchemaAdminsCsv = Get-PrivilegedGroupCsv -Identity "Schema Admins" -DomainName $DomainName

$DomainAdminsCsvJs = ConvertTo-JavaScriptString ($DomainAdminsCsv -join "`n")
$EnterpriseAdminsCsvJs = ConvertTo-JavaScriptString ($EnterpriseAdminsCsv -join "`n")
$SchemaAdminsCsvJs = ConvertTo-JavaScriptString ($SchemaAdminsCsv -join "`n")

try {
    Import-Module DnsServer -ErrorAction Stop

    $DnsZones = Get-DnsServerZone -ErrorAction Stop

    $DnsZoneCount = $DnsZones.Count
    $ADIntegratedDnsZoneCount = ($DnsZones | Where-Object { $_.IsDsIntegrated -eq $true }).Count
    $StandaloneDnsZoneCount = ($DnsZones | Where-Object { $_.IsDsIntegrated -ne $true }).Count

    $DnsZonesCsv = $DnsZones |
        Select-Object ZoneName, ZoneType, IsDsIntegrated, ReplicationScope, DynamicUpdate |
        ConvertTo-Csv -NoTypeInformation

    $DnsZonesCsvJs = ConvertTo-JavaScriptString ($DnsZonesCsv -join "`n")
}
catch {
    $DnsZones = $null
    $DnsZoneCount = "N/A"
    $ADIntegratedDnsZoneCount = "N/A"
    $StandaloneDnsZoneCount = "N/A"
    $DnsZonesCsvJs = ""
}

try {
    $Trusts = Get-ADTrust -Filter * -ErrorAction Stop
}
catch {
    $Trusts = $null
}

$html = @"
<!DOCTYPE html>
<html>
<head>
<title>Active Directory Dashboard</title>

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
        margin: 0 auto 14px auto;
        background: linear-gradient(90deg, #dbeafe, #eff6ff);
        padding: 20px;
        text-align: center;
        border-radius: 14px;
        border: 1px solid #bfdbfe;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.08);
    }

    .header-label {
        font-size: 20px;
        font-weight: 500;
        color: #475569;
        margin-right: 6px;
    }

    .header-domain {
        font-size: 34px;
        font-weight: 700;
        color: #0f172a;
        letter-spacing: 0.3px;
    }

    .functional-bar {
        max-width: 950px;
        margin: 0 auto 28px auto;
        display: grid;
        grid-template-columns: repeat(2, 1fr);
        gap: 16px;
    }

    .functional-card {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        border-left: 5px solid #2563eb;
        border-radius: 12px;
        padding: 14px 18px;
        box-shadow: 0 6px 16px rgba(15, 23, 42, 0.06);
    }

    .functional-title {
        font-size: 13px;
        color: #64748b;
        margin-bottom: 6px;
    }

    .functional-value {
        font-size: 18px;
        color: #111827;
        font-weight: 700;
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

    .tile.has-tooltip:hover .tooltip-box {
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
        .functional-bar {
            grid-template-columns: 1fr;
        }

        table,
        .page-generated {
            width: 100%;
        }

        .header-domain {
            font-size: 26px;
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
        <span class="header-label">Domain:</span>
        <span class="header-domain">$DomainName</span>
    </div>

    <div class="functional-bar">
        <div class="functional-card">
            <div class="functional-title">Forest Functional Level</div>
            <div class="functional-value">$ForestFunctionalLevel</div>
        </div>

        <div class="functional-card">
            <div class="functional-title">Domain Functional Level</div>
            <div class="functional-value">$DomainFunctionalLevel</div>
        </div>
    </div>

    <div class="functional-bar">
        <div class="functional-card">
            <div class="functional-title">AD Recycle Bin Status</div>
            <div class="functional-value">$ADRecycleBinStatus</div>
        </div>

        <div class="functional-card">
            <div class="functional-title">Tombstone Lifetime</div>
            <div class="functional-value">$TombstoneLifetime</div>
        </div>
    </div>

    <div class="dashboard">

        <div class="tile accent-blue">
            <div class="tile-top">
                <div class="tile-title">Domain Controllers</div>
                <div class="tile-icon">DC</div>
            </div>
            <div>
                <div class="tile-value">$DomainControllerCount</div>
                <div class="tile-footer">Total domain controllers</div>
            </div>
        </div>

        <div class="tile accent-cyan has-tooltip">
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

        <div class="tile accent-green has-tooltip">
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

        <div class="tile accent-red has-tooltip">
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

        <div class="tile accent-purple has-tooltip">
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

        <div class="tile accent-orange has-tooltip">
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

        <div class="tile accent-blue">
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

        <div class="tile accent-orange">
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
                    <button class="export-btn" title="Export GPOs" onclick="downloadCsv('GPO_Inventory.csv', gpoCsv)">⇩</button>
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
                <div class="tile-icon">UL</div>
            </div>
            <div>
                <div class="tile-value">$UnlinkedGPOCount</div>
                <div class="tile-footer">No Site/Domain/OU links</div>
            </div>
        </div>

        <div class="tile accent-red">
            <div class="tile-top">
                <div class="tile-title">Disabled GPOs</div>
                <div class="tile-icon">DG</div>
            </div>
            <div>
                <div class="tile-value">$DisabledGPOCount</div>
                <div class="tile-footer">User/Computer settings disabled</div>
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

        <div class="tile accent-red">
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

        <div class="tile accent-green">
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

        <div class="tile accent-blue">
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
        Active Directory Dashboard Report | Generated using PowerShell | by Azgar Mohammad
    </div>

</div>

<script>
    function downloadCsv(fileName, csvContent) {
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement("a");

        if (link.download !== undefined) {
            const url = URL.createObjectURL(blob);
            link.setAttribute("href", url);
            link.setAttribute("download", fileName);
            link.style.visibility = "hidden";
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(url);
        }
    }

    const domainAdminsCsv = '$DomainAdminsCsvJs';
    const enterpriseAdminsCsv = '$EnterpriseAdminsCsvJs';
    const schemaAdminsCsv = '$SchemaAdminsCsvJs';
    const domainControllersCsv = '$DomainControllersCsvJs';
    const dnsZonesCsv = '$DnsZonesCsvJs';
    const gpoCsv = '$GpoCsvJs';
</script>

</body>
</html>
"@

$html | Out-File -FilePath $ReportPath -Encoding UTF8

Write-Host "Report generated successfully: $ReportPath" -ForegroundColor Green

Start-Process $ReportPath