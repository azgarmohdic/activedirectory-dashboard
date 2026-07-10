<#
.SYNOPSIS
    Creates new Active Directory subnet objects and associates them to existing sites,
    driven by a CSV input file. Additive-only: never modifies or removes any existing
    site or subnet object.

.DESCRIPTION
    Enterprise-safe automation for AD Sites and Services subnet provisioning.

    Design principles (read before running in production):
      1. ADDITIVE ONLY. The script only calls New-ADReplicationSubnet for CIDR ranges
         that do not already exist in the Configuration partition. It never calls
         Set-ADReplicationSubnet or Remove-ADReplicationSubnet, and it never creates,
         renames, or modifies an AD Site. This satisfies the requirement that existing
         sites and subnets must not be disturbed.
      2. SITES MUST PRE-EXIST. The script does not create sites. If a CSV row references
         a SiteName that does not exist in AD Sites and Services, that row is skipped
         and logged as an error — it is not created, and no assumption is made about
         which site was "meant."
      3. IDEMPOTENT / RE-RUNNABLE. Running the script twice against the same CSV
         produces the same end state: the second run finds all subnets already present
         and skips them. Safe to schedule or re-run after a partial failure.
      4. DRY-RUN FIRST. The script implements ShouldProcess (-WhatIf / -Confirm). Always
         run with -WhatIf first in production and review the log before committing.
      5. NO IMPLICIT CREDENTIALS. Runs in the caller's security context by default;
         optionally accepts -Credential for a delegated account. Requires rights to
         create child objects under CN=Subnets,CN=Sites,CN=Configuration,DC=<forest>
         (delegated Enterprise Admins-equivalent permission, or explicit delegation).

.PARAMETER CsvPath
    Path to the input CSV. Required columns (header, case-insensitive):
        SubnetCIDR   - e.g. 10.20.30.0/24 or 2001:db8:abcd::/48
        SiteName     - the exact CN of an existing AD Site (as shown in Sites and Services)
    Optional columns:
        Description  - free-text description to set on the new subnet object
        Location     - value for the subnet's Location property

.PARAMETER LogPath
    Folder to write the run log and result CSV to. Defaults to .\Logs next to the script.

.PARAMETER Server
    Optional specific domain controller (or Configuration NC-writable DC) to target.
    Defaults to the domain controller AD PowerShell auto-selects.

.PARAMETER Credential
    Optional PSCredential for a delegated service/admin account.

.EXAMPLE
    .\New-ADSiteSubnets.ps1 -CsvPath .\subnets.csv -WhatIf
    Dry run: shows exactly what would be created, changes nothing.

.EXAMPLE
    .\New-ADSiteSubnets.ps1 -CsvPath .\subnets.csv -Server dc01.corp.contoso.com
    Executes for real against a specific DC, writing a log and result CSV.

.NOTES
    Author        : Enterprise AD Architecture (script generated per engagement request)
    Requires      : RSAT ActiveDirectory PowerShell module, network/LDAP reachability
                    to a writable DC for the Configuration partition, and sufficient
                    delegated rights over CN=Subnets,CN=Sites,CN=Configuration.
    Idempotent    : Yes — safe to re-run.
    Destructive   : No — creates only; never edits or deletes existing objects.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [string]$CsvPath,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = (Join-Path -Path $PSScriptRoot -ChildPath 'Logs'),

    [Parameter(Mandatory = $false)]
    [string]$Server,

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]$Credential
)

#region Setup ------------------------------------------------------------------------

$ErrorActionPreference = 'Stop'
$scriptStart = Get-Date
$timestamp   = $scriptStart.ToString('yyyyMMdd_HHmmss')

if (-not (Test-Path -Path $LogPath)) {
    New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
}

$logFile    = Join-Path $LogPath "New-ADSiteSubnets_$timestamp.log"
$resultFile = Join-Path $LogPath "New-ADSiteSubnets_Results_$timestamp.csv"

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')][string]$Level = 'INFO'
    )
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Add-Content -Path $logFile -Value $line
    switch ($Level) {
        'WARN'    { Write-Warning $Message }
        'ERROR'   { Write-Host $line -ForegroundColor Red }
        'SUCCESS' { Write-Host $line -ForegroundColor Green }
        default   { Write-Host $line }
    }
}

Write-Log "===== New-ADSiteSubnets run started ====="
Write-Log "CSV input      : $CsvPath"
Write-Log "Log folder     : $LogPath"
Write-Log "Target server  : $(if ($Server) { $Server } else { '(module default DC)' })"
Write-Log "WhatIf mode    : $($WhatIfPreference.IsPresent -or $PSBoundParameters.ContainsKey('WhatIf'))"

#endregion

#region Pre-flight checks -------------------------------------------------------------

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Log "ActiveDirectory module loaded successfully." -Level SUCCESS
}
catch {
    Write-Log "FATAL: ActiveDirectory PowerShell module is not available. Install RSAT AD DS tools and re-run. $_" -Level ERROR
    throw
}

$adParams = @{}
if ($Server)     { $adParams['Server']     = $Server }
if ($Credential) { $adParams['Credential'] = $Credential }

try {
    $null = Get-ADReplicationSite -Filter * @adParams -ErrorAction Stop
    Write-Log "Connectivity to Configuration partition confirmed." -Level SUCCESS
}
catch {
    Write-Log "FATAL: Unable to query AD Sites (Configuration partition). Check connectivity, permissions, and -Server value. $_" -Level ERROR
    throw
}

#endregion

#region Load and validate CSV ---------------------------------------------------------

$rows = Import-Csv -Path $CsvPath

$requiredColumns = @('SubnetCIDR', 'SiteName')
$csvColumns = ($rows | Get-Member -MemberType NoteProperty).Name
foreach ($col in $requiredColumns) {
    if ($col -notin $csvColumns) {
        Write-Log "FATAL: Required column '$col' not found in CSV header. Found columns: $($csvColumns -join ', ')" -Level ERROR
        throw "CSV missing required column: $col"
    }
}

Write-Log "Loaded $($rows.Count) row(s) from CSV."

# Basic CIDR validation (IPv4 and IPv6), applied before touching AD at all
$cidrPattern = '^((\d{1,3}\.){3}\d{1,3}/\d{1,2}|[0-9A-Fa-f:]+/\d{1,3})$'

$validatedRows = foreach ($row in $rows) {
    $cidr = $row.SubnetCIDR.Trim()
    $site = $row.SiteName.Trim()

    if ([string]::IsNullOrWhiteSpace($cidr) -or [string]::IsNullOrWhiteSpace($site)) {
        Write-Log "SKIP (validation): Row has blank SubnetCIDR or SiteName. Raw: '$($row.SubnetCIDR)' / '$($row.SiteName)'" -Level WARN
        continue
    }
    if ($cidr -notmatch $cidrPattern) {
        Write-Log "SKIP (validation): '$cidr' does not look like a valid CIDR (e.g. 10.0.0.0/24). Site: $site" -Level WARN
        continue
    }

    [PSCustomObject]@{
        SubnetCIDR  = $cidr
        SiteName    = $site
        Description = if ($row.PSObject.Properties.Name -contains 'Description') { $row.Description } else { $null }
        Location    = if ($row.PSObject.Properties.Name -contains 'Location')    { $row.Location }    else { $null }
    }
}

Write-Log "$($validatedRows.Count) row(s) passed format validation and will be evaluated against AD."

#endregion

#region Cache existing AD state (read-only) -------------------------------------------

Write-Log "Caching existing sites and subnets for comparison (read-only calls)..."

try {
    $existingSites = Get-ADReplicationSite -Filter * @adParams | Select-Object -ExpandProperty Name
    Write-Log "Existing sites in AD: $($existingSites.Count) found."
}
catch {
    Write-Log "FATAL: Failed to enumerate existing AD Sites. $_" -Level ERROR
    throw
}

try {
    $existingSubnets = Get-ADReplicationSubnet -Filter * @adParams | Select-Object -ExpandProperty Name
    Write-Log "Existing subnets in AD: $($existingSubnets.Count) found."
}
catch {
    Write-Log "FATAL: Failed to enumerate existing AD Subnets. $_" -Level ERROR
    throw
}

#endregion

#region Process each row (additive only) ----------------------------------------------

$results = New-Object System.Collections.Generic.List[Object]

foreach ($item in $validatedRows) {

    $resultRow = [PSCustomObject]@{
        SubnetCIDR = $item.SubnetCIDR
        SiteName   = $item.SiteName
        Action     = ''
        Detail     = ''
        Timestamp  = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    }

    # Rule: site must already exist. Never create sites.
    if ($existingSites -notcontains $item.SiteName) {
        $resultRow.Action = 'SKIPPED-SiteMissing'
        $resultRow.Detail = "Site '$($item.SiteName)' does not exist in AD Sites and Services. Subnet not created."
        Write-Log $resultRow.Detail -Level WARN
        $results.Add($resultRow)
        continue
    }

    # Rule: never touch a subnet that already exists.
    if ($existingSubnets -contains $item.SubnetCIDR) {
        $resultRow.Action = 'SKIPPED-AlreadyExists'
        $resultRow.Detail = "Subnet '$($item.SubnetCIDR)' already exists in AD. Left untouched (no modification made)."
        Write-Log $resultRow.Detail -Level INFO

        # Informational only — flag mismatch without changing anything
        try {
            $existingObj = Get-ADReplicationSubnet -Identity $item.SubnetCIDR @adParams -Properties Site
            $currentSite = ($existingObj.Site -split ',')[0] -replace '^CN=', ''
            if ($currentSite -and $currentSite -ne $item.SiteName) {
                $mismatchMsg = "NOTE: Existing subnet '$($item.SubnetCIDR)' is currently associated to site '$currentSite', CSV specifies '$($item.SiteName)'. No change made (script is additive-only)."
                Write-Log $mismatchMsg -Level WARN
                $resultRow.Detail += " $mismatchMsg"
            }
        } catch {
            Write-Log "Could not read current site association for existing subnet '$($item.SubnetCIDR)' (non-fatal): $_" -Level WARN
        }

        $results.Add($resultRow)
        continue
    }

    # New subnet — create it
    $newParams = @{
        Name        = $item.SubnetCIDR
        Site        = $item.SiteName
    }
    if ($item.Description) { $newParams['Description'] = $item.Description }
    if ($item.Location)    { $newParams['Location']    = $item.Location }
    foreach ($k in $adParams.Keys) { $newParams[$k] = $adParams[$k] }

    $target = "AD Subnet '$($item.SubnetCIDR)' -> Site '$($item.SiteName)'"

    if ($PSCmdlet.ShouldProcess($target, 'New-ADReplicationSubnet')) {
        try {
            New-ADReplicationSubnet @newParams -ErrorAction Stop
            $resultRow.Action = 'CREATED'
            $resultRow.Detail = "Subnet created and associated to site '$($item.SiteName)'."
            Write-Log "SUCCESS: $target created." -Level SUCCESS
        }
        catch {
            $resultRow.Action = 'FAILED'
            $resultRow.Detail = "Error creating subnet: $_"
            Write-Log "ERROR creating $target : $_" -Level ERROR
        }
    }
    else {
        $resultRow.Action = 'WHATIF-WouldCreate'
        $resultRow.Detail = "WhatIf: would create subnet and associate to site '$($item.SiteName)'."
        Write-Log "WHATIF: would create $target" -Level INFO
    }

    $results.Add($resultRow)
}

#endregion

#region Summary and output -------------------------------------------------------------

$results | Export-Csv -Path $resultFile -NoTypeInformation -Encoding UTF8

$summary = $results | Group-Object Action | Select-Object Name, Count
Write-Log "----- Run Summary -----"
foreach ($s in $summary) {
    Write-Log ("{0,-25} : {1}" -f $s.Name, $s.Count)
}
Write-Log "Result detail written to: $resultFile"
Write-Log "Full log written to     : $logFile"
Write-Log "===== New-ADSiteSubnets run completed in $([math]::Round(((Get-Date) - $scriptStart).TotalSeconds,1))s ====="

Write-Host "`nSummary:" -ForegroundColor Cyan
$summary | Format-Table -AutoSize
Write-Host "Detailed results: $resultFile"
Write-Host "Full log        : $logFile"

#endregion
