#requires -modules ActiveDirectory

<#
.SYNOPSIS
    Agentless Domain Controller inventory collector for the Enterprise Active Directory
    Health Dashboard.

.DESCRIPTION
    Collects, per domain controller, using CIM/WinRM (Invoke-Command) only - no agent
    software is installed on any DC:

      - Hardware inventory (manufacturer, model, VM vs physical, CPU, memory)
      - OS build, uptime, pending reboot status
      - CPU and memory utilization
      - Disk volumes, dynamically discovered and mapped to their AD role
        (Operating System / NTDS Database / NTDS Logs / SYSVOL / Data) via the
        NTDS and Netlogon registry parameters - nothing is hardcoded to C:/D:.
      - AD-related service health (NTDS, Netlogon, KDC, DNS, W32Time, DFSR, ADWS)
      - Time synchronization status (w32tm)
      - Per-DC LDAPS certificate health (port 636 + certificate inspection)
      - Network adapter / IP configuration
      - Replication partner status (repadmin /showrepl)

    Requires WinRM (PS Remoting) enabled on target domain controllers, which is the
    default for Windows Server. Designed to be dot-sourced by the dashboard generator
    (Get-ADDashboard-v1.ps1) or run standalone to produce a JSON inventory file.

.PARAMETER ComputerName
    One or more DC host names to inventory. Defaults to every DC returned by
    Get-ADDomainController -Filter * for the current domain.

.PARAMETER Credential
    Optional credential used for the remote WinRM connections.

.PARAMETER OutputJsonPath
    If specified (and the script is invoked directly, not dot-sourced), writes the
    collected inventory array as JSON to this path.

.PARAMETER ThrottleLimit
    Maximum number of concurrent WinRM connections used by Invoke-Command.
    Defaults to 10. Scales to environments with tens or hundreds of DCs.

.EXAMPLE
    # Dot-sourced by the dashboard script - functions become available, nothing runs yet
    . .\Get-DCInventory.ps1

    $inventory = Get-AllDCInventory -ComputerName $DomainControllers.HostName -ThrottleLimit 15

.EXAMPLE
    # Run directly to produce a standalone JSON file
    .\Get-DCInventory.ps1 -OutputJsonPath C:\Temp\dc-inventory.json
#>

[CmdletBinding()]
param(
    [string[]]$ComputerName,
    [System.Management.Automation.PSCredential]$Credential,
    [string]$OutputJsonPath,
    [int]$ThrottleLimit = 10
)

Import-Module ActiveDirectory -ErrorAction SilentlyContinue

# ---------------------------------------------------------------------------
# Remote collection scriptblock - executed ON each domain controller via
# Invoke-Command. Returns a single ordered hashtable of everything that can
# be determined locally on the DC.
# ---------------------------------------------------------------------------

$InventoryScriptBlock = {

    $result = [ordered]@{}

    # -------------------------------------------------------------------
    # Hardware inventory
    # -------------------------------------------------------------------
    $cs    = Get-CimInstance -ClassName Win32_ComputerSystem
    $bios  = Get-CimInstance -ClassName Win32_BIOS
    $procs = @(Get-CimInstance -ClassName Win32_Processor)
    $os    = Get-CimInstance -ClassName Win32_OperatingSystem

    $manufacturer = [string]$cs.Manufacturer
    $model        = [string]$cs.Model

    $virtualPatterns = @(
        'VMware', 'Virtual Machine', 'Xen', 'KVM', 'QEMU',
        'Amazon EC2', 'Google Compute Engine', 'Hyper-V', 'Bochs'
    )

    $isVirtual = $false
    foreach ($pattern in $virtualPatterns) {
        if ($manufacturer -match $pattern -or $model -match $pattern -or $bios.Version -match $pattern) {
            $isVirtual = $true
            break
        }
    }

    $result.Hardware = [ordered]@{
        Manufacturer              = $manufacturer
        Model                     = $model
        MachineType               = if ($isVirtual) { "Virtual Machine" } else { "Physical Machine" }
        ProcessorName             = if ($procs.Count -gt 0) { $procs[0].Name.Trim() } else { "Unknown" }
        NumberOfProcessors        = [int]$cs.NumberOfProcessors
        NumberOfCores             = [int](($procs | Measure-Object -Property NumberOfCores -Sum).Sum)
        NumberOfLogicalProcessors = [int](($procs | Measure-Object -Property NumberOfLogicalProcessors -Sum).Sum)
        TotalMemoryGB             = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
        BIOSVersion               = $bios.SMBIOSBIOSVersion
        SerialNumber              = $bios.SerialNumber
    }

    # -------------------------------------------------------------------
    # OS build, uptime, pending reboot
    # -------------------------------------------------------------------
    $lastBoot = $os.LastBootUpTime
    $uptime   = (Get-Date) - $lastBoot

    $result.System = [ordered]@{
        Caption        = $os.Caption
        Version        = $os.Version
        BuildNumber    = $os.BuildNumber
        LastBootTime   = $lastBoot.ToString("dd-MMM-yyyy HH:mm:ss")
        UptimeDays     = [int][math]::Floor($uptime.TotalDays)
        UptimeReadable = "{0}d {1}h {2}m" -f [int]$uptime.Days, [int]$uptime.Hours, [int]$uptime.Minutes
    }

    $pendingReboot = $false
    $rebootReasons = @()

    if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending") {
        $pendingReboot = $true
        $rebootReasons += "Component Based Servicing"
    }

    if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired") {
        $pendingReboot = $true
        $rebootReasons += "Windows Update"
    }

    try {
        $pfro = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" `
            -Name "PendingFileRenameOperations" -ErrorAction Stop

        if ($pfro.PendingFileRenameOperations) {
            $pendingReboot = $true
            $rebootReasons += "PendingFileRenameOperations"
        }
    }
    catch {}

    try {
        if (Test-Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Update\PendingFileRenameOperations2") {
            $pendingReboot = $true
            $rebootReasons += "PendingFileRenameOperations2"
        }
    }
    catch {}

    $result.System.PendingReboot        = $pendingReboot
    $result.System.PendingRebootReasons = ($rebootReasons -join ", ")

    # -------------------------------------------------------------------
    # CPU and memory utilization
    # -------------------------------------------------------------------
    $cpuLoad  = ($procs | Measure-Object -Property LoadPercentage -Average).Average
    $totalMem = [double]$os.TotalVisibleMemorySize
    $freeMem  = [double]$os.FreePhysicalMemory

    $usedMemPercent = if ($totalMem -gt 0) {
        [math]::Round((($totalMem - $freeMem) / $totalMem) * 100, 1)
    } else { 0 }

    $result.Performance = [ordered]@{
        CPUUtilizationPercent    = [math]::Round([double]$cpuLoad, 1)
        MemoryTotalGB            = [math]::Round($totalMem / 1MB, 2)
        MemoryUsedGB             = [math]::Round(($totalMem - $freeMem) / 1MB, 2)
        MemoryFreeGB             = [math]::Round($freeMem / 1MB, 2)
        MemoryUtilizationPercent = $usedMemPercent
    }

    # -------------------------------------------------------------------
    # Dynamic disk discovery - NO hardcoded C:/D:.
    # Every fixed local volume is enumerated, then cross-referenced against
    # the NTDS and Netlogon registry parameters to determine its AD role.
    # -------------------------------------------------------------------
    $disks = @(Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3")

    $ntdsDbPath  = $null
    $ntdsLogPath = $null
    $sysvolPath  = $null

    try {
        $ntdsParams  = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\NTDS\Parameters" -ErrorAction Stop
        $ntdsDbPath  = $ntdsParams."DSA Database file"
        $ntdsLogPath = $ntdsParams."Database log files path"
    }
    catch {}

    try {
        $netlogonParams = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters" -ErrorAction Stop
        $sysvolPath     = $netlogonParams."SysVol"
    }
    catch {}

    function Get-VolumeLetter {
        param([string]$Path)

        if ([string]::IsNullOrWhiteSpace($Path)) { return $null }

        if ($Path -match '^([A-Za-z]):') {
            return ($matches[1].ToUpper() + ":")
        }

        return $null
    }

    $ntdsDbVol  = Get-VolumeLetter -Path $ntdsDbPath
    $ntdsLogVol = Get-VolumeLetter -Path $ntdsLogPath
    $sysvolVol  = Get-VolumeLetter -Path $sysvolPath
    $osVol      = Get-VolumeLetter -Path $env:SystemRoot

    $result.Disks = @(
        foreach ($disk in $disks) {
            $drive = $disk.DeviceID.ToUpper()
            $roles = @()

            if ($drive -eq $osVol)      { $roles += "Operating System" }
            if ($drive -eq $ntdsDbVol)  { $roles += "NTDS Database" }
            if ($drive -eq $ntdsLogVol) { $roles += "NTDS Logs" }
            if ($drive -eq $sysvolVol)  { $roles += "SYSVOL" }

            if ($roles.Count -eq 0) { $roles += "Data" }

            $sizeGB = [math]::Round($disk.Size / 1GB, 2)
            $freeGB = [math]::Round($disk.FreeSpace / 1GB, 2)

            $usedPercent = if ($disk.Size -gt 0) {
                [math]::Round((($disk.Size - $disk.FreeSpace) / $disk.Size) * 100, 1)
            } else { 0 }

            [ordered]@{
                Drive       = $drive
                Label       = $disk.VolumeName
                Role        = ($roles -join " + ")
                SizeGB      = $sizeGB
                FreeGB      = $freeGB
                UsedPercent = $usedPercent
            }
        }
    )

    # -------------------------------------------------------------------
    # AD-related service health
    # -------------------------------------------------------------------
    $serviceDefinitions = @(
        @{ Name = "NTDS";     Display = "Active Directory Domain Services" },
        @{ Name = "Netlogon"; Display = "Netlogon" },
        @{ Name = "Kdc";      Display = "Kerberos Key Distribution Center" },
        @{ Name = "DNS";      Display = "DNS Server" },
        @{ Name = "W32Time";  Display = "Windows Time" },
        @{ Name = "DFSR";     Display = "DFS Replication" },
        @{ Name = "ADWS";     Display = "Active Directory Web Services" }
    )

    $result.Services = @(
        foreach ($svcDef in $serviceDefinitions) {
            $svc = Get-Service -Name $svcDef.Name -ErrorAction SilentlyContinue

            [ordered]@{
                Name        = $svcDef.Name
                DisplayName = $svcDef.Display
                Status      = if ($svc) { $svc.Status.ToString() } else { "Not Found" }
                StartType   = if ($svc) { $svc.StartType.ToString() } else { "N/A" }
            }
        }
    )

    # -------------------------------------------------------------------
    # Time synchronization (w32tm)
    # -------------------------------------------------------------------
    try {
        $w32tmOutput = w32tm /query /status 2>$null

        $source   = ($w32tmOutput | Select-String "^Source:")                -replace "^Source:\s*", ""        | Select-Object -First 1
        $stratum  = ($w32tmOutput | Select-String "^Stratum:")               -replace "^Stratum:\s*", ""       | Select-Object -First 1
        $lastSync = ($w32tmOutput | Select-String "^Last Successful Sync")   -replace "^Last Successful Sync Time:\s*", "" | Select-Object -First 1
        $offset   = ($w32tmOutput | Select-String "^Phase Offset")           -replace "^Phase Offset:\s*", ""   | Select-Object -First 1

        $sourceText = "$source".Trim()

        $timeSyncStatus = if ([string]::IsNullOrWhiteSpace($sourceText)) {
            "Unknown"
        }
        elseif ($sourceText -match "Free-running System Clock|Local CMOS Clock") {
            "Warning"
        }
        else {
            "OK"
        }

        $result.TimeSync = [ordered]@{
            Source       = $sourceText
            Stratum      = "$stratum".Trim()
            LastSyncTime = "$lastSync".Trim()
            PhaseOffset  = "$offset".Trim()
            Status       = $timeSyncStatus
        }
    }
    catch {
        $result.TimeSync = [ordered]@{
            Source = "N/A"; Stratum = "N/A"; LastSyncTime = "N/A"; PhaseOffset = "N/A"; Status = "Unknown"
        }
    }

    # -------------------------------------------------------------------
    # Per-DC LDAPS certificate health (port 636)
    # -------------------------------------------------------------------
    try {
        $tcpTest      = Test-NetConnection -ComputerName "localhost" -Port 636 -WarningAction SilentlyContinue
        $ldapsPortOpen = [bool]$tcpTest.TcpTestSucceeded

        $ldaps = [ordered]@{
            PortOpen           = $ldapsPortOpen
            CertSubject        = "N/A"
            CertIssuer         = "N/A"
            CertExpiry         = "N/A"
            DaysToExpiry       = "N/A"
            Status             = "Unknown"
            CertNotBefore      = "N/A"
            SAN                = "N/A"
            SelfSigned         = $null
            SignatureAlgorithm = "N/A"
            KeySizeBits        = "N/A"
            Thumbprint         = "N/A"
        }

        if ($ldapsPortOpen) {
            $tcpClient = New-Object System.Net.Sockets.TcpClient
            $tcpClient.Connect("localhost", 636)

            $sslStream = New-Object System.Net.Security.SslStream(
                $tcpClient.GetStream(), $false, ({ $true })
            )
            $sslStream.AuthenticateAsClient($env:COMPUTERNAME)

            $cert  = $sslStream.RemoteCertificate
            $cert2 = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($cert)

            $daysToExpiry = [int]([math]::Floor((($cert2.NotAfter) - (Get-Date)).TotalDays))

            $ldaps.CertSubject   = $cert2.Subject
            $ldaps.CertIssuer    = $cert2.Issuer
            $ldaps.CertExpiry    = $cert2.NotAfter.ToString("dd-MMM-yyyy")
            $ldaps.DaysToExpiry  = $daysToExpiry
            $ldaps.CertNotBefore = $cert2.NotBefore.ToString("dd-MMM-yyyy")
            $ldaps.SelfSigned    = ($cert2.Subject -eq $cert2.Issuer)
            $ldaps.Thumbprint    = $cert2.Thumbprint

            try {
                $ldaps.SignatureAlgorithm = $cert2.SignatureAlgorithm.FriendlyName
            }
            catch {}

            try {
                $ldaps.KeySizeBits = $cert2.PublicKey.Key.KeySize
            }
            catch {
                try {
                    $ldaps.KeySizeBits = $cert2.GetPublicKey().Length * 8
                }
                catch {}
            }

            try {
                $sanExt = $cert2.Extensions | Where-Object { $_.Oid.FriendlyName -eq "Subject Alternative Name" }
                if ($sanExt) {
                    $ldaps.SAN = $sanExt.Format($false)
                }
            }
            catch {}

            $ldaps.Status = if ($daysToExpiry -lt 0) {
                "Critical"
            } elseif ($daysToExpiry -lt 30) {
                "Warning"
            } else {
                "OK"
            }

            $sslStream.Close()
            $tcpClient.Close()
        }
        else {
            $ldaps.Status = "Critical"
        }

        $result.LDAPS = $ldaps
    }
    catch {
        $result.LDAPS = [ordered]@{
            PortOpen = $false; CertSubject = "N/A"; CertIssuer = "N/A"; CertExpiry = "N/A"; DaysToExpiry = "N/A"; Status = "Unknown"
        }
    }

    # -------------------------------------------------------------------
    # Network adapters / IP configuration
    # -------------------------------------------------------------------
    $result.Network = @(
        Get-NetIPConfiguration -ErrorAction SilentlyContinue |
            Where-Object { $_.IPv4Address } |
            ForEach-Object {
                $dnsServers = @($_.DNSServer |
                    Where-Object { $_.AddressFamily -eq 2 } |
                    Select-Object -ExpandProperty ServerAddresses)

                [ordered]@{
                    AdapterName = $_.InterfaceAlias
                    IPAddress   = ($_.IPv4Address.IPAddress -join ", ")
                    DNSServers  = ($dnsServers -join ", ")
                }
            }
    )

    # -------------------------------------------------------------------
    # Ethernet settings - IPv4/IPv6 binding status + address, and
    # preferred/alternate DNS servers for the primary "Up" adapter.
    # -------------------------------------------------------------------
    $result.EthernetSettings = [ordered]@{
        AdapterName  = $null
        IPv4Enabled  = $null
        IPv4Address  = $null
        IPv6Enabled  = $null
        IPv6Address  = $null
        PreferredDNS = $null
        AlternateDNS = $null
    }

    try {
        $primaryAdapter = Get-NetAdapter -ErrorAction Stop | Where-Object { $_.Status -eq 'Up' } | Select-Object -First 1

        if ($primaryAdapter) {
            $result.EthernetSettings.AdapterName = $primaryAdapter.Name

            $ipv4Binding = Get-NetAdapterBinding -InterfaceAlias $primaryAdapter.Name -ComponentID 'ms_tcpip'  -ErrorAction SilentlyContinue
            $ipv6Binding = Get-NetAdapterBinding -InterfaceAlias $primaryAdapter.Name -ComponentID 'ms_tcpip6' -ErrorAction SilentlyContinue

            $result.EthernetSettings.IPv4Enabled = if ($ipv4Binding) { [bool]$ipv4Binding.Enabled } else { $null }
            $result.EthernetSettings.IPv6Enabled = if ($ipv6Binding) { [bool]$ipv6Binding.Enabled } else { $null }

            if ($result.EthernetSettings.IPv4Enabled) {
                $result.EthernetSettings.IPv4Address = Get-NetIPAddress -InterfaceIndex $primaryAdapter.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                    Select-Object -First 1 -ExpandProperty IPAddress
            }

            if ($result.EthernetSettings.IPv6Enabled) {
                $result.EthernetSettings.IPv6Address = Get-NetIPAddress -InterfaceIndex $primaryAdapter.ifIndex -AddressFamily IPv6 -ErrorAction SilentlyContinue |
                    Where-Object { $_.PrefixOrigin -ne 'WellKnown' } |
                    Select-Object -First 1 -ExpandProperty IPAddress
            }

            $dnsServers = @(
                Get-DnsClientServerAddress -InterfaceIndex $primaryAdapter.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                    Select-Object -ExpandProperty ServerAddresses
            )

            if ($dnsServers.Count -ge 1) { $result.EthernetSettings.PreferredDNS = $dnsServers[0] }
            if ($dnsServers.Count -ge 2) { $result.EthernetSettings.AlternateDNS = $dnsServers[1] }
        }
    }
    catch {}

    # -------------------------------------------------------------------
    # Security posture - SMB protocol security, TLS/Schannel protocol
    # configuration, LDAP signing/channel binding, NTLM hardening, audit
    # policy coverage, and recently installed hotfixes. All read-only
    # registry/cmdlet checks - nothing is changed on the DC.
    # -------------------------------------------------------------------
    $result.SecurityPosture = [ordered]@{}

    # --- SMB protocol security ---
    try {
        $smbConfig = Get-SmbServerConfiguration -ErrorAction Stop

        $result.SecurityPosture.SMB = [ordered]@{
            SMB1Enabled     = [bool]$smbConfig.EnableSMB1Protocol
            SigningEnabled  = [bool]$smbConfig.EnableSecuritySignature
            SigningRequired = [bool]$smbConfig.RequireSecuritySignature
        }
    }
    catch {
        $result.SecurityPosture.SMB = [ordered]@{
            SMB1Enabled = $null; SigningEnabled = $null; SigningRequired = $null
        }
    }

    # --- TLS / Schannel protocol configuration ---
    $schannelProtocols = @('SSL 2.0', 'SSL 3.0', 'TLS 1.0', 'TLS 1.1', 'TLS 1.2', 'TLS 1.3')

    $result.SecurityPosture.TLS = @(
        foreach ($proto in $schannelProtocols) {
            $regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\$proto\Server"

            $enabled           = $null
            $disabledByDefault = $null

            if (Test-Path $regPath) {
                $item = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
                if ($null -ne $item.Enabled)           { $enabled           = [int]$item.Enabled }
                if ($null -ne $item.DisabledByDefault) { $disabledByDefault = [int]$item.DisabledByDefault }
            }

            [ordered]@{
                Protocol          = $proto
                Enabled           = $enabled
                DisabledByDefault = $disabledByDefault
            }
        }
    )

    # --- LDAP signing and channel binding (NTDS\Parameters) ---
    try {
        $ntdsSecParams = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\NTDS\Parameters" -ErrorAction Stop

        $result.SecurityPosture.LDAPSigning = [ordered]@{
            LDAPServerIntegrity       = $ntdsSecParams.LDAPServerIntegrity
            LdapEnforceChannelBinding = $ntdsSecParams.LdapEnforceChannelBinding
        }
    }
    catch {
        $result.SecurityPosture.LDAPSigning = [ordered]@{
            LDAPServerIntegrity = $null; LdapEnforceChannelBinding = $null
        }
    }

    # --- NTLM hardening ---
    $lmCompatibilityLevel = $null
    $restrictSendingNTLM  = $null
    $auditReceivingNTLM   = $null
    $auditNTLMInDomain    = $null

    try {
        $lsaParams = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa" -ErrorAction Stop
        $lmCompatibilityLevel = $lsaParams.LmCompatibilityLevel
    }
    catch {}

    try {
        $msv10Params = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\MSV1_0" -ErrorAction Stop
        $restrictSendingNTLM = $msv10Params.RestrictSendingNTLMTraffic
        $auditReceivingNTLM  = $msv10Params.AuditReceivingNTLMTraffic
    }
    catch {}

    try {
        $netlogonSecParams = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters" -ErrorAction Stop
        $auditNTLMInDomain = $netlogonSecParams.AuditNTLMInDomain
    }
    catch {}

    $result.SecurityPosture.NTLM = [ordered]@{
        LmCompatibilityLevel = $lmCompatibilityLevel
        RestrictSendingNTLM  = $restrictSendingNTLM
        AuditReceivingNTLM   = $auditReceivingNTLM
        AuditNTLMInDomain    = $auditNTLMInDomain
    }

    # --- Audit policy coverage (auditpol) ---
    function Get-AuditSubcategoryStatus {
        param([string]$SubcategoryName)

        try {
            $output = auditpol /get /subcategory:"$SubcategoryName" 2>$null
            $line = $output | Where-Object { $_ -match [regex]::Escape($SubcategoryName) } | Select-Object -First 1

            if ($line) {
                if ($line -match "Success and Failure") { return "Success and Failure" }
                elseif ($line -match "Success")          { return "Success" }
                elseif ($line -match "Failure")          { return "Failure" }
                elseif ($line -match "No Auditing")      { return "No Auditing" }
            }

            return "Unknown"
        }
        catch {
            return "Unknown"
        }
    }

    $result.SecurityPosture.AuditPolicy = [ordered]@{
        DirectoryServiceChanges   = Get-AuditSubcategoryStatus -SubcategoryName "Directory Service Changes"
        CredentialValidation      = Get-AuditSubcategoryStatus -SubcategoryName "Credential Validation"
        AuthorizationPolicyChange = Get-AuditSubcategoryStatus -SubcategoryName "Authorization Policy Change"
    }

    # --- NetBIOS over TCP/IP (NetBT) ---
    # Reports the TcpipNetbiosOptions setting (0=Default via DHCP, 1=Enable,
    # 2=Disable) for every IP-enabled network adapter, read via
    # Win32_NetworkAdapterConfiguration - this maps directly to the
    # "NetBIOS setting" dropdown on the adapter's WINS tab. Disabling
    # NetBIOS blocks legacy broadcast-based name resolution (NBT-NS) on
    # this DC.
    try {
        $netbtOptionsMap = @{ 0 = 'Default (via DHCP)'; 1 = 'Enabled'; 2 = 'Disabled' }
        $netbtAdapters   = @(Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -ErrorAction Stop | Where-Object { $_.IPEnabled })

        if ($netbtAdapters.Count -eq 0) {
            throw "No IP-enabled network adapters found"
        }

        $result.SecurityPosture.NetBIOS = [ordered]@{
            Adapters = @(
                $netbtAdapters | ForEach-Object {
                    $opt = $_.TcpipNetbiosOptions

                    [ordered]@{
                        Adapter = $_.Description
                        Setting = if ($null -ne $opt) { $netbtOptionsMap[[int]$opt] } else { 'Unknown' }
                    }
                }
            )
        }
    }
    catch {
        $result.SecurityPosture.NetBIOS = [ordered]@{ Adapters = @() }
    }

    # --- Recently installed hotfixes (newest first, top 5) ---
    try {
        $hotfixes = @(
            Get-HotFix -ErrorAction Stop |
                Where-Object { $_.InstalledOn } |
                Sort-Object InstalledOn -Descending |
                Select-Object -First 5
        )

        $result.SecurityPosture.Hotfixes = @(
            foreach ($hf in $hotfixes) {
                [ordered]@{
                    HotFixID    = $hf.HotFixID
                    Description = $hf.Description
                    InstalledOn = $hf.InstalledOn.ToString("dd-MMM-yyyy")
                }
            }
        )
    }
    catch {
        try {
            $hotfixes = @(
                Get-CimInstance -ClassName Win32_QuickFixEngineering -ErrorAction Stop |
                    Where-Object { $_.InstalledOn } |
                    Sort-Object { try { [datetime]$_.InstalledOn } catch { [datetime]::MinValue } } -Descending |
                    Select-Object -First 5
            )

            $result.SecurityPosture.Hotfixes = @(
                foreach ($hf in $hotfixes) {
                    [ordered]@{
                        HotFixID    = $hf.HotFixID
                        Description = $hf.Description
                        InstalledOn = "$($hf.InstalledOn)"
                    }
                }
            )
        }
        catch {
            $result.SecurityPosture.Hotfixes = @()
        }
    }

    # -------------------------------------------------------------------
    # Deep Insights: DCDIAG summary, SYSVOL/DFSR replication health,
    # DNS health, recent interactive logons, and top event log errors.
    # -------------------------------------------------------------------
    $result.DeepInsights = [ordered]@{}

    # --- DCDIAG summary (Connectivity, Replications, Advertising, FSMOCheck,
    #     KnowsOfRoleHolders, NetLogons, Services, SysVolCheck) ---
    $dcDiagTests = @('Connectivity', 'Replications', 'Advertising', 'FSMOCheck', 'KnowsOfRoleHolders', 'NetLogons', 'Services', 'SysVolCheck')

    $result.DeepInsights.DCDiag = [ordered]@{}
    foreach ($t in $dcDiagTests) { $result.DeepInsights.DCDiag[$t] = "Unknown" }

    try {
        $dcDiagArgs = $dcDiagTests | ForEach-Object { "/test:$_" }
        $dcDiagOutput = & dcdiag @dcDiagArgs 2>$null

        foreach ($line in $dcDiagOutput) {
            if ($line -match 'passed test (\w+)') {
                $result.DeepInsights.DCDiag[$Matches[1]] = "Passed"
            }
            elseif ($line -match 'failed test (\w+)') {
                $result.DeepInsights.DCDiag[$Matches[1]] = "Failed"
            }
        }
    }
    catch {}

    # --- SYSVOL/DFSR replicated folder state ---
    # Reports the DFSR replication state for each replicated folder on this DC
    # (e.g. SYSVOL). State "Normal" = healthy/in-sync; other states indicate the
    # folder is initializing, recovering, or in error.
    try {
        $dfsrStateMap = @{ 0 = 'Uninitialized'; 1 = 'Initialized'; 2 = 'Initial Sync'; 3 = 'Auto Recovery'; 4 = 'Normal'; 5 = 'In Error' }

        $result.DeepInsights.DFSRBacklog = @(
            Get-WmiObject -Namespace "root\MicrosoftDFS" -Class "DfsrReplicatedFolderInfo" -ErrorAction Stop |
                ForEach-Object {
                    [ordered]@{
                        ReplicatedFolderName = $_.ReplicatedFolderName
                        ReplicationGroupName = $_.ReplicationGroupName
                        State                = $dfsrStateMap[[int]$_.State]
                    }
                }
        )
    }
    catch {
        $result.DeepInsights.DFSRBacklog = @()
    }

    # --- SYSVOL / NETLOGON share health ---
    # SYSVOL replication (DFSR) only reports on the "SYSVOL Share" replicated
    # folder above; NETLOGON is a subfolder served from the same replica set
    # but is not reported separately by DFSR, so its health is captured here
    # as a simple share-presence check.
    $result.DeepInsights.SharesHealth = [ordered]@{}

    foreach ($shareName in 'SYSVOL', 'NETLOGON') {
        try {
            $share = Get-SmbShare -Name $shareName -ErrorAction Stop

            $result.DeepInsights.SharesHealth[$shareName] = [ordered]@{
                Status = "Shared"
                Path   = $share.Path
            }
        }
        catch {
            $result.DeepInsights.SharesHealth[$shareName] = [ordered]@{
                Status = "Not Shared"
                Path   = $null
            }
        }
    }

    # --- DNS health: SRV record registration + self-resolvability ---
    $result.DeepInsights.DNSHealth = [ordered]@{
        SRVRecordRegistered = $null
        SRVRecordCount      = $null
        SelfResolvable      = $null
    }

    try {
        $domainFqdn = (Get-CimInstance -ClassName Win32_ComputerSystem).Domain
        $srvName    = "_ldap._tcp.dc._msdcs.$domainFqdn"
        $srvRecords = @(Resolve-DnsName -Name $srvName -Type SRV -ErrorAction Stop)

        $result.DeepInsights.DNSHealth.SRVRecordCount      = $srvRecords.Count
        $result.DeepInsights.DNSHealth.SRVRecordRegistered = [bool]($srvRecords | Where-Object { $_.NameTarget -match [regex]::Escape($env:COMPUTERNAME) })
    }
    catch {}

    try {
        $selfFqdn = [System.Net.Dns]::GetHostEntry($env:COMPUTERNAME).HostName
        $result.DeepInsights.DNSHealth.SelfResolvable = [bool]@(Resolve-DnsName -Name $selfFqdn -ErrorAction Stop).Count
    }
    catch {
        $result.DeepInsights.DNSHealth.SelfResolvable = $false
    }

    # --- Last 5 interactive / RDP / unlock / cached logons (Event ID 4624) ---
    $result.DeepInsights.RecentLogons = @()

    try {
        $logonTypeMap = @{ 2 = 'Interactive'; 7 = 'Unlock'; 10 = 'RemoteInteractive (RDP)'; 11 = 'CachedInteractive' }

        $logonEvents = Get-WinEvent -FilterHashtable @{ LogName = 'Security'; Id = 4624 } -MaxEvents 500 -ErrorAction Stop |
            Where-Object { $logonTypeMap.ContainsKey([int]$_.Properties[8].Value) } |
            Select-Object -First 5

        $result.DeepInsights.RecentLogons = @(
            $logonEvents | ForEach-Object {
                [ordered]@{
                    Time      = $_.TimeCreated.ToString("dd-MMM-yyyy HH:mm:ss")
                    Account   = "$($_.Properties[6].Value)\$($_.Properties[5].Value)"
                    LogonType = $logonTypeMap[[int]$_.Properties[8].Value]
                    SourceIP  = "$($_.Properties[18].Value)"
                }
            }
        )
    }
    catch {}

    # --- Top 5 recent error-level events from Directory Service, DFS Replication, DNS Server logs ---
    function Get-TopEventErrors {
        param(
            [string]$LogName,
            [int]$MaxEvents = 5
        )

        try {
            $events = Get-WinEvent -FilterHashtable @{ LogName = $LogName; Level = 2 } -MaxEvents $MaxEvents -ErrorAction Stop

            return @(
                $events | ForEach-Object {
                    $msg = "$($_.Message)" -split "`r?`n" | Select-Object -First 1
                    if ($msg.Length -gt 200) { $msg = $msg.Substring(0, 200) + "..." }

                    [ordered]@{
                        Time    = $_.TimeCreated.ToString("dd-MMM-yyyy HH:mm:ss")
                        EventID = $_.Id
                        Source  = $_.ProviderName
                        Message = $msg
                    }
                }
            )
        }
        catch {
            return @()
        }
    }

    $result.DeepInsights.EventLogs = [ordered]@{
        DirectoryService = Get-TopEventErrors -LogName "Directory Service"
        DFSReplication   = Get-TopEventErrors -LogName "DFS Replication"
        DNSServer        = Get-TopEventErrors -LogName "DNS Server"
    }

    return $result
}

# ---------------------------------------------------------------------------
# Replication partner status - queried from the machine running this script
# via repadmin against each target DC. Scoped to that DC's own inbound
# replication partners (not a full forest-wide replication summary).
# ---------------------------------------------------------------------------

function Get-ReplicationPartners {
    param(
        [Parameter(Mandatory)]
        [string]$DCName
    )

    try {
        $csvLines = repadmin /showrepl $DCName /csv 2>$null

        if (-not $csvLines -or $csvLines.Count -lt 2) {
            return @()
        }

        $rows = $csvLines | ConvertFrom-Csv

        return @(
            $rows | Where-Object { $_."Source DSA" } | ForEach-Object {
                [ordered]@{
                    Partner             = $_."Source DSA"
                    PartnerSite         = $_."Source DSA Site"
                    Partition           = $_."Naming Context"
                    LastSuccess         = $_."Last Success Time"
                    LastAttempt         = $_."Last Failure Time"
                    ConsecutiveFailures = $_."Number of Failures"
                    LastResult          = $_."Last Failure Status"
                }
            }
        )
    }
    catch {
        return @()
    }
}

# ---------------------------------------------------------------------------
# AD backup recency - queried from the machine running this script via
# repadmin /showbackup against each target DC.
# ---------------------------------------------------------------------------

function Get-ADBackupRecency {
    param(
        [Parameter(Mandatory)]
        [string]$DCName
    )

    $result = [ordered]@{
        LastBackupTime  = $null
        DaysSinceBackup = $null
        Status          = "Unknown"
    }

    try {
        $output = repadmin /showbackup $DCName 2>$null
        $text   = $output -join "`n"

        $dateMatch = [regex]::Match($text, '\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{2}:\d{2}\s*(AM|PM)?')

        if ($dateMatch.Success) {
            $parsed = [datetime]::Parse($dateMatch.Value)

            $result.LastBackupTime  = $parsed.ToString("dd-MMM-yyyy HH:mm")
            $result.DaysSinceBackup = [math]::Round(((Get-Date) - $parsed).TotalDays, 1)

            $result.Status = if ($result.DaysSinceBackup -gt 14) { "Critical" }
                              elseif ($result.DaysSinceBackup -gt 7) { "Warning" }
                              else { "Good" }
        }
        elseif ($text -match 'never been backed up|has never been backed up') {
            $result.LastBackupTime = "Never"
            $result.Status         = "Critical"
        }
    }
    catch {}

    return $result
}

# ---------------------------------------------------------------------------
# Main collection entry point.
# ---------------------------------------------------------------------------

function Get-AllDCInventory {
    [CmdletBinding()]
    param(
        [string[]]$ComputerName,
        [System.Management.Automation.PSCredential]$Credential,
        [int]$ThrottleLimit = 10
    )

    if (-not $ComputerName -or $ComputerName.Count -eq 0) {
        $ComputerName = @(Get-ADDomainController -Filter * | Select-Object -ExpandProperty HostName)
    }

    Write-Host "Collecting agentless inventory from $($ComputerName.Count) domain controller(s) (throttle: $ThrottleLimit)..." -ForegroundColor Cyan

    $icmParams = @{
        ComputerName = $ComputerName
        ScriptBlock  = $InventoryScriptBlock
        ThrottleLimit = $ThrottleLimit
        ErrorAction  = "SilentlyContinue"
        ErrorVariable = "icmErrors"
    }

    if ($Credential) {
        $icmParams.Credential = $Credential
    }

    $icmErrors = $null
    $remoteResults = @(Invoke-Command @icmParams)

    $inventory = @()

    foreach ($dc in $ComputerName) {
        $match = $remoteResults | Where-Object { $_.PSComputerName -eq $dc }

        $dcEntry = [ordered]@{
            ComputerName = $dc
            Reachable    = $false
            Error        = $null
        }

        if ($match) {
            foreach ($section in 'Hardware','System','Performance','Disks','Services','TimeSync','LDAPS','Network','EthernetSettings','SecurityPosture','DeepInsights') {
                $dcEntry[$section] = $match.$section
            }

            $dcEntry.Reachable = $true
        }
        else {
            $dcError = $icmErrors | Where-Object {
                $_.TargetObject -eq $dc -or "$($_.Exception.Message)" -match [regex]::Escape($dc)
            } | Select-Object -First 1

            $dcEntry.Error = if ($dcError) { $dcError.Exception.Message } else { "No response from $dc (WinRM unreachable or access denied)" }
        }

        $dcEntry.Replication = Get-ReplicationPartners -DCName $dc
        $dcEntry.ADBackup    = Get-ADBackupRecency -DCName $dc

        $inventory += [pscustomobject]$dcEntry
    }

    return $inventory
}

# ---------------------------------------------------------------------------
# Standalone execution - only runs when this script is invoked directly
# (e.g. ".\Get-DCInventory.ps1"), not when dot-sourced for its functions.
# ---------------------------------------------------------------------------

if ($MyInvocation.InvocationName -ne '.') {

    $inventory = Get-AllDCInventory -ComputerName $ComputerName -Credential $Credential -ThrottleLimit $ThrottleLimit

    if ($OutputJsonPath) {
        $inventory | ConvertTo-Json -Depth 8 | Out-File -FilePath $OutputJsonPath -Encoding UTF8
        Write-Host "DC inventory written to $OutputJsonPath" -ForegroundColor Green
    }

    $inventory
}
