The dashboard was designed more as an Operations + Governance visibility layer rather than a pure health monitoring tool.

Built an “Enterprise Active Directory Operational Dashboard” using PowerShell, HTML, CSS, and native AD modules.

# Note: Updated version of Enterprise AD Health Dashboard is available in v2.ps1
# Copy Get-ADDashboard-v2.ps1 and Get-DCInventory.ps1 both the files in the same folder, run the Get-ADDashboard-v2.ps1 to get the repor
# The script is intended for lab/testing validation first. Review, test, and tune appropriately before using in production environments.

The goal was simple:
Provide a clean operational view of Active Directory infrastructure without requiring expensive monitoring platforms.
Key capabilities included:

• Forest & Domain configuration overview
• Forest Functional Level / Domain Functional Level
• UPN suffix visibility
• AD Recycle Bin and Tombstone Lifetime validation
• FSMO role holder visibility
• Domain Controller inventory with ICMP health check
• AD Sites / Sites with No DCs / Subnet mapping

• Privileged group monitoring:
   Domain Admins
   Enterprise Admins
   Schema Admins

• DNS operational visibility:
   Total DNS Zones
   AD Integrated Zones
   Standalone Zones

• GPO operational visibility:
   Total GPOs
   Unlinked GPOs
   Disabled GPOs

• Identity inventory:
   User Objects
   Computer Objects
   Security / Distribution Groups

• Trust relationship visibility across domains/forests

Features intentionally designed for operational usability:

• Hover-based contextual details without cluttering the dashboard
• Health indicators for operational warnings
• Export buttons on critical sections for CSV reporting
• Lightweight design with minimal performance impact
• Microsoft-inspired floating tile UI for better readability

Note: 
