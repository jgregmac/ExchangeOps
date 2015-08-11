<#
.SYNOPSIS
Displays the current maintenance status of an Exchange 2013 DAG member.

.PARAMETER Server
Specifies the DAG member server 

.EXAMPLE
PS> .\Get-DAGServerMaintenanceMode.ps1
#>

#Requires -version 3.0

[CmdletBinding()]
Param(
    [Parameter(Position=0, Mandatory = $true,
    HelpMessage="Enter the name of DAG Server to check for Maintenance mode.")]
    [string]$Server
)

#--------------------------------------------------------------------------------
function Show-MaintenanceMode {

    write-host "Component State on $Server" -ForegroundColor Cyan
    Get-ServerComponentState $Server | 
        Where {$_.Component -ne "Monitoring" -and $_.Component -ne "RecoveryActionsEnabled"} |
        format-table Component,State -Autosize

    write-host "Database Copy settings on $Server" -ForegroundColor Cyan
    Get-MailboxServer $Server | format-list DatabaseCopy*

    write-host "Mail queues on $Server" -ForegroundColor Cyan
    Get-Queue -Server $Server | format-table -auto

    write-host "DAG/Cluster status for $Server" -ForegroundColor Cyan
    $get_node = [ScriptBlock]::Create("Get-ClusterNode $Server")
    invoke-command -ComputerName $server -ScriptBlock $get_node

    write-host "Database Copy Status on $Server" -ForegroundColor Cyan
    Get-MailboxDatabaseCopyStatus -Server $Server
}

Show-MaintenanceMode



