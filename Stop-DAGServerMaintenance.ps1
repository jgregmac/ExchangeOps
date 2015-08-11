<#
.SYNOPSIS
Put an Exchange 2013 DAG Server back into service after maintenance

.DESCRIPTION
Prepares the specified Exchange 2013 DAG-member server for active service by
 o Unset ServerWideOffline
 o Resuming the cluster node
 o Allowing database copies to activate locally
 o Re-enable transport queues

 [per https://technet.microsoft.com/en-us/library/dd298065(v=exchg.150).aspx#Pm ]

Please note: this script does not verify that the DAG member server is in a healthy state and otherwise prepared to be into service.

.PARAMETER Server
The name of the DAG Member to place in maintenance mode.

.EXAMPLE
PS> .\Stop-DAGServerMaintenance.ps1 -server MSX01

Readies the server MSX01 for maintenance, transferring messages in delivery queues to a random
server in the same site as MSX01.

.NOTES 

Based on Stop2013DagServerMaintenance.ps1 from
    https://gallery.technet.microsoft.com/office/Exchange-2013-DAG-3ac89826

Modified by: Geoffrey.Duke@uvm.edu
#>

#Requires -version 3.0

[CmdletBinding()]
Param(
    [Parameter(Position=0, Mandatory = $true,
    HelpMessage="Enter the name of the DAG Server to remove from Maintenance mode.")]
    [string]$Server
)

#--------------------------------------------------------------------------------
function Main {
    
    Test-Parameters

    write-host "Re-enabling server $Server for active service." -ForegroundColor Green

    #Designates that the server is out of maintenance mode
    Write-Verbose "Taking $Server out of maintenance mode"
    Set-ServerComponentState $Server -Component ServerWideOffline -State Active -Requester Maintenance

    #Allows the server to accept Unified Messaging calls
    Write-Verbose "$Server can now accept Unified Messaging calls."
    Set-ServerComponentState $Server -Component UMCallRouter -State Active -Requester Maintenance

    #Resumes the node in the cluster and enables full cluster functionality for the server
    Write-Verbose "Resuming the cluster node and enabling full cluster functionality."
    try {
        $resume_node = [scriptblock]::Create("Resume-ClusterNode $Server | out-null")
        invoke-command -Computer $Server -ScriptBlock $resume_node -EV SuspErr
    }
    catch {
        throw "Encountered error trying to Suspend-ClusterNode $Server.`n$($SuspErr[0])"
    }

    #Allow databases to become active on the server
    Write-Verbose "$Server can now host active database copies."
    Set-MailboxServer $Server -DatabaseCopyActivationDisabledAndMoveNow $False

    #Remove the automatic activation blocks
    Write-Verbose "$Server can now automatically host active database copies."
    Set-MailboxServer $Server -DatabaseCopyAutoActivationPolicy Unrestricted

    #Resumes the transport queues and allows the server to accept and process messages
    Write-Verbose "Transport Queues on $Server are now active."
    Set-ServerComponentState $Server -Component HubTransport -State Active -Requester Maintenance

    Write-Verbose 'Restarting the Transport Services to resume transport activity'
    invoke-command -ComputerName $server -ScriptBlock { Restart-Service MSExchangeTransport } 

    Write-Host "$Server should now be out of maintenance mode and configured for active service." -ForegroundColor Green

    Show-MaintenanceMode

    Write-Host "You should redistribute mailbox databases in the DAG."  -ForegroundColor Green
    Write-host "This command pipeline below will activate on $Server all the databases for which it is set as the first activation preference."

    $remount_dbs = "Get-MailboxDatabaseCopyStatus -server $Server | " +
                        'where ActivationPreference -eq 1 | foreach { ' +
                        'Move-ActiveMailboxDatabase $_.DatabaseName -ActivateOnServer $Server }'
    write-host $remount_dbs -ForegroundColor White

} #end function:Main

#--------------------------------------------------------------------------------
function Test-Parameters {
    try {
        $Script:maint_server = Get-ExchangeServer $Server -ErrorAction Stop
    }
    catch {
        throw "Unable to retreive MailboxServer object for $Server"
    }
}

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
}

# Run script
Main
