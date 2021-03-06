<#
.SYNOPSIS
Prepare Exchange 2013 DAG Server for maintenance

.DESCRIPTION
Prepares the specified Exchange 2013 DAG-member server for maintenance by
 o Draining transport queues
 o Redirecting messages in local delivery queues
 o Pausing the cluster node
 o Moving active mailbox database copies to other servers
 o Prevent Database copies from activating locally
 o Set ServerWideOffline

Please note: this script does not verify that the DAG is in a healthy state and otherwise prepared for a server to be taken out of service for maintenance.

.PARAMETER Server
The name of the DAG Member to place in maintenance mode.

.PARAMETER QueueTargetServer
The name of an Exchange 2013 mailbox server to which mail in local delivery queues will be transferred. If one is not specified, the script will attempt to use a random server in the same AD site as the source server.

.EXAMPLE
PS> .\Start-DAGServerMaintenance.ps1 -server MSX01

Readies the server MSX01 for maintenance, transferring messages in delivery queues to a random
server in the same site as MSX-MH01.

.EXAMPLE
PS> .\Start-DAGServerMaintenance.ps1 -Server MSX01 -QueueTarget MSX06

Readies the server MSX-MH01 for maintenance, transferring messages in delivery queues to MSX06.

.LINK
Performing maintenance on DAG members
  https://technet.microsoft.com/en-us/library/dd298065(v=exchg.150).aspx#Pm

.LINK
Script based on Exchange2013DAGMaintenanceScripts by JBeeden
  https://gallery.technet.microsoft.com/office/Exchange-2013-DAG-3ac89826

#>

#Requires -version 3.0

[CmdletBinding()]
Param(
    [Parameter(Position=0, Mandatory = $true,
    HelpMessage="Enter the name of the DAG Server to put into Maintenance mode.")]
    [string]$Server,

    [Parameter(Position=1, Mandatory = $false,
    HelpMessage="Enter FQDN of server to move mail queue to.")]
    [string]$QueueTarget
)
Set-PSDebug -Strict

#Insert ISE and EMS environment tests here.
if ((Get-PSSnapin | ? -property Name -match 'microsoft\.exchange').count -lt 4) {
    Add-PSSnapin microsoft.exchange*
}

#--------------------------------------------------------------------------------
function Main {
    #write-verbose 'Setting up implicit remoting to ensure access to Exchange Management Shell functions'
    #$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$env:computername.campus.ad.uvm.edu/PowerShell"
    #Import-PSSession $ExchangeSession

    write-verbose 'Verifying parameters, deriving other settings'
    Test-Parameters

    Write-Verbose 'Checking DAG File Share Witness'
    try {
        if ( Get-DatabaseAvailabilityGroup $dag_name -Status -EA Stop | Where WitnessShareInUse -eq 'InvalidConfiguration' ) {
            throw "File Share Witness for $dag_name shows 'Invalid Configuration'. Starting maintenance in this state could cause problems with quorum."
        }
    }
    catch {
        Return "Error executing File Share Witness Check. Error: $_ "
    }

    Write-Verbose 'Beginning the process of draining the transport queues'
    try {
        Set-ServerComponentState $Server -Component HubTransport -State Draining -Requester Maintenance -EA Stop 
    }
    catch {
        return "Error setting HubTransport component on $Server to 'Draining'. Error: $_"
    }

    Write-Verbose 'Restarting the Transport Services to initiate draining'
    try {
        invoke-command -ComputerName $server -ScriptBlock { Restart-Service MSExchangeTransport } -EA Stop 
    }
    catch {
        return "Error restarting MSExchangeTransport service on $Server. Error: $_"
    }

    Write-verbose 'Beginning the process of draining all Unified Messaging calls'
    try {
        Set-ServerComponentState $Server -Component UMCallRouter -State Draining -Requester Maintenance -EA Stop 
    }
    catch {
        return "Error setting UMCallRouter component on $Server to 'Draining'. Error: $_"
    }

    Write-Verbose "Redirecting messages pending delivery in the local queues to $queue_server"
    try {
        Redirect-Message -Server $Server -Target $queue_server.fqdn -Confirm:$false -EA Stop 
    }
    catch {
        return "Error redirecting message gueues to $( $queue_server.fqdn ). Error: $_"
    }

    Write-Verbose 'Pausing the cluster node; prevents it from being/becoming the PrimaryActiveManager'
    $suspend_node = [scriptblock]::Create("Suspend-ClusterNode $Server | out-null")
    try {
        invoke-command -Computer $Server -ScriptBlock $suspend_node -EA Stop 
    }
    catch {
        return "Encountered error trying to Suspend-ClusterNode $Server. Error: $_"
    }

    Write-Verbose "Moving all active databases currently hosted on $Server to other DAG members"
    try {
       Set-MailboxServer $Server -DatabaseCopyActivationDisabledAndMoveNow $true -EA Stop 
    }
    catch {
        return "Error setting 'DatabaseCopyActivationDisabledAndMoveNow' property on $Server. Error: $_"
    }

    Write-Verbose "Preventing $Server from hosting active database copies" 
    try {
        Set-MailboxServer $Server -DatabaseCopyAutoActivationPolicy Blocked -EA Stop 
    }
    catch {
        return "Error setting 'DatabaseCopyAutoActivationPolicy' property on $Server. Error: $( TryErr[0] )"
    }


    Write-Verbose "Placing $Server into maintenance mode"
    try {
        Set-ServerComponentState $Server -Component ServerWideOffline -State Inactive -Requester Maintenance -EA Stop 
    }
    catch {
        return "Error setting ServerWideOffline component on $Server to 'Inactive'. Error: $_"
    }

    Show-OffloadProgress

    #Remove-PSSession $ExchangeSession

    Write-Host "$Server is out of service and ready for maintenance." -ForegroundColor Green

} # function: main

#--------------------------------------------------------------------------------
function Test-Parameters {

    try {
        $Script:maint_server = Get-ExchangeServer $Server -ErrorAction Stop
    }
    catch {
        return "Unable to retrieve MailboxServer object for $Server"
    }
    try {
        $Script:dag_name = (Get-MailboxServer $Server -ErrorAction Stop ).DatabaseAvailabilityGroup.Name
    }
    catch {
        return "Unable to determine AD Site location of Server $Server"
    }

    $Script:exch_site = $maint_server.Site.Name

    # Select a target server for local queued messages, unless one was specified
    if ( $QueueTarget ) {
        Try {
            $Script:queue_server = Get-ExchangeServer $QueueTarget
        }
        catch {
            return "Unable to find the delivery queue target server $QueueTarget"
        }
    }
    else {
        Try { # Select a random Mailbox Server in the same site
            $Script:queue_server = Get-ExchangeServer | 
                                       Where { ($_.Site.Name -eq $exch_site)    -and 
                                               ($_.ServerRole -match 'Mailbox') -and
                                               ($_.Name -ne $Server) } |
                                       Get-Random
        }
        Catch {
            return "Unable to find a mailbox server in site $exch_site"
        }
    } # end if/else

    # A final sanity check
    <#
    # These tests only will work if the Exchange Assemblies have been loaded, otherwise, the "throw" is not processed.  An error still occurs, but it is non-terminating, and so the desired termination of the script does not occur.
    try {
        if ( $maint_server -isnot [Microsoft.Exchange.Data.Directory.Management.ExchangeServer] ) {
            throw "Somehow, $maint_server isn't an ExchangeServer object."
        }
        if ( $queue_server -isnot [Microsoft.Exchange.Data.Directory.Management.ExchangeServer] ) {
            throw "Somehow, $queue_server isn't an ExchangeServer object."
        }
    } 
    catch {
        throw "Unable to test target server objects.  Error: $_"
    }
    #>
} #End Test-Parameters

#--------------------------------------------------------------------------------
function Get-DeliveryQueues {
    param ( [string] $server )
    Get-Queue -Server $server | 
        where { $_.DeliveryType -ne 'ShadowRedundancy' -and $_.Identity -notlike '*\Poison' -and
                $_.NextHopDomain -ne 'donotdeliver.campus.ad.uvm.edu'
        }
}

#--------------------------------------------------------------------------------
function Get-QueueMessageCount {
    [CmdletBinding()]
    param( [parameter(Mandatory=$true,ValueFromPipeline=$true)]$queue)
    <# Originally:
    param( [parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [Microsoft.Exchange.Data.QueueViewer.ExtensibleQueueInfo]$queue
    )
    #>
    begin   { $MessageCount = 0; }
    process { $MessageCount += $queue.MessageCount }
    end     { $MessageCount }

}

#--------------------------------------------------------------------------------
function Show-OffloadProgress {

    $initial_mesg_count = Get-DeliveryQueues $Server | Get-QueueMessageCount
        write-debug "initial_mesg_count = $initial_mesg_count"
    # Tried using the -active paramenter to Get-MBDBCopyStatus, but throws error if none are found
    $initial_db_count   = ( Get-MailboxDatabaseCopyStatus -Server $Server | where Status -eq 'Mounted' ).Count
        write-debug "initial_db_count = $initial_db_count"
    $delay = 5 # seconds to sleep between checks

    <# Some Paramater hashes to make write-progress calls simpler
    $progress_mesg    = @{
        ID = 1
        Activity = 'Migrating messages from delivery queues'
    }
    $progress_db = @{
        ID = 2
        Activity = 'Moving active mailbox databases'
    }
    $progress_sleep = @{
        ID = 3
        Activity = "Waiting $delay seconds to check again"
        Status = 'Sleeping...'
    }
#>

    $migrating   = $true
    while ( $migrating ) {

        # Calculate status of message queues
        if ( $initial_mesg_count -gt 0 ) {
            $current_mesg_count = Get-DeliveryQueues $Server | Get-QueueMessageCount
            write-debug "current_mesg_count = $current_mesg_count"
<#            $moved_mesg_count   = $initial_mesg_count - $current_mesg_count
            write-debug "moved_mesg_count = $moved_mesg_count"
            [int] $percent_mesg = $moved_mesg_count / $initial_mesg_count * 100
            write-debug "percent_mesg = $percent_mesg"
            if ($percent_mesg -ge 100) { 
                $percent_mesg = 100
                $progress_mesg['Completed'] = $true
            }
            $progress_mesg['Status'] = "Messages in queues: $current_mesg_count"
            $progress_mesg['PercentComplete'] = $percent_mesg
#>        
            write-output "$current_mesg_count messages remaining"
        }
        else {
            #This must be declared here, or we will get trapped in the final if/then.
            [int32]$current_mesg_count = 0
        }
<#        else {
            $progress_mesg['Completed'] = $true
        }
#>            
        # Calculate status of mounted databases
        if ( $initial_db_count -gt 0 ) {
            $current_db_count =  ( Get-MailboxDatabaseCopyStatus -Server $Server | where Status -eq 'Mounted' ).Count
            write-debug "current_db_count = $current_db_count"
<#            $moved_db_count   = $initial_db_count - $current_db_count
            [int] $percent_db = $moved_db_count / $initial_db_count * 100
            if ($percent_db -ge 100) { 
                $percent_db = 100
                $progress_db['Completed'] = $true
            }
            $progress_db['Status'] = "Messages in queues: $current_db_count"
            $progress_db['PercentComplete'] = $percent_db

            write-progress @progress_db
#>
            #Curiously, this value does not display incremental decreases while the script is running.  18... 18... 18... 0!           
            write-output "$current_db_count mounted databases remaining"
        }
        else {
            #This must be declared here, or we will get trapped in the final if/then.
            [int32]$current_db_count = 0
        }
<#        else {
            $progress_db['Completed'] = $true
        }


        if ( $progress_mesg.completed -and $progress_db.completed ) {
#>
        if ( $current_mesg_count -eq 0 -and $current_db_count -eq 0 ) {
            $migrating = $false
        }  
        else {
            for ( $i=$delay; $i -ge 0 ; $i-- ) {
                #write-progress @progress_sleep -SecondsRemaining $i
                write-host '.' -NoNewLine
                sleep 1
            }
            write-host ''
        }
    } #end while migrating
} #end function:show-progress

# Run the script
main
