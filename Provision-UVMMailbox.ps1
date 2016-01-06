[CmdletBinding()]

param( 
    [Parameter(ValueFromPipeline=$true,Mandatory=$true)]
    [string] $userid,
    [Parameter(Mandatory=$false)]
    [string] $ErrorsTo
    )

begin {

 # Run at invocation

    set-strictmode -version latest


    new-variable -scope script -name mbxIndex -value 0

    function Fetch-NextMailboxDatabase {

        if ($script:mbxIndex -ge $mbxDbs.Count) {
            $script:mbxIndex = 0
        }

        $db = $mbxDbs[$script:mbxIndex++]

        return $db
    }

    try {
        get-command Enable-Mailbox -erroraction Stop | Out-Null
    } catch {
        throw "I can't find the command Enable-Mailbox.  Are you running this in Exchange Management Shell, connected to an Exchange server?"
    }

    write-verbose "importing user-dbmap.csv"

    # Ugh, I hate doing this, but I'm having trouble finding a variable like $PSScriptRoot that doesn't change based upon the caller.
    $installDir = "c:\local\scripts"

    # Find a list of available databases for provisioning.

    $mbxDbs = Get-mailboxdatabase | where-object { $_.IsExcludedFromProvisioning -eq $False `
                -and $_.IsSuspendedFromProvisioning -eq $False `
                -and $_.IsExcludedFromProvisioningBySpaceMonitoring -eq $False `
                -and $_.IsExcludedFromProvisioningBySchemaVersionMonitoring -eq $False `
                -and $_.IsExcludedFromInitialProvisioning -eq $False `
                -and $_.Recovery -eq $false }

    # Shuffle the list of databases.

    $shuffled = @($mbxDbs | sort-object {Get-Random})
    $mbxDbs = $shuffled

    write-verbose "This is in the begin block."

}

process {
    # Run for each pipeline object
    $upn = $userid + "@uvm.edu"
    write-verbose "I am provisioning $userid"
    try {
        $mb = Get-Mailbox -Identity $upn -ErrorAction Stop
    } catch [System.Management.Automation.RemoteException] {
        write-verbose "There was a problem with mailbox $upn.  I should create it."
        $mbxdb = Fetch-NextMailboxDatabase
        try {

            write-verbose "calling enable-mailbox -identity $upn -database $mbxdb"
            enable-mailbox -identity $upn -database $mbxdb -ErrorAction Stop
        } catch [system.exception] {
            if ($ErrorsTo) {
                echo "error trying to create $upn : $_" | out-file -append -filepath $ErrorsTo
            } else {
                throw "Cannot create $upn : $_"
            }
        } 
        # add the userid@msx-po.uvm.edu proxy address
        #echo "would call various set-mailbox commands"
        set-mailbox -identity $upn -emailaddresspolicyenabled $false
        set-mailbox -identity $upn -EmailAddresses @{add="$($userid)@msx-po.uvm.edu"}
    }
    # make sure sa_migrator can access mailbox for IMAP migration
    $oldWP = $WarningPreference
    $WarningPreference = "SilentlyContinue"
    Add-MailboxPermission -identity $upn -AccessRights FullAccess -User sa_migrator -AutoMapping:$False | Out-Null
    $WarningPreference=$oldWP
 }


end {

 # Run at the end of the pipeline, when closing down
 write-verbose "All done."

}



# Sorry, Geoff - I have to put this comment at the bottom of the script file,
# otherwise PS thinks it's documentation about an earlier object, rather than the whole script.

<#
.SYNOPSIS
Provision-UVMMailbox.ps1

.DESCRIPTION
Provision a UVM Mailbox (Enable-Mailbox) in the way we're accustomed to at UVM.

Requirements:  Be idempotent (if the mailbox is already provisioned in the way specified, just exit with 0)

.PARAMETER userid

The userid to create (typically, in netid form.)  This will get passed literally to New-Mailbox, so there needs to be an ADUser object already.

.PARAMETER ErrorsTo

Errors encountered during the Enable-Mailbox call will be written to this file, and execution will continue.
If unspecified, errors are thrown to the user (causing the script to stop).

.EXAMPLE

Provision-UVMMailbox.ps1 fcs

.EXAMPLE

(fcs, mga, gcd) | Provision-UVMMailbox.ps1

#>
