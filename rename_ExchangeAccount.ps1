[CmdletBinding()]

param( 
    [Parameter(Mandatory = $true)]
    [string] $OldNetID,
    [Parameter(Mandatory = $true)]
    [string] $NewNetID
    )

# Update the user's mailbox alias to match the new netid
write-verbose "Updating the mailbox alias"
Set-mailbox "$NewNetID@uvm.edu" -Alias "$NewNetID"

# See if there is any redirect rule to be updated
write-verbose "Looking for redirect rule"
$IBRule = Get-InboxRule -mailbox "$NewNetID@uvm.edu" | Where name -like 'DO NOt REMOVE: redirect*'
# If we found one - and it is of the form - update it
if ( $IBRule -and $IBRule.RedirectTo[0].Address -eq "$OldNetID@pobox2.uvm.edu") {
    write-verbose "We found $OldNetID@pobox2.uvm.edu => Updating!"
    Set-InboxRule -Identity $IBRule.Identity -RedirectTo "$NewNetID@pobox2.uvm.edu"
}

write-verbose "All done -- going home."


<#
.SYNOPSIS
rename_ExchangeAccount.ps1

.DESCRIPTION
When an account is renamed on rottweiler, this script is run via SSH to force the Mailbox Alias value to change to the new NetID, and the InboxFilterRules are searched and a redirect for those that are not yet really in Exchange is updated to route to the new NetID.

Requirement: do no harm.

.PARAMETER OldNetID

The userid that is being renamed.

.PARAMETER NewNetID

The new NetID that the old NetID is being converted to.

.EXAMPLE

rename_ExchangeAccount.ps1 fcs fswasey

#>
