<#
.SYNOPSIS
Modifys records of College of Medicine users by hiding them in the Global 
Address List and publish an alternate contact record.

.DESCRIPTION
Takes a list of users from a CSV in the format NetID,emailAddress:
    - Hides users in the feed from the GAL.
    - Adds a contact object for all users in the feed.
Creates a list of all mailboxes hidden in the GAL:
    - Unhides the mailboxes that are not in in the feed.
Creates a list of all contact records in the contacts OU:
    - Removes the contact records that are not in the feed.
Returns:
   - 0   - Script ran successfully.
   - 100 - Could not initialize the PowerShell environment.
   - 110 - Failed to import $the specified input as CSV.
   - 120 - Failed to import the constrained list of COM/Allied Health users.
   - 200 - Failed to get a list of GAL-hidden mailboxes.
   - 210 - Failed to get a list of current Mail Contacts Objects.

.PARAMETER penguinFeed
Name/Path of a CSV file in NetID,emailAddress format that lists current penguin
cluster users that forward their mail to @med, and thus need to be hidden.

.PARAMETER log
Name/Path of a file to which to log the actions of this script.  Default value 
is:
C:\local\temp\Hide-COMUsers.log

.PARAMETER mail
Switch value that indicates if the results of the script should be failed.  
Default value is $true.

.PARAMETER searchList
Name/Path of a list of users which will be searched for existing 
forward/redirect rules.  searchList needs to be provided from an external 
feed.  Use of this list greatly accelerates the execution time of the 
script.  If we needed to search all inbox rules, this script would run for a
minimum of two hours.

.PARAMETER servers
An array of unqualified server names that are potential targets for implicit
remoting of Exchange cmdlets.  The target server for this run of the script 
will be selected randomly from this list.

#>
param(
    [Parameter(Mandatory=$false)]
      #Validation may be causing the script to puke.
      #[ValidateScript({Test-Path $_ -PathType 'Leaf'})]
      [string]$penguinFeed = '\\files\shared\saa\Exchange\temp\med-forwards-penguinonly.csv',
    [Parameter(Mandatory=$false)]
      [string]$log = 'c:\local\temp\Hide-COMUsers.log',
    [Parameter(Mandatory=$false)]
      [Boolean]$mail = $true,
    [Parameter(Mandatory=$false)]
      #Validation may be causing the script to puke.
      #[ValidateScript({Test-Path $_ -PathType 'Leaf'})]
      [string]$searchList = '\\files\shared\saa\Exchange\temp\COM_NURSING_GAL.txt',
    [Parameter(Mandatory=$false)]
      [array]$servers = @('msx-mh01','msx-mh02','msx-mh03','msx-mh04','msx-mh05','msx-mh06', `
                          'msx-tp01','msx-tp02','msx-tp03','msx-tp04','msx-tp05','msx-tp06')
)
Set-PSDebug -Strict

# Define script-level variables:
[string]$to = "saa-msx@uvm.edu"
[string]$from = "Hide-ComUsers Scheduled Task <Hide-ComUsers@msx-mgt2.campus.ad.uvm.edu>"
[string]$ContactOU = "OU=Contacts,dc=campus,dc=ad,dc=uvm,dc=edu"
[string]$managedOU = "OU=people,DC=campus,DC=ad,DC=uvm,DC=edu" 

if (test-path -LiteralPath $log) {Remove-Item -LiteralPath $log -Force -Confirm:$false}

function writeHostAndLog{
    param(
    [Parameter(Mandatory=$true)]
      [string]$Out,
    [Parameter(Mandatory=$false)]
      [ValidateSet('Cyan','Yellow','Red','Gray')]
      [string]$Color
    )
    # Writes parameter "Out" to the console and to the global variable $log.
    #   Optionally will use the specified color for console output.
    #   Tee-Object largely replaces the need for this function, but does not support color.
	$out | Out-File -Append -FilePath $log ;
    if ($color) {
        Write-Host $out -ForegroundColor $color
    } else {
	    Write-Host $out
    } 
}

function showElapsedTime {
    param(
        [Parameter(Mandatory=$true)]
        [datetime]$startTime
    )
    writeHostAndLog -out "Script Started: $startTime"
    $currentTime = get-date
    writeHostAndLog -out "Current time: $currentTime"
    writeHostAndLog -out " "
    $elapsed = $currentTime - $startTime
    writeHostAndLog -out ("Elapsed Time: " + $elapsed.Hours + ":" + $elapsed.Minutes + ":" + $elapsed.Seconds)
}

function outMail {
param (
    [string]$to,
    [string]$from,
	[string]$Subj, 
	[string]$Body
	)
	# Sends a simple mail message using the .NET SMTP Client.
	# Routing server and to/from addresses can be changed by editing the variables in this function.
	# Subject and Body must be provided to this function in the form of string variables, 
	#	using the "-Subj" and "-Body" parameters of this function.
	# Returns: Nothing.  This is a blind send with no delivery confirmation.
	$SmtpClient = new-object system.net.mail.smtpClient
	[string]$SmtpServer = "smtp.uvm.edu"
	[string]$SmtpClient.host = $SmtpServer
	$mailMessage = New-Object system.Net.mail.MailMessage($From,$To,$Subj,$Body)
	$mailMessage.bodyEncoding = [System.Text.Encoding]::UTF8
#	$mailAttachment = new-Object System.Net.Mail.Attachment($attach)
#	$mailMessage.Attachments.Add($mailAttachment)
	$SmtpClient.Send($mailMessage) 
}

function logError {
    param(
        [parameter(Mandatory=$true)]
          $errorObj
    )
    writeHostAndLog -out "Error Exception:" -color Yellow
    writeHostAndLog -Out $errorObj.Exception.Message -color Red
    writeHostAndLog -out "Invocation Line:" -color Yellow
    writeHostAndLog -Out $errorObj.InvocationInfo.Line -color Red
    writeHostAndLog -out "Invocation Position Message:" -color Yellow
    writeHostAndLog -Out $errorObj.InvocationInfo.PositionMessage -color Red
}

#Start logging:
$startTime = Get-Date
writeHostAndLog -out "Script Started: $startTime" -Color Cyan

$server = Get-Random $servers

try {
    $re = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri ("http://" + $server + ".campus.ad.uvm.edu/powershell")
    Import-PSSession $re -ea Stop
    Import-Module ActiveDirectory -ea Stop
} catch {
    writeHostAndLog -out "Could not initialize the PowerShell environment." -color Red
    logError $_
    return 100
}

###############################################################################
# Get list of "penguin cluster" users to hide/contact-enable:
#
writeHostAndLog -Out "Importing list of penguin cluster forwarders..." -Color Cyan
$users = @()
try {
    $users = Import-Csv -Path $penguinFeed -header 'name','email' -ea Stop
} catch {
    writeHostAndLog -Out "Failed to import $penguinFeed as CSV" -Color Red
    return 110
}
writeHostAndLog -out ("Count of penguin forwarders: " + $users.count) -Color Cyan
writeHostAndLog -out " "
#
###############################################################################

###############################################################################
# Start search for Exhchange mailbox forwarders
#
writeHostAndLog -out 'Evaluating list of COM/Allied Health users for @med forwarding' -Color Cyan
try {
    [string[]]$searchUsers = Get-Content -Path $searchList -ea Stop
} catch {
    writeHostAndLog -out "  Could not load list of users to evaluate from $searchList" -color Red
    logError $_
    return 120
}

[PSCustomObject[]]$redirUsers    =  @()
[PSCustomObject[]]$fwdUsers      =  @()
[PSCustomObject[]]$redirLegacyDN =  @()
[PSCustomObject[]]$fwdLegacyDN   =  @()
[string[]]$unprovUsers = @()
[int32]$count = 0

forEach ($user in $searchUsers) {  
    # [string]$sam = $_.samAccountName;
    # Get-InboxRule -mailbox $_.DistinguishedName -ea Stop -wa SilentlyContinue |
    $count ++
    try {
        $rules = @()
        $rules = Get-InboxRule -mailbox $user -ea Stop -wa SilentlyContinue 
    } catch {
        # Uncomment to see unprovisioned users
        # writeHostAndLog -out ("    Could not get inbox rules for user: $user") -color Yellow
        $unprovUsers += $user
        continue
    }
    forEach ($rule in $rules) {
        if ($rule.RedirectTo -like "*@med.uvm.edu*") {
            [PSCustomObject]$obj = [PSCustomObject]@{
                name = $user; 
                email = ($rule.RedirectTo[0].Split('"') | select -index 1)
            }
            writeHostAndLog -out ("    Found redirected user: " + $obj.name+ " (" + $count + " of " + $searchUsers.Count +")") -Color Gray 
            $redirUsers += $obj
        }
        elseif ($rule.RedirectTo -and $rule.RedirectTo[0].Address -match 'FYDIBOHF23SPDLT') {
            [PSCustomObject]$obj = [PSCustomObject]@{
                name = $user; 
                LegacyExchangeDN = ($rule.RedirectTo[0].Address)
            }
            writeHostAndLog -out ("    Found redirected user (to LegacyExchangeDN): " + $obj.name+ " (" + $count + " of " + $searchUsers.Count +")") -Color Gray 
            $redirLegacyDN += $obj
        }
        if ($rule.ForwardTo -like "*@med.uvm.edu*") {
            [PSCustomObject]$obj = [PSCustomObject]@{ 
                name = $user; 
                email = ($rule.ForwardTo[0].Split('"') | select -index 1)
            }
            writeHostAndLog -out ("    Found forwarded user: " + $obj.name + " (" + $count + " of " + $searchUsers.Count +")") -Color Gray
            $fwdUsers += $obj
        }
		elseif ($rule.ForwardTo -and $rule.ForwardTo[0].Address -match 'FYDIBOHF23SPDLT') {
            [PSCustomObject]$obj = [PSCustomObject]@{ 
                name = $user; 
                LegacyExchangeDN = ($rule.ForwardTo[0].Address)
            }
            writeHostAndLog -out ("    Found forwarded user (to LegacyExchangeDN): " + $obj.name + " (" + $count + " of " + $searchUsers.Count +")") -Color Gray
            $fwdLegacyDN += $obj
        }
    }
}
writeHostAndLog -out " "
writeHostAndLog -out ("Count of unprovisioned users: " + $unprovUsers.count) -color Cyan 
writeHostAndLog -out ("Count of forwarding users: " + $fwdUsers.count) -color Cyan
writeHostAndLog -out ("Count of redirected users: " + $redirUsers.count) -Color Cyan
writeHostAndLog -out ("Count of forwarding users (LegacyExchangeDN): " + $fwdLegacyDN.count) -color Cyan
writeHostAndLog -out ("Count of redirected users (LegacyExchangeDN): " + $redirLegacyDN.count) -Color Cyan
writeHostAndLog -Out " "

#if (test-path $outList) {Remove-Item -Path $outList -Force -Confirm:$false}
#writeHostAndLog -Out ("Writing out currently forwarding users to: " + $outList) -Color Cyan
#$redirUsers | Out-File -FilePath $outList -Append
#$fwdUsers | Out-File -FilePath $outList -Append

#Append mailbox-enabled users with Med forwarding to the extract from the Penguins:
writeHostAndLog -out ("Appending Exchange forwardwer to the penguin forwarders list...") -Color Cyan
$users += $redirUsers
$users += $fwdUsers
$users += $redirLegacyDN
$users += $fwdLegacyDN

writeHostAndLog -Out ("New count of forwarding users: " + $users.count)

writeHostAndLog -Out " "
showElapsedTime -startTime $startTime

#
# Stop search for Exhchange mailbox forwarders
###############################################################################

###############################################################################
# Start Hide-in-GAL / Create Contact Object loop:
#
writeHostAndLog -out "Hiding from GAL and creating contact objects for users in the feed:" -Color Cyan

$noADAccount   = @()
$noMailbox     = @()
$hideFailed    = @()
$addressHidden = @()
$contactFailed = @()
$contactExists = @()

foreach ($user in $users) {
    #Uncomment the following for verbose reporting on the user currently being processed:
    #writeHostAndLog -out ("  Processing user: " + $user.name) -Color Cyan
    try {
        $aduser = Get-ADUser -Identity $user.name -Properties `
            Company,Department,DisplayName,Fax,GivenName,Initials,Office,OfficePhone,Surname `
            -ea Stop 
        #Add a check for expired users here
    } catch {
        $out = '    Failed to get AD account information for user: ' + $user.name
        writeHostAndLog -out $out -Color Red
        $noADAccount += $user.name
        continue
    }

    #Hide the users from the GAL:
    try {
        $mb = Get-Mailbox -Identity $user.name -ea Stop 
    } catch {
        writeHostAndLog -out ("    Could not retrieve a mailbox for: " + $user.name + " `r`n    Is the user provisioned?") -Color Red
        $noMailbox += $user.name
        continue
    }
    if (-not $mb.HiddenFromAddressListsEnabled) {
        writeHostAndLog -out ("    Hiding " + $user.name + " from the GAL.") -Color Gray
        try {   
            $mb | Set-Mailbox -HiddenFromAddressListsEnabled $True -ea Stop
        } catch {
            writeHostAndLog -out ("    Failed to hide " + $user.name + " from the GAL.") -Color Red
            $hideFailed += $user.name
            continue
        }
    } else {
        # writeHostAndLog -out ("    User: " + $user.name + " Already was hidden in the GAL.") -Color Gray
    }
    # Create a new AD contact object:
    [string]$mcAlias = $user.name + '-med'
    $oldMc = Get-Contact -Identity $mcAlias -ea SilentlyContinue
    if (-not $oldMc) {
        writeHostAndLog -out ("    Creating a contact record for: " + $user.name) -Color Gray
        try {
            #Add phone, department, title
            if ($aduser.DisplayName) {$cName = $aduser.DisplayName} else {$cName = $aduser.Name}
            $newMC = New-MailContact -Name $cName -Alias ($user.name + '-med') `
                -ExternalEmailAddress $user.email -OrganizationalUnit $ContactOU -ea Stop
        } catch {
            writeHostAndLog -out ("    Failed to create a contact for: " + $user.name) -color Red
            $contactFailed += $user.name
            continue
        }
        try {
            if ($aduser.Company)     { $newMC | Set-Contact -Company $aduser.Company -ea Stop }
            if ($aduser.Department)  { $newMC | Set-Contact -Department $aduser.Department -ea Stop }
            if ($aduser.Fax)         { $newMC | Set-Contact -Fax $aduser.Fax -ea Stop }
            if ($aduser.GivenName)   { $newMC | Set-Contact -FirstName $aduser.GivenName -ea Stop }
            if ($aduser.Initials)    { $newMC | Set-Contact -Initials $aduser.Initials -ea Stop }
            if ($aduser.Office)      { $newMC | Set-Contact -Office $aduser.Office -ea Stop }
            if ($aduser.OfficePhone) { $newMC | Set-Contact -Phone $aduser.OfficePhone -ea Stop }
            if ($aduser.Surname)     { $newMC | Set-Contact -LastName $aduser.Surname -ea Stop }
        } catch {
            writeHostAndLog -out ("    Failed to set extended attributes for contact: " + $aduser.SamAccountName) -color Red
            $contactFailed += $user.name
            continue
        }
    } else {
        #Mail contact object already exists in AD.
        #writeHostAndLog -out ("    Contact object already exists") -color Gray
        $contactExists += $user.name
        ### NOTE: We need to add a routine here to check for changes to the Mail Contact and update if there are changes! ###
    }
}
showElapsedTime -startTime $startTime
writeHostAndLog -out " "
#
# End Hide-in-GAL / Create Contact Object loop:
###############################################################################

###############################################################################
# Start Unhide-in-GAL / Remove Contact Object loop:
#
writeHostAndLog -out "Un-hiding mailboxes that are not in the feed:" -Color Cyan

$hiddenUsers  = @()
$unhideFailed = @()

try {
    # Find the account names for all Exchange users who are hidden in the GAL:
    $hiddenUsers = Get-Mailbox -OrganizationalUnit $managedOU -ResultSize Unlimited `
        -Filter {(HiddenFromAddressListsEnabled -eq $true)} -ea Stop |
        select -ExpandProperty SamAccountName
} catch {
    writeHostAndLog -out "Failed to get a list of GAL-hidden mailboxes." -color Red
    logError $_
    return 200
}
# Loop though all hidden users:
foreach ($huser in $hiddenUsers) {
    # Check to see if the current hidden user was in our import list:
    if (-not $users.name.Contains($huser)) {
        # Unhide the user
        try {
            writeHostAndLog -out ("    Unhiding user: "+ $huser) -Color Gray
            Set-Mailbox -Identity $huser -HiddenFromAddressListsEnabled $False -ea Stop
        } catch {
            writeHostAndLog -out ("    Failed to unhide user: " + $huser) -color Red
            $unhideFailed += $huser
            continue
        }
    }
}
showElapsedTime -startTime $startTime
writeHostAndLog -out " "

writeHostAndLog -out "Removing contacts that are not in the feed:" -Color Cyan
$allContacts     = @()
$rmContactFailed = @()
try {
    # Find the Aliases for all existing contact objects:
    $allContacts = Get-MailContact -OrganizationalUnit $ContactOU -ResultSize Unlimited -ea Stop |
        select -ExpandProperty Alias
} catch {
    writeHostAndLog -out "Failed to get a list of current Mail Contacts Objects." -color Red
    return 210
}
# Loop though all contact objects:
foreach ($contact in $allContacts) {
    #Check to see if the current contact is in the import list:
    if (-not $users.name.Contains($contact.replace('-med',''))) {
        # Remove the contact
        try {
            writeHostAndLog -out ("    Removing contact: " + $contact) -Color Gray
            Remove-MailContact -Identity $contact -Confirm:$false -ea Stop
        } catch {
            writeHostAndLog -out ("    Failed to remove contact for: " + $contact) -color Red
            $rmContactFailed += $contact
            continue
        }
    }
}
#
# End Unhide-in-GAL / Remove Contact Object loop:
###############################################################################
showElapsedTime -startTime $startTime
writeHostAndLog -out " "


writeHostAndLog -out ("List of users with no AD account: ") 
writeHostAndLog -out ('  ' + $noADAccount) -color Yellow
writeHostAndLog -out ("List of users with no Exchange mailbox: ")
writeHostAndLog -out ('  ' + $noMailbox) -color Yellow
writeHostAndLog -out ("List of users who could not be hidden in the GAL: ")
writeHostAndLog -out ('  ' + $hideFailed) -color Yellow
writeHostAndLog -out ("List of users for whom contact resource creation failed: ")
writeHostAndLog -out ('  ' + $contactFailed) -color Yellow
writeHostAndLog -out ("List of users who could not be un-hidden from the GAL: ")
writeHostAndLog -out ('  ' + $unhideFailed) -color Yellow
writeHostAndLog -out ("List of users for whom contact removal failed: ")
writeHostAndLog -out ('  ' + $rmContactFailed) -color Yellow
# Verbose information... commented out
# writeHostAndLog -out ("List of users with addresses already hidden: " + $alreadyHidden) -color Gray
# writeHostAndLog -out ("List of users who already have a contact: " + $contactExists) -color Gray
writeHostAndLog -out " "

showElapsedTime -startTime $startTime

if ($mail) {
    [string]$subj = "Hide-ComUsers Scheduled Task for: " + [string]([datetime]::now)
    [string[]]$bodyArray = Get-Content -Path $log
    [string]$body = ''
    foreach ($line in $bodyArray) {$body += $line + "`r`n"}
    outMail -to $to -from $from -Subj $subj -Body $body
}

get-pssession | remove-pssession -ea SilentlyContinue | Out-Null

Return 0