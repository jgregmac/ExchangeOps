<#
    Hide-COMUsers.ps1:
    Takes a list of users from a flat file (extracted from the penguid cluster)
     - Adds a contact object for all users in the feed.
     - Hides users in the feed from the GAL.
#>
param(
    [Parameter(Mandatory=$true)]
    [ValidateScript({Test-Path $_ -PathType 'Leaf'})]
    [string]$file,
    [Parameter(Mandatory=$false)]
    [string]$log = 'c:\local\temp\Hide-COMUsers.log'
)
Set-PSDebug -Strict

if (test-path -LiteralPath $log) {Remove-Item -LiteralPath $log -Force -Confirm:$false}

function writeHostAndLog{
    param(
    [Parameter(Mandatory=$true)]
    [string]$Out,

    [Parameter(Mandatory=$false)]
    [ValidateSet('Cyan','Yellow','Red','Gray')]
    [string]$Color
    )
	# Writes the string vailable provided by the -outString parameter to the full log path defined in
	#	"$log".  This log path needs to be defined globally.
	# Also echos the string text to the host console.
	# used in place of "tee-object" cmdlet, which does not have an "-append" parameter.
	# Returns:  Nothing documented or tested... intended as a blind write.
	$out | Out-File -Append -FilePath $log ;
    if ($color) {
        Write-Host $out -ForegroundColor $color
    } else {
	    Write-Host $out
    } 
}

function outMail {
param (
	[string]$Subj, 
	[string]$Body
	)
	# Sends a simple mail message using the .NET SMTP Client.
	# Routing server and to/from addresses can be changed by editing the variables in this function.
	# Subject and Body must be provided to this function in the form of string variables, 
	#	using the "-Subj" and "-Body" parameters of this function.
	# Returns: Nothing.  This is a blind send with no deliver confirmation.
	$SmtpClient = new-object system.net.mail.smtpClient
	$SmtpServer = "smtp.uvm.edu"
	$SmtpClient.host = $SmtpServer
	$From = "VI Script on vCenter1 <VIScript@vcenter1.campus.ad.uvm.edu>"
	$To = "saa-vmware@uvm.edu"
	$mailMessage = New-Object system.Net.mail.MailMessage($From,$To,$Subj,$Body)
	$mailMessage.bodyEncoding = [System.Text.Encoding]::UTF8
#	$mailAttachment = new-Object System.Net.Mail.Attachment($attach)
#	$mailMessage.Attachments.Add($mailAttachment)
	$SmtpClient.Send($mailMessage) 
}


[string]$ContactOU = "OU=Contacts,dc=campus,dc=ad,dc=uvm,dc=edu"
$startTime = Get-Date
writeHostAndLog -out "Script Started: $startTime" -Color Cyan

try {
    $re = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://msx-mh06.campus.ad.uvm.edu/powershell"
    Import-PSSession $re -ea Stop
    Import-Module ActiveDirectory -ea Stop
} catch {
    writeHostAndLog -out "Could not initialize the PowerShell environment." -color Red
    return 100
}

<#
#This block should gather a list of users that currently forward to the @med.uvm.ed mail domain...
# Could be useful for later revisions of this script.
[string[]]$medUsers = @()
Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited | 
    % { 
        [string]$sam = $_.samAccountName;
        Get-InboxRule -mailbox $_.DistinguishedName -ea SilentlyContinue -wa SilentlyContinue | 
            ? {$_.RedirectTo -like "*@med.uvm.edu*"} | % {
                writeHostAndLog -out "Found $sam" -Color Cyan 
                $medUsers += $sam
            }
    }
#>

$users = @()
$users = Import-Csv -Path $file -header 'name','email' -ea Stop

$noADAccount   = @()
$noMailbox     = @()
$hideFailed    = @()
$addressHidden = @()
$contactFailed = @()
$contactExists = @()

foreach ($user in $users) {
    writeHostAndLog -out ("  Processing user: " + $user.name) -Color Cyan
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
        try {   
            $mb | Set-Mailbox -HiddenFromAddressListsEnabled $True -ea Stop
        } catch {
            writeHostAndLog -out ("    Failed to hide " + $user.name + " from the GAL.") -Color Red
            $hideFailed += $user.name
            continue
        }
    } else {
        writeHostAndLog -out ("    User: " + $user.name + " Already was hidden in the GAL.") -Color Gray
    }
    # Create a new AD contact object:
    if (-not (Get-Contact -Identity $user.email -ea SilentlyContinue)) {
        try {
            #Add phone, department, title
            $mc = New-MailContact -Name $aduser.DisplayName -Alias ($user.name + '-med') `
                -ExternalEmailAddress $user.email -OrganizationalUnit $ContactOU -ea Stop
        } catch {
            writeHostAndLog -out ("    Failed to create a contact for: " + $aduser.SamAccountName) -color Red
            $contactFailed += $user.name
            continue
        }
        try {
            if ($aduser.Company)     { $mc | Set-Contact -Company $aduser.Company -ea Stop }
            if ($aduser.Department)  { $mc | Set-Contact -Department $aduser.Department -ea Stop }
            if ($aduser.Fax)         { $mc | Set-Contact -Fax $aduser.Fax -ea Stop }
            if ($aduser.GivenName)   { $mc | Set-Contact -FirstName $aduser.GivenName -ea Stop }
            if ($aduser.Initials)    { $mc | Set-Contact -Initials $aduser.Initials -ea Stop }
            if ($aduser.Office)      { $mc | Set-Contact -Office $aduser.Office -ea Stop }
            if ($aduser.OfficePhone) { $mc | Set-Contact -Phone $aduser.OfficePhone -ea Stop }
            if ($aduser.Surname)     { $mc | Set-Contact -LastName $aduser.Surname -ea Stop }
        } catch {
            writeHostAndLog -out ("    Failed to set extended attributes for contact: " + $aduser.SamAccountName) -color Red
            $contactFailed += $user.name
            continue
        }
    } else {
        #Mail contact object already exists in AD.
        writeHostAndLog -out ("    Contact object already exists") -color Gray
        $contactExists += $user.name
    }
}

writeHostAndLog -out ("List of users with no AD account: ") 
writeHostAndLog -out ('  ' + $noADAccount) -color Yellow
writeHostAndLog -out ("List of users with no Exchange mailbox: ")
writeHostAndLog -out ('  ' + $noMailbox) -color Yellow
writeHostAndLog -out ("List of users who could not be hidden in the GAL: ")
writeHostAndLog -out ('  ' + $hideFailed) -color Yellow
writeHostAndLog -out ("List of users for whom contact resource creation failed: ")
writeHostAndLog -out ('  ' + $contactFailed) -color Yellow
# writeHostAndLog -out ("List of users with addresses already hidden: " + $alreadyHidden) -color Gray
# writeHostAndLog -out ("List of users who already have a contact: " + $contactExists) -color Gray
writeHostAndLog -out " "
writeHostAndLog -out "Script Started: $startTime"
$endTime = get-date
writeHostAndLog -out "Script ended: $endTime"
writeHostAndLog -out " "
$elapsed = $endTime - $startTime
writeHostAndLog -out ("Elapsed Time: " + $elapsed.TotalSeconds + " Seconds")
