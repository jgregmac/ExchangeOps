### BEGINNING OF SCRIPT

#####
#
# Get-ActiveSyncDeviceReport
# Author: Paul Ponzeka
# Website: port25guy.com
# email ponzekap2 at gmail dot com
# Original URLs: 
#   http://port25guy.com/wp-content/uploads/2013/10/GetActiveSyncReport.ps1_.txt
#   http://port25guy.com/2013/03/25/how-to-get-a-report-of-active-sync-devices-in-exchange-2010exchange-2013/
#
######
param(
    [Parameter(Mandatory = $true)]
    [string] $SMTPServer = "",
    [Parameter(Mandatory = $true)]
    [string] $SMTPFrom = "",
    [Parameter(Mandatory = $true)]
    [string] $SMTPTo = "",
    [Parameter(Mandatory = $true)]
    [string] $exportpath = ""
    )

#######
#
# HTML Formatting Section
# Thanks to Paul Cunningham at http://exchangeserverpro.com/
#
#######
#
# 
#
######
$style = "<style>BODY{font-family: Arial; font-size: 10pt;}"
$style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
$style = $style + "TH{border: 1px solid black; background: #dddddd; padding: 5px; }"
$style = $style + "TD{border: 1px solid black; padding: 5px; }"
$style = $style + "</style>"


$messageSubject = "ActiveSync Device Report"

$message = New-Object System.Net.Mail.MailMessage $smtpfrom, $smtpto
$message.Subject = $messageSubject
$message.IsBodyHTML = $true

####  Get Mailbox

$EASDevices = ""
$AllEASDevices = @()

$EASDevices = ""| select 'User','PrimarySMTPAddress','DeviceType','DeviceModel','DeviceOS', 'LastSyncAttemptTime','LastSuccessSync'
$EasMailboxes = Get-Mailbox -ResultSize unlimited
foreach ($EASUser in $EasMailboxes) {
$EASDevices.user = $EASUser.displayname
$EASDevices.PrimarySMTPAddress = $EASUser.PrimarySMTPAddress.tostring()
    foreach ($EASUserDevices in Get-ActiveSyncDevice -Mailbox $EasUser.alias) {
	$EASDeviceStatistics = $EASUserDevices | Get-ActiveSyncDeviceStatistics
    $EASDevices.devicetype = $EASUserDevices.devicetype
    $EASDevices.devicemodel = $EASUserDevices.devicemodel
    $EASDevices.deviceos = $EASUserDevices.deviceos
	$EASDevices.lastsyncattempttime = $EASDeviceStatistics.lastsyncattempttime
	$EASDevices.lastsuccesssync = $EASDeviceStatistics.lastsuccesssync
    $AllEASDevices += $EASDevices | select user,primarysmtpaddress,devicetype,devicemodel,deviceos,lastsyncattempttime,lastsuccesssync
    }
    }
$AllEASDevices = $AllEASDevices | sort user
$AllEASDevices
$AllEASDevices | Export-Csv $exportpath\ActiveSyncReport.csv

######
#
# Send Email Report
#
########

$message.Body = $AllEasDevices | ConvertTo-Html -Head $style

$smtp = New-Object Net.Mail.SmtpClient($smtpServer)
$smtp.Send($message)

##END OF SCRIPT
