﻿<#
.SYNOPSIS
Get-EASDeviceReport.ps1 - Exchange Server ActiveSync device report

.DESCRIPTION 
Produces a report of ActiveSync device associations in the organization.

.OUTPUTS
Results are output to screen, as well as optional log file, HTML report, and HTML email

.PARAMETER SendEmail
Sends the HTML report via email using the SMTP configuration within the script.

.PARAMETER SortBy
Sorts the resultant report by "Email", "Model", "OS","LastSuccessSync" fields.  Default is "LastSuccssSync".

.EXAMPLE
.\Get-EASDeviceReport.ps1
Produces a CSV file containing stats for all ActiveSync devices.

.EXAMPLE
.\Get-EASDeviceReport.ps1 -SendEmail -MailFrom:exchangeserver@exchangeserverpro.net -MailTo:paul@exchangeserverpro.com -MailServer:smtp.exchangeserverpro.net
Sends an email report with CSV file attached for all ActiveSync devices.

.EXAMPLE
.\Get-EASDeviceReport.ps1 -Age 30
Limits the report to devices that have not attempted synced in more than 30 days.

.NOTES
Written by: Paul Cunningham and J. Greg Mackinnon

Find me on:

* My Blog:	http://paulcunningham.me
* Twitter:	https://twitter.com/paulcunningham
* LinkedIn:	http://au.linkedin.com/in/cunninghamp/
* Github:	https://github.com/cunninghamp
* Greg's Blog: http://blog.uvm.edu/jgm

For more Exchange Server tips, tricks and news
check out Exchange Server Pro.

* Website:	http://exchangeserverpro.com
* Twitter:	http://twitter.com/exchservpro

Change Log
V1.00, 25/11/2013 - Initial version
V1.01, 11/02/2014 - Added parameters for emailing the report and specifying an "age" to report on
V1.02, 17/02/2014 - Fixed missing $mydir variable and added UTF8 encoding to Export-CSV and Send-MailMessage
v2.0,  29/10/2015 + Updated for Exchange 2013 cmdlets.  
                      - No longer requires looping though all mailboxes for smaller memory footprint and faster execution
                      - Added "sort" parameter.
#>

#requires -version 2

[CmdletBinding()]
param (
	
	[Parameter( Mandatory=$false)]
	[switch]$SendEmail,

	[Parameter( Mandatory=$false)]
	[string]$MailFrom,

	[Parameter( Mandatory=$false)]
	[string]$MailTo,

	[Parameter( Mandatory=$false)]
	[string]$MailServer,

    [Parameter( Mandatory=$false)]
    [int]$Age = 0,

    [Parameter( Mandatory=$false)]
    [string]$reportFile = "c:\local\temp\Get-EASDeviceReport.csv",

    [Parameter( Mandatory=$false)]
    [ValidateSet("Email", "Model", "OS","LastSuccessSync")]
    [string]$sortBy = 'LastSuccessSync'

	)


#...................................
# Variables
#...................................

$now = Get-Date											#Used for timestamps
$date = $now.ToShortDateString()						#Short date format for email message subject

$report = @()

$stats = @("DeviceID",
            "DeviceAccessState",
            "DeviceAccessStateReason",
            "DeviceModel"
            "DeviceType",
            "DeviceFriendlyName",
            "DeviceOS",
            "LastSyncAttemptTime",
            "LastSuccessSync"
          )

$reportemailsubject = "Exchange ActiveSync Device Report - $date"
#$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path
#$reportfile = "$myDir\ExchangeActiveSyncDeviceReport.csv"


#...................................
# Email Settings
#...................................

$smtpsettings = @{
	To =  $MailTo
	From = $MailFrom
    Subject = $reportemailsubject
	SmtpServer = $MailServer
	}


#...................................
# Initialize
#...................................

#Add Exchange 2010/2013 snapin if not already loaded in the PowerShell session
if (!(Get-PSSnapin | where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"}))
{
	try
	{
		Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction STOP
	}
	catch
	{
		#Snapin was not loaded
		Write-Warning $_.Exception.Message
		EXIT
	}
	. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
	Connect-ExchangeServer -auto -AllowClobber
}


#...................................
# Script
#...................................

Write-Host "Fetching all mobile devices"
$EASDevices = Get-MobileDevice -ResultSize Unlimited

Foreach ($EASDevice in $EASDevices) {
    Write-Host -ForegroundColor Green "Processing $($EASDevice.DeviceID)"

    #Get statistics for the current device:
    $EASDeviceStats = @()
    $EASDeviceStats = @(Get-MobileDeviceStatistics -Identity $EasDevice.Identity)
    #If we have a device with no Stats, set the Stats object to the device object, to cut down on report data loss.
    if ($EASDeviceStats.count -eq 0) {
        Write-Host "No Device Stats available for this device."
        $EasDeviceStats = $EASDevice
    }

    #Get information on the owner of the device:
    $deletedUser = $false
    try { 
        $MailboxInfo = Get-Mailbox $EasDevice.UserDisplayName -ea Stop | Select DisplayName,PrimarySMTPAddress
    } catch {
        Write-Host "Could not get mailbox info for device owner."
        $deletedUser = $true
    }
    #If a device exists for a user who no longer exists in the domain, we need to capture that information:
    if ($deletedUser) {
        $displayName = $EASDevice.UserDisplayName.Split('/') | select -last 1
    } else {
        $displayName = $MailboxInfo.DisplayName
    }

    Write-Host -ForegroundColor Green "Processing $($EASDevice.DeviceID)"
    
    #Normalize the LastSyncAttemptTime statistic:    
    $lastsyncattempt = ($EASDeviceStats.LastSyncAttemptTime)
    if ($lastsyncattempt -eq $null) {
        $syncAge = "Never"
    } else {
        $syncAge = ($now - $lastsyncattempt).Days
    }

    #Add to report if last sync attempt greater than Age specified
    if ($syncAge -ge $Age -or $syncAge -eq "Never") {
        Write-Host -ForegroundColor Yellow "$($EASDevice.DeviceID) sync age of $syncAge days is greater than $age, adding to report"

        $reportObj = New-Object PSObject
        $reportObj | Add-Member NoteProperty -Name "Display Name" -Value $displayName
        $reportObj | Add-Member NoteProperty -Name "Email Address" -Value $MailboxInfo.PrimarySMTPAddress
        $reportObj | Add-Member NoteProperty -Name "Sync Age (Days)" -Value $syncAge
                
        Foreach ($stat in $stats) {
            $reportObj | Add-Member NoteProperty -Name $stat -Value $EASDeviceStats.$stat
        }

        $report += $reportObj
    }
    #Cleanup for the next loop:
    Remove-Variable mailboxInfo,displayName,lastSyncAttempt,syncAge,reportObj -ea SilentlyContinue
}

#Apply sorting options
switch ($sortBy) {
    "Email"            {$sort = "Email Address"}
    "Model"            {$sort = "DeviceModel"}
    "OS"               {$sort = "DeviceOS"}
    "LastSuccessSync"  {$sort = $sortBy}
}

$report = $report | Sort-Object -Property $sort

Write-Host -ForegroundColor White "Saving report to $reportfile"
$report | Export-Csv -NoTypeInformation $reportfile -Encoding UTF8


if ($SendEmail) {

    $reporthtml = $report | ConvertTo-Html -Fragment

	$htmlhead="<html>
				<style>
				BODY{font-family: Arial; font-size: 8pt;}
				H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
				H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
				H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
				TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
				TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
				TD{border: 1px solid #969595; padding: 5px; }
				td.pass{background: #B7EB83;}
				td.warn{background: #FFF275;}
				td.fail{background: #FF2626; color: #ffffff;}
				td.info{background: #85D4FF;}
				</style>
				<body>
                <p>Report of Exchange ActiveSync device associations with greater than $age days since last sync attempt as of $date. CSV version of report attached to this email.</p>"
		
	$htmltail = "</body></html>"	

	$htmlreport = $htmlhead + $reporthtml + $htmltail

	Send-MailMessage @smtpsettings -Body $htmlreport -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8) -Attachments $reportfile
}