param(
    [Parameter(Mandatory=$false)]
      [string]$log = 'c:\local\temp\Query-IMAPLogsForPathPrefix.ps1',
    [Parameter(Mandatory=$false)]
      [array]$servers = @('msx-mh01','msx-mh02','msx-mh03','msx-mh04','msx-mh05','msx-mh06', `
                          'msx-tp01','msx-tp02','msx-tp03','msx-tp04','msx-tp05','msx-tp06')
      #[array]$servers = @('msx-mh06')
)
Set-PSDebug -Strict

# Define script-level variables:
[string]$to = "saa-msx@uvm.edu"
[string]$from = "IMAPQuery@msx-mgt1.campus.ad.uvm.edu"
[string]$Subj = "Potential IMAP 'mail/' Path Prefix user report for: " + [string](get-date)

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

$startTime = Get-Date
writeHostAndLog -out "Script Started: $startTime" -Color Cyan
writeHostAndLog -out " "

if (test-path -LiteralPath $log) { Remove-Item -LiteralPath $log -Force -Confirm:$false }

#Query all IMAP Back End logs for results within the specified date range:
[datetime]$endSearch = (get-date) #.AddDays(-1)
[datetime]$startSearch = $endSearch.AddDays(-1)

$sessionOptions = New-PSSessionOption -MaximumReceivedObjectSize 1gb -OperationTimeout 1800000 
forEach ($server in $servers) {
    Invoke-Command -ComputerName $server -AsJob -JobName ('IMAPQuery-' + $server) -ArgumentList $startSearch,$endSearch -SessionOption $sessionOptions -ScriptBlock {
        param($startSearch,$endSearch) 
        gci "d:\exchange server\Logging\Imap4\IMAP4BE*.LOG" | ? {($_.creationTime -lt $endSearch) -and ($_.lastWriteTime -gt $startSearch)} |
            % { # Import the IMAP log, selecting only user and parameters fields, limit to rows with "mail/* or "mail/%.
                [string]$outFile = 'c:\temp\imap_path_prefix_' + $env:computername + '.txt' ;
                if (test-path $outFile) { Remove-Item $outFile -Force -Confirm:$false } ;
                Import-CSV -path $_.FullName | select -Property user, parameters -first 300000|
                    where -Property parameters -match '"mail/[*%]' |
                    select -ExpandProperty user | Sort-Object -Unique | Out-File -FilePath $outFile -Append
                    #Tried collecting "cIp", but got only IPv6 and other exchange servers.
            }
    }
}

[boolean]$done = $false
[int32]$lCount = 0
$jobs = get-job -name IMAP* 
:imapJobs while (-not $done) {
    sleep 10
    if ($jobs | ? {$_.state -eq 'Running'}) {
        $lCount++
        write-host "  Jobs still running.  Loop Count:" $lCount -ForegroundColor Gray
    } else {
        $done = $true
    }
    if ($lCount -ge 180) {
        writeHostAndLog -Out 'IMAP log collection taking too long.' -Color Red
        $jobs | Stop-Job
        break imapJobs
    }
}

#Things break here with unknown WSMan transport failure.  Maybe WSMan needs more memory on each server?
#Strangely, the problem goes away if we run "Enable-PSWSManCombinedTrace" on the target server and enable WSMan Analytic/Debug event logs!
# Heisenburg strikes again.
$results = $jobs | Receive-job 
$jobs | Remove-Job # otherwise they'll hang out holding resources
Remove-Variable jobs

# Now get the output from the log query process:
# Need error handling here... will not work if jobs failed to generate the named text files.
$users = @()
foreach ($server in $servers) {
    #[string]$outFile = '\\' + $server + '\c$\temp\imap_path_prefix_' + $server + '.txt'
    [string]$outDir = '\\' + $server + '\c$\temp\imap_path_prefix*'
    #[string]$localFile = 'c:\temp\imap-reports\imap_path_prefix_' + $server + '.txt'
    #if (test-path $outFile) { Copy-Item $outFile 'c:\temp\imap-reports\' }
    $users += get-content $outDir 
}
$prefixUsers = $users | sort -unique

writeHostAndLog -out "List of users potentially still using the IMAP path prefix:"
writeHostAndLog -out " "
$prefixUsers | % {writeHostAndLog -out $_}

outMail -to $to -from $from -Subj $subj -Body ($prefixUsers | %{Write-Output "$_`r`n"})