param(
    [Parameter(Mandatory=$false)]
      [string]$log = 'c:\local\temp\Query-IMAPLogsForPathPrefix.ps1',
    [Parameter(Mandatory=$false)]
      [array]$servers = @('msx-mh01','msx-mh02','msx-mh03','msx-mh04','msx-mh05','msx-mh06', `
                          'msx-tp01','msx-tp02','msx-tp03','msx-tp04','msx-tp05','msx-tp06')
)
Set-PSDebug -Strict

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

$startTime = Get-Date
writeHostAndLog -out "Script Started: $startTime" -Color Cyan
writeHostAndLog -out " "

if (test-path -LiteralPath $log) { Remove-Item -LiteralPath $log -Force -Confirm:$false }

#Query all IMAP Back End logs for results within the specified date range:
[datetime]$endSearch = [datetime]::today
[datetime]$startSearch = $endSearch.AddDays(-1)
#$jobs = $servers | % { 
forEach ($server in $servers) {
    Invoke-Command -ComputerName $server -AsJob -JobName ('IMAPQuery-' + $server) -ArgumentList $startSearch,$endSearch -ScriptBlock {
        param($startSearch,$endSearch) 
        gci "d:\exchange server\Logging\Imap4\IMAP4BE*.LOG" | ? {($_.creationTime -lt $endSearch) -and ($_.lastWriteTime -gt $startSearch)} |
            % { # Import the IMAP log, selecting only user and parameters fields, limit to rows with "mail/* or "mail/%.
                Import-CSV -path $_.FullName | 
                    where -Property parameters -match '"mail/[*%]' |
                    select user #Tried collecting "cIp", but got only IPv6 and other exchange servers.
            }
    } 
}

#$jobs = get-job -Name 'IMAP*'
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

$results = Receive-job $jobs
Remove-Job $jobs  # otherwise they'll hang out holding resources

#Now capture the "user" field of the CSV and sort -unique.
$prefixUsers = @()
$prefixUsers = $results | select -ExpandProperty user | sort -Unique 
$prefixUsers