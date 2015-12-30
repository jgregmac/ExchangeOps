#This is not a real script.  It is a transcript of us trying to analyze the EWS logs.
# NOTE:  This probably is not the best way to analyze loads of log data!

$servers = Get-Mailboxserver #| where-object -property Name -notin ("msx-mgt1", "msx-mgt2")
# In 2016, everyone's a mailbox server I guess...

#Example of querying all "ErrorServerBusy" events on all servers for as far back as the logs go:
$jobs = $servers | % { 
    Invoke-Command -AsJob -ComputerName $_ -ScriptBlock { 
        gci "d:\exchange server\logging\ews\*.LOG" 
            % { 
                Import-CSV -path $_.FullName | 
                    where-object -property ErrorCode -EQ "ErrorServerBusy" 
            }
    } 
} 

# Example for querying all events for an individual user for a particular day:

$jobs = $servers | % { 
    Invoke-Command -AsJob -ComputerName $_ -ScriptBlock { 
        gci "d:\exchange server\logging\ews\Ews_20151209*.LOG" | 
            % { 
                Import-CSV -path $_.FullName | 
                    ? { ($_.AuthenticatedUser -match 'Geoffrey\.Duke') }
            }
    } 
}

#Example of querying all HTTP Status 500 events on all servers for a fixed log creation time range:
[datetime]$startTime = '2015-12-24 12:00:00am'
[datetime]$endTime = '2015-12-24 11:59:59pm'
$jobs = $servers | % { 
    Invoke-Command -AsJob -ComputerName $_ -ArgumentList $startTime,$endTime -ScriptBlock {
        param($startTime,$endTime) 
        gci "d:\exchange server\logging\ews\*.LOG" | ? {($_.creationTime -lt $endTime) -and ($_.lastWriteTime -gt $startTime)} |
            % { 
                Import-CSV -path $_.FullName | 
                    where-object -property HttpStatus -eq 500
            }
    } 
}

#Example of querying all MacOutlook activity for logs written to in a specified date range:
[datetime]$startTime = '2015-12-24 12:00:00am'
[datetime]$endTime = '2015-12-24 11:59:59pm'
$jobs = $servers | % { 
    Invoke-Command -AsJob -ComputerName $_ -ArgumentList $startTime,$endTime -ScriptBlock {
        param($startTime,$endTime) 
        gci "d:\exchange server\logging\ews\*.LOG" | ? {($_.creationTime -lt $endTime) -and ($_.lastWriteTime -gt $startTime)} |
            % { 
                Import-CSV -path $_.FullName | 
                    where-object -property UserAgent -match '^MacOutlook'
            }
    } 
}

#Example of querying all activity with TotalRequestTime over 1800000ms for logs written to in a specified date range:
[datetime]$startTime = '2015-12-22 12:00:00am'
[datetime]$endTime = '2015-12-24 11:59:59pm'
$jobs = $servers | % { 
    Invoke-Command -AsJob -ComputerName $_ -ArgumentList $startTime,$endTime -ScriptBlock {
        param($startTime,$endTime) 
        gci "d:\exchange server\logging\ews\*.LOG" | ? {($_.creationTime -lt $endTime) -and ($_.lastWriteTime -gt $startTime)} |
            % { 
                Import-CSV -path $_.FullName | 
                    ? { [int32]($_.TotalRequestTime) -ge 1200000}
            }
    } 
}

#Get HTTP 500 status, just for today: (Where 'today' is 2015-12-28)
[datetime]$startTime = '2015-12-28 12:00:00am'
$jobs = $servers | % { 
    Invoke-Command -AsJob -ComputerName $_ -ArgumentList $startTime -ScriptBlock {
        param($startTime) 
        gci "d:\exchange server\logging\ews\*.LOG" | ? {($_.lastWriteTime -gt $startTime)} |
            % { 
                Import-CSV -path $_.FullName | 
                    where-object -property HttpStatus -eq 500
            }
    } 
}

#How to collect the job data:
wait-job -job $jobs #-Timeout 30 , or however long you're willing to wait for them all to finish

$results = Receive-job $jobs
Remove-Job $jobs  # otherwise they'll hang out holding resources

# So which users got the big errors?
$results | group AuthenticatedUser


# Which server(s) did it happen on?
$results | group AuthenticatedUser, ServerHostName | ft -auto Count, Name

# Other useful fields in the results:
# GenericError, HttpStatus, SoapAction, ThrottlingDelay

# Search a result set containing all events for a period of time for transactions that timed out:
# (Note that we need to force totalRequestTime to integer format)
$timeouts = $results | ? { [int32]($_.TotalRequestTime) -ge 1800000 }
$timeOuts | group SoapAction,UserAgent -NoElement