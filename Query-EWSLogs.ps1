#This is not a real script.  It is a transcript of us trying to analyze the EWS logs.
# NOTE:  This probably is not the best way to analyze loads of log data!

$servers = Get-Mailboxserver | where-object -property Name -notin ("msx-mgt1", "msx-mgt2")
# In 2016, everyone's a mailbox server I guess...

#Example of querying all "ErrorServerBusy" events on all servers for as far back as the logs go:

$jobs = $servers | % { 
    Invoke-Command -AsJob -ComputerName $_ -ScriptBlock { 
        gci "d:\exchange server\logging\ews\*.LOG" | 
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