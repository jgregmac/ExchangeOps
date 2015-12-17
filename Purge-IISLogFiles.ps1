<#
.SYNOPSIS
Deletes IIS Log files greater older than a user-specified number of days.  30 days is used by default.

.DESCRIPTION
The description is usually a longer, more detailed explanation of what the script or function does. Take as many lines as you need.

.PARAMETER RetainDays
Sets the number of days for which you wish to retain your IIS Logs.  Any logs older than this number of days will be permanently deleted.


.EXAMPLE
Purge-IISLogFiles -RetainDays 15
Will delete all IIS Logs older than 15 days.
#>
param(
    [Parameter(Mandatory=$false)][ValidateRange(1,365)][int32]$RetainDays = 30
)
Set-PSDebug -Strict

$purgeTime = [datetime]::now.AddDays(-$RetainDays)

try {
    import-module webAdministration -ErrorAction Stop
} catch {
    [string]$out = "Could not load the required webAdministration PowerShell module."
    Write-Host $out -ForegroundColor Red
    return 100
}

[array]$dirs = @()                   # Initialize array for IIS directories to purge.
[array]$sites = gci 'IIS:\Sites'     # Get an array of sites

foreach ($site in $sites) {
    [string]$dir =  $site.logFile.directory
    #When the dirctory path has an old-school environment variable "%something%", we need to expand it and sub in the real path:
    if ($dir -match '(?<envVar>%.+%)') {
        [string]$sub = [System.Environment]::ExpandEnvironmentVariables($matches['envVar']) 
        $dir = $dir.Replace($matches['envVar'],$sub)
    }
    $dirs += $dir
}

$dirs = $dirs | sort -Unique         # Remove any identical directory entries.

[int]$rc = 0
[string[]]$failedFiles = @()

foreach ($dir in $dirs) {
    write-host "Checking for old log files in $dir" -ForegroundColor Cyan
    if (test-path $dir) {
        [array]$purgeFiles = gci -LiteralPath $dir -Recurse -Filter *.log | ? {$_.LastWriteTime -lt $purgeTime}
        if ($purgeFiles.count -gt 0) {
            [string]$out = '    Purging ' + $purgeFiles.count + ' log files from ' + $dir
            write-host $out -ForegroundColor Yellow
            $purgeFiles | % {
                $file = $_
                $out = "        Deleting: " + $file.FullName
                write-host $out 
                try { 
                    Remove-Item -LiteralPath $file.FullName -Force -Confirm:$false -ea Stop | Out-Null
                } catch {
                    $out = "          Failed to delete: " + $File.FullName
                    write-host $out -ForegroundColor Yellow
                    $failedFiles += $file.FullName
                    $rc = 210
                }
            }
        } else {
            [string]$out = "    No files to purge from $dir"
            write-host $out -ForegroundColor Gray
        }
    } else {
        [string]$out = "Could not resolve the IIS logging filesystem path '$dir' specified in the IIS server configuration."
        Write-Host $out -ForegroundColor Red
        return 200
    }
}

if ($failedFiles.count -gt 0) {
    $out = "List of files that could not be deleted:"
    write-host $out -ForegroundColor Cyan
    foreach ($fail in $failedFiles) {
        write-host "    $fail" -ForegroundColor Yellow
    }
}

Return $rc