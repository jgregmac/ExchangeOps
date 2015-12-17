<#
    Set-IISLogOptions.ps1
    Sets standard logging options for all sites on the local IIS server.

    Returns:
      0   - Script ran successfully
      100 - Could not load WebAdministration PowerShell module.
      200 - Could not enumerate IIS Sites
      201 - Could not set logExtFileFlags on site.
      202 - Could not set logFormat on site.
#>

#$flags represents the standard set of logging options to be used on the server.
[string]$flags = "Date,Time,ClientIP,UserName,SiteName,ServerIP,Method," +
    "UriStem,UriQuery,HttpStatus,Win32Status,TimeTaken,ServerPort,UserAgent," +
    "Cookie,Referer,HttpSubStatus"
[string]$format = "W3C"

try {
    import-module webAdministration -ErrorAction Stop
} catch {
    [string]$out = "Could not load the required webAdministration PowerShell module."
    Write-Host $out 
    return 100
}
$sites = gci 'IIS:\Sites' 

foreach ($site in $sites) {
    [string]$path = $site.psPath
    [string]$name = $site.psChildName
    write-host Checking site: $name -ForegroundColor Cyan
    try {
        $logOptions = Get-ItemProperty -Path $path -Name logFile -ErrorAction Stop
    } catch {
        [string]$out = "Could not get logging options for for site: $name.  Exiting."
        Write-Host $out -ForegroundColor Red
        return 200
    }
    [string]$oldFlags = $logOptions.logExtFileFlags
    [string]$oldFormat = $logOptions.logFormat
    
    #write-host "  Current logging flags: $oldFlags"
    if ($flags -eq $oldFlags) {
        write-host "    LogExtFileFlags are in compliance.  No action needed." -ForegroundColor Gray
    } else {
        write-host "    LogExtFileFlags are out of compliance.  Setting new values..." -ForegroundColor Yellow
        try {
            Set-ItemProperty -Path $path -Name logFile -Value @{logExtFileFlags=$flags} -ErrorAction Stop
        } catch {
            [string]$out = "Could not update the IIS logExtFileFlags.  Exiting."
            Write-Host $out -ForegroundColor Red
            return 201
        }
    } 
    #write-host "  Current log format: $oldFormat"
    if ($format -eq $oldFormat) {
        write-host "    Log format is in compliance.  No action needed." -ForegroundColor Gray
    } else {
        write-host "    Log format is out of compliance.  Setting new values..." -ForegroundColor Yellow
        try {
            Set-ItemProperty -Path $path -Name logFile -Value @{logFormat=$format} -ErrorAction Stop
        } catch {
            [string]$out = "    Could not update the IIS log file format.  Exiting."
            Write-Host $out -ForegroundColor Red
            return 202
        }
    } 
}
return 0