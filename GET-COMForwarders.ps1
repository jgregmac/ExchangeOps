param(
    [Parameter(Mandatory=$false)]
      [string]$log = 'c:\local\temp\Get-ComForwarders.log',
    [Parameter(Mandatory=$false)]
      [string]$searchList = '\\files\shared\saa\Exchange\temp\COM_NURSING_GAL.txt',
    [Parameter(Mandatory=$false)]
      [string]$outList = '\\files\shared\saa\Exchange\temp\mailbox_enabled_forwarders.txt'
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
if (test-path -LiteralPath $log) {Remove-Item -LiteralPath $log -Force -Confirm:$false}

try {
    $re = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://msx-mh06.campus.ad.uvm.edu/powershell" -ea Stop -wa SilentlyContinue
    Import-PSSession $re -ea Stop -wa SilentlyContinue
    # Import-Module ActiveDirectory -ea Stop
} catch {
    writeHostAndLog -out "Could not initialize the PowerShell environment." -color Red
    return 100
}


try {
    [string[]]$searchUsers = Get-Content -Path $searchList -ea Stop
} catch {
    writeHostAndLog -out "Could not load list of users to process from $searchList" -color Red
    return 110
}

[string[]]$redirUsers =  @()
[string[]]$fwdUsers   =  @()
[string[]]$unprovUsers = @()

forEach ($user in $searchUsers) {  
    # [string]$sam = $_.samAccountName;
    # Get-InboxRule -mailbox $_.DistinguishedName -ea Stop -wa SilentlyContinue |

    try {
        $rules = @()
        $rules = Get-InboxRule -mailbox $user -ea Stop -wa SilentlyContinue 
    } catch {
        writeHostAndLog -out ("    Could not get inbox rules for user: $user") -color Yellow
        $unprovUsers += $user
        continue
    }
    forEach ($rule in $rules) {
        if ($rule.RedirectTo -like "*@med.uvm.edu*") {
            [string]$out = $user + ',' + ($rule.RedirectTo[0].Split('"') | select -index 1)
            writeHostAndLog -out "    Found redirected user: $out" -Color Gray 
            $redirUsers += $out
        }
        if ($rule.ForwardTo -like "*@med.uvm.edu*") {
            [string]$out = $user + ',' + ($rule.RedirectTo[0].Split('"') | select -index 1)
            writeHostAndLog -out "    Found forwarded user: $out" -Color Gray
            $fwdUsers += $user
        }
    }
}
writeHostAndLog -out " "
writeHostAndLog -out ("Count of unprovisioned users: " + $unprovUsers.count) -color Cyan 
writeHostAndLog -out ("Count of forwarding users: " + $fwdUsers.count) -color Cyan
writeHostAndLog -out ("Count of redirected users: " + $redirUsers.count) -Color Cyan
writeHostAndLog -Out " "

if (test-path $outList) {Remove-Item -Path $outList -Force -Confirm:$false}
writeHostAndLog -Out ("Writing out currently forwarding users to: " + $outList) -Color Cyan
$redirUsers | Out-File -FilePath $outList -Append
$fwdUsers | Out-File -FilePath $outList -Append

writeHostAndLog -Out " "
showElapsedTime -startTime $startTime

get-pssession | remove-pssession