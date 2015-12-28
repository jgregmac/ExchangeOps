param(
    [Parameter(Mandatory=$false)]
    [string]$log = 'c:\local\temp\Hide-COMUsers.log'
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

$startTime = Get-Date
writeHostAndLog -out "Script Started: $startTime" -Color Cyan

if (test-path -LiteralPath $log) {Remove-Item -LiteralPath $log -Force -Confirm:$false}

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

writeHostAndLog -out "Script ended: $endTime"
writeHostAndLog -out " "
$elapsed = $endTime - $startTime
writeHostAndLog -out ("Elapsed Time: " + $elapsed.TotalSeconds + " Seconds")