[CmdletBinding()]
Param (
 [string]$Identity,
 [switch]$enable
 
)

echo "I want to .."

if ($enable) {
    echo ENABLE
    } else {
    echo DISABLE
    }


echo "CAS functionality for $Identity"

if (!$enable) {
    $enable = $False
    }



set-casmailbox $Identity  -ActiveSyncEnabled:$enable `
-ECPEnabled:$enable `
-EwsAllowEntourage:$enable `
-EwsAllowMacOutlook:$enable `
-EwsAllowOutlook:$enable `
-MAPIEnabled:$enable `
-OWAEnabled:$enable `
-OWAforDevicesEnabled:$enable `
-PopEnabled:$enable

