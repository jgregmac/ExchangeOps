[cmdletBinding()]
param(
)
Set-PSDebug -Strict

if (test-path variable:vDiskID) {remove-variable vDiskID}
if (test-path variable:pDiskID) {remove-variable pDiskID}

function createVDisk {
    param(
        [Parameter(Mandatory=$true)][string]$pDiskID,
        [Parameter(Mandatory=$true)][string]$vDiskID
    )
    & omconfig storage controller controller=1 action=createvdisk raid=r0 size=max stripesize=256kb "pdisk=$pdiskID" "name=VDisk$vDiskID" readpolicy=ara writepolicy=wb diskcachepolicy=disabled
}

### Loop though vDisks to identify missing IDs: ###
:vDisk for ($i=0; $i -le 19; $i++) {
    $result = & omreport storage vdisk "controller=1" "vdisk=$i"
    foreach ($line in $result) {
        if ($line -match 'Invalid vdisk value') {
           $vDiskID = $i
           break :vDisk
        }
    }
}
# Exit if no missing disks found:
if (test-path variable:vDiskID) {
    write-host "Missing vDisk ID is: $vDiskID"
} else {
    write-host "No missing vDisk identified."
    exit
}

### Loop though pDisks to identify disk in "Ready" state: ###
:pDisk for ($i=0; $i -le 1; $i++) {
    for ($j=0; $j -le 11; $j++) {
        $testID = '0:'+$i+':'+$j
        #$OMCmd = 'omreport storage pdisk controller=1 pdisk='+$testID
        $result = & omreport storage pdisk "controller=1" "pdisk=$testID"
        foreach ($line in $result) {
            if ($line -match '^State') {
                if ($line -match 'Ready$') {
                    $pDiskID = $testID
                    break :pDisk
                }
            }
        }
    }
}
# Exit if no ready disks are available:
if (test-path variable:pDiskID) {
    write-host "'Ready' pDisk ID is: $pDiskID"
} else {
    write-host "No 'Ready' pDisk identified."
    exit
}

### Create a new vDisk using the identified ready disk: ###
$createCmd = "omconfig storage controller controller=1 action=createvdisk raid=r0 size=max stripesize=256kb pdisk=$pdiskID name=VDisk$vDiskID readpolicy=ara writepolicy=wb diskcachepolicy=disabled"
write-host "About to run command:"
write-host $createCmd

$done = $false
do {
    $response = Read-host "Do you wish to proceed with creating the vDisk? (YES/NO)"
    switch ($response.ToUpper()) {
        YES     {createVDisk -pDiskID $pDiskID -vDiskID $vDiskID; $done = $true}
        NO      {Write-Host "Exiting the script..."; exit}
        default {write-host "Invalid reponse.  Trying again..."}
    }
} until ($done)

##### Need a procedure to verify that the vDisk was created, and to force a disk rescan so that Windows knows it is avialable ####

### Partition and format the new vDisk ###
# Find the broken mountpoint:
:rPoints foreach ($rPoint in (gci C:\ExchangeVolumes)) {
    $mvOut = & mountvol
    $RegEx = $rPoint.FullName.Replace('\','\\')
    $RegEx = $regex.Insert($RegEx.Length,'\\$')
    $mounted = $false
    foreach ($line in $mvOut) {
        if ($line -match $RegEx) {
            $mounted = $true
        }
    }
    if (-not $mounted) {
        write-host "Broken Mount Point at:" $rPoint.FullName
        break :rPoints
    }
}
# Verify that the mount point is broken:
$rPointBroken = $false
try {
    gci $rPoint.FullName -ea Stop
} catch {
    $rPointBroken = $true
    write-host "This mount point is busted."
}
if (-not $rPointBroken) {write-host "The mount point is not broken.  Exiting..."; exit}

#Remove the old Mount Point:
#Remove-Item $rPoint.FullName -Force -Confirm:$false
& cmd /c rmdir /s /q $rPoint.FullName

#Fetch the Windows disk number for the new vDisk:
$winDiskNum = (omreport storage vdisk controller=1 vdisk=$vDiskID | Select-String -pattern '^Device Name').tostring().split(' ') | select -last 1
#Format a volume label for the new volume:
$label = $rPoint.Name

$disk = Get-disk -number $winDiskNum 
#$disk | clear-disk -removeData -confirm:$false   
$disk | Initialize-Disk -PartitionStyle GPT -Confirm:$false   
$part = $disk | New-Partition -UseMaximumSize  
$part | Format-Volume -FileSystem ReFS -NewFileSystemLabel $label -AllocationUnitSize 64KB -Confirm:$false
$path = Join-Path C:\ExchangeVolumes $label
$mountpoint = New-Item $path -ItemType Directory
$part | Add-PartitionAccessPath -AccessPath $mountpoint.FullName

# get-partition -DiskNumber # -PartitionNumber 2 - has an array attribute "accessPaths", with current mountpoints.

### Enable BitLocker ###
enable-bitlocker $rPoint.FullName -UsedSpaceOnly -RecoveryPasswordProtector -Confirm:$false
Enable-BitLockerAutoUnlock $rPoint.FullName
