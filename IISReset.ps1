[array]$servers = "msx-mh01","msx-mh02","msx-mh03","msx-mh04","msx-mh05","msx-mh06","msx-tp01","msx-tp02","msx-tp03","msx-tp04","msx-tp05","msx-tp06"

#Step through each server in the array and perform an IISRESET

foreach ($server in $servers)
{
    Write-Host "Resetting IIS on $server..."
    IISRESET $server 
}
Write-Host IIS has been reset on all servers.