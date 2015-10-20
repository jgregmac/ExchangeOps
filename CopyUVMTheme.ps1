[array]$servers = "msx-mh01","msx-mh02","msx-mh03","msx-mh04","msx-mh05","msx-mh06","msx-tp01","msx-tp02","msx-tp03","msx-tp04","msx-tp05","msx-tp06"

# This is the directory you want to copy to the computer (IE. c:\folder_to_be_copied)
$source = "\\files\shared\saa\Exchange\Skinning OWA\UVM custom theme\"

# On the desination computer, where do you want the folder to be copied?
$dest = "d$\Exchange Server\V15\ClientAccess\Owa\prem\15.0.1130.7\"

foreach ($server in $servers) {
    if (test-Connection -Cn $server -quiet) {
        Copy-Item $source -Destination \\$server\$dest -Recurse -Force
    } else {
        "$server is not online"
    }

}
