$computers = Get-Content .\computers.txt

foreach($computer in $computers) {
    #$public_desktop = "\\$computer\C$\Users\Pubic\Desktop"
    $check_pow_local = Get-ChildItem "\\$computer\C$" -Name | Where-Object {$_.Name -Match "POW"}
    if ($check_pow_local) {New-Item -Path "\\$computer\C$\POW" -ItemType Directory}

    $pow_home_local = "\\$computer\C$\POW"

    
    Copy-Item -Path .\pow_img.ico -Destination $pow_home_local -Force
    $wshshell = New-Object -comObject WScript.Shell 
            $shortcut = $wshshell.CreateShortcut(".\POW.lnk")
            $shortcut.TargetPath = "https://pow-prd.nordstromaws.app/"
            $shortcut.IconLocation = "\\$computer\C$\POW\pow_img.ico"
            $shortcut.Save()
}
