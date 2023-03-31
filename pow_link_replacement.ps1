$computers = Get-Content .\computers.txt

foreach($computer in $computers) {
    $public_desktop = "\\$computer\C$\Users\Public\Desktop"
    $check_pow_local = Get-ChildItem "\\$computer\C$" -Name 
    if ($check_pow_local -notcontains "POW") {New-Item -Path "\\$computer\C$\POW" -ItemType Directory}

    $pow_home_local = "\\$computer\C$\POW"

    Copy-Item -Path .\pow_img.ico -Destination $pow_home_local -Force

    $wshshell = New-Object -comObject WScript.Shell 
            $shortcut = $wshshell.CreateShortcut("$public_desktop\POW.lnk")
            $shortcut.TargetPath = "https://pow-prd.nordstromaws.app/"
            $shortcut.IconLocation = "\\$computer\C$\POW\pow_img.ico"
            $shortcut.Save()


}
