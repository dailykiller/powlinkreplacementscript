$computers = Get-Content .\computers.txt



foreach($computer in $computers) {
    $public_desktop = "\\$computer\C$\Users\Public\Desktop"
    $new_pow_icon = "\\$computer\C$\Users\Public\Desktop\POW.lnk"
    $new_pow_icon_exists = Test-Path -Path $new_pow_icon
    $existing_pow_shortcut = Get-ChildItem "\\$computer\C$\Users\Public\Desktop" | Where-Object {$_.Name -Match "POW.url"}

    $check_pow_local = Get-ChildItem "\\$computer\C$" -Name 
    if ($check_pow_local -notcontains "POW") {New-Item -Path "\\$computer\C$\POW" -ItemType Directory}

    if ($existing_pow_shortcut) {Remove-Item "\\$computer\C$\Users\Public\Desktop\POW.url"}

    $pow_home_local = "\\$computer\C$\POW"

    Copy-Item -Path .\pow_img.ico -Destination $pow_home_local -Force
    
    if (!$new_pow_icon_exists) {
    $wshshell = New-Object -comObject WScript.Shell 
            $shortcut = $wshshell.CreateShortcut("$public_desktop\POW.lnk")
            $shortcut.TargetPath = "https://pow-prd.nordstromaws.app/"
            $shortcut.IconLocation = "\\$computer\C$\POW\pow_img.ico"
            $shortcut.Save()
    }

}
