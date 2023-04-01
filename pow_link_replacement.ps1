$computers = Get-Content .\computers.txt



foreach($computer in $computers) {
    $public_desktop = "\\$computer\C$\Users\Public\Desktop"
    $new_pow_icon = "\\$computer\C$\Users\Public\Desktop\POW.lnk"
    $existing_pow_shortcut = Get-ChildItem "\\$computer\C$\Users\Public\Desktop" | Where-Object {$_.Name -Match "POW.url"}
    $pow_home_local = "\\$computer\C$\POW"
    $check_pow_local = Get-ChildItem "\\$computer\C$" -Name 
    $pow_icon_exists = Get-ChildItem "\\$computer\C$\" | Where-Object {$_.Name -Match "POW"}

    
    
    if ($check_pow_local -notcontains "POW") {
        Write-Host "Adding POW directory to local machine" -ForegroundColor Yellow
        New-Item -Path "\\$computer\C$\POW" -ItemType Directory
        Write-Host "POW Directory added" -ForegroundColor Green
    }


    if ($existing_pow_shortcut) {
        Write-Host "Deleting old POW shortcut" -ForegroundColor Yellow
        Remove-Item "\\$computer\C$\Users\Public\Desktop\POW.url"
        Write-Host "Old POW shortcut deleted" -ForegroundColor Green
    }
    

    if (!$pow_icon_exists) {
        Write-Host "Copying POW icon to local machine" -ForegroundColor Yellow
        Copy-Item -Path .\pow_img.ico -Destination $pow_home_local -Force
        Write-Host "Icon Copied" -ForegroundColor Green
    }
    

    $new_pow_icon_exists = Test-Path -Path $new_pow_icon
    if (!$new_pow_icon_exists) {
        $wshshell = New-Object -comObject WScript.Shell 
                $shortcut = $wshshell.CreateShortcut("$public_desktop\POW.lnk")
                $shortcut.TargetPath = "https://pow-prd.nordstromaws.app/"
                $shortcut.IconLocation = "\\$computer\C$\POW\pow_img.ico"
                $shortcut.Save()
    }   


    $new_pow_icon_exists = Test-Path -Path $new_pow_icon
    if (!$new_pow_icon_exists) {
        $date = Get-Date
        $log_file = ".\logs\failurelogs.csv"
        "$computer failed on $date" | Out-File $log_file -Append
        Write-Host "$computer failed to add icon" -ForegroundColor Red
    }
    else {Write-Host "POW icon already on Public Desktop" -ForegroundColor Green}


}
