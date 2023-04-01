$computers = Get-Content .\computers.txt

#deletes all versions of POW from computers
foreach ($computer in $computers) {
    
    
    $pow_public_items = Get-ChildItem "\\$computer\C$\Users\Public\Desktop" -Filter "POW*" -Recurse | ForEach-Object { $_.FullName }
    Remove-Item $pow_public_items -Recurse
    
    
    $pow_default_items = Get-ChildItem "\\$computer\C$\Users\Default\Desktop" -Filter POW* -Recurse | ForEach-Object { $_.FullName }
    Remove-Item $pow_default_items -Recurse


}

#Adds shortcuts and local files for POW icon to computers exports failure log if fails
foreach($computer in $computers) {
    
    
    $public_desktop = "\\$computer\C$\Users\Public\Desktop"
    $new_pow_icon = "\\$computer\C$\Users\Public\Desktop\POW.lnk"
    $pow_home_local = "\\$computer\C$\POW"
    $check_pow_local = Get-ChildItem "\\$computer\C$" -Name 
    $pow_icon_exists = Get-ChildItem "\\$computer\C$\" | Where-Object {$_.Name -Match "POW"}
    
    if ($check_pow_local -notcontains "POW") {
        Write-Host "Adding POW directory to local machine" -ForegroundColor Yellow
        New-Item -Path "\\$computer\C$\POW" -ItemType Directory
        Write-Host "POW Directory added" -ForegroundColor Green
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
        $log_file = ".\logs\failurelogs.txt"
        "$computer failed on $date" | Out-File $log_file -Append
        Write-Host "$computer failed to add icon" -ForegroundColor Red
    }
    else {Write-Host "POW icon already on Public Desktop" -ForegroundColor Green}


}
