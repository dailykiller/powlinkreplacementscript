$computers = Get-Content .\computers.txt
$error_log_file = ".\logs\error_logs.txt"
$date = Get-Date


#deletes all versions of POW from computers
foreach ($computer in $computers) {
    $connection_test = Test-Connection $computer -Count 1 -Quiet 
    if ($connection_test) {  
        $user_profiles = Get-ChildItem "\\$computer\C$\Users" | ForEach-Object {$_.Fullname}
        foreach ($user_profile in $user_profiles) { 
            $user_profile = "$user_profile\Desktop"
            $pow_users_items = Get-ChildItem "$user_profile" -Filter POW* -Recurse -ErrorVariable error_log | ForEach-Object { $_.FullName }
            $error_log | Out-File $error_log_file -Append
            if (!$null -eq $pow_users_items) {
                Remove-Item $pow_users_items 
                }
            }
        

        $pow_default_items = Get-ChildItem "\\$computer\C$\Users\Default\Desktop" -Filter POW* -Recurse | ForEach-Object { $_.FullName }
        if (!$null -eq $pow_default_items) {
        Remove-Item $pow_default_items -Recurse
        }
    }
}


#Adds shortcuts and local files for POW icon to computers exports failure log if fails
foreach($computer in $computers) {
    $connection_test = Test-Connection $computer -Count 1 -Quiet 
    if ($connection_test) {
    
        $public_desktop = "\\$computer\C$\Users\Public\Desktop"
        $new_pow_icon = "\\$computer\C$\Users\Public\Desktop\POW.lnk"
        $pow_home_local = "\\$computer\C$\POW"
        $check_pow_local = Get-ChildItem "\\$computer\C$" -Name 
        
        
        if ($check_pow_local -notcontains "POW") {
            Write-Host "Adding POW directory to local machine" -ForegroundColor Yellow
            New-Item -Path "\\$computer\C$\POW" -ItemType Directory
            Write-Host "POW Directory added" -ForegroundColor Green
        }

        $pow_icon_exists = Get-ChildItem "\\$computer\C$\POW" | Where-Object {$_.Name -Match "POW"}
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
            $log_file = ".\logs\logs.txt"
            "$computer failed on $date" | Out-File $log_file -Append
            Write-Host "$computer failed to add icon" -ForegroundColor Red
            
            if ($env:COMPUTERNAME -eq $computer) {
                $date = Get-Date
                $log_file = ".\logs\logs.txt"
                "$computer is local computer, functionality of this script is not supported yet. Failed on $date" | Out-File $log_file -Append
                Write-Host "$computer added icon" -ForegroundColor Green
            }
        }
        else {
            $date = Get-Date
            $log_file = ".\logs\logs.txt"
            "$computer succeeded on $date" | Out-File $log_file -Append
            Write-Host "$computer added icon" -ForegroundColor Green
        }
    }


    else {
        $date = Get-Date
        $log_file = ".\logs\logs.txt"
        "$computer couldn't connect on $date" | Out-File $log_file -Append
        Write-Host "$computer failed to connect to $computer" -ForegroundColor Red
    }
    

}
