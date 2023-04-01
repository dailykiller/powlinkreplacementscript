$computer = "W-FHDT2T3"


$pow_public_items = Get-ChildItem "\\$computer\C$\Users\Public\Desktop" -Filter POW* -Recurse | ForEach-Object { $_.FullName }
    Remove-Item $pow_public_items -Recurse -ErrorAction Ignore
    
    
$pow_default_items = Get-ChildItem "\\$computer\C$\Users\Default\Desktop" -Filter POW* -Recurse | ForEach-Object { $_.FullName }
Remove-Item $pow_default_items -Recurse -ErrorAction Ignore


$pow_users_items = Get-ChildItem "\\$computer\C$\Users" -Filter POW* -Recurse | ForEach-Object {$_.Fullname}
Remove-Item $pow_users_items -Recurse