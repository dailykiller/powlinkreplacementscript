$comp = "W-fhdt2t3"

#$pow_items = Select-Object {$_.FullName}
#Get-ChildItem "\\$comp\C$\Users\Public\Desktop" | Where-Object {$_.Name -eq "POW"} | ForEach-Object {Write-Host $_.FullName}

#$pow_public_items = Get-ChildItem "\\$comp\C$\Users\Public\Desktop" -Filter POW* -Recurse | ForEach-Object { $_.FullName }
#$pow_default_items = Get-ChildItem "\\$comp\C$\Users\Default\Desktop" -Filter POW* -Recurse | ForEach-Object { $_.FullName }

#foreach ($item in $pow_public_items) {Remove-Item $item}
#foreach ($item in $pow_default_items) {Remove-Item $item}

#Remove-Item $pow_public_items -Recurse

$pow_public_items = Get-ChildItem "\\$comp\C$\Users\Public\Desktop" -Filter POW* -Recurse | ForEach-Object { $_.FullName }
Remove-Item $pow_public_items -Recurse