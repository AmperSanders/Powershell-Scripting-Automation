$userName = Read-Host "Please enter a username"
$user = Get-ADUser -Identity $userName -Properties samAccountName, MemberOf, DistinguishedName
$edited = $user.MemberOf -replace ",.*" -replace "CN="
"`r`n"
Write-Host $userName "belongs to these security groups:" -ForegroundColor Yellow
Write-Output $edited
"`r`n"
Write-Host "If nothing is returned then the username may be incorrect or the user does not exist in Active Directory." -ForegroundColor Red
################# Export CSV as backup #########################
$Export = Read-Host "Export results to a CSV file? (y/n)"
if ($Export -eq 'y'){
    Write-Output $edited | Out-File .\DisabledUserLogs\$userName"_Exported_Groups".csv
}
else{}
################# Remove from groups ###########################
$removeChoice = Read-Host "Remove user from groups? (y/n)"
if ($removeChoice -eq 'y'){
    $ADgroups = Get-ADPrincipalGroupMembership -Identity $userName | Where-Object {$_.Name -ne "Domain Users"}
    Remove-ADPrincipalGroupMembership -Identity $userName -MemberOf $ADgroups -Confirm: $false
}
else{}
################# Disable user account #########################
$disableChoice = Read-Host "Disable this user's account? (y/n)"
if ($disableChoice -eq 'y') {
    Disable-ADAccount -Identity $userName   
}
else{}
################ Move user to NA Disabled ######################
$DisabledUsersOUPath = "OU=...,OU=...,DC=..,DC=..."
$moveChoice = Read-Host "Move user to NA Disabled OU? (y/n)"
if ($moveChoice -eq 'y'){
    Move-ADObject -Identity $user.DistinguishedName -TargetPath $DisabledUsersOUPath 
}
else{}
Read-Host "Press enter to exit"