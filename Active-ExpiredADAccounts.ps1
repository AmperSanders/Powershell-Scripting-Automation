#Documentation: http://www.rlmueller.net/AccountExpires.htm 
#https://learn.microsoft.com/en-us/windows/win32/adschema/a-accountexpires

#Expired-attribute:     132096240000000000 -- Most expired accounts begin with 13
#Not-expired-attribute: 9223372036854775807
$SearchBase = "OU=This is an Example,DC=...,DC=..."
$currentDate = Get-Date
$DisabledUsersOUPath = "OU=Another example,OU=...,DC=...,DC=..."
$enabledExpired = Get-ADUser -Filter {Enabled -eq $true} -properties accountExpires, AccountExpirationDate -SearchBase $SearchBase | Select-Object Name, samAccountName, Enabled, accountExpires, AccountExpirationDate
New-Item -Path .\Active-Expired_Exported_Groups.txt
foreach($user in $enabledExpired){
    if($user.accountExpires -like "13*" -and $user.AccountExpirationDate -lt $currentDate){
        #Confirm
        Write-Host $user.samAccountName "removed"
        Add-Content -Path .\Active-Expired_Exported_Groups.txt -Value $user.samAccountName

        #Remove user from groups
        $ADgroups = Get-ADPrincipalGroupMembership -Identity $user.samAccountName | Where-Object {$_.Name -ne "Domain Users"}
        Remove-ADPrincipalGroupMembership -Identity $user.samAccountName -MemberOf $ADgroups -Confirm: $false

        #Disable the user
        Disable-ADAccount -Identity $user.samAccountName

        #Move user to  Disabled OU
        Move-ADObject -Identity $user.DistinguishedName -TargetPath $DisabledUsersOUPath
    }
    else{}
}