#Array of AD groups with OU exclusion
$SECGroup = Get-ADGroup -Filter * -Properties Name, samaccountname, DistinguishedName | Select Name, samaccountname, DistinguishedName | Sort-Object Name | Where {$_.DistinguishedName -notlike <#"*CN=...,DC=...,DC=...*"#>}

#log file creation
$fileDate = (Get-Date -Format "MM-dd-yyyy").ToString()
$filePath = 'C:\This\is\an\example'
$logFile = "$($filepath)\$($filedate)RemovedSecurityGroups_log.txt"
try{
New-Item -Path . -Name $filepath$fileDate"RemovedSecurityGroups_log.txt" -ErrorAction Stop
New-Item -Path $logFile -ErrorAction Stop
}catch [System.IO.IOException]{
Remove-Item $logFile
New-Item -Path $logFile
}

foreach($group in $SECGroup){
    try{
    $memberCount = Get-ADGroupMember -Identity $group.samaccountname | Measure-Object | Select count
    }
    catch [Microsoft.ActiveDirectory.Management.ADException]{
        "GROUP: "+$group.Name
        #Determines if group contains contacts or other object types that cannot be checked for active status
        Write-Host "Contains objects that cannot be assessed. Moving to next object." -ForegroundColor Magenta
        "`r`n"
        continue
    }
    "GROUP: "+$group.Name
    "COUNT: "+$memberCount.Count
    if($memberCount.Count -lt 1){ #Change this to 2 when checking groups that is 1 or less
        Write-Host $group.Name"Group Removed" -ForegroundColor Cyan
        Remove-ADObject -Identity $group.DistinguishedName -Recursive -Confirm:$false
        Add-Content -Path $logfile -Value ("Deleted $($group.Name) ") -NoNewline
        Add-Content -Path $logfile -Value (Get-Date)
    }
    else{
        $i = get-adgroupmember -Identity $group.samaccountname | Select samaccountname
        $c = $p = 0 #Counter to determine if object is user or computer
        foreach($user in $i){
            $tempPC = $temp = $null #Will hold value of user or computer to be assessed
            try{
                $temp = Get-ADUser -Identity $user.SamAccountName -Properties samaccountname, enabled, sAMAccountType, msExchRecipientTypeDetails `
                | Select SamAccountName, enabled, sAMAccountType, msExchRecipientTypeDetails -ErrorAction SilentlyContinue
            }
            catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]{
                try{ #checks if the object is a computer instead
                    $tempPC = Get-ADComputer -Identity $user.samaccountname -Properties samaccountname, enabled, sAMAccountType `
                    | Select SamAccountName, enabled, sAMAccountType -ErrorAction SilentlyContinue
                }
                catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]{}
            }
            #checks if user is disabled & is user account & is not a room resource & is not a sharedMBX
            if($temp.enabled -eq $false -and $temp.sAMAccountType -eq "805306368" -and $temp.msExchRecipientTypeDetails -ne "8589934592" -and $temp.msExchRecipientTypeDetails -ne "34359738368"){
                $c++
            }elseif ($tempPC.enabled -eq $false -and $tempPC.sAMAccountType -eq "805306369"){
                $p++
            }
        }

        if($p -eq 0){ #checks if the group had computers/groups or users and will calculate accordingly
                if($c/$memberCount.Count*100 -eq 100){ #CHANGE BOTH IF STATEMENTS TO -GT 50 TO CHECK GROUP MEMBERSHIP W/HALF OF USERS DISABLED
               "Percentage: "+$c/$memberCount.Count*100
                Write-Host "All users are disabled in this group. Delete this group" -ForegroundColor Yellow
                Remove-ADObject -Identity $group.DistinguishedName -Recursive -Confirm:$false
                Add-Content -Path $logfile -Value ("Deleted $($group.Name) ") -NoNewline
                Add-Content -Path $logfile -Value (Get-Date)
            }else{
                "Percentage: "+$c/$memberCount.Count*100
                Write-Host "At least one member is active. Keep Group." -ForegroundColor Green
            }
        }else{
            if($p/$memberCount.Count*100 -eq 100){
                "Percentage: " + $p/$memberCount.Count*100
                Write-Host "All objects are disabled in this group. Delete this group" -ForegroundColor Yellow
                Remove-ADObject -Identity $group.DistinguishedName -Recursive -Confirm:$false
                Add-Content -Path $logfile -Value ("Deleted $($group.Name) ") -NoNewline
                Add-Content -Path $logfile -Value (Get-Date)
            }else{
                "Percentage: " + $p/$memberCount.Count*100
                Write-Host "At least one object is active. Keep Group." -ForegroundColor Green
            }
        }
    }
   "`r`n"
}