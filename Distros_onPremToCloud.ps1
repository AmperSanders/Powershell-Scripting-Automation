﻿<#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://ONPREM_EXCHANGE_SERVER/PowerShell/" -Name onprem
Import-PSSession $Session

Connect-ExchangeOnline -Prefix cloud#>

#Move distribution groups to unsync
foreach($item in Get-Content -Path ".\Desktop\Batch_DistrosToBeMigrated.txt"){
    Write-Host $item -ForegroundColor Cyan
    $OnPremGroup = Get-DistributionGroup -Identity $item | select *
    Move-ADObject -Identity $OnPremGroup.DistinguishedName -TargetPath #unsynced OU to separate groups for migration
}

foreach($item in Get-Content -Path ".\Desktop\Batch_DistrosToBeMigrated.txt"){
    Write-Host $item -ForegroundColor Cyan

    $OnPremGroup = Get-DistributionGroup -Identity $item | select *
    if($OnPremGroup.GroupType.Length -le 9){
        $type = "Distribution"
    }else{
        $type = "Security"
    }
    if($OnPremGroup.Name -like "*–*"){
        $GroupName = $OnPremGroup.Name.Replace("–","-")

    }else{
        $GroupName = $OnPremGroup.DisplayName
    }
    if([string]$OnPremGroup.ManagedBy -eq '' -or [string]$OnPremGroup.ManagedBy -like "*Disabled*" -or [string]$OnPremGroup.ManagedBy -like "*9*" `
    -or [string]$OnPremGroup.ManagedBy -like "*Group*"){
        $owner = "USE GENERIC EMAIL"
    }else{
        $a = [string]$OnPremGroup.ManagedBy[0]
        $n = $a.Split("/")[5]
        $email = $n.Split(",")[1]+"."+$n.Split(",")[0]+"@domain.com"
        $owner = $email.Replace(" ","")
    }

    #create cloud distro mirroring onprem properties
    try{
        New-cloudDistributionGroup -Name $GroupName -Alias $OnPremGroup.Alias -DisplayName $GroupName -ManagedBy $owner `
        -PrimarySmtpAddress $OnPremGroup.PrimarySmtpAddress -Type $type -Description $OnPremGroup.info -ErrorAction Stop
    }catch{
        #checks to see if owner's name was returned in different format like sameaccountname instead of email address
        $owner = Get-ADUser -Identity $n -Properties * | select mail
        try{
            New-cloudDistributionGroup -Name $GroupName -Alias $OnPremGroup.Alias -DisplayName $GroupName -ManagedBy $owner.mail `
            -PrimarySmtpAddress $OnPremGroup.PrimarySmtpAddress -Type $type -Description $OnPremGroup.info -ErrorAction Stop
        } catch{
            New-cloudDistributionGroup -Name $GroupName -Alias $OnPremGroup.Alias -DisplayName $GroupName -ManagedBy "GENERIC EMAIL" `
            -PrimarySmtpAddress $OnPremGroup.PrimarySmtpAddress -Type $type -Description $OnPremGroup.info -ErrorAction Stop     
        }
    }
    
    Start-Sleep -Seconds 5
    Get-cloudDistributionGroup -Identity $GroupName | Set-cloudDistributionGroup -RequireSenderAuthenticationEnabled $false
    #Add Members back    
    foreach($user in Get-DistributionGroupMember -Identity $OnPremGroup.Name | select Name, DistinguishedName){
        $email = Get-ADUser -Identity $user.DistinguishedName -Properties * | select mail
        Add-cloudDistributionGroupMember -Identity $GroupName -Member $email.mail -BypassSecurityGroupManagerCheck
    }#>
}