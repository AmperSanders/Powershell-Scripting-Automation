Add-Type -AssemblyName System.Windows.Forms
$DisabledUsersPath = "Disabled OU Location"

#####Window Form#####
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "User Termination"
$mainForm.Height = 280
$mainForm.Width = 241
$mainForm.MaximizeBox = $false
$mainForm.FormBorderStyle = "FixedDialog"
$mainForm.TopMost = $false
$mainForm.StartPosition = 'CenterScreen'
$mainForm.ForeColor = "black"
#$mainForm.BackColor = "#141414"

#####Name#####
$NameLabel = New-Object System.Windows.Forms.Label
$NameLabel.Text = "Enter user's first of last name"
$NameLabel.Location = New-Object System.Drawing.Point(40,20)
$NameLabel.AutoSize = $true
$mainForm.Controls.Add($NameLabel)

$NameTextbox = New-Object System.Windows.Forms.TextBox
$NameTextbox.Location = New-Object System.Drawing.Point(65,40)
$NameTextbox.Size = New-Object System.Drawing.Size(100,20)
$mainForm.Controls.Add($NameTextbox)

#Export Groups
$ExportLabel = New-Object System.Windows.Forms.Label
$ExportLabel.Text = "Export Group Memberships?"
$ExportLabel.Location = New-Object System.Drawing.Point(45,68)
$ExportLabel.AutoSize = $true
$mainForm.Controls.Add($ExportLabel)

$ExportCheckbox = New-Object System.Windows.Forms.CheckBox
$ExportCheckbox.Location = New-Object System.Drawing.Point(30,66)
$ExportCheckbox.Size = New-Object System.Drawing.Size(20,20)
$mainForm.Controls.Add($ExportCheckbox)

#Share Onedrive
$ShareLabel = New-Object System.Windows.Forms.Label
$ShareLabel.Text = "Share Onedrive with Manager?"
$ShareLabel.Location = New-Object System.Drawing.Point(45,95)
$ShareLabel.AutoSize = $true
$mainForm.Controls.Add($ShareLabel)

$OnedriveCheckbox = New-Object System.Windows.Forms.CheckBox
$OnedriveCheckbox.Location = New-Object System.Drawing.Point(30,93)
$OnedriveCheckbox.Size = New-Object System.Drawing.Size(20,20)
$mainForm.Controls.Add($OnedriveCheckbox)

#Additional Onedrive Access
$AdditionalAccessLabel = New-Object System.Windows.Forms.Label
$AdditionalAccessLabel.Text = "Share OneDrive to additional user?"
$AdditionalAccessLabel.Location = New-Object System.Drawing.Point(30,125)
$AdditionalAccessLabel.AutoSize = $true
$mainForm.Controls.Add($AdditionalAccessLabel)

$AdditionalNameTextbox = New-Object System.Windows.Forms.TextBox
$AdditionalNameTextbox.Location = New-Object System.Drawing.Point(65,145)
$AdditionalNameTextbox.Size = New-Object System.Drawing.Size(100,20)
$mainForm.Controls.Add($AdditionalNameTextbox)

#####Confirmation#####
$TermBtn = New-Object System.Windows.Forms.Button
$TermBtn.Location = New-Object System.Drawing.Point(20,180)
$TermBtn.Text = "Terminate"
#$TermBtn.BackColor = "#a10000"
$mainForm.Controls.Add($TermBtn)

####Close#####
$cancelBtn = New-Object System.Windows.Forms.Button
$cancelBtn.Location = New-Object System.Drawing.Point(125,180)
$cancelBtn.Text = "Cancel"
#$cancelBtn.BackColor = "#a10000"
$mainForm.Controls.Add($cancelBtn)

$cancelBtn.add_click({
$mainForm.Close()
})

function ClearForm{
    $NameTextbox.Text = ""
    $ExportCheckbox.Checked = $false
    $OnedriveCheckbox.Checked = $false
    $AdditionalNameTextbox.Text = ""
}

$TermBtn.add_click({
    if($NameTextbox.Text -eq ""){
        [System.Windows.Forms.MessageBox]::show("Please enter terminated user's name", "Invalid Values", "OK", "Question")
        return
    }else {
        try{
        $user = Get-ADUser -Filter "mail -like '*$($NameTextbox.Text)*'" -Server "server" -Properties Name, displayName, samaccountname, mail, Manager, DistinguishedName, MemberOf | Select-Object Name, displayName, samAccountName, mail, Manager, DistinguishedName, MemberOf | Out-GridView -PassThru
        Get-ADUser -Identity $user.samAccountName
        }catch{
            [System.Windows.Forms.MessageBox]::show("User not found. Try again.", "Invalid Values", "OK", "Question")
            return
        }
        $NameTextbox.Text = $($user.displayName).toString()
        $Result = [System.Windows.Forms.MessageBox]::show("Are you sure you want to terminate "+$user.displayName+"?", "Confirm Termination", "YesNo", "Question")
        if(!$Result){

        }else{
            $currentUser = whoami | %{$_.remove(0,4)}

            if($ExportCheckbox.Checked -eq $false){
                #do nothing
            }else{
                Write-Output $edited | Out-File "C:\Users\$($currentUser)\Downloads\$($NameTextbox.Text).csv"
                [System.Windows.Forms.MessageBox]::show("Group Export csv placed in Downloads folder", "Group Exportation", "OK", "Question")
            }

            if($OnedriveCheckbox.Checked -eq $false){
                #do nothing
            }else{
                Connect-SPOService -Url "https://company-admin.sharepoint.com/"
                $userSite = "https://company-my.sharepoint.com/personal/"+$($user.mail.Replace(".","_").Replace("@","_"))
                $termedUser = (Get-Culture).TextInfo.ToTitleCase($userSite.Split('/')[4].Replace("_company_com", "").Replace("_"," "))
                $delegatedUser =  Get-ADUser -Identity $user.manager -Properties mail | select mail

                if($AdditionalNameTextbox.Text -ne ""){
                    try{
                        Set-SPOUser -Site $userSite -LoginName $delegatedUser.mail -IsSiteCollectionAdmin $true
                        Set-SPOUser -Site $userSite -LoginName $AdditionalNameTextbox.Text -IsSiteCollectionAdmin $true

                        Send-MailMessage -SmtpServer "internal" -Port 25 -Subject "Terminated User Onedrive Access" -From "internal_email" -To $delegatedUser.mail -Cc $AdditionalNameTextbox.Text -BodyAsHtml `
                        -Body ("The user:<i style='color:blue'> $termedUser </i>is no longer working at company and HR has requested that their onedrive contents be shared." + "`r`n" + `
                        "<br>Please click the <a href='$userSite'> link </a> and download the needed as soon as possible.<br><br>")
                        [System.Windows.Forms.MessageBox]::show("OneDrive invitation link shared with requested user(s)", "OneDrive Shared", "OK", "Question")            
    
                    }catch{
                        [System.Windows.Forms.MessageBox]::show("User has no OneDrive content/data", "No OneDrive", "OK", "Question")
                    }
                }else{
                    try{
                        Set-SPOUser -Site $userSite -LoginName $delegatedUser.mail -IsSiteCollectionAdmin $true

                        Send-MailMessage -SmtpServer "internal" -Port 25 -Subject "Terminated User Onedrive Access" -From "internal_email" -To $delegatedUser.mail -BodyAsHtml `
                        -Body ("The user:<i style='color:blue'> $termedUser </i>is no longer working at company and HR has requested that their onedrive contents be shared." + "`r`n" + `
                        "<br>Please click the <a href='$userSite'> link </a> and download the needed as soon as possible.<br><br>")
                        [System.Windows.Forms.MessageBox]::show("OneDrive invitation link shared with requested user(s)", "OneDrive Shared", "OK", "Question")                
                    }catch{
                        [System.Windows.Forms.MessageBox]::show("User has no OneDrive content/data", "No OneDrive", "OK", "Question")
                    }
                }
            }
        }

        foreach($group in $user.memberOf){
            Remove-ADPrincipalGroupMembership -Identity $user.samAccountName -MemberOf $group -Confirm: $false
        }

        Disable-ADAccount -Identity $user.samAccountName   
        Move-ADObject -Identity $user.DistinguishedName -TargetPath $DisabledUsersPath
        
        [System.Windows.Forms.MessageBox]::show("User terminated.", "Complete", "OK", "Question")
        ClearForm           
    }
})
$mainForm.ShowDialog()
$mainForm.Dispose()