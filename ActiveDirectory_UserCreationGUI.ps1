﻿Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope Process -Confirm:$false

Add-Type -AssemblyName System.Windows.Forms
$Timer = New-Object System.Windows.Forms.Timer

#Global Variables
$domainController = "PRIMARY_DOMAIN_CONTROLLER"
$newUserPassword = "STANDARD_PASSWORD_POLICY"
$sendFrom = "FOR_AUTOMATED_SEND_FROM"
$morVonPhones = "DEDICATED_SITE_RESOURCE"
$oracleContact = "DEDICATED_ORACLE_RESOURCE"
$OUs = Get-ADOrganizationalUnit -Filter {(Name -like "*Users*")} -SearchBase "OU LOCATION" -Properties DistinguishedName -SearchScope Subtree `
| Select DistinguishedName `
| Where {$_.DistinguishedName -notlike "DISABLED OU" -and `
$_.DistinguishedName -notlike "*Generic*" -and $_.DistinguishedName -notlike "*Administration*" -and `
$_.DistinguishedName -notlike "*TestCompany*" -and $_.DistinguishedName -notlike "*Migrated*"} | Sort-Object DistinguishedName

$domainNames = @(   '@LIST_OF.COM',
                    '@DOMAIN.mx',
                    '@NAMES.com.br')

#Use for Password generation -- Can add 8 character word minimum
$1stWord = #Needs an uppercase letter and special charater
"Galactical_","Decimeter_",
"Inactivity!", "Resentful_",
"Enchanted_", "Incognito-",
"Approached_", "Croquette!",
"Watchlist_", "Daylight_",
"Correlation-", "Logitech_",
"Dependency_", "Etiquette_"
$2ndWord = #All lowercase
"perfectionist", "hydration",
"digestion", "starvation",
"flashcard", "celestial",
"consonant", "cartridge",
"extension", "barracuda",
"armadillo", "botanical",
"resourceful", "skydriving"
$specNum =
"!21", "-47",
"808", "#93",
"@73", "!76",
"+44", "!37",
"_11", "+24",
"@00", "404"

function ClearForm{
$fNameTextbox.Clear()
$lNameTextbox.Clear()
$locationCombo.Text = "OU="
$titleTextbox.Clear()
$mgrTextbox.Clear()
$oracleTextbox.Clear()
$ADMirrorTextbox.Clear()
$VPNCheckbox.Checked = $false
$VendorCheckbox.Checked = $false
$O365Checkbox.Checked = $false
$PWGenTextbox.Text = "Generate Password"
$cellCheckbox.Checked = $false
$startDateTextbox.Text = "MM/DD/YYYY"
$costCenterTextbox.Text = "XXX-XXX-XXXX"
$addressTextbox.Clear()
$ticketTextbox.Clear()
}

function FieldUpdates {
    $samNameTextbox.Text = $fNameTextbox.Text+"."+$lNameTextbox.Text
    $EmailTextbox.Text = $($samNameTextbox.Text+$domainCombo.SelectedItem)

    if($cellCheckbox.Checked -ne $true){
        $startDateTextbox.Visible = $false
        $startDateLabel.Visible = $false
        $costCenterTextbox.Visible = $false
        $costCenterLabel.Visible = $false
        $addressTextbox.Visible = $false
        $addressLabel.Visible = $false
        $ticketLabel.Visible = $false
        $ticketTextbox.Visible = $false
    }else{
        $startDateTextbox.Visible = $true
        $startDateLabel.Visible = $true
        $costCenterTextbox.Visible = $true
        $costCenterLabel.Visible = $true
        $addressTextbox.Visible = $true
        $addressLabel.Visible = $true
        $ticketLabel.Visible = $true
        $ticketTextbox.Visible = $true
    }
    if($VendorCheckbox.Checked){
        $locationCombo.Text = "VENDOR_OU_LOCATION"
        $ADMirrorTextbox.BackColor = "184, 57, 57"
        $ADMirrorTextbox.Enabled = $false
        $CellLabel.BackColor = "184, 57, 57"
        $cellCheckbox.Enabled = $false
        $oracleTextbox.BackColor = "184, 57, 57"
        $oracleTextbox.Enabled = $false
        $VendorO365Label.Visible = $true
        $O365Checkbox.Visible = $true
        $PWGenTextbox.Visible = $true
        $PasswordBtn.Visible = $true
        $startDateTextbox.enabled = $false
        $costCenterTextbox.enabled = $false
        $addressTextbox.enabled = $false
        $ticketTextbox.enabled = $false
        $startDateTextbox.BackColor = "184, 57, 57"
        $costCenterTextbox.BackColor = "184, 57, 57"
        $addressTextbox.BackColor = "184, 57, 57"
        $ticketTextbox.BackColor = "184, 57, 57"

    }else{
        $VendorO365Label.Visible = $false
        $O365Checkbox.Visible = $false
        $PWGenTextbox.Visible = $false
        $PasswordBtn.Visible = $false
        $ADMirrorTextbox.Enabled = $true
        $ADMirrorTextbox.ResetBackColor()
        $cellCheckbox.Enabled = $true
        $oracleTextbox.Enabled = $true
        $CellLabel.ResetBackColor()
        $oracleTextbox.ResetBackColor()
        $startDateTextbox.enabled = $true
        $costCenterTextbox.enabled = $true
        $addressTextbox.enabled = $true
        $ticketTextbox.enabled = $true
        $startDateTextbox.ResetBackColor()
        $costCenterTextbox.ResetBackColor()
        $addressTextbox.ResetBackColor()
        $ticketTextbox.ResetBackColor()
    }
}

#####Window Form#####
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "User creation"
$mainForm.Height = 450
$mainForm.Width = 595
$mainForm.MaximizeBox = $false
$mainForm.FormBorderStyle = "FixedDialog"
$mainForm.TopMost = $false
$mainForm.StartPosition = 'CenterScreen'
$mainForm.ForeColor = "#ffffff"
$mainForm.BackColor = "#141414"

#####First Name#####
$FNameLabel = New-Object System.Windows.Forms.Label
$FNameLabel.Text = "User's First Name"
$FNameLabel.Location = New-Object System.Drawing.Point(10,20)
$FNameLabel.AutoSize = $true
$mainForm.Controls.Add($FNameLabel)

$fNameTextbox = New-Object System.Windows.Forms.TextBox
$fNameTextbox.Location = New-Object System.Drawing.Point(125,18)
$fNameTextbox.Size = New-Object System.Drawing.Size(100,20)
$mainForm.Controls.Add($fNameTextbox)

#####Last Name####
$lNameLabel = New-Object System.Windows.Forms.Label
$lNameLabel.Text = "User's Last Name"
$lNameLabel.Location = New-Object System.Drawing.Point(10,50)
$lNameLabel.AutoSize = $true
$mainForm.Controls.Add($lNameLabel)

$lNameTextbox = New-Object System.Windows.Forms.TextBox
$lNameTextbox.Location = New-Object System.Drawing.Point(125,48)
$lNameTextbox.Size = New-Object System.Drawing.Size(100,20)
$mainForm.Controls.Add($lNameTextbox)

#####SamaccountName#####
$samNameLabel = New-Object System.Windows.Forms.Label
$samNameLabel.Text = "SamAccountName"
$samNameLabel.Location = New-Object System.Drawing.Point(10,78)
$samNameLabel.AutoSize = $true
$mainForm.Controls.Add($samNameLabel)

$samNameTextbox = New-Object System.Windows.Forms.TextBox
$samNameTextbox.Location = New-Object System.Drawing.Point(125,78)
$samNameTextbox.Size = New-Object System.Drawing.Size(120,20)
$samNameTextbox.Enabled = $false
$mainForm.Controls.Add($samNameTextbox)

#####Email Address#####
$EmailLabel = New-Object System.Windows.Forms.Label
$EmailLabel.Text = "Email Address"
$EmailLabel.Location = New-Object System.Drawing.Point(10,108)
$EmailLabel.AutoSize = $true
$mainForm.Controls.Add($EmailLabel)

$EmailTextbox = New-Object System.Windows.Forms.TextBox
$EmailTextbox.Location = New-Object System.Drawing.Point(125,108)
$EmailTextbox.Size = New-Object System.Drawing.Size(180,20)
$EmailTextbox.Enabled = $false
$mainForm.Controls.Add($EmailTextbox)

$domainCombo = New-Object System.Windows.Forms.ComboBox
$domainCombo.Location = New-Object System.Drawing.Point(325,108)
$domainCombo.Width = 200
$domainCombo.Items.AddRange($domainNames)
$domainCombo.SelectedIndex = 0
$mainForm.Controls.Add($domainCombo)

#####User location#####
$locationLabel = New-Object System.Windows.Forms.Label
$locationLabel.Text = "Location"
$locationLabel.Location = New-Object System.Drawing.Point(10,140)
$locationLabel.AutoSize = $true
$mainForm.Controls.Add($locationLabel)

$locationCombo = New-Object System.Windows.Forms.ComboBox
$locationCombo.Width = 450
foreach($ou in $OUs){
    $locationCombo.Items.Add($ou.distinguishedName)
}
$locationCombo.AutoCompleteMode = 'SuggestAppend'
$locationCombo.AutoCompleteSource = 'ListItems'
$locationCombo.Text = "OU="
$locationCombo.Location = New-Object System.Drawing.Point(125, 138)
$mainForm.Controls.Add($locationCombo)

#####Title and Manager#####
$TitleLabel = New-Object System.Windows.Forms.Label
$TitleLabel.Text = "Title/Job Position"
$TitleLabel.Location = New-Object System.Drawing.Point(10,170)
$TitleLabel.AutoSize = $true
$mainForm.Controls.Add($TitleLabel)

$titleTextbox = New-Object System.Windows.Forms.TextBox
$titleTextbox.Location = New-Object System.Drawing.Point(125,168)
$titleTextbox.Size = New-Object System.Drawing.Size(180,20)
$mainForm.Controls.Add($titleTextbox)

$mgrLabel = New-Object System.Windows.Forms.Label
$mgrLabel.Text = "Manager"
$mgrLabel.Location = New-Object System.Drawing.Point(325,170)
$mgrLabel.AutoSize = $true
$mainForm.Controls.Add($mgrLabel)

$mgrTextbox = New-Object System.Windows.Forms.TextBox
$mgrTextbox.Location = New-Object System.Drawing.Point(395,168)
$mgrTextbox.Size = New-Object System.Drawing.Size(180,325)
$mgrTextbox
$mainForm.Controls.Add($mgrTextbox)

#####Oracle#####
$oracleLabel = New-Object System.Windows.Forms.Label
$oracleLabel.Text = "Mirror Oracle Like"
$oracleLabel.Location = New-Object System.Drawing.Point(10,200)
$oracleLabel.AutoSize = $true
$mainForm.Controls.Add($oracleLabel)

$oracleTextbox = New-Object System.Windows.Forms.TextBox
$oracleTextbox.Location = New-Object System.Drawing.Point(125,198)
$oracleTextbox.Size = New-Object System.Drawing.Size(150,20)
$mainForm.Controls.Add($oracleTextbox)

#####Security Groups#####
$ADMirrorLabel = New-Object System.Windows.Forms.Label
$ADMirrorLabel.Text = "Mirror AD Like"
$ADMirrorLabel.Location = New-Object System.Drawing.Point(290,200)
$ADMirrorLabel.AutoSize = $true
$mainForm.Controls.Add($ADMirrorLabel)

$ADMirrorTextbox = New-Object System.Windows.Forms.TextBox
$ADMirrorTextbox.Location = New-Object System.Drawing.Point(395,198)
$ADMirrorTextbox.Size = New-Object System.Drawing.Size(150,20)
$mainForm.Controls.Add($ADMirrorTextbox)

$VPNLabel = New-Object System.Windows.Forms.Label
$VPNLabel.Text = "VPN Access?"
$VPNLabel.Location = New-Object System.Drawing.Point(10,230)
$VPNLabel.AutoSize = $true
$mainForm.Controls.Add($VPNLabel)

$VPNCheckbox = New-Object System.Windows.Forms.CheckBox
$VPNCheckbox.Location = New-Object System.Drawing.Point(125,229)
$VPNCheckbox.Size = New-Object System.Drawing.Size(20,20)
$mainForm.Controls.Add($VPNCheckbox)

#Vendor Options
$Vendorlabel = New-Object System.Windows.Forms.Label
$Vendorlabel.Text = "Vendor Account?"
$Vendorlabel.Location = New-Object System.Drawing.Point(10,260)
$Vendorlabel.AutoSize = $true
$mainForm.Controls.Add($Vendorlabel)

$VendorCheckbox = New-Object System.Windows.Forms.CheckBox
$VendorCheckbox.Location = New-Object System.Drawing.Point(125,258)
$VendorCheckbox.Size = New-Object System.Drawing.Size(20,20)
$mainForm.Controls.Add($VendorCheckbox)

$VendorO365Label = New-Object System.Windows.Forms.Label
$VendorO365Label.Text = "O365 Access?"
$VendorO365Label.Location = New-Object System.Drawing.Point(10,285)
$VendorO365Label.AutoSize = $true
$mainForm.Controls.Add($VendorO365Label)

$O365Checkbox = New-Object System.Windows.Forms.CheckBox
$O365Checkbox.Location = New-Object System.Drawing.Point(125,282)
$O365Checkbox.Size = New-Object System.Drawing.Size(20,20)
$mainForm.Controls.Add($O365Checkbox)

$PWGenTextbox = New-Object System.Windows.Forms.TextBox
$PWGenTextbox.Text = "Generate Password"
$PWGenTextbox.TextAlign = "Center"
$PWGenTextbox.Location = New-Object System.Drawing.Point(10,320)
$PWGenTextbox.Size = New-Object System.Drawing.Size(150,20)
$PWGenTextbox.AutoSize = $true
$mainForm.Controls.Add($PWGenTextbox)

$PasswordBtn = New-Object System.Windows.Forms.Button
$PasswordBtn.Location = New-Object System.Drawing.Point(45,340)
$PasswordBtn.Text = "Generate"
$PasswordBtn.BackColor = "#a10000"
$mainForm.Controls.Add($PasswordBtn)

$CellLabel = New-Object System.Windows.Forms.Label
$CellLabel.Text = "Order Cellphone"
$CellLabel.Location = New-Object System.Drawing.Point(200,232)
$CellLabel.AutoSize = $true
$mainForm.Controls.Add($CellLabel)

$cellCheckbox = New-Object System.Windows.Forms.CheckBox
$cellCheckbox.Location = New-Object System.Drawing.Point(290,231)
$cellCheckbox.Size = New-Object System.Drawing.Size(20,20)
$mainForm.Controls.Add($cellCheckbox)

$startDateLabel = New-Object System.Windows.Forms.Label
$startDateLabel.Text = "User's Start Date"
$startDateLabel.Location = New-Object System.Drawing.Point(320,232)
$startDateLabel.AutoSize = $true
$mainForm.Controls.Add($startDateLabel)

$startDateTextbox = New-Object System.Windows.Forms.TextBox
$startDateTextbox.Location = New-Object System.Drawing.Point(420,230)
$startDateTextbox.Size = New-Object System.Drawing.Size(100,20)
$startDateTextbox.Text = "MM/DD/YYYY"
$mainForm.Controls.Add($startDateTextbox)

$costCenterLabel = New-Object System.Windows.Forms.Label
$costCenterLabel.Text = "User's Cost Center"
$costCenterLabel.Location = New-Object System.Drawing.Point(320,262)
$costCenterLabel.AutoSize = $true
$mainForm.Controls.Add($costCenterLabel)

$costCenterTextbox = New-Object System.Windows.Forms.TextBox
$costCenterTextbox.Location = New-Object System.Drawing.Point(420,260)
$costCenterTextbox.Size = New-Object System.Drawing.Size(100,20)
$costCenterTextbox.Text = "XXX-XXX-XXXX"
$mainForm.Controls.Add($costCenterTextbox)

$addressLabel = New-Object System.Windows.Forms.Label
$addressLabel.Text = "Shipping Address"
$addressLabel.Location = New-Object System.Drawing.Point(320,292)
$addressLabel.AutoSize = $true
$mainForm.Controls.Add($addressLabel)

$addressTextbox = New-Object System.Windows.Forms.TextBox
$addressTextbox.Location = New-Object System.Drawing.Point(420,290)
$addressTextbox.Size = New-Object System.Drawing.Size(100,20)
$mainForm.Controls.Add($addressTextbox)

$ticketLabel = New-Object System.Windows.Forms.Label
$ticketLabel.Text = "Ticket URL"
$ticketLabel.Location = New-Object System.Drawing.Point(320,322)
$ticketLabel.AutoSize = $true
$mainForm.Controls.Add($ticketLabel)

$ticketTextbox = New-Object System.Windows.Forms.TextBox
$ticketTextbox.Location = New-Object System.Drawing.Point(420,320)
$ticketTextbox.Size = New-Object System.Drawing.Size(100,20)
$mainForm.Controls.Add($ticketTextbox)

#####Confirmation#####
$creationBtn = New-Object System.Windows.Forms.Button
$creationBtn.Location = New-Object System.Drawing.Point(220,380)
$creationBtn.Text = "Create"
$creationBtn.BackColor = "#a10000"
$mainForm.Controls.Add($creationBtn)

####Close#####
$closeBtn = New-Object System.Windows.Forms.Button
$closeBtn.Location = New-Object System.Drawing.Point(300,380)
$closeBtn.Text = "Close"
$closeBtn.BackColor = "#a10000"
$mainForm.Controls.Add($closeBtn)

#Generates password on each button press
$PasswordBtn.Add_click({
try{
$pwAPI = Invoke-RestMethod -Uri "https://random-word-api.herokuapp.com/word?number=2&length=8" -Method Get -TimeoutSec 10 -ErrorAction continue
$PWGenTextbox.Text = $textInfo.ToTitleCase($pwAPI[0].ToLower())+"-"+$pwAPI[1]+('!', '@', '+' | Get-Random)+(get-random -Minimum 1000 -Maximum 9999)
}catch{ #If API fails
$password = $1stWord[(0..($1stWord.Count-1) | Get-Random)]+$2ndWord[(0..($2ndWord.Count-1) | Get-Random)]+$specNum[(0..($specNum.Count-1) | Get-Random)]
$PWGenTextbox.Text = $password
}
})

$closeBtn.add_click({
try{
    Disconnect-MgGraph -ErrorAction Stop
} catch{}
$mainForm.Close()})

$creationBtn.add_click({
if(($fNameTextbox.Text -eq "") -or ($lNameTextbox.Text -eq "") -or ($locationCombo.Text -eq "OU=") -or $mgrTextbox.Text -eq ""){
    [System.Windows.Forms.MessageBox]::show("Please fill out the user's name, user's manager, and/or select an OU location.", "Invalid Values", "OK", "Question")
    return
}else{

    if($VendorCheckbox.Checked){
        $displayName = $mgrTextbox.Text
        $MgrSearch = $mgrTextbox.Text.Split(",",2).Replace(" ","")
        $mgr = $MgrSearch[1]+"."+$MgrSearch[0]
        $Manager = Get-ADUser -Filter "mail -like '$mgr*'" -Properties Name, samaccountname, mail | Select-Object Name, samAccountName, mail
        $MangerDN = Get-ADUser -Filter "DisplayName -like '$displayName'" -Properties Name, samaccountname, mail | select Name, samaccountName, mail

        if($Manager -ne $null){
            $Manager = Get-ADUser -Filter "mail -like '$mgr*'" -Properties Name, samaccountname, mail | Select-Object Name, samAccountName, mail
        }elseif($MangerDN -ne $null){
            $Manager = Get-ADUser -Filter "DisplayName -like '$displayName'" -Properties Name, samaccountname, mail | select Name, samaccountName, mail
        }
        $msgBody = "Confirm user creation?"+"`r`n"+"`r`n"+
        "Display Name: "+$lNameTextbox.Text+", "+$fNameTextbox.Text+"`r`n"+
        "Email Address: "+$EmailTextbox.Text+"`r`n"+
        "Manager: "+$Manager.mail

        $givenName = $lNameTextbox.Text+", "+$fNameTextbox.Text
        $msgBtn = "OKCancel"
        $Result = [System.Windows.Forms.MessageBox]::show($msgBody, "Create User in Active Directory", $msgBtn, "Question")

        if ($Result -eq 1) {
                #checks if user already exists before creation#
                $confirm = Get-ADUser -Filter{SamAccountName -eq $samNameTextbox.Text} -Server $domainController
                if($PWGenTextbox.Text.Length -lt 20){
                    [System.Windows.Forms.MessageBox]::show("Please generate or create 20 character password", "Password Error", "OK","Error")
                    return
                }
                if($confirm -ne $null){
                        [System.Windows.Forms.MessageBox]::show("User already exists within AD.", "User Exists", "OK","Error")
                        return
                    }else{
                       try{
                       New-ADUser  -Name $givenName -GivenName $fNameTextbox.Text -Surname $lNameTextbox.Text -SamAccountName $samNameTextbox.Text.ToLower() `
                      -DisplayName $givenName -Description $titleTextbox.Text -Title $titleTextbox.Text -Company "Vendor" -Manager $Manager.samAccountName -UserPrincipalName $EmailTextbox.Text `
                      -Enabled $true -CannotChangePassword $true -AccountPassword (ConvertTo-SecureString -AsPlainText $PWGenTextbox.Text -Force) -ChangePasswordAtLogon $false `
                      -Path $locationCombo.Text -Server $domainController
                      }catch{
                        [System.Windows.Forms.MessageBox]::show("User already exists within AD.", "User Exists", "OK","Error")
                        return
                      }

                      Add-ADGroupMember -Identity "GROUP_ACCESS" -Members $samNameTextbox.Text -Server $domainController

                      if($VPNCheckbox.Checked -eq $true){
                          Add-ADGroupMember -Identity "GROUP_ACCESS" -Members $samNameTextbox.Text -Server $domainController
                      }else{}
  
                      if($O365Checkbox.Checked){
                          Add-ADGroupMember -Identity "GROUP_ACCESS" -Members $samNameTextbox.Text -Server $domainController
                          Add-ADGroupMember -Identity "GROUP_ACCESS" -Members $samNameTextbox.Text -Server $domainController
                          Add-ADGroupMember -Identity "GROUP_ACCESS" -Members $samNameTextbox.Text -Server $domainController
                          Add-ADGroupMember -Identity "GROUP_ACCESS" -Members $samNameTextbox.Text -Server $domainController

                          $RemoteRouting = $samNameTextbox.Text + "@COMPANY'S ONPREM ROUTING ADDRESS"
                          Enable-RemoteMailbox -Identity $samNameTextbox.Text -PrimarySmtpAddress $EmailTextbox.Text -RemoteRoutingAddress $EmailTextbox.Text -DomainController $domainController
                          Set-RemoteMailbox -Identity $samNameTextbox.Text -EmailAddresses @{add=$RemoteRouting} -RemoteRoutingAddress $RemoteRouting -DomainController $domainController
                      }else{}
  
                      $confirm = Get-ADUser -Filter{SamAccountName -eq $samNameTextbox.Text} -Server $domainController
                      if($confirm -ne $null){ 
                          [System.Windows.Forms.MessageBox]::show("User has been successfully created!`nUser Information has been copied to clipboard to paste in ticket.", "Creation Successful", "OK")
                          Set-Clipboard -Value ("Username will be: " + $samNameTextbox.Text + `
                          "`r`nPassword is: "+ $PWGenTextbox.Text)
                          ClearForm
                      }else{
                          [System.Windows.Forms.MessageBox]::show("User creation has failed!", "No User Found", "OK")
                      }
                  }
                }else{}
            }#End of Vendor Creation
    else{#Start of regular user

        #Searches for manager by either legal name or displayed nickname
        $displayName = $mgrTextbox.Text
        $MgrSearch = $mgrTextbox.Text.Split(",",2).Replace(" ","")
        $mgr = $MgrSearch[1]+"."+$MgrSearch[0]
        $Manager = Get-ADUser -Filter "mail -like '$mgr*'" -Properties Name, samaccountname, mail | Select-Object Name, samAccountName, mail
        $MangerDN = Get-ADUser -Filter "DisplayName -like '$displayName'" -Properties Name, samaccountname, mail | select Name, samaccountName, mail

        if($Manager -ne $null){
            $Manager = Get-ADUser -Filter "mail -like '$mgr*'" -Properties Name, samaccountname, mail | Select-Object Name, samAccountName, mail
        }elseif($MangerDN -ne $null){
            $Manager = Get-ADUser -Filter "DisplayName -like '$displayName'" -Properties Name, samaccountname, mail | select Name, samaccountName, mail
        }

        #Copy AD user account
        if($ADMirrorTextbox.Text -ne ""){
            $displayName = $ADMirrorTextbox.Text
            $ADMirrorSearch = $ADMirrorTextbox.Text.Split(",",2).Replace(" ","")
            $mirroredUser = $ADMirrorSearch[1]+"."+$ADMirrorSearch[0]
            $UserMirror = Get-ADUser -Filter "mail -like '$mirroredUser*'" -Properties Name, samaccountname, mail | Select-Object Name, samAccountName, mail
            $MirroredDN = Get-ADUser -Filter "DisplayName -like '$displayName'" -Properties Name, samaccountname, mail | select Name, samaccountName, mail

            if($UserMirror -ne $null){
                $UserMirror = Get-ADUser -Filter "mail -like '$mirroredUser*'" -Properties Name, samaccountname, mail, displayName | Select-Object Name, samAccountName, mail, displayName
            }elseif($MirroredDN -ne $null){
                $UserMirror = Get-ADUser -Filter "DisplayName -like '$displayName'" -Properties Name, samaccountname, mail, displayName | select Name, samaccountName, mail, displayName
            }
            #checks if connection is already active to MgGraph in order to check cloud group membership
            try{
                Get-MgOrganization -ErrorAction Stop
            } catch{
                Connect-MgGraph -Scopes "Group.Read.All"
            }
            $AADGroupMembership = Get-MgUserMemberOf -UserId $UserMirror.mail | ForEach-Object `
            {Get-MgGroup -GroupId $_.Id | select DisplayName, Id, Mail, MailEnabled, Visibility `
            | where {$_.MailEnabled -eq $true -and $_.Mail -notlike "*msteam*" -and $_.Visibility -eq $null}}
            $MirroredGroups = Get-ADUser -Identity $UserMirror.samaccountname -Properties MemberOf, samaccountName | Select-Object samaccountname -ExpandProperty MemberOf | foreach{(get-adgroup $_).samaccountname}
        }else{}


        $msgBody = "Confirm user creation?"+"`r`n"+"`r`n"+
        "Display Name: "+$lNameTextbox.Text+", "+$fNameTextbox.Text+"`r`n"+
        "Email Address: "+$EmailTextbox.Text+"`r`n"+
        "Manager: "+$Manager.mail+"`r`n"+
        "Mirror AD user: "+$UserMirror.displayName

        $givenName = $lNameTextbox.Text+", "+$fNameTextbox.Text
        $msgBtn = "OKCancel"
        $Result = [System.Windows.Forms.MessageBox]::show($msgBody, "Create User in Active Directory", $msgBtn, "Question")

        if ($Result -eq 1) {
            #checks if user already exists before creation
            $confirm = Get-ADUser -Filter "(proxyAddresses -like '*$($EmailTextbox.Text)*')" -Server $domainController
            if($confirm -eq $null){
                #User does not exist. Continue silently.
            }else{
                [System.Windows.Forms.MessageBox]::show($samNameTextbox.Text + " already exists within AD.", "User Exists", "OK","Error")
                return
            }
            try{
            New-ADUser  -Name $givenName -GivenName $fNameTextbox.Text -Surname $lNameTextbox.Text -SamAccountName $samNameTextbox.Text.ToLower() `
            -DisplayName $givenName -Description $titleTextbox.Text -Title $titleTextbox.Text -Company "COMPANY_NAME" -Manager $Manager.samAccountName -UserPrincipalName $EmailTextbox.Text `
            -Enabled $true -AccountPassword (ConvertTo-SecureString -AsPlainText $newUserPassword -Force) -ChangePasswordAtLogon $true `
            -Path $locationCombo.Text -Server $domainController -ErrorAction Stop
            }catch{
                [System.Windows.Forms.MessageBox]::show("SamAccountName exceeds Active Directory's Max length of 20 characters", "Creation Error", "OK","Error")
                return
            }
            
            #####Await timer##### -- NOT NEEDED but ensures that account creation is done before adding groups
            Start-Sleep -Seconds 5

            Add-ADGroupMember -Identity "GROUP_ACCESS" -Members $samNameTextbox.Text -Server $domainController
            Add-ADGroupMember -Identity "GROUP_ACCESS" -Members $samNameTextbox.Text -Server $domainController
            Add-ADGroupMember -Identity "GROUP_ACCESS" -Members $samNameTextbox.Text -Server $domainController

            if($MirroredGroups -ne $null -and $ADMirrorTextbox.Text -ne ""){
                foreach($group in $MirroredGroups){
                    Add-ADGroupMember -Identity $group -Members $samNameTextbox.Text -Server $domainController
                }
            }else{}

            if($locationCombo.Text -like "*SPECIFIC_SITE_LOCATION*"){
                Add-ADGroupMember -Identity "GROUP_ACCESS" -Members $samNameTextbox.Text -Server $domainController
            }else{}
            
            if ($VPNCheckbox.Checked -eq $true) {
                Add-ADGroupMember -Identity VPNUsers -Members $samNameTextbox.Text -Server $domainController
            }else{}
            if($cellCheckbox.Checked -eq $true){
                 if((($locationCombo.Text -like "*SPECIFIC_SITE_LOCATION*") -or ($locationCombo.Text -like "*SPECIFIC_SITE_LOCATION*"))){
                        #Sends to DEDICATED_USER if location is SITE or SITE
                        Add-ADGroupMember -Identity "GROUP_ACCESS"-Members $samNameTextbox.Text -Server $domainController
                        $url = ($ticketTextbox.Text)

                        Send-MailMessage -From $sendFrom -To $morVonPhones -Cc @($Manager.mail,$EmailTextbox.Text)  -Bcc @('OVERSEER') -SmtpServer "SMTP_MAIL_SERVER" `
                        -Subject "CellPhone Order Placement" -BodyAsHtml `
                        -Body ("Hello,<br>An order for a cellphone is needed for this new employee: " + "$givenName" + "`r`n" +
                        "<br>User's Expected Start Date: " + $startDateTextbox.Text +
                        "<br>User's Cost Center: " + $costCenterTextbox.Text + 
                        "<br>Ship Device to: " + $addressTextbox.Text +
                        "<br>Please visit <a href='$url'> Link </a> to view ticket." +
                        "<br>Thank you!") -Attachments ".\Zero Touch Phone Enrollment.pdf"
                }else{ 
                        Add-ADGroupMember -Identity "GROUP_ACCESS" -Members $samNameTextbox.Text -Server $domainController
                        $url = ($ticketTextbox.Text)

                        Send-MailMessage -From $sendFrom -To "DISTRIBUTION_GROUP" -Cc @($Manager.mail,$EmailTextbox.Text)  -Bcc @("OVERSEER") -SmtpServer "SMTP_MAIL_SERVER" `
                        -Subject "CellPhone Order Placement" -BodyAsHtml `
                        -Body ("Hello,<br>An order for a cellphone is needed for this new employee: " + "$givenName" + "`r`n" +
                        "<br>User's Expected Start Date: " + $startDateTextbox.Text +
                        "<br>User's Cost Center: " + $costCenterTextbox.Text + 
                        "<br>Ship Device to: " + $addressTextbox.Text +
                        "<br>Please visit <a href='$url'> Link </a> to view ticket." +
                        "<br>Thank you!") -Attachments ".\Zero Touch Phone Enrollment.pdf"
                }
            }else{}
            if ($oracleTextbox.Text -ne "") {
                Send-MailMessage -From $sendFrom -To $oracleContact -Cc $Manager.mail -Bcc "OVERSEER" -SmtpServer smtpmail `
                -Subject "Oracle Account Creation" `
                -Body ("Hello $($oracleContact.Split('.')[0]), `nPlease create an oracle account for: " + "$givenName" + " that mirrors: " + $oracleTextbox.Text.toString() + "`r`n" +
                "`r`nPlease reach out to: " + $mgrTextbox.Text + " if additional information is needed. Thank you!")
            }else{}

            $confirm = Get-ADUser -Filter{SamAccountName -eq $samNameTextbox.Text} -Server $domainController
            if($confirm -ne $null){
                if($AADGroupMembership -ne $null){
                    [System.Windows.Forms.MessageBox]::show("User added to On-Prem groups/resources. A script with user's name has been created on your desktop to run in about an hour to automatically add user to the O365 Groups below.`
Simply right-click > Run w/Powershell and then use CLOUD_ACCOUNT when prompted. Once successfully ran, you may delete the script from your desktop."+"`r`n"+"`r`n"+($AADGroupMembership.DisplayName -join "`n"), "Add User to Cloud Distribution Groups", "OK")

                    $currentUser = whoami | %{$_.remove(0,4)}
                    #Auto Create script file
                    #1
                    Add-Content -Path "C:\Users\$($currentUser)\Desktop\$($samNameTextbox.Text).ps1" -Value 'Connect-ExchangeOnline -ShowBanner:$false'
                    #2
                    Add-Content -path "C:\Users\$($currentUser)\Desktop\$($samNameTextbox.Text).ps1" -value 'Connect-MgGraph -Scopes "Group.Read.All" -NoWelcome'
                    #3
                    '$AADuserId = "'+$UserMirror.mail+'"' | Add-Content -path "C:\Users\$($currentUser)\Desktop\$($samNameTextbox.Text).ps1"
                    #4
                    '$NewUser ="'+$EmailTextbox.Text+'"' | Add-Content -path "C:\Users\$($currentUser)\Desktop\$($samNameTextbox.Text).ps1"
                    Add-Content -path "C:\Users\$($currentUser)\Desktop\$($samNameTextbox.Text).ps1" -encoding string -value '$AADGroupMembership = Get-MgUserMemberOf -UserId $AADuserId | ForEach-Object `
                    {Get-MgGroup -GroupId $_.Id | select DisplayName, Id, Mail, MailEnabled, Visibility `
                    | where {$_.MailEnabled -eq $true -and $_.Mail -notlike "*msteam*" -and $_.Visibility -eq $null}}'

                    Add-Content -path "C:\Users\$($currentUser)\Desktop\$($samNameTextbox.Text).ps1" -value 'foreach($AADid in $AADGroupMembership){
                        Add-DistributionGroupMember -Identity $AADid.Mail -Member $NewUser -BypassSecurityGroupManagerCheck
                    }' -nonewline
                    #End of automatic script creation

                }else{}
                [System.Windows.Forms.MessageBox]::show("User has been successfully created!`nUser Information has been copied to clipboard to paste in ticket.", "Creation Successful", "OK")
                Set-Clipboard -Value ("Username will be: " + $samNameTextbox.Text + " `r`nEmail address will be: " + $EmailTextbox.Text +`
                "`r`nPassword is: "+ $newUserPassword + " `r`nUser will be prompted to change password after logging in.")
                ClearForm
            }else{
                [System.Windows.Forms.MessageBox]::show("User creation has failed!", "No User Found", "OK")
            }
        }
    }
 }
})
$startDateTextbox.add_click({$startDateTextbox.SelectAll()})
$costCenterTextbox.add_click({$costCenterTextbox.SelectAll()})
$addressTextbox.add_click({$addressTextbox.SelectAll()})
$locationCombo.add_click({$locationCombo.Select($locationCombo.Text.Length,0)})
$locationCombo.Add_GotFocus({$locationCombo.Select($locationCombo.Text.Length,0)})

#For real-time changes
$Timer.Interval = 2
$Timer.Add_Tick({FieldUpdates})
$Timer.Enabled = $True

$mainForm.ShowDialog()
$mainForm.Dispose()