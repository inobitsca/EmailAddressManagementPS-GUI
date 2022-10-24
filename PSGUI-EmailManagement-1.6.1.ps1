# GUI to Manage email addresses using Active Directory PowerShell.
#
#Create/Change Primary email address
#Add an email alias
#Remove an email alias
#Hide from Address List
#
#It does not require an Exchange server. 
##Requires ACTIVEDIRECTORY PowerShell module. 
#You will need User Admin rights.
#Created by Cedric Abrahams - cedric@inobits.com
#
#Version 1.6.1 2022-10-24

<##########-Add a Logo (64x64) - Does not work with all systems.
$base64ImageString = "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"
$imageBytes = [Convert]::FromBase64String($base64ImageString)
$ms = New-Object IO.MemoryStream($imageBytes, 0, $imageBytes.Length)
$ms.Write($imageBytes, 0, $imageBytes.Length);
$img = [System.Drawing.Image]::FromStream($ms, $true)
#$img = [System.Drawing.Image]::Fromfile('c:\temp4\ok_small.png') #>

Import-module Activedirectory
$allUsers = Get-adobject -properties mail,proxyaddresses,msExchHideFromAddressLists -Filter 'mail -like "*@*"'



#Variables
$result = ""
$res = $false
$IT = 0
$user = 'None'
$NP =''
$NewEmail =''
$check = $null

#Menu Options
$NewPrimaryForm = 'Change the Primary Address'
$NewAliasForm = 'Add an Alias Address'
$DeleteAliasForm = 'Remove an Alias address'
$CheckEmailForm = 'Check Existing Email Addresess'
$HideFromGAL = "Hide User from GAL"

############Code Functions Start
######## Set new Primary Address Function ########
function SetNewPrimary {
$SourceFunction = 1

  Write-host "SNP" $ChangeEmail.text
  Write-host $ID.DistinguishedName
  
$NewEmail = $ChangeEmail.Text
$checkmail = "*:" + $NewEmail + "*"
Write-host "Checkmail: " $checkmail
$Check = $null
$Check = $allUsers |where {$_.DistinguishedName -ne $User -and $_.proxyaddresses -like $checkmail} 
if ($check -eq $null) {Write-host "Check is Null"}
Write-host "Check " $Check
If (!$NewEmail) {Write-host "No Address entered. NewPrimaryForm" -fore Red
$NP = "No email address entered!"
 }
  If ($Check -ne $null) {
 Write-host "Email Address already existsin the organization. NewPrimaryForm" -fore Red
$NP = "Email Address already exists"}
If ($Check -eq $null ){
Write-host "Setting Primary Email"
if ( $Obj -eq "User"){
set-ADuser -Identity $CN -remove @{Proxyaddresses="SMTP:"+$user.mail}
set-ADuser -Identity $CN -remove @{Proxyaddresses="smtp:"+$ChangeEmail.text}
set-ADuser -Identity $CN -Add @{Proxyaddresses="smtp:"+$user.mail}
set-ADuser -Identity $CN -Add @{Proxyaddresses="SMTP:"+$ChangeEmail.text}
set-ADuser -Identity $CN -emailaddress $ChangeEmail.text
Write-Host "Primary User Email Set"
Write-host "Refreshing User Data"
$user = Get-ADObject $CN -Properties mail,proxyaddresses,msExchHideFromAddressLists
ResultsForm
}

if ( $Obj -eq "Group"){
set-ADGroup -Identity $CN -remove @{Proxyaddresses="SMTP:"+$user.mail}
set-ADGroup -Identity $CN -remove @{Proxyaddresses="smtp:"+$ChangeEmail.text}
set-ADGroup -Identity $CN -Add @{Proxyaddresses="smtp:"+$user.mail}
set-ADGroup -Identity $CN -Add @{Proxyaddresses="SMTP:"+$ChangeEmail.text}
set-ADGroup -Identity $CN -emailaddress $ChangeEmail.text
Write-Host "Primary Group Email Set"
Write-host "Refreshing User Data"
$user = Get-ADObject $CN -Properties mail,proxyaddresses,msExchHideFromAddressLists
ResultsForm
}
}
$NewPrimaryForm.close()
}

######### Add Alias Function ######### 
 
Function AddAlias {
$SourceFunction = 2
Write-host $NewAlias.text
$NewEmail = $NewAlias.Text
$checkmail = "*:" + $NewEmail + "*"
$check = $null
$Check = $allUsers |where {$_.proxyaddresses -like $checkmail} 
If (!$NewEmail) {NoAddressForm
$Check = $null
}

If ($Check) {UsedAddressForm
Write-host "Email Address already existsing the organization. NewPrimaryForm" -fore Red}
Else {
if ( $Obj -eq "User"){
set-ADuser -Identity $CN -Add @{Proxyaddresses="smtp:" + $NewAlias.text }
Write-Host "Alias Added"
Write-host "Refreshing User Data"
$user = Get-ADObject $CN -Properties mail,proxyaddresses,msExchHideFromAddressLists
}

}

if ( $Obj -eq "Group"){
set-ADGroup -Identity $CN -Add @{Proxyaddresses="smtp:"+ $NewAlias.text}
Write-Host "Alias Added"
Write-host "Refreshing User Data"
$user = Get-ADObject $CN -Properties mail,proxyaddresses,msExchHideFromAddressLists
}
ResultsForm
$AddAliasForm.close()
}

######### Remove Alias Function #########

Function RemoveAlias {
$SourceFunction = 3
Write-host $RemoveAlias.text
$NewEmail = $RemoveAlias.Text
If (!$RemoveAlias) {NoAddressForm
}
Else {
if ( $Obj -eq "User") {
set-ADuser -Identity $CN -Remove @{Proxyaddresses="smtp:" + $RemoveAlias.text }
Write-Host "Alias Removed"
Write-host "Refreshing User Data"

$user = Get-ADObject $CN -Properties mail,proxyaddresses,msExchHideFromAddressLists
ResultsForm
}

if ( $Obj -eq "Group"){
set-ADGroup -Identity $CN -remove @{Proxyaddresses="smtp:"+ $RemoveAlias.text}
Write-Host "Alias Removed"
Write-host "Refreshing User Data"

$user = Get-ADObject $CN -Properties mail,proxyaddresses,msExchHideFromAddressLists
ResultsForm
}
}
$RemoveAliasForm.close()  
}

#########-Hide from GAL Function #########
<#
Function HideFromGAL {
$SourceFunction = 4

 If ($checkbox1.CheckState -eq $False) {set-ADuser -Identity $CN -Replace @{msExchHideFromAddressLists=$False}
 Write-Host "User is not hidden from GAL" -fore Green}
 If ($checkbox1.CheckState -eq $True) {set-ADuser -Identity $CN -Replace @{msExchHideFromAddressLists=$True}
Write-Host "User is now hidden from GAL" -fore Yellow }
$user = Get-ADObject $CN -Properties mail,proxyaddresses,msExchHideFromAddressLists
ResultsForm
$HideForm.close() 
$ActionForm.close()
ActionForm
}
#>
############Code Functions End 


############Form Functions Start

function NewPrimaryForm  {
Add-Type -AssemblyName System.Windows.Forms
$user = Get-ADObject $CN -Properties mail,proxyaddresses,msExchHideFromAddressLists
# Result form
$NewPrimaryForm                    = New-Object system.Windows.Forms.Form
$NewPrimaryForm.ClientSize         = '500,600'
$NewPrimaryForm.text               = "Email Result Management"
$NewPrimaryForm.BackColor          = "#bababa"


if ($Valid -eq 1)  { [void]$ResultForm2.Close() }

#####Add a Logo - Image detils at the start of the script
$pictureBox = new-object Windows.Forms.PictureBox
$pictureBox.Location = new-object System.Drawing.Size(430,5)
$pictureBox.Width = $img.Size.Width
$pictureBox.Height = $img.Size.Height
$pictureBox.Image = $img
$NewPrimaryForm.controls.add($pictureBox)

#Account Name Heading
$NewPrimaryText                           = New-Object system.Windows.Forms.Label
$NewPrimaryText.text                      = 'You have chosen to edit user:'
$NewPrimaryText.AutoSize                  = $true
$NewPrimaryText.width                     = 25
$NewPrimaryText.height                    = 10
$NewPrimaryText.location                  = New-Object System.Drawing.Point(20,10)
$NewPrimaryText.Font                      = 'Microsoft Sans Serif,13'
$NewPrimaryForm.controls.AddRange(@($NewPrimaryText))



# Account Name.
$ResultChoice                           = New-Object system.Windows.Forms.Label

$ResultChoice.text                      = $SearchName.Text + " ("+ $user.Name +")"
$ResultChoice.AutoSize                  = $true
$ResultChoice.width                     = 25
$ResultChoice.height                    = 10
$ResultChoice.location                  = New-Object System.Drawing.Point(20,35)
$ResultChoice.Font                      = 'Microsoft Sans Serif,13,style=bold'
$NewPrimaryForm.controls.AddRange(@($ResultChoice))




# Show Email Address Heading.
$EmailTitle                           = New-Object system.Windows.Forms.Label
$EmailTitle.text                      = 'Current Primary Email Address:'
$EmailTitle.AutoSize                  = $true
$EmailTitle.width                     = 25
$EmailTitle.height                    = 10
$EmailTitle.location                  = New-Object System.Drawing.Point(20,75)
$EmailTitle.Font                      = 'Microsoft Sans Serif,13'
$NewPrimaryForm.controls.AddRange(@($EmailTitle))



# Show Email Address
$EmailChoice                          = New-Object system.Windows.Forms.Label
$EmailChoice.text                      = $user.mail
$EmailChoice.AutoSize                  = $true
$EmailChoice.width                     = 25
$EmailChoice.height                    = 10
$EmailChoice.location                  = New-Object System.Drawing.Point(20,100)
$EmailChoice.Font                      = 'Microsoft Sans Serif,13,style=Bold'
$NewPrimaryForm.controls.AddRange(@($EmailChoice))



# Show Proxy Addresses Heading.
$ProxyTitle                           = New-Object system.Windows.Forms.Label
$ProxyTitle.text                      = 'Current Proxy Addresses:'
$ProxyTitle.AutoSize                  = $true
$ProxyTitle.width                     = 25
$ProxyTitle.height                    = 10
$ProxyTitle.location                  = New-Object System.Drawing.Point(20,125)
$ProxyTitle.Font                      = 'Microsoft Sans Serif,13'
$NewPrimaryForm.controls.AddRange(@($ProxyTitle))

$VL = 150
Foreach ($P in $user.proxyaddresses) {
If ($P -clike "*smtp*") {$Pr = $P 
$PR =$PR -replace 'smtp:',''
# Show Proxy Addresses
$EmailChoice                          = New-Object system.Windows.Forms.Label
$EmailChoice.text                      = $PR
$EmailChoice.AutoSize                  = $true
$EmailChoice.width                     = 25
$EmailChoice.height                    = 10
$EmailChoice.location                  = New-Object System.Drawing.Point(20,$VL)
$EmailChoice.Font                      = 'Microsoft Sans Serif,13,style=Bold'
$NewPrimaryForm.controls.AddRange(@($EmailChoice))
$VL = $VL +25
}
else {}
}
#NO Email TextBoxLable
$NoEmailLable               = New-Object system.Windows.Forms.Label
$NoEmailLable.text           = $NP
$NoEmailLable.AutoSize       = $true
$NoEmailLable.width          = 25
$NoEmailLable.height         = 20
$NoEmailLable.location       = New-Object System.Drawing.Point(20,475)
$NoEmailLable.Font           = 'Microsoft Sans Serif,16,style=Bold'
$NoEmailLable.ForeColor      = "#ff0000"
$NoEmailLable.Visible        = $True
$NewPrimaryForm.Controls.Add($NoEmailLable)

#New Email TextBoxLable
$ChangeEmailLabel                = New-Object system.Windows.Forms.Label
$ChangeEmailLabel.text           = "Enter New Primary Email:"
$ChangeEmailLabel.AutoSize       = $true
$ChangeEmailLabel.width          = 25
$ChangeEmailLabel.height         = 20
$ChangeEmailLabel.location       = New-Object System.Drawing.Point(20,500)
$ChangeEmailLabel.Font           = 'Microsoft Sans Serif,10,style=Bold'
$ChangeEmailLabel.Visible        = $True
$NewPrimaryForm.Controls.Add($ChangeEmailLabel)

#New Email TextBox
$ChangeEmail                     = New-Object system.Windows.Forms.TextBox
$ChangeEmail.multiline           = $false
$ChangeEmail.width               = 314
$ChangeEmail.height              = 20
$ChangeEmail.location            = New-Object System.Drawing.Point(20,525)
$ChangeEmail.Font                = 'Microsoft Sans Serif,10'
$ChangeEmail.Visible             = $True
$NewPrimaryForm.Controls.Add($ChangeEmail)
$ChangeEmail.Add_KeyDown({})

Write-host "NewPrimaryForm" $ChangeEmail.text

#Execute Button
$ExecBtn                   = New-Object system.Windows.Forms.Button
$ExecBtn.BackColor         = "#ff0000"
$ExecBtn.text              = "Set Primary Address"
$ExecBtn.width             = 180
$ExecBtn.height            = 30
$ExecBtn.location          = New-Object System.Drawing.Point(20,560)
$ExecBtn.Font              = 'Microsoft Sans Serif,10'
$ExecBtn.ForeColor         = "#ffffff"
$NewPrimaryForm.CancelButton   = $cancelBtn3
$NewPrimaryForm.Controls.Add($ExecBtn)

$ExecBtn.Add_Click({SetNewPrimary})

#Cancel Button
$cancelBtn3                       = New-Object system.Windows.Forms.Button
$cancelBtn3.BackColor             = "#000000"
$cancelBtn3.text                  = "Exit"
$cancelBtn3.width                 = 90
$cancelBtn3.height                = 30
$cancelBtn3.location              = New-Object System.Drawing.Point(210,560)
$cancelBtn3.Font                  = 'Microsoft Sans Serif,10'
$cancelBtn3.ForeColor             = "#ffffff"
$cancelBtn3.DialogResult          = [System.Windows.Forms.DialogResult]::Cancel
$NewPrimaryForm.CancelButton   = $cancelBtn3
$NewPrimaryForm.Controls.Add($cancelBtn3)

$cancelBtn3.Add_Click({ $NewPrimaryForm.Close() })

Write-host "NewPrimaryForm open Resultsform"

# Display the form
$result = $NewPrimaryForm.ShowDialog()
}

#---------------------------------#

#Add an Alias Form
function NewAliasForm {
Write-Host "NewAliasForm Add Alias"
$user = Get-ADObject $CN -Properties mail,proxyaddresses,msExchHideFromAddressLists
Add-Type -AssemblyName System.Windows.Forms
# Add Alias form
$AddAliasForm                    = New-Object system.Windows.Forms.Form
$AddAliasForm.ClientSize         = '500,600'
$AddAliasForm.text               = "Email Management - Add an Alias"
$AddAliasForm.BackColor          = "#bababa"


if ($Valid -eq 1)  { [void]$ResultForm2.Close() }

#####Add a Logo - Image detils at the start of the script
$pictureBox = new-object Windows.Forms.PictureBox
$pictureBox.Location = new-object System.Drawing.Size(430,5)
$pictureBox.Width = $img.Size.Width
$pictureBox.Height = $img.Size.Height
$pictureBox.Image = $img
$AddAliasForm.controls.add($pictureBox)



#Account Name Heading
$NewAliasText                           = New-Object system.Windows.Forms.Label
$NewAliasText.text                      = 'You have chosen to edit user:'
$NewAliasText.AutoSize                  = $true
$NewAliasText.width                     = 25
$NewAliasText.height                    = 10
$NewAliasText.location                  = New-Object System.Drawing.Point(20,10)
$NewAliasText.Font                      = 'Microsoft Sans Serif,13'
$AddAliasForm.controls.AddRange(@($NewAliasText))

# Account Name.
$ResultChoice                           = New-Object system.Windows.Forms.Label
$ResultChoice.text                      = $SearchName.Text + " ("+ $ID.Name +")"
$ResultChoice.AutoSize                  = $true
$ResultChoice.width                     = 25
$ResultChoice.height                    = 10
$ResultChoice.location                  = New-Object System.Drawing.Point(20,35)
$ResultChoice.Font                      = 'Microsoft Sans Serif,13,style=bold'
$AddAliasForm.controls.AddRange(@($ResultChoice))

# Show Email Address Heading.
$EmailTitle                           = New-Object system.Windows.Forms.Label
$EmailTitle.text                      = 'Current Primary Email Address:'
$EmailTitle.AutoSize                  = $true
$EmailTitle.width                     = 25
$EmailTitle.height                    = 10
$EmailTitle.location                  = New-Object System.Drawing.Point(20,75)
$EmailTitle.Font                      = 'Microsoft Sans Serif,13'
$AddAliasForm.controls.AddRange(@($EmailTitle))



# Show Email Address
$EmailChoice                          = New-Object system.Windows.Forms.Label
$EmailChoice.text                      = $user.mail
$EmailChoice.AutoSize                  = $true
$EmailChoice.width                     = 25
$EmailChoice.height                    = 10
$EmailChoice.location                  = New-Object System.Drawing.Point(20,100)
$EmailChoice.Font                      = 'Microsoft Sans Serif,13,style=Bold'
$AddAliasForm.controls.AddRange(@($EmailChoice))



# Show Proxy Addresses Heading.
$ProxyTitle                           = New-Object system.Windows.Forms.Label
$ProxyTitle.text                      = 'Current Alias Addresses:'
$ProxyTitle.AutoSize                  = $true
$ProxyTitle.width                     = 25
$ProxyTitle.height                    = 10
$ProxyTitle.location                  = New-Object System.Drawing.Point(20,125)
$ProxyTitle.Font                      = 'Microsoft Sans Serif,13'
$AddAliasForm.controls.AddRange(@($ProxyTitle))

$VL = 150
Foreach ($P in $user.proxyaddresses) {
If ($P -clike "*smtp*") {$Pr = $P 
$PR =$PR -replace 'smtp:',''
# Show Proxy Addresses
$EmailChoice                          = New-Object system.Windows.Forms.Label
$EmailChoice.text                      = $PR
$EmailChoice.AutoSize                  = $true
$EmailChoice.width                     = 25
$EmailChoice.height                    = 10
# Position the element
$EmailChoice.location                  = New-Object System.Drawing.Point(20,$VL)
# Define the font type and size
$EmailChoice.Font                      = 'Microsoft Sans Serif,13,style=Bold'
$AddAliasForm.controls.AddRange(@($EmailChoice))
$VL = $VL +25
}
else {}
}
#NO Email TextBoxLable
$NoEmailLable               = New-Object system.Windows.Forms.Label
$NoEmailLable.text           = $NP
$NoEmailLable.AutoSize       = $true
$NoEmailLable.width          = 25
$NoEmailLable.height         = 20
$NoEmailLable.location       = New-Object System.Drawing.Point(20,475)
$NoEmailLable.Font           = 'Microsoft Sans Serif,16,style=Bold'
$NoEmailLable.ForeColor      = "#ff0000"
$NoEmailLable.Visible        = $True
$AddAliasForm.Controls.Add($NoEmailLable)

#New Email TextBoxLable
$NewAliasLabel                = New-Object system.Windows.Forms.Label
$NewAliasLabel.text           = "Enter New Alias Email Address:"
$NewAliasLabel.AutoSize       = $true
$NewAliasLabel.width          = 25
$NewAliasLabel.height         = 20
$NewAliasLabel.location       = New-Object System.Drawing.Point(20,500)
$NewAliasLabel.Font           = 'Microsoft Sans Serif,10,style=Bold'
$NewAliasLabel.Visible        = $True
$AddAliasForm.Controls.Add($NewAliasLabel)

#New Email TextBox
$NewAlias                     = New-Object system.Windows.Forms.TextBox
$NewAlias.multiline           = $false
$NewAlias.width               = 314
$NewAlias.height              = 20
$NewAlias.location            = New-Object System.Drawing.Point(20,525)
$NewAlias.Font                = 'Microsoft Sans Serif,10'
$NewAlias.Visible             = $True
$AddAliasForm.Controls.Add($NewAlias)
Write-host $NewAlias.text
$NewAlias.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") 
    {    
    AddAlias
    }
})

#Result Buttons
$ExecBtn                   = New-Object system.Windows.Forms.Button
$ExecBtn.BackColor         = "#ff0000"
$ExecBtn.text              = "Add Email Alias"
$ExecBtn.width             = 180
$ExecBtn.height            = 30
$ExecBtn.location          = New-Object System.Drawing.Point(20,560)
$ExecBtn.Font              = 'Microsoft Sans Serif,10'
$ExecBtn.ForeColor         = "#ffffff"
$AddAliasForm.CancelButton   = $cancelBtn4
$AddAliasForm.Controls.Add($ExecBtn)
$ExecBtn.Add_Click({AddAlias})

#Cancel Button
$cancelBtn4                       = New-Object system.Windows.Forms.Button
$cancelBtn4.BackColor             = "#000000"
$cancelBtn4.text                  = "Exit"
$cancelBtn4.width                 = 90
$cancelBtn4.height                = 30
$cancelBtn4.location              = New-Object System.Drawing.Point(210,560)
$cancelBtn4.Font                  = 'Microsoft Sans Serif,10'
$cancelBtn4.ForeColor             = "#ffffff"
$cancelBtn4.DialogResult          = [System.Windows.Forms.DialogResult]::Cancel
$AddAliasForm.CancelButton   = $cancelBtn4
$AddAliasForm.Controls.Add($cancelBtn4)
$cancelBtn4.Add_Click({ $AddAliasForm.close() })

# Display the form
$result = $AddAliasForm.ShowDialog()
}

#-------------------------------------#

#Remove an Alias Form
function DeleteAliasForm {Add-Type -AssemblyName System.Windows.Forms
$user = Get-ADObject $CN -Properties mail,proxyaddresses,msExchHideFromAddressLists
# Result form
$RemoveAliasForm                    = New-Object system.Windows.Forms.Form
$RemoveAliasForm.ClientSize         = '500,600'
$RemoveAliasForm.text               = "Email Result Management"
$RemoveAliasForm.BackColor          = "#bababa"

if ($Valid -eq 1)  { [void]$ResultForm2.Close() }

#####Add a Logo - Image detils at the start of the script
$pictureBox = new-object Windows.Forms.PictureBox
$pictureBox.Location = new-object System.Drawing.Size(430,5)
$pictureBox.Width = $img.Size.Width
$pictureBox.Height = $img.Size.Height
$pictureBox.Image = $img
$RemoveAliasForm.controls.add($pictureBox)

#Account Name Heading
$RemoveAliasText                           = New-Object system.Windows.Forms.Label
$RemoveAliasText.text                      = 'You have chosen to edit user:'
$RemoveAliasText.AutoSize                  = $true
$RemoveAliasText.width                     = 25
$RemoveAliasText.height                    = 10
$RemoveAliasText.location                  = New-Object System.Drawing.Point(20,10)
$RemoveAliasText.Font                      = 'Microsoft Sans Serif,13'
$RemoveAliasForm.controls.AddRange(@($RemoveAliasText))

# Account Name.
$ResultChoice                           = New-Object system.Windows.Forms.Label
$ResultChoice.text                      = $SearchName.Text + " ("+ $user.Name +")"
$ResultChoice.AutoSize                  = $true
$ResultChoice.width                     = 25
$ResultChoice.height                    = 10
$ResultChoice.location                  = New-Object System.Drawing.Point(20,35)
$ResultChoice.Font                      = 'Microsoft Sans Serif,13,style=bold'
$RemoveAliasForm.controls.AddRange(@($ResultChoice))

# Show Email Address Heading.
$EmailTitle                           = New-Object system.Windows.Forms.Label
$EmailTitle.text                      = 'Current Primary Email Address:'
$EmailTitle.AutoSize                  = $true
$EmailTitle.width                     = 25
$EmailTitle.height                    = 10
$EmailTitle.location                  = New-Object System.Drawing.Point(20,75)
$EmailTitle.Font                      = 'Microsoft Sans Serif,13'
$RemoveAliasForm.controls.AddRange(@($EmailTitle))

# Show Email Address
$EmailChoice                          = New-Object system.Windows.Forms.Label
$EmailChoice.text                      = $user.mail
$EmailChoice.AutoSize                  = $true
$EmailChoice.width                     = 25
$EmailChoice.height                    = 10
# Position the element
$EmailChoice.location                  = New-Object System.Drawing.Point(20,100)
# Define the font type and size
$EmailChoice.Font                      = 'Microsoft Sans Serif,13,style=Bold'
$RemoveAliasForm.controls.AddRange(@($EmailChoice))

# Show Proxy Addresses Heading.
$ProxyTitle                           = New-Object system.Windows.Forms.Label
$ProxyTitle.text                      = 'Current Alias Addresses:'
$ProxyTitle.AutoSize                  = $true
$ProxyTitle.width                     = 25
$ProxyTitle.height                    = 10
# Position the element
$ProxyTitle.location                  = New-Object System.Drawing.Point(20,125)
# Define the font type and size
$ProxyTitle.Font                      = 'Microsoft Sans Serif,13'
$RemoveAliasForm.controls.AddRange(@($ProxyTitle))

$VL = 150
$PRcount = 0
Foreach ($P in $user.proxyaddresses) {
If ($P -clike "*smtp*") {
$PRcount = $PRcount + 1
$Pr = $P 
$PR =$PR -replace 'smtp:',''
if ($PRcount -eq 1) {$prx1 = $PR}
if ($PRcount -eq 2) {$prx2 = $PR}
if ($PRcount -eq 3) {$prx3 = $PR}
if ($PRcount -eq 4) {$prx4 = $PR}
if ($PRcount -eq 5) {$prx5 = $PR}
if ($PRcount -eq 6) {$prx6 = $PR}
if ($PRcount -eq 7) {$prx7 = $PR}
if ($PRcount -eq 8) {$prx8 = $PR}
if ($PRcount -eq 9) {$prx9 = $PR}
if ($PRcount -eq 10) {$prx10 = $PR}
if ($PRcount -eq 11) {$prx11 = $PR}
if ($PRcount -eq 12) {$prx12 = $PR}
if ($PRcount -gt 1) {$prxExtra1 = "Only first 12 Addresses Shown. Type address to remove"}


# Show Proxy Addresses
$EmailChoice                          = New-Object system.Windows.Forms.Label
$EmailChoice.text                      = $PR
$EmailChoice.AutoSize                  = $true
$EmailChoice.width                     = 25
$EmailChoice.height                    = 10
# Position the element
$EmailChoice.location                  = New-Object System.Drawing.Point(20,$VL)
# Define the font type and size
$EmailChoice.Font                      = 'Microsoft Sans Serif,13,style=Bold'
$RemoveAliasForm.controls.AddRange(@($EmailChoice))
$VL = $VL +25
}
else {}
}
#NO Email TextBoxLable
$NoEmailLable               = New-Object system.Windows.Forms.Label
$NoEmailLable.text           = $NP
$NoEmailLable.AutoSize       = $true
$NoEmailLable.width          = 25
$NoEmailLable.height         = 20
$NoEmailLable.location       = New-Object System.Drawing.Point(20,475)
$NoEmailLable.Font           = 'Microsoft Sans Serif,16,style=Bold'
$NoEmailLable.ForeColor      = "#ff0000"
$NoEmailLable.Visible        = $True
$RemoveAliasForm.Controls.Add($NoEmailLable)

#New Email TextBoxLable
$RemoveAliasLabel                = New-Object system.Windows.Forms.Label
$RemoveAliasLabel.text           = "Select the Alias Email Address to be removed:"
$RemoveAliasLabel.AutoSize       = $true
$RemoveAliasLabel.width          = 250
$RemoveAliasLabel.height         = 20
$RemoveAliasLabel.location       = New-Object System.Drawing.Point(20,500)
$RemoveAliasLabel.Font           = 'Microsoft Sans Serif,10,style=Bold'
$RemoveAliasLabel.Visible        = $True
$RemoveAliasForm.Controls.Add($RemoveAliasLabel)

#New Email TextBox
$RemoveAlias                     = New-Object system.Windows.Forms.ComboBox
$RemoveAlias.text                = "Choose"
$RemoveAlias.width               = 400
$RemoveAlias.autosize            = $true
$RemoveAlias.Visible             = $true  
      
# Add the items in the dropdown list
@($prx1,$prx2,$Prx3,$Prx4,$prx5,$prx6,$prx7,$prx8,$prx9,$prx10,$prx11,$prx12,$prxExtra1) | ForEach-Object {if($_) {[void] $RemoveAlias.Items.Add($_)}}
# Select the default value
$RemoveAlias.SelectedIndex       = 0
$RemoveAlias.location            = New-Object System.Drawing.Point(20,525)
$RemoveAlias.Font                = 'Microsoft Sans Serif,10'
$RemoveAliasForm.Controls.Add($RemoveAlias)
Write-host $RemoveAlias.text
$RemoveAlias.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") 
    {    
    RemoveAlias
    }
    })

#Result Buttons
$ExecBtn                   = New-Object system.Windows.Forms.Button
$ExecBtn.BackColor         = "#FF0000"
$ExecBtn.text              = "Delete Email Alias"
$ExecBtn.width             = 180
$ExecBtn.height            = 30
$ExecBtn.location          = New-Object System.Drawing.Point(20,560)
$ExecBtn.Font              = 'Microsoft Sans Serif,10'
$ExecBtn.ForeColor         = "#ffffff"
$RemoveAliasForm.CancelButton   = $cancelBtn5
$RemoveAliasForm.Controls.Add($ExecBtn)
$ExecBtn.Add_Click({ RemoveAlias })

#Cancel Button
$cancelBtn5                       = New-Object system.Windows.Forms.Button
$cancelBtn5.BackColor             = "#000000"
$cancelBtn5.text                  = "Exit"
$cancelBtn5.width                 = 90
$cancelBtn5.height                = 30
$cancelBtn5.location              = New-Object System.Drawing.Point(210,560)
$cancelBtn5.Font                  = 'Microsoft Sans Serif,10'
$cancelBtn5.ForeColor             = "#ffffff"
$cancelBtn5.DialogResult          = [System.Windows.Forms.DialogResult]::Cancel
$RemoveAliasForm.CancelButton   = $cancelBtn5
$RemoveAliasForm.Controls.Add($cancelBtn5)
$cancelBtn5.Add_Click({ $RemoveAliasForm.close() })

# Display the form
$result = $RemoveAliasForm.ShowDialog()
}

#------------------------------#

#Check Existing Addresses Form
function CheckEmailForm {
ResultsForm
Write-host "Check Existing Addresses"
}

Function ResultsForm {
if ($user -ne "None") {Write-host "Check Address" $NewEmail -ForegroundColor Green
$res = $true

$user = get-adobject $CN -Properties mail,proxyaddresses,msExchHideFromAddressLists
Write-host "Resultsform"

# Process form
$ResultsForm                    = New-Object system.Windows.Forms.Form
$ResultsForm.ClientSize         = '500,600'
$ResultsForm.text               = "Email Process Management"
$ResultsForm.BackColor          = "#abdbff"

#####Add a Logo - Image detils at the start of the script
$pictureBox = new-object Windows.Forms.PictureBox
$pictureBox.Location = new-object System.Drawing.Size(430,5)
$pictureBox.Width = $img.Size.Width
$pictureBox.Height = $img.Size.Height
$pictureBox.Image = $img
$ResultsForm.controls.add($pictureBox)

#Create a Title for our form. We will use a label for it.
$EmailTitle                           = New-Object system.Windows.Forms.Label
$EmailTitle.text                      = 'Email Addresses for ' + $user.name
$EmailTitle.AutoSize                  = $true
$EmailTitle.location                  = New-Object System.Drawing.Point(20,10)
$EmailTitle.Font                      = 'Microsoft Sans Serif,13,style=bold'
$EmailTitle.ForeColor                 = "#151051"
$ResultsForm.controls.AddRange(@($EmailTitle))

$HideTitle                           = New-Object system.Windows.Forms.Label
If($user.msExchHideFromAddressLists -eq "True"){$HideTitle.text =  "User is Hidden from Address Book"}
$HideTitle.AutoSize                  = $true
$HideTitle.location                  = New-Object System.Drawing.Point(20,40)
$HideTitle.Font                      = 'Microsoft Sans Serif,13,style=bold'
$HideTitle.ForeColor                 = "#ff0000"
$ResultsForm.controls.AddRange(@($HideTitle))


$EmailTitle1                           = New-Object system.Windows.Forms.Label
$EmailTitle1.text                      = 'Primary Email Address:'
$EmailTitle1.AutoSize                  = $true
$EmailTitle1.width                     = 25
$EmailTitle1.height                    = 10
$EmailTitle1.location                  = New-Object System.Drawing.Point(20,75)
$EmailTitle1.Font                      = 'Microsoft Sans Serif,13'
$ResultsForm.controls.AddRange(@($EmailTitle1))


# Show Email Address
$EmailChoice                          = New-Object system.Windows.Forms.Label
$EmailChoice.text                      = $user.mail
$EmailChoice.AutoSize                  = $true
$EmailChoice.width                     = 25
$EmailChoice.height                    = 10
$EmailChoice.location                  = New-Object System.Drawing.Point(20,100)
$EmailChoice.Font                      = 'Microsoft Sans Serif,13,style=Bold'
$ResultsForm.controls.AddRange(@($EmailChoice))

# Show Proxy Addresses Heading.
$ProxyTitle                           = New-Object system.Windows.Forms.Label
$ProxyTitle.text                      = 'Alias Addresses:'
$ProxyTitle.AutoSize                  = $true
$ProxyTitle.width                     = 25
$ProxyTitle.height                    = 10
$ProxyTitle.location                  = New-Object System.Drawing.Point(20,150)
$ProxyTitle.Font                      = 'Microsoft Sans Serif,13'
$ResultsForm.controls.AddRange(@($ProxyTitle))

$VL = 175
Foreach ($P in $user.proxyaddresses) {
If ($P -clike "*smtp*") {$Pr = $P 
$PR =$PR -replace 'smtp:',''
# Show Proxy Addresses
$EmailChoice                          = New-Object system.Windows.Forms.Label

$EmailChoice.text                      = $PR
$EmailChoice.AutoSize                  = $true
$EmailChoice.width                     = 25
$EmailChoice.height                    = 10
$EmailChoice.location                  = New-Object System.Drawing.Point(20,$VL)
$EmailChoice.Font                      = 'Microsoft Sans Serif,13,style=Bold'
$ResultsForm.controls.AddRange(@($EmailChoice))
$VL = $VL +25}}

#Cancel Button
$cancelBtn                       = New-Object system.Windows.Forms.Button
$cancelBtn.BackColor             = "#ffffff"
$cancelBtn.text                  = "Exit"
$cancelBtn.width                 = 90
$cancelBtn.height                = 30
$cancelBtn.location              = New-Object System.Drawing.Point(220,560)
$cancelBtn.Font                  = 'Microsoft Sans Serif,10'
$cancelBtn.ForeColor             = "#000000"
$cancelBtn.DialogResult          = [System.Windows.Forms.DialogResult]::Cancel
$ResultsForm.CancelButton   = $cancelBtn
$ResultsForm.Controls.Add($cancelBtn)





# Display the form
[void]$ResultsForm.ShowDialog()
}}

#--------------------------------#
#Main Action Form

Function ActionForm {
$user = Get-ADObject $CN -Properties mail,proxyaddresses,msExchHideFromAddressLists
Add-Type -AssemblyName System.Windows.Forms
# Create a new form
$ActionForm                    = New-Object system.Windows.Forms.Form
# Define the size, title and background color
$ActionForm.ClientSize         = '400,400'
$ActionForm.text               = "Email Address Management"
$ActionForm.BackColor          = "#ffffff"

# Create a Title for our form. We will use a label for it.
$TitleOperationChoice                           = New-Object system.Windows.Forms.Label
$TitleOperationChoice.text                      = "Email Address Management"
$TitleOperationChoice.AutoSize                  = $true
$TitleOperationChoice.width                     = 25
$TitleOperationChoice.height                    = 10
$TitleOperationChoice.location                  = New-Object System.Drawing.Point(20,5)
$TitleOperationChoice.Font                      = 'Microsoft Sans Serif,13'

# Other elemtents
#####Add a Logo - Image detils at the start of the script
$pictureBox = new-object Windows.Forms.PictureBox
$pictureBox.Location = new-object System.Drawing.Size(330,5)
$pictureBox.Width = $img.Size.Width
$pictureBox.Height = $img.Size.Height
$pictureBox.Image = $img
$ActionForm.controls.add($pictureBox)


$Name = $ID.name

#TextLable
$SearchNameLabel                = New-Object system.Windows.Forms.Label
$SearchNameLabel.text           = "You are editing user: "
$SearchNameLabel.AutoSize       = $true
$SearchNameLabel.width          = 25
$SearchNameLabel.height         = 20
$SearchNameLabel.ForeColor      = "#0000ff"
$SearchNameLabel.location       = New-Object System.Drawing.Point(20,60)
$SearchNameLabel.Font           = 'Microsoft Sans Serif,10,style=Bold'
$SearchNameLabel.Visible        = $True
$ActionForm.Controls.Add($SearchNameLabel)

#TextLable
$SearchNameName                = New-Object system.Windows.Forms.Label
$SearchNameName.text           = "$Name"
$SearchNameName.AutoSize       = $true
$SearchNameName.width          = 50
$SearchNameName.height         = 20
$SearchNameName.ForeColor      = "#ff0000"
$SearchNameName.location       = New-Object System.Drawing.Point(160,60)
$SearchNameName.Font           = 'Microsoft Sans Serif,10,style=Bold'
$SearchNameName.Visible        = $True
$ActionForm.Controls.Add($SearchNameName)

$HideTitle                           = New-Object system.Windows.Forms.Label
If($user.msExchHideFromAddressLists -eq "True"){$HideTitle.text =  "User is Hidden from Address Book"}
$HideTitle.AutoSize                  = $true
$HideTitle.location                  = New-Object System.Drawing.Point(20,85)
$HideTitle.Font                      = 'Microsoft Sans Serif,13,style=bold'
$HideTitle.ForeColor                 = "#ff0000"
$ActionForm.controls.AddRange(@($HideTitle))


#Buttons

#--------------------------------------#
# New Primary Button

$NewPrimBtn                    = New-Object system.Windows.Forms.Button
$NewPrimBtn.BackColor         = "#0000ff"
$NewPrimBtn.text              = "Set Primary Address"
$NewPrimBtn.width             = 220
$NewPrimBtn.height            = 30
$NewPrimBtn.location          = New-Object System.Drawing.Point(50,120)
$NewPrimBtn.Font              = 'Microsoft Sans Serif,10'
$NewPrimBtn.ForeColor         = "#ffffff"
$ActionForm.CancelButton   = $cancelBtn
$ActionForm.Controls.Add($NewPrimBtn )
$NewPrimBtn.Add_Click({ NewPrimaryForm })
#--------------------------------------#

#New Alias Button

$NewAliasBtn                     = New-Object system.Windows.Forms.Button
$NewAliasBtn.BackColor         = "#0000ff"
$NewAliasBtn.text              = "Add Email Alias"
$NewAliasBtn.width             = 220
$NewAliasBtn.height            = 30
$NewAliasBtn.location          = New-Object System.Drawing.Point(50,155)
$NewAliasBtn.Font              = 'Microsoft Sans Serif,10'
$NewAliasBtn.ForeColor         = "#ffffff"
$ActionForm.CancelButton   = $cancelBtn
$ActionForm.Controls.Add($NewAliasBtn )
$NewAliasBtn.Add_Click({ NewAliasForm })
#--------------------------------------#

#Delete Alias Button

$DelAliasBtn                     = New-Object system.Windows.Forms.Button
$DelAliasBtn.BackColor         = "#0000ff"
$DelAliasBtn.text              = "Delete Email Alias"
$DelAliasBtn.width             = 220
$DelAliasBtn.height            = 30
$DelAliasBtn.location          = New-Object System.Drawing.Point(50,190)
$DelAliasBtn.Font              = 'Microsoft Sans Serif,10'
$DelAliasBtn.ForeColor         = "#ffffff"
$ActionForm.CancelButton   = $cancelBtn
$ActionForm.Controls.Add($DelAliasBtn  )
$DelAliasBtn.Add_Click({ DeleteAliasForm })
#--------------------------------------#

#Hide From GAL Button

$HideGALBtn                     = New-Object system.Windows.Forms.Button
$HideGALBtn.BackColor         = "#0000ff"
$HideGALBtn.text              = "Hide/Unhide user from GAL"
$HideGALBtn.width             = 220
$HideGALBtn.height            = 30
$HideGALBtn.location          = New-Object System.Drawing.Point(50,225)
$HideGALBtn.Font              = 'Microsoft Sans Serif,10'
$HideGALBtn.ForeColor         = "#ffffff"
$ActionForm.CancelButton   = $cancelBtn
$ActionForm.Controls.Add($HideGALBtn  )
$HideGALBtn.Add_Click({ HideFromGALForm})
#--------------------------------------#

#Check Addresses Button

$CheckMailBtn                     = New-Object system.Windows.Forms.Button
$CheckMailBtn.BackColor         = "#3b003b"
$CheckMailBtn.text              = "Check Email Addresses"
$CheckMailBtn.width             = 220
$CheckMailBtn.height            = 30
$CheckMailBtn.location          = New-Object System.Drawing.Point(50,260)
$CheckMailBtn.Font              = 'Microsoft Sans Serif,10'
$CheckMailBtn.ForeColor         = "#ffffff"
$ActionForm.CancelButton   = $cancelBtn
$ActionForm.Controls.Add($CheckMailBtn  )
$CheckMailBtn.Add_Click({ CheckEmailForm })
#--------------------------------------#



#Cancel Button
$cancelBtn                       = New-Object system.Windows.Forms.Button
$cancelBtn.BackColor             = "#000000"
$cancelBtn.text                  = "Change user or Exit"
$cancelBtn.width                 = 150
$cancelBtn.height                = 30
$cancelBtn.location              = New-Object System.Drawing.Point(50,360)
$cancelBtn.Font                  = 'Microsoft Sans Serif,10'
$cancelBtn.ForeColor             = "#ffffff"
$cancelBtn.DialogResult          = [System.Windows.Forms.DialogResult]::Cancel
$ActionForm.CancelButton   = $cancelBtn
$ActionForm.Controls.Add($cancelBtn)

$ActionForm.controls.AddRange(@($TitleOperationChoice,$Description,$Status))


# Display the form
$result = $ActionForm.ShowDialog()




}

Function UsedAddressForm {
Add-Type -AssemblyName System.Windows.Forms
# Action Form
$UsedEmailForm                    = New-Object system.Windows.Forms.Form
$UsedEmailForm.ClientSize         = '400,100'
$UsedEmailForm.text               = "Invalid User"
$UsedEmailForm.BackColor          = "#bababa"

#Account Name Heading
$UsedEmailText                          = New-Object system.Windows.Forms.Label
$UsedEmailText.text                      = 'Email address is already in use.'
$UsedEmailText.AutoSize                  = $true
$UsedEmailText.width                     = 25
$UsedEmailText.height                    = 10
$UsedEmailText.ForeColor                 = "#ff0000"
$UsedEmailText.location                  = New-Object System.Drawing.Point(20,10)
$UsedEmailText.Font                      = 'Microsoft Sans Serif,13'
$UsedEmailForm.controls.AddRange(@($UsedEmailText))
$UsedEmailForm.ShowDialog()
#$OperationChoice.text =''
}

Function NoAddressForm {
Add-Type -AssemblyName System.Windows.Forms
# Result form2
$NoEmailForm                    = New-Object system.Windows.Forms.Form
$NoEmailForm.ClientSize         = '400,100'
$NoEmailForm.text               = "Invalid User"
$NoEmailForm.BackColor          = "#bababa"

#Account Name Heading
$NoEmailText                          = New-Object system.Windows.Forms.Label
$NoEmailText.text                      = 'No Email Address was Entered.'
$NoEmailText.AutoSize                  = $true
$NoEmailText.width                     = 25
$NoEmailText.height                    = 10
$NoEmailText.ForeColor                 = "#ff0000"
$NoEmailText.location                  = New-Object System.Drawing.Point(20,10)
$NoEmailText.Font                      = 'Microsoft Sans Serif,13'
$NoEmailForm.controls.AddRange(@($NoEmailText))
$NoEmailForm.ShowDialog()
$OperationChoice.text =''
}

Function InvalidUserForm {
# Invalid User Form
$InvalidUserForm                    = New-Object system.Windows.Forms.Form
$InvalidUserForm.ClientSize         = '450,100'
$InvalidUserForm.text               = "Invalid User"
$InvalidUserForm.BackColor          = "#bababa"

#Account Name Heading
$InvalidUserText                           = New-Object system.Windows.Forms.Label
$InvalidUserText.text                      = 'Username is invalid.'
$InvalidUserText.AutoSize                  = $true
$InvalidUserText.width                     = 25
$InvalidUserText.height                    = 10
$InvalidUserText.ForeColor                 = "#ff0000"
$InvalidUserText.location                  = New-Object System.Drawing.Point(20,10)
$InvalidUserText.Font                      = 'Microsoft Sans Serif,13'
$InvalidUserForm.controls.AddRange(@($InvalidUserText))


$InvalidUserForm.ShowDialog()

}
############################## Main Landing Form (Address Form)
Function AddressForm {
Add-Type -AssemblyName System.Windows.Forms
# Create a new form
$AddressForm                    = New-Object system.Windows.Forms.Form
# Define the size, title and background color
$AddressForm.ClientSize         = '400,200'
$AddressForm.text               = "Email Address Management"
$AddressForm.BackColor          = "#c2fcf6"

# Create a Title for our form. We will use a label for it.
$TitleOperationChoice                           = New-Object system.Windows.Forms.Label
$TitleOperationChoice.text                      = "Email Address Management Tool"
$TitleOperationChoice.AutoSize                  = $true
$TitleOperationChoice.width                     = 25
$TitleOperationChoice.height                    = 10
$TitleOperationChoice.location                  = New-Object System.Drawing.Point(20,1)
$TitleOperationChoice.Font                      = 'Microsoft Sans Serif,13'

# Other elemtents
$Description                     = New-Object system.Windows.Forms.Label
$Description.text                = "This tool is for managing email addresses in Active Directory without the need for an Exchange server."
$Description.AutoSize            = $false
$Description.width               = 320
$Description.height              = 35
$Description.location            = New-Object System.Drawing.Point(20,30)
$Description.Font                = 'Microsoft Sans Serif,10'
#>

#Buttons
#Choose a User
$FinduserBtn                   = New-Object system.Windows.Forms.Button
$FinduserBtn.BackColor         = "#0000ff"
$FinduserBtn.text              = "Choose a user to manage"
$FinduserBtn.width             = 190
$FinduserBtn.height            = 30
$FinduserBtn.location          = New-Object System.Drawing.Point(40,100)
$FinduserBtn.Font              = 'Microsoft Sans Serif,10'
$FinduserBtn.ForeColor         = "#ffffff"
$AddressForm.CancelButton   = $cancelBtn
$AddressForm.Controls.Add($FinduserBtn)

$FinduserBtn.Add_Click({ FinduserForm })

#Bulk User Management button - FOR FUTURE USE
<#
$BulkBtn                   = New-Object system.Windows.Forms.Button
$BulkBtn.BackColor         = "#0000ff"
$BulkBtn.text              = "Bulk Management"
$BulkBtn.width             = 190
$BulkBtn.height            = 30
$BulkBtn.location          = New-Object System.Drawing.Point(40,100)
$BulkBtn.Font              = 'Microsoft Sans Serif,10'
$BulkBtn.ForeColor         = "#ffffff"
$AddressForm.CancelButton   = $cancelBtn
$AddressForm.Controls.Add($BulkBtn)

$BulkBtn.Add_Click({ BulkActionsForm })

#>


###################### add an image/logo

$pictureBox = new-object Windows.Forms.PictureBox
$pictureBox.Location = new-object System.Drawing.Size(330,5)
$pictureBox.Width = $img.Size.Width
$pictureBox.Height = $img.Size.Height
$pictureBox.Image = $img
$AddressForm.controls.add($pictureBox)

#Cancel Button
$cancelBtn                       = New-Object system.Windows.Forms.Button
$cancelBtn.BackColor             = "#000000"
$cancelBtn.text                  = "Exit"
$cancelBtn.width                 = 90
$cancelBtn.height                = 30
$cancelBtn.location              = New-Object System.Drawing.Point(40,165)
$cancelBtn.Font                  = 'Microsoft Sans Serif,10'
$cancelBtn.ForeColor             = "#ffffff"
$cancelBtn.DialogResult          = [System.Windows.Forms.DialogResult]::Cancel
$AddressForm.CancelButton   = $cancelBtn
$AddressForm.Controls.Add($cancelBtn)

    if ($_.KeyCode -eq "Enter") 
    {    
    FinduserForm
    }

$AddressForm.controls.AddRange(@($TitleOperationChoice,$Description,$Status))
# Display the form
$result = $AddressForm.ShowDialog()

}

 ################ Choose the User to edit ###############
function FindUserForm { 
  Write-host "Select User" -fore Yellow
  #Username to be used in code
  #$Search = $SearchName.text
  #$searchString = "*" + $search + "*"
  #$ID=get-adobject -filter 'cn -like $searchstring' |Out-GridView -PassThru
  #$ID=$allusers |where {$_.DistinguishedName -like $searchstring} |Out-GridView -PassThru
$ID=$allusers |Out-GridView -PassThru
$CN = $ID.DistinguishedName
$Obj = $ID.ObjectClass
  #Get user data
Write-host "CN: " $CN
If ($res -eq $true){actionform}
if (!$CN) {write-host "Username invalid:" $SearchName.Text -ForegroundColor Red
$Valid = 0
InvalidUserForm 
}
if ($CN) {write-host "Username Valid:" $SearchName.Text -ForegroundColor Cyan
$Valid = 0
ActionForm 
}
}
#>
 ################ Hide from GAL Form ###############
function HideFromGALForm{
	
	Write-host "Hide From GAL Form" -Fore Green
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
     $Value = "False"
        # Set the size of your form
    $HideForm = New-Object System.Windows.Forms.Form
    $HideForm.width = 500
    $HideForm.height = 200
    $HideForm.BackColor          = "#bababa"
    $HideForm.Text = ”Hidden from Global Address List”
 
    # Set the font of the text to be used within the form
    $Font = New-Object System.Drawing.Font("Sans Serif",12)
    $HideForm.Font = $Font

    $HideTitle                           = New-Object system.Windows.Forms.Label
If($user.msExchHideFromAddressLists -eq "True"){$HideTitle.text =  "User is Hidden from Address Book"}
$HideTitle.AutoSize                  = $true
$HideTitle.location                  = New-Object System.Drawing.Point(20,40)
$HideTitle.Font                      = 'Microsoft Sans Serif,13,style=bold'
$HideTitle.ForeColor                 = "#ff0000"
$HideForm.controls.AddRange(@($HideTitle))
 
 #####Add a Logo - Image detils at the start of the script
$pictureBox = new-object Windows.Forms.PictureBox
$pictureBox.Location = new-object System.Drawing.Size(430,5)
$pictureBox.Width = $img.Size.Width
$pictureBox.Height = $img.Size.Height
$pictureBox.Image = $img
$HideForm.controls.add($pictureBox)

 
    # create your checkbox 
    $checkbox1 = new-object System.Windows.Forms.checkbox
    $checkbox1.Location = new-object System.Drawing.Size(30,50)
    $checkbox1.Size = new-object System.Drawing.Size(450,50)
    $checkbox1.Text = "Hide User from Global Address List"
    if ($User.msExchHideFromAddressLists -ne $true) {$checkbox1.Checked = $False
     }
	if ($User.msExchHideFromAddressLists -eq $true) {$checkbox1.Checked = $True}
    $HideForm.Controls.Add($checkbox1)  
	
	$Name = $ID.name

#TextLable
$SearchNameLabel                = New-Object system.Windows.Forms.Label
$SearchNameLabel.text           = "You are editing user: "
$SearchNameLabel.AutoSize       = $true
$SearchNameLabel.width          = 25
$SearchNameLabel.height         = 20
$SearchNameLabel.ForeColor      = "#0000ff"
$SearchNameLabel.location       = New-Object System.Drawing.Point(30,20)
$SearchNameLabel.Font           = 'Microsoft Sans Serif,13'
$SearchNameLabel.Visible        = $True
$HideForm.Controls.Add($SearchNameLabel)

#TextLable
$SearchNameName                = New-Object system.Windows.Forms.Label
$SearchNameName.text           = "$Name"
$SearchNameName.AutoSize       = $true
$SearchNameName.width          = 50
$SearchNameName.height         = 20
$SearchNameName.ForeColor      = "#ff0000"
$SearchNameName.location       = New-Object System.Drawing.Point(200,20)
$SearchNameName.Font           = 'Microsoft Sans Serif,13,style=Bold'
$SearchNameName.Visible        = $True
$HideForm.Controls.Add($SearchNameName)

    write-host $checkbox1
 
    # Add an OK button
    $OKButton = new-object System.Windows.Forms.Button
    $OKButton.Location = new-object System.Drawing.Size(130,100)
    $OKButton.Size = new-object System.Drawing.Size(100,40)
    $OKButton.Text = "OK"
    $OKButton.Add_Click({
    
    If ($checkbox1.CheckState -eq "Checked"){Write-host $checkbox1.CheckState -fore Green
     Get-ADuser -Identity $CN -property msExchHideFromAddressLists |Set-ADObject -Replace @{msExchHideFromAddressLists=$true}}
    If ($checkbox1.CheckState -eq "Unchecked"){ Write-host $checkbox1.CheckState -fore Red
    Get-ADuser -Identity $CN -property msExchHideFromAddressLists |Set-ADObject -Replace @{msExchHideFromAddressLists=$False
    }}
    
    })
    $HideForm.Controls.Add($OKButton)
	$OKButton.Add_Click({
    $HideForm.Close()
    ActionForm
    })
    
 
    #Add a cancel button
    $CancelButton = new-object System.Windows.Forms.Button
    $CancelButton.Location = new-object System.Drawing.Size(255,100)
    $CancelButton.Size = new-object System.Drawing.Size(100,40)
    $CancelButton.Text = "Cancel"
    $CancelButton.Add_Click({$HideForm.Close()})
    $HideForm.Controls.Add($CancelButton)
    
    ###########  This is the important piece ##############
    #                                                     #
    # Do something when the state of the checkbox changes #
    #######################################################
   <#
    $checkbox1.Add_CheckStateChanged({
   If ($checkbox1.CheckState -eq $False) { Write-host  "Hidden From GAL" $checkbox1.CheckState -fore Cyan}
   else { Write-host "Hidden From GAL" $checkbox1.CheckState -fore Yellow}
    })
    #>
    
    # Activate the form
    $HideForm.Add_Shown({$HideForm.Activate()})
    [void] $HideForm.ShowDialog() 
    }




####Bulk Actions Form #######
function BulkActionsForm {

}



############Form Functions End


# Init PowerShell GUI
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$script =  $scriptPath + "\" + $MyInvocation.MyCommand.name 
AddressForm