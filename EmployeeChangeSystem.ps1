################################################################################ 
#
#  Name    : C:\Users\kody.oneill\Documents\Script1\Form1.ps1  
#  Version : 0.1
#  Author  : Kody O'Neill
#  Date    : 9/7/2021
#
 #  Generated with ConvertForm module version 2.0.0
#  PowerShell version 5.1.18362.1171
################################################################################



# Loading external assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ECSmain = New-Object System.Windows.Forms.Form

$EmployeeChangeTabs = New-Object System.Windows.Forms.TabControl
$NewHireTab = New-Object System.Windows.Forms.TabPage
$TransferTab = New-Object System.Windows.Forms.TabPage
$TerminationTab = New-Object System.Windows.Forms.TabPage
$TerminationUsernameBox = New-Object System.Windows.Forms.TextBox
$TerminationTicketNumberBox = New-Object System.Windows.Forms.TextBox
$TermTicketNumberLabel = New-Object System.Windows.Forms.Label
$TermUsernameLabel = New-Object System.Windows.Forms.Label
$TerminationSubmitButton = New-Object System.Windows.Forms.Button
$TransferSubmitButton = New-Object System.Windows.Forms.Button
$TransferUsernameLabel = New-Object System.Windows.Forms.Label
$TransferTicketNumberLabel = New-Object System.Windows.Forms.Label
$TransferUsernameBox = New-Object System.Windows.Forms.TextBox
$TransferTicketNumberBox = New-Object System.Windows.Forms.TextBox
$TransferMirrorLabel = New-Object System.Windows.Forms.Label
$TransferMirrorBox = New-Object System.Windows.Forms.TextBox
$TransferDescriptionLabel = New-Object System.Windows.Forms.Label
$TransferDescription = New-Object System.Windows.Forms.TextBox
$TransferTransferRadioButton = New-Object System.Windows.Forms.RadioButton
$TransferRehireRadioButton = New-Object System.Windows.Forms.RadioButton
$NewHireRosemountRadioButton = New-Object System.Windows.Forms.RadioButton
$NewHirePlant1RadioButton = New-Object System.Windows.Forms.RadioButton
$NewHireDescriptionLabel = New-Object System.Windows.Forms.Label
$NewHireDescriptionBox = New-Object System.Windows.Forms.TextBox
$NewHireMirrorLabel = New-Object System.Windows.Forms.Label
$NewHireMirrorBox = New-Object System.Windows.Forms.TextBox
$NewHireSubmitButton = New-Object System.Windows.Forms.Button
$NewHireTicketNumberLabel = New-Object System.Windows.Forms.Label
$NewHireTicketNumberBox = New-Object System.Windows.Forms.TextBox
$NewHireMiddleInitialLabel = New-Object System.Windows.Forms.Label
$NewHireMiddleInitialBox = New-Object System.Windows.Forms.TextBox
$NewHireLastNameLabel = New-Object System.Windows.Forms.Label
$NewHireFirstNameLabel = New-Object System.Windows.Forms.Label
$NewHireLastNameBox = New-Object System.Windows.Forms.TextBox
$NewHireFirstNameBox = New-Object System.Windows.Forms.TextBox
$NewHireCaryRadioButton = New-Object System.Windows.Forms.RadioButton
$NewHireHQRadioButton = New-Object System.Windows.Forms.RadioButton
$NewHireRaleighRadioButton = New-Object System.Windows.Forms.RadioButton
$NewHireBrooklynParkRadioButton = New-Object System.Windows.Forms.RadioButton
$NewHireNashuaRadioButton = New-Object System.Windows.Forms.RadioButton
$NewHirePlymouthRadioButton = New-Object System.Windows.Forms.RadioButton
#
# EmployeeChangeTabs
#
$EmployeeChangeTabs.Controls.Add($NewHireTab)
$EmployeeChangeTabs.Controls.Add($TransferTab)
$EmployeeChangeTabs.Controls.Add($TerminationTab)
$EmployeeChangeTabs.Location = New-Object System.Drawing.Point(12, 12)
$EmployeeChangeTabs.Name = "EmployeeChangeTabs"
$EmployeeChangeTabs.SelectedIndex = 0
$EmployeeChangeTabs.Size = New-Object System.Drawing.Size(443, 610)
$EmployeeChangeTabs.TabIndex = 0
#
# NewHireTab
#
$NewHireTab.Controls.Add($NewHireNashuaRadioButton)
$NewHireTab.Controls.Add($NewHirePlymouthRadioButton)
$NewHireTab.Controls.Add($NewHireRaleighRadioButton)
$NewHireTab.Controls.Add($NewHireBrooklynParkRadioButton)
$NewHireTab.Controls.Add($NewHireCaryRadioButton)
$NewHireTab.Controls.Add($NewHireHQRadioButton)
$NewHireTab.Controls.Add($NewHireMiddleInitialLabel)
$NewHireTab.Controls.Add($NewHireMiddleInitialBox)
$NewHireTab.Controls.Add($NewHireLastNameLabel)
$NewHireTab.Controls.Add($NewHireFirstNameLabel)
$NewHireTab.Controls.Add($NewHireLastNameBox)
$NewHireTab.Controls.Add($NewHireFirstNameBox)
$NewHireTab.Controls.Add($NewHireRosemountRadioButton)
$NewHireTab.Controls.Add($NewHirePlant1RadioButton)
$NewHireTab.Controls.Add($NewHireDescriptionLabel)
$NewHireTab.Controls.Add($NewHireDescriptionBox)
$NewHireTab.Controls.Add($NewHireMirrorLabel)
$NewHireTab.Controls.Add($NewHireMirrorBox)
$NewHireTab.Controls.Add($NewHireSubmitButton)
$NewHireTab.Controls.Add($NewHireTicketNumberLabel)
$NewHireTab.Controls.Add($NewHireTicketNumberBox)
$NewHireTab.Location = New-Object System.Drawing.Point(4, 24)
$NewHireTab.Name = "NewHireTab"
$NewHireTab.Padding = New-Object System.Windows.Forms.Padding(3)
$NewHireTab.Size = New-Object System.Drawing.Size(435, 582)
$NewHireTab.TabIndex = 0
$NewHireTab.Text = "NewHire"
$NewHireTab.UseVisualStyleBackColor = $true



#
# TransferTab
#
$TransferTab.Controls.Add($TransferRehireRadioButton)
$TransferTab.Controls.Add($TransferTransferRadioButton)
$TransferTab.Controls.Add($TransferDescriptionLabel)
$TransferTab.Controls.Add($TransferDescription)
$TransferTab.Controls.Add($TransferMirrorLabel)
$TransferTab.Controls.Add($TransferMirrorBox)
$TransferTab.Controls.Add($TransferSubmitButton)
$TransferTab.Controls.Add($TransferUsernameLabel)
$TransferTab.Controls.Add($TransferTicketNumberLabel)
$TransferTab.Controls.Add($TransferUsernameBox)
$TransferTab.Controls.Add($TransferTicketNumberBox)
$TransferTab.Location = New-Object System.Drawing.Point(4, 24)
$TransferTab.Name = "TransferTab"
$TransferTab.Padding = New-Object System.Windows.Forms.Padding(3)
$TransferTab.Size = New-Object System.Drawing.Size(435, 582)
$TransferTab.TabIndex = 1
$TransferTab.Text = "Transfer"
$TransferTab.UseVisualStyleBackColor = $true
#
# TerminationTab
#
$TerminationTab.Controls.Add($TerminationSubmitButton)
$TerminationTab.Controls.Add($TermUsernameLabel)
$TerminationTab.Controls.Add($TermTicketNumberLabel)
$TerminationTab.Controls.Add($TerminationUsernameBox)
$TerminationTab.Controls.Add($TerminationTicketNumberBox)
$TerminationTab.Location = New-Object System.Drawing.Point(4, 24)
$TerminationTab.Name = "TerminationTab"
$TerminationTab.Padding = New-Object System.Windows.Forms.Padding(3)
$TerminationTab.Size = New-Object System.Drawing.Size(435, 582)
$TerminationTab.TabIndex = 2
$TerminationTab.Text = "Termination"
$TerminationTab.UseVisualStyleBackColor = $true
#
# TerminationUsernameBox
#
$TerminationUsernameBox.Location = New-Object System.Drawing.Point(90, 46)
$TerminationUsernameBox.Name = "TerminationUsernameBox"
$TerminationUsernameBox.Size = New-Object System.Drawing.Size(100, 23)
$TerminationUsernameBox.TabIndex = 1
#
# TerminationTicketNumberBox
#
$TerminationTicketNumberBox.Location = New-Object System.Drawing.Point(90, 17)
$TerminationTicketNumberBox.Name = "TerminationTicketNumberBox"
$TerminationTicketNumberBox.Size = New-Object System.Drawing.Size(100, 23)
$TerminationTicketNumberBox.TabIndex = 0
#
# TermTicketNumberLabel
#
$TermTicketNumberLabel.AutoSize = $true
$TermTicketNumberLabel.Location = New-Object System.Drawing.Point(12, 25)
$TermTicketNumberLabel.Name = "TermTicketNumberLabel"
$TermTicketNumberLabel.Size = New-Object System.Drawing.Size(45, 15)
$TermTicketNumberLabel.TabIndex = 2
$TermTicketNumberLabel.Text = "Ticket#"
#
# TermUsernameLabel
#
$TermUsernameLabel.AutoSize = $true
$TermUsernameLabel.Location = New-Object System.Drawing.Point(12, 54)
$TermUsernameLabel.Name = "TermUsernameLabel"
$TermUsernameLabel.Size = New-Object System.Drawing.Size(60, 15)
$TermUsernameLabel.TabIndex = 3
$TermUsernameLabel.Text = "Username"
#
# TerminationSubmitButton
#
$TerminationSubmitButton.Location = New-Object System.Drawing.Point(218, 17)
$TerminationSubmitButton.Name = "TerminationSubmitButton"
$TerminationSubmitButton.Size = New-Object System.Drawing.Size(80, 52)
$TerminationSubmitButton.TabIndex = 7
$TerminationSubmitButton.Text = "Submit"
$TerminationSubmitButton.UseVisualStyleBackColor = $true
#
# TransferSubmitButton
#
$TransferSubmitButton.Location = New-Object System.Drawing.Point(215, 17)
$TransferSubmitButton.Name = "TransferSubmitButton"
$TransferSubmitButton.Size = New-Object System.Drawing.Size(80, 52)
$TransferSubmitButton.TabIndex = 12
$TransferSubmitButton.Text = "Submit"
$TransferSubmitButton.UseVisualStyleBackColor = $true
#
# TransferUsernameLabel
#
$TransferUsernameLabel.AutoSize = $true
$TransferUsernameLabel.Location = New-Object System.Drawing.Point(9, 54)
$TransferUsernameLabel.Name = "TransferUsernameLabel"
$TransferUsernameLabel.Size = New-Object System.Drawing.Size(60, 15)
$TransferUsernameLabel.TabIndex = 11
$TransferUsernameLabel.Text = "Username"


#
# TransferTicketNumberLabel
#
$TransferTicketNumberLabel.AutoSize = $true
$TransferTicketNumberLabel.Location = New-Object System.Drawing.Point(9, 25)
$TransferTicketNumberLabel.Name = "TransferTicketNumberLabel"
$TransferTicketNumberLabel.Size = New-Object System.Drawing.Size(45, 15)
$TransferTicketNumberLabel.TabIndex = 10
$TransferTicketNumberLabel.Text = "Ticket#"
#
# TransferUsernameBox
#
$TransferUsernameBox.Location = New-Object System.Drawing.Point(87, 46)
$TransferUsernameBox.Name = "TransferUsernameBox"
$TransferUsernameBox.Size = New-Object System.Drawing.Size(100, 23)
$TransferUsernameBox.TabIndex = 9
#
# TransferTicketNumberBox
#
$TransferTicketNumberBox.Location = New-Object System.Drawing.Point(87, 17)
$TransferTicketNumberBox.Name = "TransferTicketNumberBox"
$TransferTicketNumberBox.Size = New-Object System.Drawing.Size(100, 23)
$TransferTicketNumberBox.TabIndex = 8
#
# TransferMirrorLabel
#
$TransferMirrorLabel.AutoSize = $true
$TransferMirrorLabel.Location = New-Object System.Drawing.Point(9, 83)
$TransferMirrorLabel.Name = "TransferMirrorLabel"
$TransferMirrorLabel.Size = New-Object System.Drawing.Size(40, 15)
$TransferMirrorLabel.TabIndex = 14
$TransferMirrorLabel.Text = "Mirror"
#
# TransferMirrorBox
#
$TransferMirrorBox.Location = New-Object System.Drawing.Point(87, 75)
$TransferMirrorBox.Name = "TransferMirrorBox"
$TransferMirrorBox.Size = New-Object System.Drawing.Size(100, 23)
$TransferMirrorBox.TabIndex = 13
#
# TransferDescriptionLabel
#
$TransferDescriptionLabel.AutoSize = $true
$TransferDescriptionLabel.Location = New-Object System.Drawing.Point(9, 162)
$TransferDescriptionLabel.Name = "TransferDescriptionLabel"
$TransferDescriptionLabel.Size = New-Object System.Drawing.Size(67, 15)
$TransferDescriptionLabel.TabIndex = 16
$TransferDescriptionLabel.Text = "Description"
#
# TransferDescription
#
$TransferDescription.Location = New-Object System.Drawing.Point(87, 154)
$TransferDescription.Multiline = $true
$TransferDescription.Name = "TransferDescription"
$TransferDescription.Size = New-Object System.Drawing.Size(208, 345)
$TransferDescription.TabIndex = 15
#
# TransferTransferRadioButton
#
$TransferTransferRadioButton.AutoSize = $true
$TransferTransferRadioButton.Location = New-Object System.Drawing.Point(9, 118)
$TransferTransferRadioButton.Name = "TransferTransferRadioButton"
$TransferTransferRadioButton.Size = New-Object System.Drawing.Size(66, 19)
$TransferTransferRadioButton.TabIndex = 17
$TransferTransferRadioButton.TabStop = $true
$TransferTransferRadioButton.Text = "Transfer"
$TransferTransferRadioButton.UseVisualStyleBackColor = $true
#
# TransferRehireRadioButton
#
$TransferRehireRadioButton.AutoSize = $true
$TransferRehireRadioButton.Location = New-Object System.Drawing.Point(114, 118)
$TransferRehireRadioButton.Name = "TransferRehireRadioButton"
$TransferRehireRadioButton.Size = New-Object System.Drawing.Size(58, 19)
$TransferRehireRadioButton.TabIndex = 18
$TransferRehireRadioButton.TabStop = $true
$TransferRehireRadioButton.Text = "Rehire"
$TransferRehireRadioButton.UseVisualStyleBackColor = $true
#
# NewHireRosemountRadioButton
#
$NewHireRosemountRadioButton.AutoSize = $true
$NewHireRosemountRadioButton.Location = New-Object System.Drawing.Point(326, 84)
$NewHireRosemountRadioButton.Name = "NewHireRosemountRadioButton"
$NewHireRosemountRadioButton.Size = New-Object System.Drawing.Size(86, 19)
$NewHireRosemountRadioButton.TabIndex = 29
$NewHireRosemountRadioButton.TabStop = $true
$NewHireRosemountRadioButton.Text = "Rosemount"
$NewHireRosemountRadioButton.UseVisualStyleBackColor = $true
#
# NewHirePlant1RadioButton
#
$NewHirePlant1RadioButton.AutoSize = $true
$NewHirePlant1RadioButton.Location = New-Object System.Drawing.Point(218, 84)
$NewHirePlant1RadioButton.Name = "NewHirePlant1RadioButton"
$NewHirePlant1RadioButton.Size = New-Object System.Drawing.Size(88, 19)
$NewHirePlant1RadioButton.TabIndex = 28
$NewHirePlant1RadioButton.TabStop = $true
$NewHirePlant1RadioButton.Text = "MPL Plant 1"
$NewHirePlant1RadioButton.UseVisualStyleBackColor = $true
#
# NewHireDescriptionLabel
#
$NewHireDescriptionLabel.AutoSize = $true
$NewHireDescriptionLabel.Location = New-Object System.Drawing.Point(12, 201)
$NewHireDescriptionLabel.Name = "NewHireDescriptionLabel"
$NewHireDescriptionLabel.Size = New-Object System.Drawing.Size(67, 15)
$NewHireDescriptionLabel.TabIndex = 27
$NewHireDescriptionLabel.Text = "Description"
#
# NewHireDescriptionBox
#
$NewHireDescriptionBox.Location = New-Object System.Drawing.Point(85, 193)
$NewHireDescriptionBox.Multiline = $true
$NewHireDescriptionBox.Name = "NewHireDescriptionBox"
$NewHireDescriptionBox.Size = New-Object System.Drawing.Size(208, 307)
$NewHireDescriptionBox.TabIndex = 26
#
# NewHireMirrorLabel
#
$NewHireMirrorLabel.AutoSize = $true
$NewHireMirrorLabel.Location = New-Object System.Drawing.Point(12, 56)
$NewHireMirrorLabel.Name = "NewHireMirrorLabel"
$NewHireMirrorLabel.Size = New-Object System.Drawing.Size(40, 15)
$NewHireMirrorLabel.TabIndex = 25
$NewHireMirrorLabel.Text = "Mirror"
#
# NewHireMirrorBox
#
$NewHireMirrorBox.Location = New-Object System.Drawing.Point(90, 48)
$NewHireMirrorBox.Name = "NewHireMirrorBox"
$NewHireMirrorBox.Size = New-Object System.Drawing.Size(100, 23)
$NewHireMirrorBox.TabIndex = 24
#
# NewHireSubmitButton
#
$NewHireSubmitButton.Location = New-Object System.Drawing.Point(218, 18)
$NewHireSubmitButton.Name = "NewHireSubmitButton"
$NewHireSubmitButton.Size = New-Object System.Drawing.Size(80, 52)
$NewHireSubmitButton.TabIndex = 23
$NewHireSubmitButton.Text = "Submit"
$NewHireSubmitButton.UseVisualStyleBackColor = $true
#
# NewHireTicketNumberLabel
#
$NewHireTicketNumberLabel.AutoSize = $true
$NewHireTicketNumberLabel.Location = New-Object System.Drawing.Point(12, 26)
$NewHireTicketNumberLabel.Name = "NewHireTicketNumberLabel"
$NewHireTicketNumberLabel.Size = New-Object System.Drawing.Size(45, 15)
$NewHireTicketNumberLabel.TabIndex = 21
$NewHireTicketNumberLabel.Text = "Ticket#"
#
# NewHireTicketNumberBox
#
$NewHireTicketNumberBox.Location = New-Object System.Drawing.Point(90, 18)
$NewHireTicketNumberBox.Name = "NewHireTicketNumberBox"
$NewHireTicketNumberBox.Size = New-Object System.Drawing.Size(100, 23)
$NewHireTicketNumberBox.TabIndex = 19
#
# NewHireMiddleInitialLabel
#
$NewHireMiddleInitialLabel.AutoSize = $true
$NewHireMiddleInitialLabel.Location = New-Object System.Drawing.Point(12, 144)
$NewHireMiddleInitialLabel.Name = "NewHireMiddleInitialLabel"
$NewHireMiddleInitialLabel.Size = New-Object System.Drawing.Size(76, 15)
$NewHireMiddleInitialLabel.TabIndex = 35
$NewHireMiddleInitialLabel.Text = "Middle Initial"
#
# NewHireMiddleInitialBox
#
$NewHireMiddleInitialBox.Location = New-Object System.Drawing.Point(90, 136)
$NewHireMiddleInitialBox.Name = "NewHireMiddleInitialBox"
$NewHireMiddleInitialBox.Size = New-Object System.Drawing.Size(100, 23)
$NewHireMiddleInitialBox.TabIndex = 34
#
# NewHireLastNameLabel
#
$NewHireLastNameLabel.AutoSize = $true
$NewHireLastNameLabel.Location = New-Object System.Drawing.Point(12, 115)
$NewHireLastNameLabel.Name = "NewHireLastNameLabel"
$NewHireLastNameLabel.Size = New-Object System.Drawing.Size(63, 15)
$NewHireLastNameLabel.TabIndex = 33
$NewHireLastNameLabel.Text = "Last Name"
#
# NewHireFirstNameLabel
#
$NewHireFirstNameLabel.AutoSize = $true
$NewHireFirstNameLabel.Location = New-Object System.Drawing.Point(12, 86)
$NewHireFirstNameLabel.Name = "NewHireFirstNameLabel"
$NewHireFirstNameLabel.Size = New-Object System.Drawing.Size(64, 15)
$NewHireFirstNameLabel.TabIndex = 32
$NewHireFirstNameLabel.Text = "First Name"
#
# NewHireLastNameBox
#
$NewHireLastNameBox.Location = New-Object System.Drawing.Point(90, 107)
$NewHireLastNameBox.Name = "NewHireLastNameBox"
$NewHireLastNameBox.Size = New-Object System.Drawing.Size(100, 23)
$NewHireLastNameBox.TabIndex = 31
#
# NewHireFirstNameBox
#
$NewHireFirstNameBox.Location = New-Object System.Drawing.Point(90, 78)
$NewHireFirstNameBox.Name = "NewHireFirstNameBox"
$NewHireFirstNameBox.Size = New-Object System.Drawing.Size(100, 23)
$NewHireFirstNameBox.TabIndex = 30
#
# NewHireCaryRadioButton
#
$NewHireCaryRadioButton.AutoSize = $true
$NewHireCaryRadioButton.Location = New-Object System.Drawing.Point(326, 107)
$NewHireCaryRadioButton.Name = "NewHireCaryRadioButton"
$NewHireCaryRadioButton.Size = New-Object System.Drawing.Size(49, 19)
$NewHireCaryRadioButton.TabIndex = 37
$NewHireCaryRadioButton.TabStop = $true
$NewHireCaryRadioButton.Text = "Cary"
$NewHireCaryRadioButton.UseVisualStyleBackColor = $true
#
# NewHireHQRadioButton
#
$NewHireHQRadioButton.AutoSize = $true
$NewHireHQRadioButton.Location = New-Object System.Drawing.Point(218, 107)
$NewHireHQRadioButton.Name = "NewHireHQRadioButton"
$NewHireHQRadioButton.Size = New-Object System.Drawing.Size(70, 19)
$NewHireHQRadioButton.TabIndex = 36
$NewHireHQRadioButton.TabStop = $true
$NewHireHQRadioButton.Text = "MPL HQ"
$NewHireHQRadioButton.UseVisualStyleBackColor = $true
#
# NewHireRaleighRadioButton
#
$NewHireRaleighRadioButton.AutoSize = $true
$NewHireRaleighRadioButton.Location = New-Object System.Drawing.Point(326, 132)
$NewHireRaleighRadioButton.Name = "NewHireRaleighRadioButton"
$NewHireRaleighRadioButton.Size = New-Object System.Drawing.Size(64, 19)
$NewHireRaleighRadioButton.TabIndex = 39
$NewHireRaleighRadioButton.TabStop = $true
$NewHireRaleighRadioButton.Text = "Raleigh"
$NewHireRaleighRadioButton.UseVisualStyleBackColor = $true
#
# NewHireBrooklynParkRadioButton
#
$NewHireBrooklynParkRadioButton.AutoSize = $true
$NewHireBrooklynParkRadioButton.Location = New-Object System.Drawing.Point(218, 132)
$NewHireBrooklynParkRadioButton.Name = "NewHireBrooklynParkRadioButton"
$NewHireBrooklynParkRadioButton.Size = New-Object System.Drawing.Size(98, 19)
$NewHireBrooklynParkRadioButton.TabIndex = 38
$NewHireBrooklynParkRadioButton.TabStop = $true
$NewHireBrooklynParkRadioButton.Text = "Brooklyn Park"
$NewHireBrooklynParkRadioButton.UseVisualStyleBackColor = $true
#
# NewHireNashuaRadioButton
#
$NewHireNashuaRadioButton.AutoSize = $true
$NewHireNashuaRadioButton.Location = New-Object System.Drawing.Point(326, 157)
$NewHireNashuaRadioButton.Name = "NewHireNashuaRadioButton"
$NewHireNashuaRadioButton.Size = New-Object System.Drawing.Size(65, 19)
$NewHireNashuaRadioButton.TabIndex = 41
$NewHireNashuaRadioButton.TabStop = $true
$NewHireNashuaRadioButton.Text = "Nashua"
$NewHireNashuaRadioButton.UseVisualStyleBackColor = $true
#
# NewHirePlymouthRadioButton
#
$NewHirePlymouthRadioButton.AutoSize = $true
$NewHirePlymouthRadioButton.Location = New-Object System.Drawing.Point(218, 157)
$NewHirePlymouthRadioButton.Name = "NewHirePlymouthRadioButton"
$NewHirePlymouthRadioButton.Size = New-Object System.Drawing.Size(77, 19)
$NewHirePlymouthRadioButton.TabIndex = 40
$NewHirePlymouthRadioButton.TabStop = $true
$NewHirePlymouthRadioButton.Text = "Plymouth"
$NewHirePlymouthRadioButton.UseVisualStyleBackColor = $true
#
# ECSmain
#
$ECSmain.ClientSize = New-Object System.Drawing.Size(467, 634)
$ECSmain.Controls.Add($EmployeeChangeTabs)
$ECSmain.Name = "ECSmain"
$ECSmain.Text = "EmployeeChangeSystem"


#-----------------------------------------------------------------------------------------------------


#Setting defualt Radio Button
$NewHireHQRadioButton.Checked = $True
$NewHireSubmitButton.Add_Click( {

    if([string]::IsNullOrEmpty($NewHireTicketNumberBox.Text) -or [string]::IsNullOrEmpty($NewHireMirrorBox.Text) -or [string]::IsNullOrEmpty($NewHireFirstNameBox.Text) -or [string]::IsNullOrEmpty($NewHireLastNameBox.Text) -or [string]::IsNullOrEmpty($NewHireMiddleInitialBox.Text)){
        [void][System.Windows.Forms.MessageBox]::Show("Not all input boxes filled out")
    }else{
        SubmitNewHireForm
        #clears fields after new hire is complete
        $NewHireTicketNumberBox.Text = ""
        $NewHireMirrorBox.Text = ""
        $NewHireFirstNameBox.Text = ""
        $NewHireLastNameBox.Text = ""
        $NewHireMiddleInitialBox.Text = ""
        $NewHireDescriptionBox.Text = ""
        $NewHireHQRadioButton.Checked = $True
    
        [void][System.Windows.Forms.MessageBox]::Show("New Hire Complete")
    }
        
    })
    $TransferTransferRadioButton.Checked = $true
#Setting default Radio Button
$TransferSubmitButton.Add_Click( {
    
    if([string]::IsNullOrEmpty($TransferTicketNumberBox.Text) -or [string]::IsNullOrEmpty($TransferUsernameBox.Text) -or [string]::IsNullOrEmpty($TransferMirrorBox.Text)){
        [void][System.Windows.Forms.MessageBox]::Show("Not all input boxes filled out")
    }else {
        SubmitTransferForm
        $TransferTicketNumberBox.Text = ""
        $TransferUsernameBox.Text = ""
        $TransferMirrorBox.Text = ""
        $TransferTransferRadioButton.Checked = $True
        $TransferDescriptionBox.Text = ""
    }
    })

$TerminationSubmitButton.Add_Click( {

    if([string]::IsNullOrEmpty($TerminationTicketNumberBox.Text) -or [string]::IsNullOrEmpty($TerminationUsernameBox.Text)){
        [void][System.Windows.Forms.MessageBox]::Show("Not all input boxes filled out")
    }else {
        SubmitTerminationForm
        $TerminationTicketNumberBox.Text = ""
        $TerminationUsernameBox.Text = ""
    }
        
    })

    #Setting globale vairables and getting configuration from JSON file
    $GetDate = Get-Date -Format "_dddd_MM.dd.yyyy_HH_mm_ss"
    $LoggedOnUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $LoggedOnUser = $LoggedOnUser.Split("\")
    $LoggedOnUser = $LoggedOnUser[1] 

    #Import Json configuration file
    $JsonFilePath = Get-Content -Raw -Path #Path to JSON Cofnig file
    $JsonFile = $JsonFilePath | ConvertFrom-Json

    #Pulling data from file
    $HelpDeskEmail = $JsonFile[0].HelpdeskEmail
    $DefaultPasswordUnenecrypted = $JsonFile[0].DefaultPassword
    $CommonLoggingCore = $JsonFile[0].CommonLoggingCorePath
    $CommingLogging = $JsonFile[0].CommonLoggingPath
    $iTextIO = $JsonFile[0].iTextIOPath
    $iTextKernal = $JsonFile[0].iTextKernalPath
    $iTextForm = $JsonFile[0].iTextFormPath
    $iTextLayout = $JsonFile[0].iTextLayoutPath
    $BouncyCastle = $JsonFile[0].iTextBouncyCastlePath
    $NewHireHirePDF = $JsonFile[0].NewHireDefaultPDFPath
    $NewHireOutput = $JsonFile[0].NewHirePDFOutputPath
    $TransferPDF = $JsonFile[0].TransferDefaultPDFPath
    $TransferOutput = $JsonFile[0].TransferPDFOutputPath
    $TerminationPDF = $JsonFile[0].TerminationDefaultPDFPath
    $TerminationOutput = $JsonFile[0].TermiationPDFOutputPath
    $SpecialApprover1 = $JsonFile[0].LocationSpecificApprover1
    $SpecialApprover2 = $JsonFile[0].LocationSpecificApprover2
    $DefaultMESSpecialApprover = $JsonFile[0].LocationSpecificDefault
    $EmployeeChangeLog = $JsonFile[0].EmployeeChangeLogPath
    $GroupOwnerApprovalFile = $JsonFile[0].ApprovalCSVFilePath
    $RelayServer = $JsonFile[0].RelayServer
    $SMTPServer = $JsonFile[0].SMTPServer


    #Adding necessary dll's
    Add-Type -Path $CommonLoggingCore
    Add-Type -Path $CommingLogging
    Add-Type -Path $iTextIO
    Add-Type -Path $iTextKernal
    Add-Type -Path $iTextForm
    Add-Type -Path $iTextLayout
    Add-Type -Path $BouncyCastle

function NewHirePDFForm($FinalUserName,$NewHireGroups,$NewHireDisplayName,$MirrorTitle,$NewHirePDFManagerDisplayName,$MirrorDepartment,$UserNewOfficeLocation,$UserCopy,$FileName, $NewHireAdditionalNotes) {

       #Initialize PDFReader to access PDF
        $Reader = [iText.Kernel.Pdf.PdfReader]::new($NewHireHirePDF)
        $Reader.SetUnethicalReading($True)
    
       #Initialize PDFWriter to write our modified PDF to. Needs to be different name than above
        $Writer = [iText.Kernel.Pdf.PdfWriter]::new("$NewHireOutput\$FileName.pdf")
    
        #Now we generate a PdfDocument object to tie our PdfReader and PdfWriter together and allow us to start our work
        $NewHirePdfDoc = [iText.Kernel.Pdf.PdfDocument]::new($Reader, $Writer)
        $Form = [iText.Forms.PdfAcroForm]::getAcroForm($NewHirePdfDoc, $True)
        $Fields = $Form.getFormFields()

        #Added groups to AD field in form
        $NewHireADGroupField = ($Fields | Where-Object {$_.key -eq "Active Directory Changes"}).Value
        $NewHireADGroupField.SetValue($NewHireGroups)
    
        #Added additional notes in form
        $NewHireADGroupField = ($Fields | Where-Object {$_.key -eq "Active Directory Changes"}).Value
        $NewHireADGroupField.SetValue($NewHireAdditionalNotes)

        #Fill out forms on the right hand side of form
        $NewHirePDFNameBox = ($Fields | Where-Object {$_.key -eq "Name"}).Value
        $NewHirePDFNameBox.SetValue($NewHireDisplayName)

        $NewHirePDFUsernameBox = ($Fields | Where-Object {$_.key -eq "Username"}).Value
        $NewHirePDFUsernameBox.SetValue($FinalUserName)
    
        $NewHirePDFPositionBox = ($Fields | Where-Object {$_.key -eq "Position"}).Value
        $NewHirePDFPositionBox.SetValue($MirrorTitle)
    
        $NewHirePDFDepartmentBox = ($Fields | Where-Object {$_.key -eq "Department"}).Value
        $NewHirePDFDepartmentBox.SetValue($MirrorDepartment)
    
        $NewHirePDFManagerBox = ($Fields | Where-Object {$_.key -eq "Manager"}).Value
        $NewHirePDFManagerBox.SetValue($NewHirePDFManagerDisplayName)
    
        $NewHirePDFOfficeBox = ($Fields | Where-Object {$_.key -eq "Office"}).Value
        $NewHirePDFOfficeBox.SetValue($UserNewOfficeLocation)
    
        $NewHirePDFMirrorBox = ($Fields | Where-Object {$_.key -eq "Mirror"}).Value
        $NewHirePDFMirrorBox.SetValue($UserCopy)
    
        #Checking check boxes that the script completes
        $NewHireActiveDirectoryAccountCheckBox = ($Fields | Where-Object {$_.key -eq "Active Directory Account"}).Value
        $NewHireActiveDirectoryAccountCheckBox.SetValue($True)

        $NewHireLowerCaseLettersCheckBox = ($Fields | Where-Object {$_.key -eq "Lowercase Letters for Login"}).Value
        $NewHireLowerCaseLettersCheckBox.SetValue($True)

        $NewHireProperOUCheckBox = ($Fields | Where-Object {$_.key -eq "In Proper OU"}).Value
        $NewHireProperOUCheckBox.SetValue($True)

        $NewHireSecurityGroupsCheckBox = ($Fields | Where-Object {$_.key -eq "Mirrored DistributionSecurity Groups Present"}).Value
        $NewHireSecurityGroupsCheckBox.SetValue($True)

        $NewHireHomeDriveCheckBox = ($Fields | Where-Object {$_.key -eq "Home Directory Created All but shop"}).Value
        $NewHireHomeDriveCheckBox.SetValue($True)

        $NewHireOrgFieldsCheckBox = ($Fields | Where-Object {$_.key -eq "Organization Fields Title Office Manager etc"}).Value
        $NewHireOrgFieldsCheckBox.SetValue($True)

        $NewHireCoarseGranAXGroupsCheckBox = ($Fields | Where-Object {$_.key -eq "Get approval for all APPAX and coarse grain groups"}).Value
        $NewHireCoarseGranAXGroupsCheckBox.SetValue($True)

    $NewHirePdfDoc.Close()
    
}

function TransferPDFForm($UserDisplayName, $Username, $UserPosition, $UserDepartment, $UserManager, $UserOffice, $UserMirror, $UserGroupChanges, $FileName, $AdditionalNotes){
    #Initialize PDFReader to access PDF
    $Reader = [iText.Kernel.Pdf.PdfReader]::new($TransferPDF)
    $Reader.SetUnethicalReading($True)

    #Initialize PDFWriter to write our modified PDF to. Needs to be different name than above
    $Writer = [iText.Kernel.Pdf.PdfWriter]::new("$TransferOutput\$FileName.pdf")

    #Now we generate a PdfDocument object to tie our PdfReader and PdfWriter together and allow us to start our work
    $TransferPdfDoc = [iText.Kernel.Pdf.PdfDocument]::new($Reader, $Writer)
    $Form = [iText.Forms.PdfAcroForm]::getAcroForm($TransferPdfDoc, $True)
    $Fields = $Form.getFormFields()

    #Adding name to the Name box
    $TransferName = ($Fields | Where-Object {$_.key -eq "Name"}).Value
    $TransferName.SetValue($UserDisplayName)

    #Adding username to the Username box
    $TransferUsername = ($Fields | Where-Object {$_.key -eq "Username"}).Value
    $TransferUsername.SetValue($Username)

    #Adding position to the Position box
    $TransferPosition = ($Fields | Where-Object {$_.key -eq "Position"}).Value
    $TransferPosition.SetValue($UserPosition)

    #Adding department to the department box
    $TransferDepartment = ($Fields | Where-Object {$_.key -eq "Department"}).Value
    $TransferDepartment.SetValue($UserDepartment)

    #Adding manager to the manager box
    $TransferManager = ($Fields | Where-Object {$_.key -eq "Manager"}).Value
    $TransferManager.SetValue($UserManager)

    #Adding office to the office box
    $TransferOffice = ($Fields | Where-Object {$_.key -eq "Office"}).Value
    $TransferOffice.SetValue($UserOffice)

    #Adding mirror to Mirror box
    $TransferMirror = ($Fields | Where-Object {$_.key -eq "Mirror"}).Value
    $TransferMirror.SetValue($UserMirror)

    #Adding group changes to text box
    $TransferGroupChanges = ($Fields | Where-Object {$_.key -eq "Active Directory Changes"}).Value
    $TransferGroupChanges.SetValue($UserGroupChanges)

    #Adding additional notes to text box
    $TransferAdditionalNotes = ($Fields | Where-Object {$_.key -eq "Additional Notes"}).Value
    $TransferAdditionalNotes.SetValue($AdditionalNotes)


    #Checking check boxes for tasks completed by the script
    $TransferActiveDirectoryAccount = ($Fields | Where-Object {$_.key -eq "Active Directory Account"}).Value
    $TransferActiveDirectoryAccount.SetValue($True)

    $TransferedMirroredGroups = ($Fields | Where-Object {$_.key -eq "Mirrored DistributionSecurity Groups Present"}).Value
    $TransferedMirroredGroups.SetValue($True)

    $TransferMovedOU = ($Fields | Where-Object {$_.key -eq "Moved to Proper OU"}).Value
    $TransferMovedOU.SetValue($True)

    $TransferOrgFieldsUpdated = ($Fields | Where-Object {$_.key -eq "Verify Organizational Fields are Updated"}).Value
    $TransferOrgFieldsUpdated.SetValue($True)

    $TransferAXCoarseGranApproval = ($Fields | Where-Object {$_.key -eq "Get approval for all APPAX and coarse grain groups"}).Value
    $TransferAXCoarseGranApproval.SetValue($True)


    $TransferPdfDoc.Close()
}
function TerminationPDFForm($TerminatedUserName,$RemovedGroups,$TerminationDisplayName,$PDFTerminationManagerDisplayName,$TerminationDepartment,$TerminationOfficeLocation,$FileName,$TerminaitonPhoneNumber, $AdditionalNotes){

    #Initialize PDFReader to access PDF
    $Reader = [iText.Kernel.Pdf.PdfReader]::new($TerminationPDF)
    $Reader.SetUnethicalReading($True)

    #Initialize PDFWriter to write our modified PDF to. Needs to be different name than above
    $Writer = [iText.Kernel.Pdf.PdfWriter]::new("$TerminationOutput\$FileName.pdf")

    #Now we generate a PdfDocument object to tie our PdfReader and PdfWriter together and allow us to start our work
    $TermPdfDoc = [iText.Kernel.Pdf.PdfDocument]::new($Reader, $Writer)
    $Form = [iText.Forms.PdfAcroForm]::getAcroForm($TermPdfDoc, $True)
    $Fields = $Form.getFormFields()

    #Adding name to the Name box
    $TermName = ($Fields | Where-Object {$_.key -eq "Name"}).Value
    $TermName.SetValue($TerminationDisplayName)

    #Adding username to the Username box
    $TermUsername = ($Fields | Where-Object {$_.key -eq "Username"}).Value
    $TermUsername.SetValue($TerminatedUserName)

    #Adding department to the department box
    $TermDepartment = ($Fields | Where-Object {$_.key -eq "Department"}).Value
    $TermDepartment.SetValue($TerminationDepartment)

    #Adding manager to the manager box
    $TermManager = ($Fields | Where-Object {$_.key -eq "Manager"}).Value
    $TermManager.SetValue($PDFTerminationManagerDisplayName)

    #Adding office to the office box
    $TermOffice = ($Fields | Where-Object {$_.key -eq "Office"}).Value
    $TermOffice.SetValue($TerminationOfficeLocation)

    #Determining if there is a phone number and putting in correct value
    $TermPhone = ($Fields | Where-Object {$_.key -eq "Phone"}).Value
    if([string]::IsNullOrEmpty($TerminaitonPhoneNumber)){
        $TermPhone.SetValue("No Phone")
    }else {
        $TermPhone.SetValue($TerminaitonPhoneNumber)
    }

    #Adding removed groups to active directory changes
    $TermRemovedGroups = ($Fields | Where-Object {$_.key -eq "Active Directory Changes"}).Value
    $TermRemovedGroups.SetValue($RemovedGroups)

    #Adding additional Notes
    $TermAdditionalNotes = ($Fields | Where-Object {$_.key -eq "Additional Notes"}).Value
    $TermAdditionalNotes.SetValue($AdditionalNotes)

    #Checking check boxes for tasks completed by the script
    $TermDisableADAccount = ($Fields | Where-Object {$_.key -eq "Disable AD Account"}).Value
    $TermDisableADAccount.SetValue($True)

    $TermMovedToDisabledOU = ($Fields | Where-Object {$_.key -eq "Move User to Disabled Users OU"}).Value
    $TermMovedToDisabledOU.SetValue($True)

    $TermRemoveGroups = ($Fields | Where-Object {$_.key -eq "Remove  Document DistributionSecurity Groups"}).Value
    $TermRemoveGroups.SetValue($True)

    $TermManagerFieldCleared = ($Fields | Where-Object {$_.key -eq "Manager Field cleared Automated"}).Value
    $TermManagerFieldCleared.SetValue($True)

    $TermPdfDoc.Close()
    
}

function SendApprovalEmail($OwnerEmailList,$OriginalMessage,$UserDisplayName, $MirrorDisplayName,$SubjectVariable){

    foreach ($Owner in $OwnerEmailList) {

        $AppOwnerEmail = $Owner.GroupOwner
        $OwnerUserName = $Owner.GroupOwner -split "@"
        $OwnerUserName = $OwnerUserName[0]
        $OwnerDisplayName = Get-ADUser  $OwnerUserName -Properties DisplayName | Select-Object -ExpandProperty DisplayName
        $RequestedGroups = $Owner.GroupList -join "`n"
        $bodyvariable = "Hello $OwnerDisplayName
    
    Do you approve of $UserDisplayName being given access to the following groups based on the mirror of $MirrorDisplayName :
    
    $RequestedGroups
        
    Thanks,
    Helpdesk
    
    --original message--
    $OriginalMessage
    "
        invoke-command -computername $RelayServer -scriptblock { Send-MailMessage -Subject  $using:Subjectvariable -Body $using:bodyvariable -From $using:HelpDeskEmail -SmtpServer $using:SMTPServer -To $using:AppOwnerEmail -Bcc $using:HelpDeskEmail }
    }
}

function FindGroups($File1,$File2,$File3,$File4) {

    #Finds unique owners and stores in array
    $OwnersFound = @();
    #Finding unique owners
    foreach ($Owner in $File1) {
        if ($OwnersFound -notcontains $Owner) {
            $OwnersFound += $Owner;
        }
    }
    foreach ($Owner in $File2) {
        if ($OwnersFound -notcontains $Owner) {
            $OwnersFound += $Owner;
        }
    }
    foreach ($Owner in $File3) {
        if ($OwnersFound -notcontains $Owner) {
            $OwnersFound += $Owner;
        }
    }
    
    #Create list of PS objects of approvers found
    $OwnerObjectList = @();
    #Creates PS custom Object for each owner
    foreach ($UniqueOwner in $OwnersFound) {
        #create object pairs to link group to owner
        $TempPSCustomObject2 = [PSCustomObject]@{
            GroupOwner = $UniqueOwner
            GroupList  = @()
        }
        $OwnerObjectList += $TempPSCustomObject2;
    }
    
    #Creates array to store custom objects
    $ParentObjectArray = @();
    #Creates PS custom Object for each group owner pair
    foreach ($Group in $File4) {
        $TempPSCustomObject = [PSCustomObject]@{
            Owner  = $Group.Owner
            Owner2 = $Group.Owner2
            Owner3 = $Group.Owner3
            Group  = $Group.SamAccountName
        }
        $ParentObjectArray += $TempPSCustomObject;
    }
    
    #Finding PS custom objects for 2.0 groups
    $FoundGroupObjects = @()
    foreach ($Group in $FoundGroups) {
        $x = 0;
        while ($Group -ne $ParentObjectArray[$x].Group) {
            $x++;
        }
        $FoundGroupObjects += $ParentObjectArray[$x]
    }
    
    #setting counters used below
    
    $y = 0;
    $z = 0;
    $defaultMESApprover = 0;
    
    while ($OwnerObjectList[$y].GroupOwner -ne $SpecialApprover1) {
        $y++;
    }
    while ($OwnerObjectList[$z].GroupOwner -ne $SpecialApprover2) {
        $z++;
    }
    while ($OwnerObjectList[$defaultMESApprover].GroupOwner -ne $DefaultMESSpecialApprover) {
        $defaultMESApprover++;
    }
    
    #Adding found 2.0 group names to owner objects GroupList property
    foreach ($Group in $FoundGroupObjects) {
        

        #Special group approval based on location
        if ($Group.Group -eq "" -or $Group.Group -eq "") {
            $JerrodBoosObject;
            $KennyCappsObject;

        
            switch ($UserNewOfficeLocation) {
                "Brooklyn Park, MN" { $OwnerObjectList[$y].GroupList += $Group.Group; break; }
                "Cary, NC" { $OwnerObjectList[$z].GroupList += $Group.Group; break; }
                "Raleigh, NC" { $OwnerObjectList[$z].GroupList += $Group.Group; break; }
                Default { $OwnerObjectList[$defaultMESApprover].GroupList += $Group.Group; break }
            }
        }
        else {
            $x = 0;
            while ($Group.Owner -ne $OwnerObjectList[$x].GroupOwner) {
                $x++;
            }
            $OwnerObjectList[$x].GroupList += $Group.Group;
        }

    }
    #creates list of Owner objects that need to be emailed
    $OwnerEmailList = @();
    foreach ($Owner in $OwnerObjectList) {
        if ($Owner.GroupList -ne $null) {
            $OwnerEmailList += $Owner;
        }
    }

    return $OwnerEmailList
}

function SubmitNewHireForm() {

    #Logging for errors
    $TickerNumberNewHireForLog = $NewHireTicketNumberBox.Text
    $NewHireLogFileName = "NewHire - $TickerNumberNewHireForLog - $GetDate" 
    Start-Transcript -Path "$EmployeeChangeLog\$NewHireLogFileName.txt"
    #New hire script to create account, add to correct OU, and add groups
    #Getting info for the ticket
    $TicketNumber = $NewHireTicketNumberBox.Text

    #Enter body of new hire
    $OriginalMessage = $NewHireDescriptionBox.Text

    #checking to make sure mirror exists
    do {
        $UserCopy = $NewHireMirrorBox.Text
        try {
            $temp = Get-ADUser $UserCopy
            $Exists = $True
        }
        catch {
            [void][System.Windows.Forms.MessageBox]::Show("Mirror not found. Try again")
            $Exists = $False
        }
    } while ($Exists -ne $True)

    #------------------------------------------------------- Creating User account ---------------------------------------------
    #gathering normal info
    $UserNewFirst = $NewHireFirstNameBox.Text
    #removing spaces from first name
    $UserNewFirst = $UserNewFirst.Trim()
    #Making first letters uppercase
    $UserNewFirst = (get-culture).TextInfo.ToTitleCase($UserNewFirst.ToLower())

    #For users with two first names that you need to combine for the username
    $UserNewFirstNoWhiteSpace = $UserNewFirst.replace(' ', '')

    $UserNewLast = $NewHireLastNameBox.Text
    #removing spaces from last name
    $UserNewLast = $UserNewLast.Trim()
    #Making first letters uppercase
    $UserNewLast = (get-culture).TextInfo.ToTitleCase($UserNewLast.ToLower())

    #For users with two last names that you need to combine for the username
    $UserNewLastNoWhiteSpace = $UserNewLast.replace(' ', '')

    #Grabbing office location
    $UserNewOfficeLocation = ""
    if ($NewHirePlant1RadioButton.Checked) {
        $UserNewOfficeLocation = "Maple Plain Plant 1"
    }
    elseif ($NewHireHQRadioButton.Checked) {
        $UserNewOfficeLocation = "Maple Plain HQ, MN"
    }
    elseif ($NewHireBrooklynParkRadioButton.Checked) {
        $UserNewOfficeLocation = "Brooklyn Park, MN"
    }
    elseif ($NewHirePlymouthRadioButton.Checked) {
        $UserNewOfficeLocation = "Plymouth, MN"
    }
    elseif ($NewHireRosemountRadioButton.Checked) {
        $UserNewOfficeLocation = "Rosemount, MN"
    }
    elseif ($NewHireCaryRadioButton.Checked) {
        $UserNewOfficeLocation = "Cary, NC"
    }
    elseif ($NewHireRaleighRadioButton.Checked) {
        $UserNewOfficeLocation = "Raleigh, NC"
    }
    elseif ($NewHireNashuaRadioButton.Checked) {
        $UserNewOfficeLocation = "Nashua, NH"
    }
    else {
        $UserNewOfficeLocation = ""
    }

    $FullName = $UserNewFirst + " " + $UserNewLast

    #Converting for username
    $FirstLower = $UserNewFirstNoWhiteSpace.ToLower()
    $LastLower = $UserNewLastNoWhiteSpace.ToLower()

    #Parsing out apostrphes 
    $FirstSplitOnSpecial1 = $FirstLower -split "'"
    $JoinedFirstName = $FirstSplitOnSpecial1[0] + $FirstSplitOnSpecial1[1]

    #Parsing out hyphens
    $FirstSplitOnSpecial2 = $JoinedFirstName -split "-"
    $JoinedFirstName2 = $FirstSplitOnSpecial2[0] + $FirstSplitOnSpecial2[1]

    #Parsing out apostrphes 
    $LastSplitOnSpecial1 = $LastLower -split "'"
    $JoinedLastName = $LastSplitOnSpecial1[0] + $LastSplitOnSpecial1[1]

    #Parsing out hyphens
    $LastSplitOnSpecial2 = $JoinedLastName -split "-"
    $JoinedLastName2 = $LastSplitOnSpecial2[0] + $LastSplitOnSpecial2[1]

    #Joining username after parsing
    $NewUserName = $JoinedFirstName2 + "." + $JoinedLastName2

    #Cropping user name down to 20 character limit if needed
    If ($NewUserName.Length -gt 20) {
        $FinalUserName = $NewUserName.Substring(0, 19)
    }
    else {
        $FinalUserName = $NewUserName
    }

    #checking to see if user already exists
    try {
        #variable to just capture the get-aduser
        $CheckForCurrentUser = Get-ADUser $FinalUserName
        $MiddleInitial = $NewHireMiddleInitialBox.Text

        #Adds initial to username
        $JoinedFirstName2 = $JoinedFirstName2 + $MiddleInitial
        $NewUserName = $JoinedFirstName2 + "." + $JoinedLastName2

        #Adds initial to display name
        $FullName = $UserNewFirst + " " + $MiddleInitial.ToUpper() + ". " + $UserNewLast

        #Cropping user name down to 20 character limit if needed
        If ($NewUserName.Length -gt 20) {
            $FinalUserName = $NewUserName.Substring(0, 19)
        }
        else {
            $FinalUserName = $NewUserName
        }
    }
    catch {
        #no user found and will continue as normal with the remainder of the script
        
    }

    #Getting OU from mirror
    $UserCopyADObject = Get-ADUser $UserCopy
    $GetOUPath = $UserCopyADObject.DistinguishedName
    $NewUserOUPath = $GetOUPath.Substring($GetOUPath.IndexOf('OU=', [System.StringComparison]::CurrentCultureIgnoreCase));
    #Importing 2.0/AX groups
    $File4 = Import-Csv -Path $GroupOwnerApprovalFile   
    $File1 = $File4.Owner 
    $File2 = $File4.Owner2 
    $File3 = $File4.Owner3 
    $ImportFile1 = $File4.SamAccountName

    #Converting secure string
    $DefaultPassword = ConvertTo-SecureString $DefaultPasswordUnenecrypted -AsPlainText -Force

    #Creating new AD account without groups
    New-ADUser -Name $FullName -GivenName $UserNewFirst -Surname $UserNewLast -SamAccountName $FinalUserName -Path $NewUserOUPath -AccountPassword $DefaultPassword -Enabled $true
    #Getting Title, Department, and manager from mirror
    $MirrorTitle = Get-ADUser $UserCopy -Properties Title | Select-Object -ExpandProperty Title
    $MirrorDepartment = Get-ADUser $UserCopy -Properties Department | Select-Object -ExpandProperty Department
    $MirrorManagerProperty = Get-ADUser $UserCopy -Properties Manager | Select-Object -ExpandProperty Manager
    $MirrorManager = Get-ADUser $MirrorManagerProperty

    #Setting org fields to the mirrors
    Set-ADUser $FinalUserName -Title $MirrorTitle -Department $MirrorDepartment -Manager $MirrorManager -Company "Proto Labs, Inc." -DisplayName $FullName -UserPrincipalName "$FinalUserName@protolabs.com" -Office $UserNewOfficeLocation
    #-------------------------------------------------------------- Adding non 2.0/AX groups based on mirror -----------------------------------------------------
    #Grabs groups from mirror
    $UserMirrorGroups = GET-ADUser -Identity $UserCopy -Properties MemberOf | Select-Object -ExpandProperty MemberOf | Get-ADGroup | Select-Object -ExpandProperty samaccountname

    #Finding 2.0/AX groups
    $FoundGroups = $UserMirrorGroups | Where-Object -FilterScript { $_ -in $ImportFile1 } | Sort-Object

    #Finding non 2.0/AX groups
    $FilterAddGroups = $UserMirrorGroups | Where-Object -FilterScript { $_ -notin $FoundGroups } | Sort-Object
    #creates list to add all groups found that they couldnt be removed from
    $g2 = @();

    #Adding groups to user
    foreach ($Member1 in $FilterAddGroups) { 
        try {
            #try to add the user to the group
            Add-ADGroupMember $Member1 -Members $FinalUserName -Confirm:$false
        }
        catch {
            #if you cant add them then you add that group to the list $g2
            $g2 += $Member;
        }
    }


    #Creating additional Notes
    $NewHireAdditionalNotes = @()
    $NewHireAdditionalNotes + "-------Lowercase letters for login-------"
    $NewHireAdditionalNotes + "-------OU: $NewUserOUPath-------"
    $NewHireAdditionalNotes + "-------Org Fields set based on mirror-------"
    $NewHireAdditionalNotes + "-------Approval emails sent if needed-------"
    $NewHireAdditionalNotes = $NewHireAdditionalNotes -join "`n"

    #Getting Display name
    $NewHireDisplayName = Get-ADUser $FinalUserName -Properties DisplayName | Select-Object -ExpandProperty DisplayName

    #Finding groups that need approvals
    $OwnerEmailList = FindGroups $File1 $File2 $File3 $File4
  
    #sending out approval emails 
    $UserCopyDisplayName = Get-ADUser $UserCopy -Properties DisplayName | Select-Object -ExpandProperty DisplayName
    $SubjectVariable = "Service Request# $TicketNumber New Hire $NewHireDisplayName"

    #Utilizing the send approval email function
    SendApprovalEmail $OwnerEmailList $OriginalMessage $NewHireDisplayName $UserCopyDisplayName $SubjectVariable

    #Runs function that creates PDF and fills it out
    #Creating file name
    $PDFFileName = "$TicketNumber $FinalUserName - New Hire $GetDate"

    $NewHirePDFManagerDisplayName = Get-ADUser $MirrorManager -Properties DisplayName | Select-Object -ExpandProperty DisplayName
    $PDFGroupList = @()
    $PDFGroupList += "-----------Groups added-----------"
    $PDFGroupList += $FilterAddGroups
    $PDFGroupList += "-----------2.0/AX groups that need approval-----------"
    $PDFGroupList += $FoundGroups
    $PDFGroupList += "-----------Groups not authorized to be added-----------"
    $PDFGroupList += $g2
    $PDFGroupList = $PDFGroupList -join "`n"

    NewHirePDFForm $FinalUserName $PDFGroupList $NewHireDisplayName $MirrorTitle $NewHirePDFManagerDisplayName $MirrorDepartment $UserNewOfficeLocation $UserCopy $PDFFileName $NewHireAdditionalNotes
    Stop-Transcript  
    }

    #------------------------------------------------------------------Transfer----------------------------
function SubmitTransferForm() {

    $TicketNumberTransferForLog = $TransferTicketNumberBox.Text
    $TransferLogFileName = "Transfer - $TicketNumberTransferForLog - $GetDate" 
    Start-Transcript -Path "$EmployeeChangeLog\$TransferLogFileName.txt"
    #Getting info for the ticket
    $TicketNumber = $TransferTicketNumberBox.Text
    
    # grabs the user who is transferring -------------------
    $u1 = $TransferUsernameBox.Text

    try {
        Enable-ADAccount -Identity $u1
    }
    catch {
        #User already enabled
    }


    $u1groups = @()
    #Adding "Proto Labs US" group if rehire as it is meant for all users in the US for transfer to work
    if ($TransferRehireRadioButton.Checked) {
        $FileName = "$u1 - Rehire - $GetDate"
        Write-Host "Rehire Checked"
    }else {
        $FileName = "$u1 - Transfer - $GetDate"
        $u1groups = GET-ADUser -Identity $u1 –Properties MemberOf | Select-Object -ExpandProperty MemberOf | Get-ADGroup | Select-Object -ExpandProperty samaccountname
        Write-Host "Transfer Checked"
    }
  
    # grabs the mirror -------------------------------------
    $u2 = $TransferMirrorBox.Text
    $u2groups = GET-ADUser -Identity $u2 –Properties MemberOf | Select-Object -ExpandProperty MemberOf | Get-ADGroup | Select-Object -ExpandProperty samaccountname
    
    # compares the 2 users memberships ---------------------
    
    $removegroups = Compare-Object $u1groups $u2groups  | Where-Object { $_.sideindicator -eq "<=" } | Sort-Object samaccountname | Select-Object -ExpandProperty InputObject
    $addgroups = Compare-Object $u1groups $u2groups | Where-Object { $_.sideindicator -eq "=>" } | Sort-Object samaccountname | Select-Object -ExpandProperty InputObject
        
    #Importing 2.0/AX groups
    $File4 = Import-Csv -Path $GroupOwnerApprovalFile   
    $File1 = $File4.Owner 
    $File2 = $File4.Owner2 
    $File3 = $File4.Owner3 
    $ImportFile1 = $File4.SamAccountName
    #compares the addgroups to the 2.0 groups
    $FoundGroups = $addgroups | Where-Object -FilterScript { $_ -in $ImportFile1 } | Sort-Object
    $FilterAddGroups = $addgroups | Where-Object -FilterScript { $_ -notin $FoundGroups } | Sort-Object
    
    #Remove the protect from accidental deletion
    Get-ADUser $u1 | Set-ADObject -ProtectedFromAccidentalDeletion:$false

    #Getting OU from mirror
    $UserCopyADObject = Get-ADUser $u2
    $GetOUPath = $UserCopyADObject.DistinguishedName
    $NewUserOUPath = $GetOUPath.Substring($GetOUPath.IndexOf('OU=', [System.StringComparison]::CurrentCultureIgnoreCase));

    #Getting Title, Department, and manager from mirror
    $MirrorTitle = Get-ADUser $u2 -Properties Title | Select-Object -ExpandProperty Title
    $MirrorDepartment = Get-ADUser $u2 -Properties Department | Select-Object -ExpandProperty Department
    $MirrorManagerProperty = Get-ADUser $u2 -Properties Manager | Select-Object -ExpandProperty Manager
    $MirrorManager = Get-ADUser $MirrorManagerProperty

    #Setting Org fileds and moving OU
    Set-ADUser $u1 -Manager $MirrorManager -Title $MirrorTitle -Department $MirrorDepartment
    $moving = Get-ADUser $u1
    $moving2 = $moving.DistinguishedName
    Move-ADObject -Identity $moving2 -TargetPath $NewUserOUPath

    #Adding group to list of added groups if rehire
    if ($TransferRehireRadioButton.Checked) {
        $FilterAddGroups += ""#Need to add group here before rehire/transfer script can compare groups
    }

    #-------------------------Updating Group Membership--------------------------------------------------------
    #creates list to add all groups found that they couldnt be removed from
    $g = @();
    $g2 = @();
    
    #If it does then it sets variable $repeat to True and exists loop
    foreach ($Member2 in $removegroups) { 
        try {
            #try to remove the user from the group
            Remove-ADGroupMember $Member2 -Members $u1 -Confirm:$false
        }
        catch {
            #if you cant remove them then you add that group to the list $g
            $g += $Member;
        }
    }
    foreach ($Member1 in $FilterAddGroups) { 
        try {
            #try to add the user to the group
            Add-ADGroupMember $Member1 -Members $u1 -Confirm:$false
        }
        catch {
            #if you cant add them then you add that group to the list $g2
            $g2 += $Member;
        }
    }

    #Getting new OU for notes
    $OUPathForNotes = Get-ADUser $u1
    $OUPathForNotes = $OUPathForNotes.DistinguishedName
    $g = $g -join "`n"
    $g2 = $g2 -join "`n"

    #Setting list for additional notes box in the PDF
    $TransferAdditionalNotes = @()
    $TransferAdditionalNotes += "Account moved to $OUPathForNotes"
    $TransferAdditionalNotes += "---------------------------------"
    $TransferAdditionalNotes += "Org fields updated"
    $TransferAdditionalNotes = $TransferAdditionalNotes -Join "`n"

    #--------------------------------------------------------------------------------------------------------------------------------
    $TransferNewLocation = Get-ADUser $u2 -Properties Office | Select-Object -ExpandProperty Office
    $TransferPDFManagerDisplayName = Get-ADUser $MirrorManager -Properties DisplayName | Select-Object -ExpandProperty DisplayName
    $TransferPosition = Get-ADUser $u1 -Properties Title | Select-Object -ExpandProperty Title
    $TransferDepartment = Get-ADUser $u1 -Properties Department | Select-Object -ExpandProperty Department

    #Seeting lists for the Active directory box in the PDF
    $TransferGroupChanges = @()
    $TransferGroupChanges += "-------Memberships Removed-------"
    $TransferGroupChanges += $removegroups
    $TransferGroupChanges += "-------Memberships Added-------"
    $TransferGroupChanges += $Filteraddgroups
    $TransferGroupChanges += "-------PL2.0 Groups that need approval and have not been added-------"
    $TransferGroupChanges += $FoundGroups
    $TransferGroupChanges += "-------Groups not authorized to be removed-------"
    $TransferGroupChanges += $g
    $TransferGroupChanges += "-------Groups not authorized to be added-------"
    $TransferGroupChanges += $g2
    $TransferGroupChanges = $TransferGroupChanges -join "`n"

    $RehireGroupChanges = @()
    $RehireGroupChanges += "-------Memberships Added-------"
    $RehireGroupChanges += $Filteraddgroups
    $RehireGroupChanges += "-------PL2.0 Groups that need approval and have not been added-------"
    $RehireGroupChanges += $FoundGroups
    $RehireGroupChanges += "-------Groups not authorized to be added-------"
    $RehireGroupChanges += $g2
    $RehireGroupChanges = $RehireGroupChanges -join "`n"

    #Finds groups that need approvals
        $TransferOwnerEmailList = FindGroups $File1 $File2 $File3 $File4

    #sending out approval emails 
        $OriginalTransferMessage = $TransferDescriptionBox.Text
        $TransferDisplayName = Get-ADUser $u1 -Properties DisplayName | Select-Object -ExpandProperty DisplayName
        $MirrorDisplayName = Get-ADUser $u2 -Properties DisplayName | Select-Object -ExpandProperty DisplayName
        $TransferSubjectVariable = "Service Request# $TicketNumber - Transfer/Rehire"

    #Utilizing the send approval email function
        SendApprovalEmail $TransferOwnerEmailList $OriginalTransferMessage $TransferDisplayName $MirrorDisplayName $TransferSubjectVariable

    if ($TransferRehireRadioButton.Checked){
        #Creating file name
        $PDFFileName = "$TicketNumber $TransferDisplayName - New Hire $GetDate"
        
        #Since it is a rehire we will use the New Hire PDF 
        NewHirePDFForm $u1 $RehireGroupChanges $TransferDisplayName $MirrorTitle $TransferPDFManagerDisplayName $MirrorDepartment $TransferNewLocation $u2 $PDFFileName
    }else {
        #Creating file name
        $PDFFileName = "$TicketNumber $TransferDisplayName - Transfer $GetDate"

        TransferPDFForm $TransferDisplayName $u1 $TransferPosition $TransferDepartment $TransferPDFManagerDisplayName $TransferNewLocation $MirrorDisplayName $TransferGroupChanges $PDFFileName $TransferAdditionalNotes
        
    }
    


    [void][System.Windows.Forms.MessageBox]::Show('Transfer Complete')

   
    Stop-Transcript
    
}

function SubmitTerminationForm() {

    $TicketNumberTermForLog = $TerminationTicketNumberBox.Text
    $TermLogFileName = "Transfer - $TicketNumberTermForLog - $GetDate" 
    Start-Transcript -Path "$EmployeeChangeLog\$TermLogFileName.txt"

    #Getting info for the ticket
    $TicketNumber = $TerminationTicketNumberBox.Text
    $TicketNumber = $TicketNumber.Trim()
    $UserName = $TerminationUsernameBox.Text

    $Search1 = Get-ADUser -Filter { sAMAccountName -eq $UserName }

    #Remove the protect from accidental deletion
    try {
        Get-ADUser $Search1 | Set-ADObject -ProtectedFromAccidentalDeletion:$false
    }
    catch {
    }

    #Get users groups to be removed
    $UserGroups = GET-ADUser -Identity $Search1 –Properties MemberOf | Select-Object -ExpandProperty MemberOf | Get-ADGroup | Select-Object -ExpandProperty samaccountname | Sort-Object 

    #-------------------------Disabling account-------------------------------------------
    Disable-ADAccount -Identity $UserName

    #---------------------Moving to disabled OU-------------------------------------------
    $moving = Get-ADUser $UserName
    $moving2 = $moving.DistinguishedName
    Move-ADObject -Identity $moving2 -TargetPath "OU=Disabled Users,OU=Departments,DC=pmdomhq,DC=protomold,DC=com"
        
    #Setting attributes for auto email termination script
    Set-ADUser -Identity $UserName -Replace @{extensionAttribute1 = "Terminated" }
    $GetTodaysDate = Get-Date -Format "MM/dd/yyyy"
    $Attribute3Variable = "Disabled: $GetTodaysDate"
    Set-ADUser -Identity $UserName -Replace @{extensionAttribute3 = $Attribute3Variable }

    #creates list to add all groups found that they couldnt be removed from
    $g = @();

    foreach ($Member in $UserGroups) {
        try {
            #try to remove the user from the group
            Remove-ADGroupMember $Member -Members $Search1 -Confirm:$false 
        }
        catch {
            #if you cant remove them then you add that group to the list $g
            $g += $Member;
        }
    }

    #adding the list of groups that the user wasnt able to be removed from
    $g = $g -join "`n"

    #Creating and formating additional notes
    $TermAdditionalNotes = @()
    $TermAdditionalNotes += "Account moved to disabled OU"
    $TermAdditionalNotes += "-----------------------------"
    $TermAdditionalNotes = $TermAdditionalNotes -join "`n" 

    $Subjectvariable = "Service Request# $TicketNumber  -Termination Notes"
    $bodyvariable = [System.IO.File]::ReadAllText("$TerminationOutput\$FileName.txt")
    invoke-command -computername $SMTPSeverName -scriptblock { Send-MailMessage -Subject  $using:Subjectvariable -Body $using:bodyvariable -From $using:HelpDeskEmail -SmtpServer $using:SMTPServer -To $using:HelpDeskEmail }

    $Phone = Get-ADUser $UserName -Properties Telephonenumber | Select-Object -ExpandProperty telephonenumber

    if ($Phone -eq $null) {
        [void][System.Windows.Forms.MessageBox]::Show('No Phone found');
    }
    else {
        [void][System.Windows.Forms.MessageBox]::Show("Phone found $Phone");
    }

    #Formating groups to be added to PDF
    $PDFGroupFields = @()
    $PDFGroupFields += "-------Account disabled-------"
    $PDFGroupFields += "-------Groups removed-------"
    $PDFGroupFields += $UserGroups
    $PDFGroupFields += "-------Groups not authorized to be removed-------"
    $PDFGroupFields += $g
    $PDFGroupFields += "-------------------------------------"
    $PDFGroupFields += $UKACC
    $PDFGroupFields += "-------------------------------------"
    $PDFGroupFields += $JPACC
    $PDFGroupFields = $PDFGroupFields -join "`n"

    #Getting manager display name
    $ManagerProperty = Get-ADUser $UserName -Properties Manager | Select-Object -ExpandProperty Manager
    $ManagerDisplayName = Get-ADUser $ManagerProperty -Properties DisplayName | Select-Object -ExpandProperty DisplayName

    #Getting Termed users department
    $DepartmentProperty = Get-ADUser $UserName -Properties Department | Select-Object -ExpandProperty Department

    #Getting Termed office location
    $LocationProperty = Get-ADUser $UserName -Properties Office | Select-Object -ExpandProperty Office

    #Setting PDF file name
    $TermPDFFileName = "$TicketNumber $UserName - Termination $GetDate"

    #Creating PDF file
    TerminationPDFForm $UserName $PDFGroupFields $SubjectDisplayName $ManagerDisplayName $DepartmentProperty $LocationProperty $TermPDFFileName $Phone $TermAdditionalNotes
    
    Stop-Transcript
    [void][System.Windows.Forms.MessageBox]::Show('Termination Complete');

}

#-----------------------------------------------------------------------------------------------------
function OnFormClosing_ECSmain{ 
	# $this parameter is equal to the sender (object)
	# $_ is equal to the parameter e (eventarg)

	# The CloseReason property indicates a reason for the closure :
	#   if (($_).CloseReason -eq [System.Windows.Forms.CloseReason]::UserClosing)

	#Sets the value indicating that the event should be canceled.
	($_).Cancel= $False
}

$ECSmain.Add_FormClosing( { OnFormClosing_ECSmain} )

$ECSmain.Add_Shown({$ECSmain.Activate()})
$ModalResult=$ECSmain.ShowDialog()
# Release the Form
$ECSmain.Dispose()
