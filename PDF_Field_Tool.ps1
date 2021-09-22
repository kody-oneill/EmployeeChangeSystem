################################################################################ 
#
#  Name    : C:\Users\kody.oneill\OneDrive - Proto Labs, Inc\Powershell\GitHubTailoredScripts\Form1.ps1  
#  Version : 0.1
#  Author  :
#  Date    : 9/21/2021
#
 #  Generated with ConvertForm module version 2.0.0
#  PowerShell version 5.1.18362.1171
#
#  Invocation Line   : Convert-Form -Path $Source -Destination $Destination -Encoding UTF8 -force
#  Source            : C:\Users\kody.oneill\source\repos\PDFFieldTool\Form1.Designer.cs
################################################################################

function Get-ScriptDirectory
{ #Return the directory name of this script
  $Invocation = (Get-Variable MyInvocation -Scope 1).Value
  Split-Path $Invocation.MyCommand.Path
}

$ScriptPath = Get-ScriptDirectory

# Loading external assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$Form1 = New-Object System.Windows.Forms.Form

$ResultsTextBox = New-Object System.Windows.Forms.TextBox
$SubmitButton = New-Object System.Windows.Forms.Button
$BrowseButton = New-Object System.Windows.Forms.Button
$FilePathTextBox = New-Object System.Windows.Forms.TextBox
$label1 = New-Object System.Windows.Forms.Label
#
# ResultsTextBox
#
$ResultsTextBox.Location = New-Object System.Drawing.Point(12, 84)
$ResultsTextBox.Multiline = $true
$ResultsTextBox.Name = "ResultsTextBox"
$ResultsTextBox.Size = New-Object System.Drawing.Size(776, 578)
$ResultsTextBox.TabIndex = 0
$ResultsTextBox.ScrollBars = "Vertical"
#
# SubmitButton
#
$SubmitButton.Location = New-Object System.Drawing.Point(613, 40)
$SubmitButton.Name = "SubmitButton"
$SubmitButton.Size = New-Object System.Drawing.Size(75, 23)
$SubmitButton.TabIndex = 1
$SubmitButton.Text = "Submit"
$SubmitButton.UseVisualStyleBackColor = $true
#
# BrowseButton
#
$BrowseButton.Location = New-Object System.Drawing.Point(694, 40)
$BrowseButton.Name = "BrowseButton"
$BrowseButton.Size = New-Object System.Drawing.Size(75, 23)
$BrowseButton.TabIndex = 2
$BrowseButton.Text = "Browse"
$BrowseButton.UseVisualStyleBackColor = $true
#
# FilePathTextBox
#
$FilePathTextBox.Location = New-Object System.Drawing.Point(12, 41)
$FilePathTextBox.Name = "FilePathTextBox"
$FilePathTextBox.Size = New-Object System.Drawing.Size(595, 23)
$FilePathTextBox.TabIndex = 3
#
# label1
#
$label1.AutoSize = $true
$label1.Location = New-Object System.Drawing.Point(12, 20)
$label1.Name = "label1"
$label1.Size = New-Object System.Drawing.Size(52, 15)
$label1.TabIndex = 4
$label1.Text = "File Path"
#
# Form1
#
$Form1.ClientSize = New-Object System.Drawing.Size(800, 674)
$Form1.Controls.Add($label1)
$Form1.Controls.Add($FilePathTextBox)
$Form1.Controls.Add($BrowseButton)
$Form1.Controls.Add($SubmitButton)
$Form1.Controls.Add($ResultsTextBox)
$Form1.Name = "Form1"
$Form1.Text = "PDF Field Tool"


#Setting global variables
$LoggedOnUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$LoggedOnUser = $LoggedOnUser.Split("\")
$LoggedOnUser = $LoggedOnUser[1] 

#Import Json configuration file
$JsonFilePath = Get-Content -Raw -Path #Path to JSON Cofnig file
$JsonFile = $JsonFilePath | ConvertFrom-Json

#Pulling data from file

$CommonLoggingCore = $JsonFile[0].OneDriveFilePathStart+$LoggedOnUser+$JsonFile[0].CommonLoggingCorePath
$CommingLogging = $JsonFile[0].OneDriveFilePathStart+$LoggedOnUser+$JsonFile[0].CommonLoggingPath
$iTextIO = $JsonFile[0].OneDriveFilePathStart+$LoggedOnUser+$JsonFile[0].iTextIOPath
$iTextKernal = $JsonFile[0].OneDriveFilePathStart+$LoggedOnUser+$JsonFile[0].iTextKernalPath
$iTextForm = $JsonFile[0].OneDriveFilePathStart+$LoggedOnUser+$JsonFile[0].iTextFormPath
$iTextLayout = $JsonFile[0].OneDriveFilePathStart+$LoggedOnUser+$JsonFile[0].iTextLayoutPath
$BouncyCastle = $JsonFile[0].OneDriveFilePathStart+$LoggedOnUser+$JsonFile[0].iTextBouncyCastlePath

#Adding necessary dll's
Add-Type -Path $CommonLoggingCore
Add-Type -Path $CommingLogging
Add-Type -Path $iTextIO
Add-Type -Path $iTextKernal
Add-Type -Path $iTextForm
Add-Type -Path $iTextLayout
Add-Type -Path $BouncyCastle

$BrowseButton.Add_Click({

	$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') }
	$FileBrowser.ShowDialog()
	$FilePathTextBox.Text = $FileBrowser.FileName
})

#$FilePathTextBox.Text = "Test"
#$PDFFilePath = $FilePathTextBox.Text
#Write-Host $PDFFilePath

$SubmitButton.Add_Click({

	$PDFFilePath = $FilePathTextBox.Text
	#Initialize PDFReader to access PDF
	$Reader = [iText.Kernel.Pdf.PdfReader]::new($PDFFilePath)
	$Reader.SetUnethicalReading($True)
	$GetDate = Get-Date -Format "_dddd_MM.dd.yyyy_HH_mm_ss"
	$FileName = "Testing - $GetDate"

	#Initialize PDFWriter to write our modified PDF to. Needs to be different name than above
	$Writer = [iText.Kernel.Pdf.PdfWriter]::new(#output file path)

	#Now we generate a PdfDocument object to tie our PdfReader and PdfWriter together and allow us to start our work
	$NewHirePdfDoc = [iText.Kernel.Pdf.PdfDocument]::new($Reader, $Writer)
	$Form = [iText.Forms.PdfAcroForm]::getAcroForm($NewHirePdfDoc, $True)
	$Fields = $Form.getFormFields()
	
	$SplitFields = $Fields -split ","
	$x = 0
	$ShowResults = @()
	$ShowResults += "PDF Fields"
	$ShowResults += [System.Environment]::NewLine + "------------------------------"
	while($x -lt $SplitFields.Length){
		$TableLeft = $SplitFields[$x]
		$x++
		$x++
		$ShowResults += [System.Environment]::NewLine + $TableLeft.Substring(1)
	}
	$ResultsTextBox.Text = $ShowResults
	
})






function OnFormClosing_Form1{ 
	# $this parameter is equal to the sender (object)
	# $_ is equal to the parameter e (eventarg)

	# The CloseReason property indicates a reason for the closure :
	#   if (($_).CloseReason -eq [System.Windows.Forms.CloseReason]::UserClosing)

	#Sets the value indicating that the event should be canceled.
	($_).Cancel= $False
}

$Form1.Add_FormClosing( { OnFormClosing_Form1} )

$Form1.Add_Shown({$Form1.Activate()})
$ModalResult=$Form1.ShowDialog()
# Release the Form
$Form1.Dispose()
