<#
Get-EOLBatchToCSV.ps1
Author - Kristopher Roy
Date Created - 11/21/2017
The purpose of this script is to dynamically grab a list of all EOL Migration Batches that currently exist in an environment
It then allows you to select one of the batches and then export the list of users to a csv
Currently the script returns; EmailAddress,BatchID,Migration Status, and LastSuccessful Sync time
The Script must be run in an EOL Powershell Session
#>

#This function lets you build an array of specific list items you wish
Function MultipleSelectionBox ($inputarray,$prompt,$listboxtype) {
 
	# Taken from Technet - http://technet.microsoft.com/en-us/library/ff730950.aspx
	# This version has been updated to work with Powershell v3.0.
	# Had to replace $x with $Script:x throughout the function to make it work. 
	# This specifies the scope of the X variable.  Not sure why this is needed for v3.
	# http://social.technet.microsoft.com/Forums/en-SG/winserverpowershell/thread/bc95fb6c-c583-47c3-94c1-f0d3abe1fafc
	#
	# Function has 3 inputs:
	#     $inputarray = Array of values to be shown in the list box.
	#     $prompt = The title of the list box
	#     $listboxtype = system.windows.forms.selectionmode (None, One, MutiSimple, or MultiExtended)
 
	$Script:x = @()
 
	[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
	[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
 
	$objForm = New-Object System.Windows.Forms.Form 
	$objForm.Text = $prompt
	$objForm.Size = New-Object System.Drawing.Size(300,600) 
	$objForm.StartPosition = "CenterScreen"
 
	$objForm.KeyPreview = $True
 
	$objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
		{
			foreach ($objItem in $objListbox.SelectedItems)
				{$Script:x += $objItem}
			$objForm.Close()
		}
		})
 
	$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
		{$objForm.Close()}})
 
	$OKButton = New-Object System.Windows.Forms.Button
	$OKButton.Location = New-Object System.Drawing.Size(75,520)
	$OKButton.Size = New-Object System.Drawing.Size(75,23)
	$OKButton.Text = "OK"
 
	$OKButton.Add_Click(
	   {
			foreach ($objItem in $objListbox.SelectedItems)
				{$Script:x += $objItem}
			$objForm.Close()
	   })
 
	$objForm.Controls.Add($OKButton)
 
	$CancelButton = New-Object System.Windows.Forms.Button
	$CancelButton.Location = New-Object System.Drawing.Size(150,520)
	$CancelButton.Size = New-Object System.Drawing.Size(75,23)
	$CancelButton.Text = "Cancel"
	$CancelButton.Add_Click({$objForm.Close()})
	$objForm.Controls.Add($CancelButton)
 
	$objLabel = New-Object System.Windows.Forms.Label
	$objLabel.Location = New-Object System.Drawing.Size(10,20) 
	$objLabel.Size = New-Object System.Drawing.Size(280,20) 
	$objLabel.Text = "Please make a selection from the list below:"
	$objForm.Controls.Add($objLabel) 
 
	$objListbox = New-Object System.Windows.Forms.Listbox 
	$objListbox.Location = New-Object System.Drawing.Size(10,40) 
	$objListbox.Size = New-Object System.Drawing.Size(260,20) 
 
	$objListbox.SelectionMode = $listboxtype
 
	$inputarray | ForEach-Object {[void] $objListbox.Items.Add($_)}
 
	$objListbox.Height = 470
	$objForm.Controls.Add($objListbox) 
	$objForm.Topmost = $True
 
	$objForm.Add_Shown({$objForm.Activate()})
	[void] $objForm.ShowDialog()
 
	Return $Script:x
}

#This Function creates a dialogue to return a Folder Path
function Get-Folder {
    param([string]$Description="Select Folder to place results in",[string]$RootFolder="Desktop")

 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
     Out-Null     

   $objForm = New-Object System.Windows.Forms.FolderBrowserDialog
        $objForm.Rootfolder = $RootFolder
        $objForm.Description = $Description
        $Show = $objForm.ShowDialog()
        If ($Show -eq "OK")
        {
            Return $objForm.SelectedPath
        }
        Else
        {
            Write-Error "Operation cancelled by user."
        }
}

#Return all Migration Batches
$batches = Get-MigrationBatch
#Creates list for selection
$options01 = ($batches|select Identity).identity.name
#Presents list to user for selection
$batchselection = MultipleSelectionBox -listboxtype one -inputarray $options01

#Builds the Report and exports it to CSV
Get-MigrationUser -BatchId $batchselection|select @{N="EmailAddress";E={$_.MailboxEmailAddress}},BatchId,Status,LastSuccessfulSyncTime|export-csv ((get-folder -Description "Select Folder to Place CSV")+"\"+$batchselection+".csv") -NoTypeInformation