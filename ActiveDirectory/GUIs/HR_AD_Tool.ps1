Function HR_AD_Form
{
	#Listbox
	#region Import the Assemblies
	[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
	[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
	#endregion
	
	#region Generate form objects
	$HRForm = New-Object System.Windows.Forms.Form
	$accountLabel = New-Object System.Windows.Forms.Label
	$accountFirstLabel = New-Object System.Windows.Forms.Label
	$accountLastLabel = New-Object System.Windows.Forms.Label
	$accountFirstTB = New-Object System.Windows.Forms.TextBox
	$accountLastTB = New-Object System.Windows.Forms.TextBox
	$accountButton = New-Object System.Windows.Forms.Button
	$titleLabel = New-Object System.Windows.Forms.Label
	$titleTB = New-Object System.Windows.Forms.TextBox
	$titleButton = New-Object System.Windows.Forms.Button
	$managerLabel = New-Object System.Windows.Forms.Label
	$managerFirstLabel = New-Object System.Windows.Forms.Label
	$managerLastLabel = New-Object System.Windows.Forms.Label
	$managerFirstTB = New-Object System.Windows.Forms.TextBox
	$managerLastTB = New-Object System.Windows.Forms.Textbox
	$managerButton = New-Object System.Windows.Forms.Button
	$logRTB = New-Object System.Windows.Forms.RichTextBox
	$okButton = New-Object System.Windows.Forms.Button
	$cancelButton = New-Object System.Windows.Forms.Button
	#endregion Generate form objects
	
	#region Form actions
	$accountButton_Click=
	{#Loads the user info
		$account = CheckAccountTBs -Properties manager,title
		If($account.exists)
		{
			If($account.account.title)
			{
				$titleTB.text = $account.account.title
			}
			If($account.account.manager)
			{
				$man = Get-ADUser $account.account.manager
				If($man)
				{
					$managerFirstTB.text = $man.givenname
					$managerLastTB.text = $man.surname
				}
			}
			Else
			{
				$managerFirstTB.text = $null
				$managerLastTB.text = $null
			}
			Log "$($account.account.name)`'s info located"
		}
	}
	$accountTB_GotFocus=
	{
		$HRForm.acceptbutton = $accountButton
	}
	$titleButton_Click=
	{#sets the user title
		$account = CheckAccountTBs
		If($account.exists)
		{
			If($titleTB.text.length -eq 0)
			{
				Log "No title entered. To enter a blank title, enter one space"
			}
			Else
			{
				Set-ADUser $account.account -Title $titleTB.text -Description $titleTB.text
				Log "$($account.account.name)`'s title updated"
			}
		}
	}
	$titleTB_GotFocus=
	{
		$HRForm.acceptbutton = $titleButton
	}
	$managerButton_Click=
	{#sets the user manager
		$account = CheckAccountTBs
		If($account.exists)
		{
			If(($managerFirstTB.text.length -eq 0) -or ($managerLastTB.text.length -eq 0))
			{
				Log "Please enter a first and last name for the manager"
			}
			Else
			{
				Try
				{
					$filter = "surname -like `"$($managerLastTB.text)`" -and givenname -like `"$($managerFirstTB.text)`""
					$man = Get-ADUser -Filter $filter
					If($man)
					{
						Set-ADUser $account.account -Manager $man
						Log "$($account.account.name)`'s manager set to $($man.name)"
					}
					Else
					{
						Log "Could not locate `'$($managerFirstTB.text) $($managerLastTB.text)`'"
					}
				}
				Catch
				{
					Log "No `'$($managerFirstTB.text) $($managerLastTB.text)`' found in AD"
				}
			}
		}
	}
	$manager_GotFocus=
	{
		$HRForm.acceptbutton = $managerButton
	}
	$OnLoadForm_StateCorrection=
	{	#Correct the initial state of the form to prevent the .Net maximized form issue
		$HRForm.WindowState = $InitialFormWindowState
		#$SAMAccountName.Focus()
	}
	#endregion Form actions
	
	#region Functions
	Function Log #writes to the richtexbox
	{
		Param
		(
			[string]$text
		)
		Begin{}
		Process
		{
			If($logRTB.text.length -gt 0)
			{
				$logRTB.text += "`n$text"
			}
			Else
			{
				$logRTB.text += "$text"
			}
		}
		End
		{
			$logRTB.select($logRTB.text.length,1)
			$logRTB.ScrollToCaret()
		}
	}
	Function CheckAccountTBs #checks for input, validates, and returns the user
	{
		Param
		(
			[String[]]$Properties
		)
		If(($accountFirstTB.text.length -eq 0) -or ($accountLastTB.text.length -eq 0))
		{
			Log "Please enter a first and last name for the user"
			Return New-Object PSObject -Property @{
					"Exists" = $false
			}
		}
		$filter = "surname -like `"$($accountLastTB.text.trim())`" -and givenname -like `"$($accountFirstTB.text.trim())`""
		If($Properties)
		{
			$account = Get-ADUser -Filter $filter -Properties $Properties
		}
		Else
		{
			$account = Get-ADUser -Filter $filter
		}
		If($account)
		{
			Return New-Object PSObject -Property @{
					"Account" = $account
					"Exists" = $true
			}
		}
		Else
		{
			Log "`'$($accountFirstTB.text) $($accountLastTB.text)`' does not exist"
		}
		Return New-Object PSObject -Property @{
				"Exists" = $false
		}
	}
	#endregion Functions
	
	#region Form building
	$HRForm.text = "HR AD Tool"
	$HRForm.name = "HRForm"
	$HRForm.ClientSize = New-Object System.Drawing.Size(400,517)
	$HRForm.acceptButton = $okButton
	$HRForm.cancelButton = $cancelButton
	
	$accountLabel.name = "AccountLabel"
	$accountLabel.text = "Employee:"
	$accountLabel.location = New-Object System.Drawing.Size(13,35)
	$accountLabel.autosize = $true
	
	$accountFirstLabel.name = "AccountFirstLabel"
	$accountFirstLabel.text = "First name:"
	$accountFirstLabel.location = New-Object System.Drawing.Size(90,13)
	$accountFirstLabel.autosize = $true
	
	$accountLastLabel.name = "AccountLastLabel"
	$accountLastLabel.text = "Last name:"
	$accountLastLabel.location = New-Object System.Drawing.Size(200,13)
	$accountLastLabel.autosize = $true
	
	$accountFirstTB.name = "AccountTB"
	$accountFirstTB.add_GotFocus($accountTB_GotFocus)
	$accountFirstTB.location = New-Object System.Drawing.Size(90,30)
	$accountFirstTB.size = New-Object System.Drawing.Size(100,23)
	
	$accountLastTb.Name = "AccountLastTB"
	$accountLastTB.add_GotFocus($accountTB_GotFocus)
	$accountLastTB.location = New-Object System.Drawing.Size(200,30)
	$accountLastTB.size = New-Object System.Drawing.Size(100,23)
	
	$accountButton.name = "AccountButton"
	$accountButton.text = "Get"
	$accountButton.location = New-Object System.Drawing.Size(310,30)
	$accountButton.size = New-Object System.Drawing.Size(60,23)
	$accountButton.add_Click($accountButton_Click)
	
	$titleLabel.name = "TitleLabel"
	$titleLabel.text = "Title:"
	$titleLabel.location = New-Object System.Drawing.Size(13,77)
	$titleLabel.autosize = $true
	
	$titleTB.name = "TitleTB"
	$titleTB.add_GotFocus($titleTB_GotFocus)
	$titleTB.location = New-Object System.Drawing.Size(90,72)
	$titleTB.size = New-Object System.Drawing.Size(210,23)
	
	$titleButton.name = "TitleButton"
	$titleButton.text = "Set"
	$titleButton.location = New-Object System.Drawing.Size(310,72)
	$titleButton.size = New-Object System.Drawing.Size(60,23)
	$titleButton.add_Click($titleButton_Click)
	
	$managerLabel.name = "ManagerLabel"
	$managerLabel.text = "Manager:"
	$managerLabel.location = New-Object System.Drawing.Size(13,115)
	$managerLabel.autosize = $true
	
	$managerFirstLabel.name = "ManagerFirstLabel"
	$managerFirstLabel.text = "First name:"
	$managerFirstLabel.location = New-Object System.Drawing.Size(90,97)
	$managerFirstLabel.autosize = $true
	
	$managerFirstTB.name = "ManagerTB"
	$managerFirstTB.add_GotFocus($manager_GotFocus)
	$managerFirstTB.location = New-Object System.Drawing.Size(90,114)
	$managerFirstTB.size = New-Object System.Drawing.Size(100,23)
	
	$managerLastLabel.name = "ManagerLastLabel"
	$managerLastLabel.text = "Last name:"
	$managerLastLabel.location = New-Object System.Drawing.Size(200,97)
	$managerLastLabel.autosize = $true
	
	$managerLastTB.name = "ManagerLastTB"
	$managerLastTB.add_GotFocus($manager_GotFocus)
	$managerLastTB.location = New-Object System.Drawing.Size(200,114)
	$managerLastTB.size = New-Object System.Drawing.Size(100,23)
	
	$managerButton.name = "ManagerButton"
	$managerButton.text = "Set"
	$managerButton.location = New-Object System.Drawing.Size(310,114)
	$managerButton.size = New-Object System.Drawing.Size(60,23)
	$managerButton.add_Click($managerButton_Click)
	
	$logRTB.name = "LogRTB"
	$logRTB.readonly = $true
	$logRTB.location = New-Object System.Drawing.Size(13,167)
	$logRTB.size = New-Object System.Drawing.Size(274,300)
	
	$okButton.name = "OKButton"
	$okButton.text = "OK"
	$okButton.location = New-Object System.Drawing.Size(70,481)
	
	$cancelButton.name = "CancelButton"
	$cancelButton.Text = "Cancel"
	$cancelButton.location = New-Object System.Drawing.Size(150,481)
	
	$HRForm.Controls.Add($accountLabel)
	$HRForm.Controls.Add($accountFirstLabel)
	$HRForm.Controls.Add($accountLastLabel)
	$HRForm.Controls.Add($accountFirstTB)
	$HRForm.Controls.Add($accountLastTB)
	$HRForm.Controls.Add($accountButton)
	$HRForm.Controls.Add($titleLabel)
	$HRForm.Controls.Add($titleTB)
	$HRForm.Controls.Add($titleButton)
	$HRForm.Controls.Add($managerLabel)
	$HRForm.Controls.Add($managerFirstLabel)
	$HRForm.Controls.Add($managerFirstTB)
	$HRForm.Controls.Add($managerLastLabel)
	$HRForm.Controls.Add($managerLastTB)
	$HRForm.Controls.Add($managerButton)
	$HRForm.Controls.Add($logRTB)
	$HRForm.Controls.Add($okButton)
	$HRForm.Controls.Add($cancelButton)
	#endregion Form building
	
	#Save the initial state of the form
	$InitialFormWindowState = $HRForm.WindowState
	#Init the OnLoad event to correct the initial state of the form
	$HRForm.add_Load($OnLoadForm_StateCorrection)
	#Show the Form
	$HRForm.ShowDialog()| Out-Null
}
Import-Module ActiveDirectory
HR_AD_Form