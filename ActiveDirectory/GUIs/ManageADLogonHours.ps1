#Working with individual users and with groups
Function SetAll
{
	Param
	(
		[Microsoft.ActiveDirectory.Management.ADUser[]]$users
		,
		[ValidateSet("Template","Individual")]
		[string]$Type
		,
		[string]$TemplateUser
	)
	$form = New-Object System.Windows.Forms.Form
	$rtb = New-Object System.Windows.Forms.RichTextBox
	$ok = New-Object System.Windows.Forms.Button
	$cancel = New-Object System.Windows.Forms.Button
	$warning = New-Object System.Windows.Forms.Label
	
	$ok_click=
	{
		$form.close()
	}
	$OnLoadForm_StateCorrection=
	{
		$Form.WindowState = $InitialFormWindowState #Correct the initial state of $the form to prevent the .Net maximized form issue
		$rtb.top = $warning.height + 26
		$form.size = New-Object System.Drawing.Size(250,$($warning.height+$rtb.height+$ok.height+$cancel.height+56)) #variable height depending on the number of users being set
		$ok.top = $form.clientsize.height - 13 - $ok.height
		$ok.left = ($form.clientsize.width / 2) - $ok.width - 2
		$cancel.top = $form.clientsize.height - 13 - $cancel.height
		$cancel.left = ($form.clientsize.width / 2) + 2
	}
	
	Switch($Type)
	{
		"Template" {$warning.text = "-WARNING-:`nYou are about to set the logon hours for all of the following users based on the hours set on the $TemplateUser tab:"}
		"Individual" {$warning.text = "-WARNING-:`nYou are about to set the logon hours for all of the following users based on the hours set on each tab:"}
	}
	$warning.autosize = $true
	$warning.maximumsize = New-Object System.Drawing.Size(224,0)
	$warning.top = 13
	$warning.left = 13
	
	$rtb.left = 13
	$rtb.readonly = $true
	$rtb.text = $users[0].name
	For($x=1;$x -lt $users.count;$x++)
	{
		$rtb.text += "`n$($users[$x].name)"
	}
	$rtb.height = ($rtb.GetLineFromCharIndex($rtb.Text.Length)+1) * $rtb.Font.Height + 1 + $rtb.Margin.Vertical
	$rtb.width = 224
	
	$form.cancelbutton = $cancel
	$form.acceptbutton = $ok
	
	$ok.text = "OK"
	$ok.dialogresult = "ok"
	$ok.add_click($ok_click)
	
	$cancel.text = "Cancel"
	$cancel.dialogresult = "cancel"
	
	$form.controls.add($warning)
	$form.controls.add($ok)
	$form.controls.add($cancel)
	$form.controls.add($rtb)
	
	$InitialFormWindowState = $Form.WindowState
	$Form.add_Load($OnLoadForm_StateCorrection)
	$form.startposition = "CenterScreen"
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	$form.FormBorderStyle = "FixedSingle"
	$form.showdialog()
}
Function LogonHours_Tabbed
{
	#region Import the Assemblies
	[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
	[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
	#endregion
	
	#region declare form objects
	$form = New-Object System.Windows.Forms.Form
	$newButton = New-Object System.Windows.Forms.Button
	$tabControl = New-Object System.Windows.Forms.TabControl
	$accountLabel = New-Object System.Windows.Forms.Label
	$accountFirstLabel = New-Object System.Windows.Forms.Label
	$accountLastLabel = New-Object System.Windows.Forms.Label
	$accountFirstTB = New-Object System.Windows.Forms.TextBox
	$accountLastTB = New-Object System.Windows.Forms.TextBox
	$groupTB = New-Object System.Windows.Forms.TextBox
	$groupLabel = New-Object System.Windows.Forms.Label
	$groupButton = New-Object System.Windows.Forms.Button
	$logRTB = New-Object System.Windows.Forms.RichTextBox
	#endregion declare form objects
	
	#region actions
	$newButton_Click=
	{
		$user = CheckAccount -FName $accountFirstTB.text -LName $accountLastTB.text -properties logonhours #checks to make sure the user exists
		If($user.exists)
		{
			ForEach($tab in $tabControl.tabpages) #make sure there isn't already a tab open for this user
			{
				If($tab.text -eq $user.account.samaccountname)
				{
					Log "$($user.account.name) already has a tab open."
					Return
				}
			}
			New-Tab -tabcontrolobject $tabControl -name $user.account.samaccountname -user $user.account #create a new tab for that user
			Log "$($user.account.name)`'s logon hours loaded."
		}
		Else
		{
			Log "`'$($accountFirstTB.text) $($accountLastTB.text)`' does not exist."
		}
	}
	$groupButton_Click=
	{
		Try
		{
			$group = Get-ADGroup $groupTB.text #check to see if the group exists
		}
		Catch
		{
			Log "`'$($groupTB.text)`' does not exist"
		}
		If($group)
		{
			ForEach($user in Get-ADGroupMember $group) #grab the group members
			{
				$exists = $false
				ForEach($tab in $tabControl.tabpages)
				{
					If($tab.text -eq $user.samaccountname) #check to see if this user already has a tab open
					{
						Log "$($user.name) already has a tab open."
						$exists = $true
					}
				}
				If(!($exists))
				{
					New-Tab -tabcontrolobject $tabControl -name $user.samaccountname -user $user #create a tab for the user
					Log "$($user.name)`'s logon hours loaded."
				}
			}
		}
	}
	$closeButton_Click=
	{
		Remove-Tab -tabcontrolobject $tabControl
	}
	$setButton_Click=
	{
		$user = CheckAccount -SAMAccountName $this.parent.text -properties logonhours
		$dataGridView = $this.parent.controls | ?{$_.gettype().fullname -eq "System.Windows.Forms.DataGridView"} #get the datagridview on this tab
		If($user.exists)
		{
			$binLogonHours = Get-BinADUserLogonHours -User $user.account -logonhours $user.account.logonhours #get the logon hours formatted to easily be used in the datagridview
			For($a=0;$a -lt 7;$a++)
			{
				For($b=0;$b -lt 24;$b++)
				{
					If($dataGridView[$a,$b].selected -eq $true)
					{
						$binLogonHours[$a,$b] = "1"
					}
					Else
					{
						$binLogonHours[$a,$b] = "0"
					}
				}
			}
			Set-BinADUserLogonHours -user $user.account -binlogonhours $binLogonHours
			Log "$($user.account.name)`'s logon hours set."
		}
	}
	$reloadButton_Click= #refreshes the current tab to show any hour changes
	{
		$user = CheckAccount -SAMAccountName $this.parent.text -properties logonhours
		If($user.exists)
		{
			Reload-Tab -tabPage $this.parent -user $user.account
		}
		Log "$($user.account.name)`'s logonhours reloaded"
	}
	$setAllIndButton_click= #this button will set all tabs as they have been individually set
	{
		$users = @() #declare the hashtable to send to SetAll
		ForEach($tab in $tabControl.tabpages)
		{
			$dataGridView = $tab.controls | ?{$_.gettype().fullname -eq "System.Windows.Forms.DataGridView"} #get this tab's datagridview
			$user = checkaccount -SAMAccountName $tab.text -properties logonhours
			$binLogonHours = Get-BinADUserLogonHours -User $user.account -logonhours $user.account.logonhours #get the logon hours formatted to easily be used in the datagridview
			For($a=0;$a -lt 7;$a++)
			{
				For($b=0;$b -lt 24;$b++)
				{
					If($dataGridView[$a,$b].selected -eq $true)
					{
						$binLogonHours[$a,$b] = "1"
					}
					Else
					{
						$binLogonHours[$a,$b] = "0"
					}
				}
			}
			$user | Add-Member -MemberType NoteProperty -Name BinLogonHours -Value $binLogonHours #add the formatted logon hours
			$users += $user
		}
		$result = SetAll -users $users.account -Type "Individual" #sends all the info to SetAll to throw a warning dialog before any hours are changed
		If($result.tostring() -eq "OK")
		{
			ForEach($user in $users)
			{
				Set-BinADUserLogonHours -user $user.account -binlogonhours $user.binlogonhours
				Log "$($user.account.name)`'s logon hours set."
			}
		}
		Else
		{
			""
		}
	}
	$setAllTempButton_click= #this button will set all tabs exactly as the currently open tab is set
	{
		$users = @() #declare the hashtable to be sent to SetAll
		$dataGridView = $tabControl.selectedtab.controls | ?{$_.gettype().fullname -eq "System.Windows.Forms.DataGridView"} #get this tab's datagridview
		$templateUser = $tabControl.selectedtab.text #grab the name of the user hours to be used as a template
		$user = checkaccount -SAMAccountName $templateUser -properties logonhours
		$binLogonHours = Get-BinADUserLogonHours -User $user.account -logonhours $user.account.logonhours #get the logon hours formatted to easily be used in the datagridview
		For($a=0;$a -lt 7;$a++)
		{
			For($b=0;$b -lt 24;$b++)
			{
				If($dataGridView[$a,$b].selected -eq $true)
				{
					$binLogonHours[$a,$b] = "1"
				}
				Else
				{
					$binLogonHours[$a,$b] = "0"
				}
			}
		}
		ForEach($tab in $tabControl.tabpages)
		{
			$user = checkaccount -SAMAccountName $tab.text -properties logonhours
			$user | Add-Member -MemberType NoteProperty -Name BinLogonHours -Value $binLogonHours #uses the $binLogonHours set to the template user
			$users += $user
		}
		$result = SetAll -users $users.account -type "Template" -TemplateUser $templateUser #sends all the info to SetAll to throw a warning dialog before any hours are changed
		If($result.tostring() -eq "OK")
		{
			ForEach($user in $users)
			{
				Set-BinADUserLogonHours -user $user.account -binlogonhours $user.binlogonhours #set the hours
				Log "$($user.account.name)`'s logon hours set."
			}
		}
		Else
		{
			""
		}
	}
	$accountTB_GotFocus= #when focus is on the account tb, the accept button becomes the 'Load User' button
	{
		$form.acceptbutton = $newButton
	}
	$groupTB_GotFocus= #when focus is on the group tb, the accept button becomes the 'Load Group' button
	{
		$form.acceptbutton = $groupButton
	}
	$tabControl_ControlAdded= #whenever a control is added, set the following settings
	{
		$dataGridView = $this.tabpages[$this.tabcount-1].Controls | ?{$_.gettype().fullname -eq "System.Windows.Forms.DataGridView"}
		$columnSortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
		ForEach($column in $dataGridView.columns) #set the columns to the correct width and sort mode
		{
			$column.width = 35
			$column.sortmode = $columnSortMode
		}
		For($x=0;$x -lt 24;$x++)#add the headercell values
		{
			If($x -eq 0)
			{
				$hourInDay = "12AM"
			}
			ElseIf($x -lt 12)
			{
				$hourInDay = "$($x)AM"
			}
			ElseIf($x -eq 12)
			{
				$hourInDay = "12PM"
			}
			Else
			{
				$hourInDay = "$($x - 12)PM"
			}
			$dataGridView.rows[$x].HeaderCell.Value = $hourInDay
		}
		$selectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::ColumnHeaderSelect
		$datagridview.SelectionMode = $selectionMode
		$dataGridView.RowHeadersWidth = 60
		$dataGridView.readonly = $true
		$dataGridView.AllowDrop = $false
		$dataGridView.AllowUserToAddRows = $false
		$dataGridView.AllowUserToDeleteRows = $false
		$dataGridView.AllowUserToOrderColumns = $false
		$dataGridView.AllowUserToResizeColumns = $false
		$dataGridView.AllowUserToResizeRows = $false
	}
	$OnLoadForm_StateCorrection=
	{#Correct the initial state of $the form to prevent the .Net maximized form issue
		$Form.WindowState = $InitialFormWindowState
	}
	#endregion actions
	
	#region functions
	Function New-Tab
	{ #creates each new tab including the datagridview and buttons
		Param
		(
			[System.Windows.Forms.TabControl]$tabControlObject
			,
			[string]$name
			,
			[Microsoft.ActiveDirectory.Management.ADUser]$user
		)
		Begin
		{
			$newTP = New-Object System.Windows.Forms.TabPage
			$newDataTable = New-Object System.Data.DataTable
			$newDataGridView = New-Object System.Windows.Forms.DataGridView
			$newSetButton = New-Object System.Windows.Forms.Button
			$newCloseButton = New-Object System.Windows.Forms.Button
			$newReloadButton = New-Object System.Windows.Forms.Button
			$newSetAllIndButton = New-Object System.Windows.Forms.Button
			$newSetAllTempButton = New-Object System.Windows.Forms.Button
			$newGroupBox = New-Object System.Windows.Forms.GroupBox
		}
		Process
		{
			#create the tab page
			$newTP.text = $name
			#create the datatable
			$daysOfWeek = "Sun","Mon","Tue","Wed","Thu","Fri","Sat"
			For($x=0;$x -lt 7;$x++)
			{
				$newDataTable.Columns.Add($daysOfWeek[$x]) | Out-Null
			}
			For($h=0;$h -lt 24;$h++)
			{
				$row = $newDataTable.NewRow()
				For($d=0;$d -lt 7; $d++)
				{
					$row[$daysOfWeek[$d]] = ""
				}
				$newDataTable.Rows.Add($row)
			}
			#create the datagrid
			$newDataGridView.location = New-Object System.Drawing.Size(13,13)
			$newDataGridView.size = New-Object System.Drawing.Size(307,553)
			$newDataGridView.DataSource = $newDataTable
			$newDataGridView.Add_ColumnHeaderMouseClick($newDataGridView_ColumnHeaderMouseClick)
			$newDataGridView.Add_GotFocus($newDataGridView_GotFocus)
			#create the set button
			$newSetButton.location = New-Object System.Drawing.Size(330,104)
			$newSetButton.Text = "Set"
			$newSetButton.Add_Click($setButton_Click)
			#create the reload button
			$newReloadButton.location = New-Object System.Drawing.Size(330,62)
			$newReloadButton.Text = "Reload"
			$newReloadButton.Add_Click($reloadButton_Click)
			#create the close button
			$newCloseButton.text = "Close"
			$newCloseButton.location  = New-Object System.Drawing.Size(330,20)
			$newCloseButton.Add_Click($closeButton_Click)
			#create the set all group box
			$newGroupBox.location = New-Object System.Drawing.Size(328,146)
			$newGroupBox.size = New-Object System.Drawing.Size(78,121)
			$newGroupBox.text = "Set All"
			#create the set all individually button
			$newSetAllIndButton.location = New-Object System.Drawing.Size(2,70)
			$newSetAllIndButton.text = "Individually"
			$newSetAllIndButton.Add_Click($setAllIndButton_click)
			$newGroupBox.Controls.Add($newSetAllIndButton)
			#create the set all template button
			$newSetAllTempButton.location = New-Object System.Drawing.Size(2,28)
			$newSetAllTempButton.text = "Template"
			$newSetAllTempButton.Add_Click($setAllTempButton_click)
			$newGroupBox.Controls.Add($newSetAllTempButton)
		}
		End
		{
			#add the controls to the tab page
			$newTP.Controls.Add($newDataGridView)
			$newTP.Controls.Add($newSetButton)
			$newTP.Controls.Add($newCloseButton)
			$newTP.Controls.Add($newReloadButton)
			$newTP.Controls.Add($newGroupBox)
			#add the tab page to the tab control object
			$tabControlObject.controls.add($newTP)
			#get user logonhours data
			$binLogonHours = New-Object 'object[,]' 7,24
			$binLogonHours = Get-BinADUserLogonHours -User $user -logonHours $user.logonhours
			For($a=0;$a -lt 7;$a++) #set the datagridview
			{
				For($b=0;$b -lt 24;$b++)
				{
					If($binLogonHours[$a,$b] -eq "1")
					{
						$newDataGridView[$a,$b].selected = $true
					}
					Else
					{
						$newDataGridView[$a,$b].selected = $false
					}
				}
			}
		}
	}
	Function Remove-Tab ([System.Windows.Forms.TabControl]$tabControlObject) #closes the open tab
	{
		$selectedIndex = $tabControlObject.selectedindex
		$tabControlObject.tabPages.RemoveAt($selectedIndex)
	}
	Function Reload-Tab
	{#function to reload the user's logon hours
		Param
		(
			[System.Windows.Forms.TabPage]$tabPage
			,
			[Microsoft.ActiveDirectory.Management.ADUser]$user
		)
		Begin
		{
			$dataGridView = $tabPage.controls | ?{$_.gettype().fullname -eq "System.Windows.Forms.DataGridView"} #grabs the existing datagridview
		}
		Process
		{
			$binLogonHours = New-Object 'object[,]' 7,24 #creates the array for the logon hours, hereto referred to as $binLogonHours in the rest of the script, not just this function
			$binLogonHours = Get-BinADUserLogonHours -User $user -logonHours $user.logonhours
			For($a=0;$a -lt 7;$a++) #update the gridview
			{
				For($b=0;$b -lt 24;$b++)
				{
					If($binLogonHours[$a,$b] -eq "1")
					{
						$DataGridView[$a,$b].selected = $true
					}
					Else
					{
						$DataGridView[$a,$b].selected = $false
					}
				}
			}
		}
		End
		{
		
		}
	}
	Function CheckAccount #checks for input, validates, and returns the user
	{
		[CmdletBinding(DefaultParameterSetName="FnameLname")]
		Param
		(
			[Parameter(ParameterSetName="FnameLname")]
			[string]$FName
			,
			[Parameter(ParameterSetName="FnameLname")]
			[string]$LName
			,
			[Parameter(ParameterSetName="SAMAccountName")]
			[string]$SAMAccountName
			,
			[String[]]$Properties
		)
		#create the filter
		If($SAMAccountName)
		{
			$filter = "SAMAccountName -like `"$SAMAccountName`""
		}
		ElseIf($FName -and $LName)
		{
			$filter = "surname -like `"$LName`" -and givenname -like `"$FName`""
		}
		Else
		{
			Return New-Object PSObject -Property @{
					"Exists" = $false
			}
		}
		#get the user, including properties if specified
		If($Properties)
		{
			$account = Get-ADUser -Filter $filter -Properties $Properties
		}
		Else
		{
			$account = Get-ADUser -Filter $filter
		}
		#return the user
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
	Function Get-BinADUserLogonHours
	{ #this is where the magic with the user logon hours happens
		Param
		(
			[Microsoft.ActiveDirectory.Management.ADUser]$User
			,
			[Byte[]]$logonhours
		)
		Begin
		{
			If(!($logonhours))#if it wasn't passed
			{
				$logonhours = (Get-ADUser $User -Properties logonhours).logonhours
			}
			If(!($logonhours))#if it is null in AD, set to a default of 1 for everything
			{
				[byte[]]$logonhours = @(255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,255,255)
			}
			$arrLogonHours = New-Object 'object[,]' 7,24 #7x24 array, each bit equals one hour. 1 for logon, 0 for no logon
		}
		Process
		{
			$tmp = $null
			$arrIndex = 0;
			For($x=1;$x -lt $logonhours.count;$x++) #start at 1 because the first byte set actually belongs to Sat for whatever reason
			{
				[string]$t = [Convert]::ToString($logonHours[$x], 2) #convert it to a string for ease of use
				For($y=$t.length;$y -lt 8;$y++) #if the length is less than 8, we need the leading zeroes
				{
					[string]$t = "0$t" #add leading zeros
				}
				#reverse the binary string
				$arr = $t -split ""
				[array]::Reverse($arr)
				$t = $arr -join ""
				#add the new byte to the tmp string
				$tmp += $t
				If($tmp.length -eq 24)
				{ #once the string is 24 hours long add it as a row to the array. Each row is a day
					For($a=0;$a -lt $tmp.length;$a++)
					{
						$arrLogonHours[$arrIndex,$a] = $tmp[$a]
					}
					$arrIndex++
					$tmp = $null
				}
			}
			$t = [Convert]::ToString($logonhours[0], 2) #grab the first byte now, since it is the last byte of Sat
			For($y=$t.length;$y -lt 8;$y++)
			{
				[string]$t = "0$t" #add leading zeros
			}
			#reverse the binary string
			$arr = $t -split ""
			[array]::Reverse($arr)
			$t = $arr -join ""
			#add the new byte to the tmp string
			$tmp += $t
			For($a=0;$a -lt $tmp.length;$a++)
			{ #complete the 3 bytes into a string
				$arrLogonHours[6,$a] = $tmp[$a]
			}
		}
		End
		{
			Return , $arrLogonHours #comma returns the array exactly as it is
		}
	}
	Function Set-BinADUserLogonHours
	{
		Param
		(
			[Parameter(Mandatory=$true)]
			[Microsoft.ActiveDirectory.Management.ADUser]$user
			,
			[Parameter(Mandatory=$true)]
			$binLogonHours
		)
		Begin
		{
			Function Get-NextByte
			{ #quick and easy to grab the next byte, so eight chars from the array in order
				Param
				(
					$binArr
					,
					[int]$x
					,
					[int]$y
				)
				$string = $null
				For($a=7;$a -ge 0;$a--)
				{
					$string += $binArr[$x,$($y+$a)]
				}
				Return $string
			}
		}
		Process
		{
			[String[]]$tmpLogonHours = @()
			For($x=0;$x -lt 7;$x++)
			{
				For($y=0;$y -lt 24;$y = $y+8) #incremented by 8 because of the need for a whole byte
				{
					[string[]]$tmplogonhours += Get-NextByte $binLogonHours -x $x -y $y
				}
			}
			[Byte[]]$logonhours = @() #declare the byte hashtable
			$logonHours += [Convert]::ToInt32($tmpLogonHours[20],2) #there are 21 bytes in 7 24 hour days (3 per day), grab the last one first since AD formats it funky
			For($x=0;$x -lt 20;$x++)
			{
				$logonhours += [Convert]::ToInt32($tmpLogonHours[$x],2) #add another byte
			}
			$user.logonhours = $logonhours #set the logon hours on the $user object
			Try
			{
				Set-ADUser -Instance $user #set the changes in AD
			}
			Catch
			{
				Log "Can't set $($user.name)`'s logon hours"
			}
		}
		End{}
	}
	Function Log #writes to the richtexbox
	{
		Param
		(
			[string]$text
		)
		Begin{}
		Process
		{
			If($logRTB.text.length -gt 0)#if there is already text in the rtb, put in a line break
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
	#endregion functions
	
	#region define form objects
	$form.size = New-Object system.Drawing.Size(450,850)
	
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
	
	$newButton.name = "NewButton"
	$newButton.location = New-Object System.Drawing.Size(333,30)
	$newButton.text = "Load User"
	$newButton.add_Click($newButton_Click)
	
	$groupTB.name = "GroupTB"
	$groupTB.location = New-Object System.Drawing.Size(90,66)
	$groupTB.size = New-Object System.Drawing.Size(150,23)
	$groupTB.Add_GotFocus($groupTB_GotFocus)
	
	$groupButton.name = "GroupButton"
	$groupButton.text = "Load Group"
	$groupButton.location = New-Object System.Drawing.Size(333,66)
	$groupButton.Add_Click($groupButton_Click)
	
	$groupLabel.name = "GroupLabel"
	$groupLabel.text = "Group:"
	$groupLabel.location = New-Object System.Drawing.Size(13,71)
	$groupLabel.autosize = $true
	
	$logRTB.name = "LogRTB"
	$logRTB.readonly = $true
	$logRTB.location = New-Object System.Drawing.Size(13,717)
	$logRTB.size = New-Object System.Drawing.Size(420,100)
	
	$tabControl.location = New-Object System.Drawing.Size(13,104)
	$tabControl.Size = New-Object System.Drawing.Size(420,600)
	$tabControl.Add_ControlAdded($tabControl_ControlAdded)
	#endregion define form objects
	
	$form.Controls.Add($accountFirstTB)
	$form.Controls.Add($accountLastTB)
	$form.controls.add($newButton)
	$form.Controls.Add($groupTB)
	$form.Controls.Add($groupButton)
	$form.controls.add($tabControl)
	$form.Controls.Add($accountLabel)
	$form.Controls.Add($accountFirstLabel)
	$form.Controls.Add($accountLastLabel)
	$form.Controls.Add($groupLabel)
	$form.controls.add($tabControl)
	$form.controls.add($logRTB)
	
	$InitialFormWindowState = $Form.WindowState
	$Form.add_Load($OnLoadForm_StateCorrection)
	$Form.ShowDialog()| Out-Null
	Import-Module ActiveDirectory
}
LogonHours_Tabbed #call the main form