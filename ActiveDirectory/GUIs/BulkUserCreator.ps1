<#
Wishlist:
    User count selection
    Add option to load existing users for editing
	Add option to load from CSV
#>
import-module activedirectory
Add-Type -AssemblyName PresentationFramework
#region Functions
Function SwitchAttribs{ #This is where validation happens!
    Param(
        $attrib
    )
    $good = "+"
    $bad = "!"
    Switch($attrib){
        "Manager" {
            Try{
                $man = Get-ADUser $tb.Text #if we can pull the manager
                Return $good #passed validation
            }catch{
                Return $bad #failed validation
            }
        }
        {"Mobile","Fax","TelephoneNumber" -contains $_} {#all phone number
            If($tb.text -match "^\d{3}\-\d{3}\-\d{4}$"){ #xxx-xxx-xxxx
                Return $good #passed validation
            }Else{
                Return $bad #failed validation
            }
        }
        {"GivenName","SurName","SamAccountName","Title","Department","Company" -Contains $_} {
            If($tb.text)
            {
                Return $good
            }Else{
                Return $bad
            }
        
        }
        Default {
            Write-Host "Not yet implemented: $attrib"
            Return "" #no validation
        }
    }
}

Function New-BUCGridElement{ #BUC: bulk user creator
    Param(
        [xml]$xaml
        ,
        [string]$Type
        ,
        [string]$Name
        ,
        [int]$Column
        ,
        [int]$Row
        ,
        [string]$Content
        ,
        [System.Collections.Hashtable]$Attributes
        ,
		[System.Xml.XmlElement]$grid
		,
        [bool]$ReadOnly
    )
    #row addition if necessary
    $rowcount = $grid.'Grid.RowDefinitions'.RowDefinition.count
    If($Row -ge $rowcount){
        For($x=0;$x -lt $Row-$rowcount+1;$x++){ #allows addition of multiple rows, up to the row index provided
            $newRow = $xaml.CreateElement("RowDefinition")
            $newRow.SetAttribute("Height","Auto") | Out-Null
            $grid.'Grid.RowDefinitions'.AppendChild($newRow) | Out-Null
        }
    }

    #column addition if necessary
    $colcount = $grid.'Grid.ColumnDefinitions'.ColumnDefinition.count
    If($column -ge $colcount)
    {
        For($x=0;$x -lt $Column-$colcount+1;$x++){ #allows addition of multiple columns, up to the column index provided
            $newCol = $xaml.CreateElement("ColumnDefinition")
            $newCol.SetAttribute("Width","Auto") | Out-Null
            $grid.'Grid.ColumnDefinitions'.AppendChild($newCol) | Out-Null
        }
    }#>

    $newEl = $xaml.CreateElement("$Type")
    $newEl.SetAttribute("Name", $nsm.LookupNamespace("x"), $Name) | Out-Null
    $newEl.SetAttribute("Grid.Column",$Column) | Out-Null
    $newEl.SetAttribute("Grid.Row",$Row) | Out-Null
    If($Content){
        $newEl.SetAttribute("Content",$Content) | Out-Null
    }
    If($attributes)
    {
        ForEach($attrib in $Attributes.GetEnumerator()){
            $newEl.SetAttribute($attrib.Name, $attrib.Value) | Out-Null
        }
    }
    If($ReadOnly)
    {
        $newEl.SetAttribute("IsReadOnly", "True") | Out-Null
    }

    $grid.AppendChild($newEl) | Out-Null
}

Function New-GroupTemplatizer{
    Param(
        [psobject[]]$users
    )
    [xml]$xaml = @"
<Window
	xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
	xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
	x:Name="Window"
	Title="Group Templatizer"
	WindowStartupLocation = "CenterScreen"
	SizeToContent = "WidthAndHeight"
	ShowInTaskbar = "True">
	
    <DockPanel>
	    <TabControl x:Name="TabControl" DockPanel.Dock="Top">
	    </TabControl>
        <StackPanel Orientation="Horizontal" DockPanel.Dock="Bottom" HorizontalAlignment="Center">
            <Button x:Name="ButtonOK"
                Content="OK"
                Margin="3"
                Width="50"
            />
            <Button x:Name="ButtonCancel"
                Content="Cancel"
                Margin="3"
                Width="50"
                IsCancel="True"
            />
        </StackPanel>
    </DockPanel>


	
</Window>
"@
    $rows = 6
    $columns = 2
    [System.Xml.XmlNamespaceManager]$nsm = New-Object System.Xml.XmlNamespaceManager $xaml.NameTable
    $nsm.AddNamespace("x", "http://schemas.microsoft.com/winfx/2006/xaml")
    #create each tabitem
    Foreach($user in $users)
    {
        $newEl = $xaml.CreateNode("element","TabItem","")
	    $newEl.SetAttribute("Name", $nsm.LookupNamespace("x"), "TabItem$user") | Out-Null
	    $newEl.SetAttribute("Header", $user) | Out-Null
        $xaml.Window.DockPanel.TabControl.AppendChild($newEl) | Out-Null
    }
    Function New-GridElement{
        Param(
            [xml]$xaml
            ,
            [string]$type
            ,
            [string]$name
            ,
            [int]$column
            ,
            [int]$row
            ,
            [string]$content
            ,
            [System.Collections.Hashtable]$attributes
            ,
            [System.Xml.XmlElement]$grid
        )
        $newEl = $xaml.CreateElement("$Type")
        $newEl.SetAttribute("Name", $nsm.LookupNamespace("x"), $Name) | Out-Null
        $newEl.SetAttribute("Grid.Column",$Column) | Out-Null
        $newEl.SetAttribute("Grid.Row",$Row) | Out-Null
        If($Content){
            $newEl.SetAttribute("Content",$Content) | Out-Null
        }
        If($attributes)
        {
            ForEach($attrib in $Attributes.GetEnumerator()){
                $newEl.SetAttribute($attrib.Name, $attrib.Value) | Out-Null
            }
        }

        $grid.AppendChild($newEl) | Out-Null
    }

    #loadout each tab
    If($xaml.Window.DockPanel.TabControl.TabItem.Count)
    {
        $counter = $xaml.Window.DockPanel.TabControl.TabItem.Count
    }Else{
        $counter = 1
    }

    For($y=0;$y -lt $counter;$y++)
    {
        #grid
        $newGrid = $xaml.CreateNode("element","Grid","")
        #grid row and col defs
        $rowDef = $xaml.CreateNode("element","Grid.RowDefinitions","")
        $colDef = $xaml.CreateNode("element","Grid.ColumnDefinitions","")
        For($x=0;$x -lt $rows-1;$x++)
        {
            $newRow = $xaml.CreateElement("RowDefinition")
            $newRow.SetAttribute("Height","Auto") | Out-Null
            $rowDef.AppendChild($newRow) | Out-Null
        }
        $newRow = $xaml.CreateElement("RowDefinition")
        $newRow.SetAttribute("Height","*") | Out-Null
        $rowDef.AppendChild($newRow) | Out-Null
        For($x=0;$x -lt $columns;$x++)
        {
            $newCol = $xaml.CreateElement("ColumnDefinition")
            $newCol.SetAttribute("Width","Auto") | Out-Null
            $colDef.AppendChild($newcol) | Out-Null
        }
        $newGrid.AppendChild($rowDef) | Out-Null
        $newGrid.AppendChild($colDef) | Out-Null

        $primaryScreenHeight = [System.Windows.SystemParameters]::PrimaryScreenHeight
        $LBHeight = [System.Math]::Max(200,$primaryScreenHeight/2)

        If($counter -eq 1)
        {
            #create listbox
            New-GridElement -xaml $xaml -type ListBox -name "ListBoxGroups$($xaml.Window.DockPanel.TabControl.TabItem.Header)" -column 0 -row 0 -attributes @{"MinWidth"=150;"Grid.RowSpan"=$Rows;"SelectionMode"="Multiple";"MaxHeight"="$LBHeight"} -grid $newGrid
            #create text box
            New-GridElement -xaml $xaml -type TextBox -name "TextBoxUser$($xaml.Window.DockPanel.TabControl.TabItem.Header)" -column 1 -row 0 -attributes @{"Margin"=3;"MinWidth"=60} -grid $newGrid
            #create load button
            New-GridElement -xaml $xaml -type Button -name "ButtonLoad$($xaml.Window.DockPanel.TabControl.TabItem.Header)" -column 1 -row 1 -content Load -attributes @{"Margin"=3} -grid $newGrid
            #create add button
            New-GridElement -xaml $xaml -type Button -name "ButtonAdd$($xaml.Window.DockPanel.TabControl.TabItem.Header)" -column 1 -row 2 -content + -attributes @{"Margin"=3} -grid $newGrid
            #create remove button
            New-GridElement -xaml $xaml -type Button -name "ButtonRemove$($xaml.Window.DockPanel.TabControl.TabItem.Header)" -column 1 -row 3 -content - -attributes @{"Margin"=3} -grid $newGrid
            
            $xaml.Window.DockPanel.TabControl.TabItem.AppendChild($newGrid) | Out-Null
        }Else{
            #create listbox
            New-GridElement -xaml $xaml -type ListBox -name "ListBoxGroups$($xaml.Window.DockPanel.TabControl.TabItem[$y].Header)" -column 0 -row 0 -attributes @{"MinWidth"=150;"Grid.RowSpan"=$Rows;"SelectionMode"="Multiple";"MaxHeight"="$LBHeight"} -grid $newGrid
            #create text box
            New-GridElement -xaml $xaml -type TextBox -name "TextBoxUser$($xaml.Window.DockPanel.TabControl.TabItem[$y].Header)" -column 1 -row 0 -attributes @{"Margin"=3;"MinWidth"=60} -grid $newGrid
            #create load button
            New-GridElement -xaml $xaml -type Button -name "ButtonLoad$($xaml.Window.DockPanel.TabControl.TabItem[$y].Header)" -column 1 -row 1 -content Load -attributes @{"Margin"=3} -grid $newGrid
            #create add button
            New-GridElement -xaml $xaml -type Button -name "ButtonAdd$($xaml.Window.DockPanel.TabControl.TabItem[$y].Header)" -column 1 -row 2 -content + -attributes @{"Margin"=3} -grid $newGrid
            #create remove button
            New-GridElement -xaml $xaml -type Button -name "ButtonRemove$($xaml.Window.DockPanel.TabControl.TabItem[$y].Header)" -column 1 -row 3 -content - -attributes @{"Margin"=3} -grid $newGrid
            
            $xaml.Window.DockPanel.TabControl.TabItem[$y].AppendChild($newGrid) | Out-Null
        }
    }

    [xml]$xaml = $xaml.OuterXml.Replace(" xmlns=`"`"","")
    $reader = (New-Object System.Xml.XmlNodeReader $xaml)
    $Window = [Windows.Markup.XamlReader]::Load($reader)
    $xaml.save("C:\temp\tabs.xml")

    #add any actions below
    For($y=0;$y -lt $counter; $y++) #$xaml.Window.DockPanel.TabControl.TabItem.Count;$y++)
    {
        #add the groupLB loading (buttonload click)
        If($counter -eq 1)
        {
            $loadButton = $Window.FindName("ButtonLoad$($xaml.Window.DockPanel.TabControl.TabItem.Header)")
            $userTB = $Window.FindName("TextBoxUser$($xaml.Window.DockPanel.TabControl.TabItem.Header)")
            $removeButton = $window.FindName("ButtonRemove$($xaml.Window.DockPanel.TabControl.TabItem.Header)")
            $addButton = $window.FindName("ButtonAdd$($xaml.Window.DockPanel.TabControl.TabItem.Header)")
        }Else{
            $loadButton = $Window.FindName("ButtonLoad$($xaml.Window.DockPanel.TabControl.TabItem[$y].Header)")
            $userTB = $Window.FindName("TextBoxUser$($xaml.Window.DockPanel.TabControl.TabItem[$y].Header)")
            $removeButton = $window.FindName("ButtonRemove$($xaml.Window.DockPanel.TabControl.TabItem[$y].Header)")
            $addButton = $window.FindName("ButtonAdd$($xaml.Window.DockPanel.TabControl.TabItem[$y].Header)")
        }


        $loadButton.Add_Click({
            $userTB = $this.Parent.Children | ?{$_.name -like "TextBoxUser*"}
            $groupLB = $this.Parent.Children | ?{$_.name -like "ListBoxGroups*"}
            $groups = @()
            If($userTB.Text)
            {
                Try{
                    (Get-ADUser $userTB.text -prop memberof).memberof | sort | %{$groups += $_.split(",")[0].split("=")[1]} #get each group name without the entire dn
                    $groupLB.Items.Clear()
                }Catch{}
                ForEach($group in $groups)
                {
                    $groupLB.Items.Add($group)
                }
            }
        })

        #add the got focus to the textboxuser
        $userTB.Add_GotFocus({
            $loadButton = $this.Parent.Children | ?{$_.name -like "ButtonLoad*"}
            $loadButton.IsDefault = $true
        })

        #add the remove action
        $removeButton.Add_Click({
            $groupLB = $this.Parent.Children | ?{$_.name -like "ListBoxGroups*"}
            For($x=$groupLB.SelectedItems.Count-1;$x -ge 0;$x--)
            {
                $groupLB.Items.Remove($groupLB.SelectedItems[$x])
            }
        })

        #Add the add button actiov
        $addButton.Add_Click({
            $userTB = $this.Parent.Children | ?{$_.name -like "TextBoxUser*"}
            $groupLB = $this.Parent.Children | ?{$_.name -like "ListBoxGroups*"}
            $groups = @()
            If($userTB.Text)
            {
                (Get-ADUser $userTB.text -prop memberof).memberof | sort | %{$groups += $_.split(",")[0].split("=")[1]} #get each group name without the entire dn
                ForEach($group in $groups)
                {
                    If($groupLB.Items -notcontains $group)
                    {
                        $groupLB.Items.Add($group)
                    }
                }
            }
        })
    }

    $okButton = $window.FindName("ButtonOK")
    $okButton.Add_Click({
        Write-Host "click"
        $return = @()
        ForEach($name in $xaml.Window.DockPanel.TabControl.TabItem.Header)
        {
            $groupLB = $Window.FindName("ListBoxGroups$name")
            $return += New-Object PSObject -Property @{
                "User" = $name
                "Groups" = $groupLB.Items
            }
        }
        $window | Add-Member -MemberType NoteProperty -Name "ReturnValue" -Value $return
        $window.DialogResult = $true
        $window.Close()
    })

    If($window.showdialog()){
        Return $window.ReturnValue
    }


    <#
         0           1
    0 ListBox | TextBoxName
    1   \/    | ButtonLoad
    2   \/    | ButtonAdd
    3   \/    | ButtonRemove
    4   \/    | 

    ListBoxGroups<SamAccountName>

    #>
}

#this is here in case a SQL query is needed for user validation, such as was in the original requirements
Function Execute-SQLQuery{

    Param(
        [string]$SQLquery
        ,
        [string]$server
    )

    $connection = New-Object System.Data.SqlClient.SQLConnection("Server=$server;Integrated Security=TRUE");
    $cmd = new-object System.Data.SqlClient.SqlCommand($sqlQuery, $connection);

    $connection.Open();
    $reader = $cmd.ExecuteReader()

    $results = @()
    while ($reader.Read())
    {
        $row = @{}
        for ($i = 0; $i -lt $reader.FieldCount; $i++)
        {
            $row[$reader.GetName($i)] = $reader.GetValue($i)
        }
        $results += new-object psobject -property $row            
    }
    $connection.Close();
    $results
}

#endregion Functions

#Add or remove attributes from here
#Other validation settings to be marked
$attributes = "GivenName","SurName","SamAccountName","Title","Manager","Department","Company","Mobile","Fax","TelephoneNumber","OU"
$OURoot = "OU=People,DC=Domain,DC=org" #Set base User OU here

#region Build XAML

[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    x:Name="Window"
    Title="Bulk User Creator"
    WindowStartupLocation = "CenterScreen"
    SizeToContent = "WidthAndHeight"
    ShowInTaskbar = "True">
    
    <DockPanel>
    
        <StackPanel x:Name="StackPanel" Orientation="Horizontal" DockPanel.Dock="Bottom">
    
        </StackPanel>
    </DockPanel>

</Window>
"@

[System.Xml.XmlNamespaceManager]$nsm = New-Object System.Xml.XmlNamespaceManager $xaml.NameTable
$nsm.AddNamespace("x", "http://schemas.microsoft.com/winfx/2006/xaml")

#create alphabet array for use in row element names
$alph = [char[]]([int][char]'A'..[int][char]'Z')

#################################################################################################
$a = 5 #stand in for the number of users to be created. to be replaced with a prompt of some kind
#################################################################################################
#the foreach to create users depends on this

$newGrid = $xaml.CreateNode("element","Grid","")
$newGrid.SetAttribute("Name", $nsm.LookupNamespace("x"), "Grid") | Out-Null
$newGrid.SetAttribute("DockPanel.Dock","Top")
#grid row and col defs
$rowDef = $xaml.CreateNode("element","Grid.RowDefinitions","")
$newRow = $xaml.CreateElement("RowDefinition")
$newRow.SetAttribute("Height","Auto") | Out-Null
$rowDef.AppendChild($newRow) | Out-Null

$colDef = $xaml.CreateNode("element","Grid.ColumnDefinitions","")
$newCol = $xaml.CreateElement("ColumnDefinition")
$newCol.SetAttribute("Width","Auto") | Out-Null
$colDef.AppendChild($newcol) | Out-Null
    
$newGrid.AppendChild($rowDef) | Out-Null
$newGrid.AppendChild($colDef) | Out-Null

For($y=0;$y -lt $a;$y++){
    #creating 2 rows, 1 for the text boxes and one for the info row

    #add the user label to each row
    New-BUCGridElement -xaml $xaml -grid $newGrid -Type Label -Name "LabelUser$($alph[$y])" -Column 0 -Row (($y*2)+1) -Content "User$($alph[$y])"

    #for each user, create a control and validate label for each attribute
    For($x=0;$x -lt $attributes.count;$x++){
        Switch($($attributes[$x].ToString().Replace(' ', ''))) #combobox for OU, text box for everything else. Creating as aswitch statement just in case future addition
        {
            "OU" {  New-BUCGridElement -xaml $xaml -grid $newGrid -Type ComboBox -Name "ComboBox$($attributes[$x].ToString().Replace(' ', ''))$($alph[$y])" -Column (($x*2)+1) -Row (($y*2)+1)}
            Default {New-BUCGridElement -xaml $xaml -grid $newGrid -Type TextBox -Name "TextBox$($attributes[$x].ToString().Replace(' ', ''))$($alph[$y])" -Column (($x*2)+1) -Row (($y*2)+1) -ReadOnly ($($attributes[$x].ToString().Replace(' ', '')) -eq "SamAccountName") <#readonly is true for samaccountname#>}
        }

        #validate label
        New-BUCGridElement -xaml $xaml -grid $newGrid -Type Label -Name "LabelValidate$($attributes[$x].ToString().Replace(' ', ''))$($alph[$y])" -Column (($x*2)+2) -Row (($y*2)+1)

        #information labels, one below the TBs and validate labels
        New-BUCGridElement -xaml $xaml -grid $newGrid -Type Label -Name "LabelInformation$($attributes[$x].ToString().Replace(' ',''))$($alph[$y])" -Column (($x*2)+1) -Row (($y+1)*2)
    }
}

#for each attribute, create 2 columns, one for the TB and one for the validation label
For($x=0;$x -lt $attributes.count;$x++){

    #create a label per column, naming it the attribute name
    New-BUCGridElement -xaml $xaml -grid $newGrid -Type Label -Name "LabelColumn$x" -Column (($x*2)+1) -Row 0 -Content $attributes[$x]
}

For($y=0;$y -lt $a;$y++){
    #create the validate button
    New-BUCGridElement -xaml $xaml -grid $newGrid -Type Button -Name "ButtonValidate$($alph[$y])" -Column (($x*2)+1) -Row (($y*2)+1) -Content Validate -Attributes @{"Margin"="3";"MinWidth"="50"}
}

#action buttons
New-BUCGridElement -xaml $xaml -grid $newGrid -Type Button -Name "ButtonGroupTemplatizer" -Column 1 -Row (($a*2)+1) -Content Groups -Attributes @{"Margin"="3"}
New-BUCGridElement -xaml $xaml -grid $newGrid -Type Button -Name "ButtonUserCreate" -Column 3 -Row (($a*2)+1) -Content Create -Attributes @{"Margin"="3"}

$xaml.Window.DockPanel.AppendChild($newgrid) | Out-Null

[xml]$xaml = $xaml.OuterXml.Replace(" xmlns=`"`"","")
$xaml.Save("C:\temp\test.xml") #saving to disk for easier debug



#endregion Build XAML

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$Window = [Windows.Markup.XamlReader]::Load($reader)

$grid = $Window.Content.Children | ?{$_.name -eq "Grid"}
$vButtons = $grid.Children | ?{$_.name -like "ButtonValidate*"}
$vLabels = $grid.Children | ?{$_.name -like "LabelValidate*"}
$OUCBs = $grid.Children | ?{$_.name -like "ComboBoxOU*"}
ForEach($OU in (Get-ChildItem "AD:$OURoot" | ?{$_.objectclass -eq "organizationalunit"}).name)
{
    ForEach($combobox in $OUCBs)
    {
        $combobox.Items.Add($OU) | Out-Null
    }
}


ForEach($vButton in $vButtons){
    $vButton.Add_Click({
        #grab the row letter
        $letter = $this.Name[$($this.name.length)-1]
        $textboxes = $grid.Children | ?{$_.name -like "TextBox*$letter"}

        #validate templates
        ForEach($tb in $textboxes){
            $attrib = $tb.name -replace "^TextBox(\w*)$letter", '$1' #get the attribute name
            ($grid.Children | ?{$_.name -eq "LabelValidate$attrib$letter"}).content = (SwitchAttribs $attrib) #+ is passed, ! is failed, '' is no validation defined
            If("SurName","GivenName" -contains $attrib)
            {
                #convert names to title case
                ($grid.Children | ?{$_.name -eq "TextBox$attrib$letter"}).Text = (Get-Culture).TextInfo.ToTitleCase("$(($grid.Children | ?{$_.name -eq "TextBox$attrib$letter"}).Text)".tolower())
            }
        }
    })
}

#set the validate button as default when focus is gotten in a textbox in the same row
For($y=0;$y -lt $a;$y++){
    $rowTBs = $grid.Children | ?{$_.name -like "TextBox*$($alph[$y])"}
    ForEach($TB in $rowTBs){
        $TB.Add_GotFocus({
            ($grid.Children | ?{$_.name -eq "ButtonValidate$($this.Name[$($this.name.length-1)])"}).IsDefault = $true #set the row's validate button as default for each TB
            $buttons = $grid.Children | ?{$_.name -like "ButtonValidate*"} | ?{$_.name -notlike "*$($this.Name[$($this.name.length-1)])"} #turn off default for other buttons
            ForEach($button in $buttons){
                $button.IsDefault = $false
            }
        })
    }
}

#group templatizer action
$gtemplatizerButton = $grid.Children | ?{$_.name -eq "ButtonGroupTemplatizer"}
$gtemplatizerButton.Add_Click({
    [string[]]$users = ($grid.Children | ?{$_.name -like "TextBoxSamAccountName*"}).text | ?{$_}
    If($users)
    {
        $global:groups = New-GroupTemplatizer $users
    }Else{
        #Write-Host "none"
    }
})

$createButton = $grid.Children | ?{$_.name -eq "ButtonUserCreate"}
$createButton.Add_Click({
    #make users here!
    
	$DC = Get-ADDomainController | Select-object -First 1 -ExpandProperty HostName
	
    While(!($connected))
    {
        $creds = Get-Credential -Message "Credentials for creating mailbox in Exchange"
        Try
        {
            $exchSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$exchCASServ/PowerShell/" -Authentication Kerberos -Credential $creds -ErrorAction stop
            Import-PSSession $exchSession -commandname Enable-Mailbox,Set-Mailbox,Get-Mailbox,Set-CASMailbox,New-Mailbox -AllowClobber | Out-Null
            $connected = $true
        }Catch{
            $creds = Get-Credential -Message "Incorrect credentials, try again"
        }
    }

    For($x=0;$x -lt $a;$x++)
    {
        $fname = $lname = $username = $title = $manager = $department = $mobile = $fax = $phone = $OUCB = $null

        $letter = $alph[$x]
        $textBoxes = $grid.Children | ?{$_.name -like "TextBox*$letter"}
        
        $fname = ($textBoxes | ?{$_.name -like "*givenname*"}).Text
        $lname = ($textBoxes | ?{$_.name -like "*surname*"}).Text
        $username = ($textBoxes | ?{$_.name -like "*samaccountname*"}).Text
        $title = ($textBoxes | ?{$_.name -like "*title*"}).Text
        $manager = ($textBoxes | ?{$_.name -like "*manager*"}).Text
        $department = ($textBoxes | ?{$_.name -like "*department*"}).Text
        $company = ($textBoxes | ?{$_.name -like "*company*"}).Text
        $mobile = ($textBoxes | ?{$_.name -like "*mobile*"}).Text
        $fax = ($textBoxes | ?{$_.name -like "*fax*"}).Text
        $phone = ($textBoxes | ?{$_.name -like "*telephonenumber*"}).Text
        $OUCB = ($grid.Children | ?{$_.name -eq "ComboBoxOU$letter"}).SelectedItem

        Write-Host "fname: $fname"
        Write-Host "lname: $lname"
        Write-Host "username: $username"
        Write-Host "title: $title"
        Write-Host "manager: $manager"
        Write-Host "department: $department"
        Write-Host "mobile: $mobile"
        Write-Host "fax: $fax"
        Write-Host "phone: $phone"
        Write-Host "ou: $OUCB"

        If(!($username))
        {
            Break
        }

        #create the user
        Write-Host "Creating user with username: $username..."
        New-ADUser -Name "$fname $lname" -GivenName $fname -Surname $lname -SAMAccountName $username -UserPrincipalName "$username`@$domainA" -EmailAddress "$username`@$emailDomain" -Server $DC

        #create the password
        [char[]]$chars = 'a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z','0','1','2','3','4','5','6','7','8','9','!','@','#','$','%','^','&','*','<','>','/','?','=','+','\','-','_'
        $validPassword = $false
        While(!$validPassword)
        {
            $password = ($chars | Get-Random -count 10) -join "" #Passwordlength
            Try
            {
                Set-ADAccountPassword $username -NewPassword ($password | ConvertTo-SecureString -AsPlainText -Force) -Server $DC
                Set-ADUser $username -Enabled $true -Server DC
		        $validPassword = $true
            }Catch{}
        }

        #updating the user object
        $user = Get-ADUser $username -Server $DC

        #phone numbers
        Write-Host "Setting phone numbers..."
        If($mobile)
        {
            Try{Set-ADUser $username -Mobile $mobile -ErrorAction SilentlyContinue -Server $DC
            }Catch{}
        }Else{
            Write-Host " No mobile."
        }
        If($fax)
        {
            Try{Set-ADUser $username -Fax $fax -ErrorAction SilentlyContinue -Server $DC
            }Catch{}
        }Else{
            Write-Host " No fax."
        }
        If($phone)
        {
            Try{$tUser = Get-ADUser $username -Properties TelephoneNumber -Server $DC
                $tUser.TelephoneNumber = $phone
                Set-ADUser -Instance $tUser -ErrorAction SilentlyContinue -Server $DC
            }Catch{}
        }Else{
            Write-Host " No phone."
        }

        If($title)
        {
            Write-Host "Setting title..."
            Set-ADUser $username -Title "$title" -Server $DC
        }

        If($manager)
        {
            Write-Host "Setting manager..."
            Set-ADUser $username -Manager $manager -Server $DC
        }

        If($department)
        {
            Write-Host "Setting department..."
            Set-ADUser $username -Department "$department" -Server $DC
        }

        If($company)
        {
            Write-Host "Setting company..."
            Set-ADUser $username -Company "$company" -Server $DC
        }

        Write-Host "Moving user..."
        Move-ADObject (Get-ADUser $username -Server $DC).DistinguishedName -TargetPath "OU=$OUCB,$OURoot"

        $userGroups = ($groups | ?{$_.user -eq $username}).Groups
        If($userGroups)
        {
            "Adding to groups..."
            Foreach($group in $userGroups)
            {
                Write-Host " Adding to $group"
                Add-ADGroupMember $group -Members $username -Server $DC
            }
        }


        #New-Mailbox -UserPrincipalName $upn -Name "$fname $lname" -LinkedDomainController $domainBDC -LinkedMasterAccount $domainB\$username -LinkedCredential $domainBCreds
        Set-ADUser $username -Enabled $false -Server $domainADC
        Write-Host "Enabling mailbox..."
        Enable-Mailbox $username -DomainController $DC
        Set-ADUser $username -Enabled $true -Server $ADC
        Write-Host "$fname's password set to: $password"
    }
    Remove-PSSession $exchSession
})

$window.showdialog()


<#
    0         1                2                3                4                  5                 6
0       | LabelCol1 |                      | LabelCol2 |                      | LabelCol3 |
1 UserA | TBAttr01A | LabelValidateAttr01A | TBAttr02A | LabelValidateAttr02A | TBAttr03A | LabelValidateAttr03A
2       | InfLabelA |                      | InfLabelA |                      | InfLabelA | 
3 UserB | TBAttr01B | LabelValidateAttr01B | TBAttr02B | LabelValidateAttr02B | TBAttr03B | LabelValidateAttr03B
4       | InfLabelB |                      | InfLabelB |                      | InfLabelB | 
5 UserC  etc
6
7

#>