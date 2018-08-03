#requires -module GroupPolicy
<#
.SYNOPSIS
	Returns the drivemaps from the specific group policy name.
.DESCRIPTION
	Exports a GPO as XML and parses the xml to return the drive mappings contained there in.
.PARAMETER GPOname
	The name of the group policy object to be used. A list can be found using Get-GPO -All
.EXAMPLE
	Get-GPODriveMaps "Users"
    
    PAth                                                                             Group                                        Letter Order
    ----                                                                             -----                                        ------ -----
    \\domain.org\dfs\Users\%username%\My documents                                                                                H      1
    \\domain.org\dfs\Groups                                                                                                       G      2
    \\domain.org\dfs\IT                                                            DOMAIN\IT Department                           J      3
    \\domain.org\dfs\Groups\Sales                                                  DOMAIN\Sales Department                        J      4
    \\domain.org\dfs\Groups\HR                                                     DOMAIN\HR Department                           J      5    
    ...
.INPUTS
    The string name of the GPO to query
.OUTPUTS
    A custom PSObject with the following properties:

    [string]Path - The path of the drive to be mapped
    [string]Group - The group that it applies to, if any
    [string]Letter - The letter assigned to the drive
    [string]Order - The order in which the drive is mapped
.NOTES
	Author: Anthony Howell
.LINK
	Get-GPO
#>
Function Get-GPODriveMaps
{
    Param
    (
        [string]$GPOname = "Users"
    )
    [xml]$gpo = Get-GPOReport -Name $GPOName -ReportType XML
    $driveMapSetings = ($gpo.GPO.user.ExtensionData | ?{$_.name -eq "Drive Maps"}).Extension.DriveMapSettings
    ForEach($drivemap in $driveMapSetings.drive)
    {
        If($drivemap.Filters.FilterGroup)
        {
            $groups = $drivemap.Filters.FilterGroup.Name
        }Else{
            $groups = $null
        }
        New-Object PSObject -Property @{
            "Letter" = $drivemap.Properties.Letter
            "Path" = $drivemap.Properties.Path
            "Group" = $groups
            "Order" = $drivemap.GPOSettingOrder
        }
    }
}