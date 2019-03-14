<#
.SYNOPSIS
	Get-InstalledSoftware retrieves a list of installed software
.DESCRIPTION
	Get-InstalledSoftware opens up the specified (remote) registry and scours it for installed software. When found it returns a list of the software and it's version.
.PARAMETER ComputerName
	The computer from which you want to get a list of installed software. Defaults to the local host.
.EXAMPLE
	Get-InstalledSoftware DC1
	
	This will return a list of software from DC1. Like:
	Name			Version		Computer  UninstallCommand
	----			-------     --------  ----------------
	7-Zip 			9.20.00.0	DC1       MsiExec.exe /I{23170F69-40C1-2702-0920-000001000000}
	Google Chrome	65.119.95	DC1       MsiExec.exe /X{6B50D4E7-A873-3102-A1F9-CD5B17976208}
	Opera			12.16		DC1		  "C:\Program Files (x86)\Opera\Opera.exe" /uninstall
.EXAMPLE
	Import-Module ActiveDirectory
	Get-ADComputer -filter 'name -like "DC*"' | Get-InstalledSoftware
	
	This will get a list of installed software on every AD computer that matches the AD filter (So all computers with names starting with DC)
.INPUTS
	[string[]]Computername
.OUTPUTS
	PSObject with properties: Name,Version,Computer,UninstallCommand
.NOTES
	Author: ThePoShWolf
	
	To add registry directories, add to the lmKeys (LocalMachine)
.LINK
	[Microsoft.Win32.RegistryHive]
    [Microsoft.Win32.RegistryKey]
    https://github.com/theposhwolf/utilities
#>
Function Get-InstalledSoftware {
    Param(
        [Alias('Computer','ComputerName','HostName')]
        [Parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$true,Mandatory=$true,Position=1)]
        [string[]]$Name = $env:COMPUTERNAME
    )
    Begin{
        $lmKeys = "Software\Microsoft\Windows\CurrentVersion\Uninstall","SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        $lmReg = [Microsoft.Win32.RegistryHive]::LocalMachine
        $cuKeys = "Software\Microsoft\Windows\CurrentVersion\Uninstall"
        $cuReg = [Microsoft.Win32.RegistryHive]::CurrentUser
    }
    Process{
        if (!(Test-Connection -ComputerName $Computer -count 1 -quiet)) {
            Write-Error -Message "Unable to contact $Computer. Please verify its network connectivity and try again." -Category ObjectNotFound -TargetObject $Computer
            Break
        }
        $masterKeys = @()
        $remoteCURegKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($cuReg,$computer)
        $remoteLMRegKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($lmReg,$computer)
        foreach ($key in $lmKeys) {
            $regKey = $remoteLMRegKey.OpenSubkey($key)
            foreach ($subName in $regKey.GetSubkeyNames()) {
                foreach($sub in $regKey.OpenSubkey($subName)) {
                    $masterKeys += (New-Object PSObject -Property @{
                        "ComputerName" = $Computer
                        "Name" = $sub.GetValue("displayname")
                        "SystemComponent" = $sub.GetValue("systemcomponent")
                        "ParentKeyName" = $sub.GetValue("parentkeyname")
                        "Version" = $sub.GetValue("DisplayVersion")
                        "UninstallCommand" = $sub.GetValue("UninstallString")
                        "InstallDate" = $sub.GetValue("InstallDate")
                        "RegPath" = $sub.ToString()
                    })
                }
            }
        }
        foreach ($key in $cuKeys) {
            $regKey = $remoteCURegKey.OpenSubkey($key)
            if ($regKey -ne $null) {
                foreach ($subName in $regKey.getsubkeynames()) {
                    foreach ($sub in $regKey.opensubkey($subName)) {
                        $masterKeys += (New-Object PSObject -Property @{
                            "ComputerName" = $Computer
                            "Name" = $sub.GetValue("displayname")
                            "SystemComponent" = $sub.GetValue("systemcomponent")
                            "ParentKeyName" = $sub.GetValue("parentkeyname")
                            "Version" = $sub.GetValue("DisplayVersion")
                            "UninstallCommand" = $sub.GetValue("UninstallString")
                            "InstallDate" = $sub.GetValue("InstallDate")
                            "RegPath" = $sub.ToString()
                        })
                    }
                }
            }
        }
        $woFilter = {$null -ne $_.name -AND $_.SystemComponent -ne "1" -AND $null -eq $_.ParentKeyName}
        $props = 'Name','Version','ComputerName','Installdate','UninstallCommand','RegPath'
        $masterKeys = ($masterKeys | Where-Object $woFilter | Select-Object $props | Sort-Object Name)
        $masterKeys
    }
    End{}
}