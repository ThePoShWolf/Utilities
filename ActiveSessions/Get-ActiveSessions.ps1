<#
.SYNOPSIS
	Returns open sessions of a remote or local workstations
.DESCRIPTION
	Get-ActiveSessions uses the command line tool qwinsta to retrieve all open user sessions on a computer regardless of how they are connected.
.PARAMETER $ComputerName
	The name of the computer that you would like to know who is logged into.
.EXAMPLE
	Get-ActiveSessions PC1

	Returns all users logged onto PC1 as well as pointing out which one is actually active.
.EXAMPLE
	Get-ADComputer -filter 'name -like "IT*"' | select name | Get-ActiveSessions

	This will go through each and every computer that has a name like "IT*" and return the users that are logged into it.
.INPUTS
	[string[]]$ComputerName
.OUTPUTS
	A custom object with the following members:
		UserName: [string]
		SessionName: [string]
		ID: [string]
		Type: [string]
		State: [string]
.NOTES
	Author: Anthony Howell
.LINK
	qwinsta
	http://stackoverflow.com/questions/22155943/qwinsta-error-5-access-is-denied
	https://theposhwolf.com
#>
Function Get-ActiveSessions{
	Param(
		[Parameter(
			Mandatory = $true,
			ValueFromPipeline = $true
		)]
		[ValidateNotNullOrEmpty()]
		[string]$ComputerName
        ,
        [switch]$Quiet
	)
	Begin{
		$return = @()
	}
	Process{
		If(!(Test-Connection $ComputerName -Quiet -Count 1)){
			Write-Error -Message "Unable to contact $ComputerName. Please verify its network connectivity and try again." -Category ObjectNotFound -TargetObject $ComputerName
			Return
		}
		If([bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544")){
			#the following registry key is necessary to avoid the error 5 access is denied error
			$LMtype = [Microsoft.Win32.RegistryHive]::LocalMachine
			$LMkey = "SYSTEM\CurrentControlSet\Control\Terminal Server"
			$LMRegKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($LMtype,$ComputerName)
			$regKey = $LMRegKey.OpenSubKey($LMkey,$true)
			If($regKey.GetValue("AllowRemoteRPC") -ne 1){
				$regKey.SetValue("AllowRemoteRPC",1)
				Start-Sleep -Seconds 1
			}
			$regKey.Dispose()
			$LMRegKey.Dispose()
		}
		$result = qwinsta /server:$ComputerName
		If($result){
			ForEach($line in $result[1..$result.count]){ #avoiding the line 0, don't want the headers
				$tmp = $line.split(" ") | ?{$_.length -gt 0}
				If(($line[19] -ne " ")){ #username starts at char 19
					If($line[48] -eq "A"){ #means the session is active ("A" for active)
						$return += New-Object PSObject -Property @{
							"ComputerName" = $ComputerName
							"SessionName" = $tmp[0]
							"UserName" = $tmp[1]
							"ID" = $tmp[2]
							"State" = $tmp[3]
							"Type" = $tmp[4]
						}
					}Else{
						$return += New-Object PSObject -Property @{
							"ComputerName" = $ComputerName
							"SessionName" = $null
							"UserName" = $tmp[0]
							"ID" = $tmp[1]
							"State" = $tmp[2]
							"Type" = $null
						}
					}
				}
			}
		}Else{
			Write-Error "Unknown error, cannot retrieve logged on users"
		}
	}
	End{
		If($return){
			If($Quiet){
                Return $true
            }
            Else{
                Return $return
            }
		}Else{
		    If(!($Quiet)){
                Write-Host "No active sessions."
            }
		    Return $false
        }
	}
}