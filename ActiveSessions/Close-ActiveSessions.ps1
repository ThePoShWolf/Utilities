<#
.SYNOPSIS
	Closes the specified open sessions of a remote or local workstations
.DESCRIPTION
	Close-ActiveSessions uses the command line tool rwinsta to close the specified open sessions on a computer whether they are RDP or logged on locally.
.PARAMETER $Name
	The name of the computer that you would like to close the active sessions on.
.PARAMETER ID
    The ID number of the session to be closed.
.EXAMPLE
	Close-ActiveSessions -ComputerName PC1 -ID 1

	Closes the session with ID of 1 on PC1
.EXAMPLE
	Get-ActiveSessions DC01 | ?{$_.State -ne "Active"} | Close-ActiveSessions

	Closes all sessions that are not active on DC01
.INPUTS
	[string]$Name
    [int]ID
.OUTPUTS
	Progress of the rwinsta command.
.NOTES
	Author: Anthony Howell
.LINK
	rwinsta
    http://stackoverflow.com/questions/22155943/qwinsta-error-5-access-is-denied
    https://theposhwolf.com
#>
Function Close-ActiveSessions {
    Param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 1
        )]
        [string]$ComputerName,
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 2
        )]
        [int]$ID
    )
    Begin {}
    Process {
        If (!(Test-Connection $ComputerName -Quiet -Count 1)) {
            Write-Error -Message "Unable to contact $ComputerName. Please verify its network connectivity and try again." -Category ObjectNotFound -TargetObject $ComputerName
            Return 
        }
        If([bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544")){ #check if user is admin, otherwise no registry work can be done
            #the following registry key is necessary to avoid the error 5 access is denied error
            $LMtype = [Microsoft.Win32.RegistryHive]::LocalMachine
            $LMkey = "SYSTEM\CurrentControlSet\Control\Terminal Server"
            $LMRegKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($LMtype,$Name)
            $regKey = $LMRegKey.OpenSubKey($LMkey,$true)
            If($regKey.GetValue("AllowRemoteRPC") -ne 1){
                $regKey.SetValue("AllowRemoteRPC",1)
                Start-Sleep -Seconds 1
            }
            $regKey.Dispose()
            $LMRegKey.Dispose()
        }
        rwinsta.exe /server:$ComputerName $ID /V
    }
    End {}
}