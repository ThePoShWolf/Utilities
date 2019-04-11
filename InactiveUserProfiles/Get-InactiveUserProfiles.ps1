Function Get-InactiveUserProfiles {
    [cmdletbinding()]
    param (
        [Parameter(
            Position = 1,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('Computer','ComputerName','HostName')]
        [ValidateNotNullOrEmpty()]
        [string]$Name = $env:COMPUTERNAME
        ,
        [Parameter()]
        [string[]]$include # Supports ? and *
        ,
        [Parameter()]
        [string[]]$exclude # Supports ? and *
        ,
        [Parameter()]
        [int]$olderThan
        ,
        [string]$delprof2 = "c:\windows\system32\delprof2.exe"
    )
    Begin {
        $switches = & {
            "/l"
            if ($include) {
                foreach ($inc in $include) {
                    "/id:$inc"
                }
            }
            if ($exclude) {
                foreach ($exc in $exclude) {
                    "/ed:$exc"
                }
            }
            if ($olderThan) {
                "/d:$olderThan"
            }
        }
    }
    Process {
        if ($PSBoundParameters.ContainsKey('Name') -and ($Name -ne $env:COMPUTERNAME)) {
            if (Test-Connection $Name -Count 1 -Quiet) {
                $computer = "/c:$Name"
            } else {
                Throw "Cannot ping $Name"
            }
        }
        Write-Verbose "CMD: $delprof2 $computer $switches"
        $return = & $delprof2 $computer $switches
        foreach ($line in $return) {
            if ($line -match "^Ignoring profile \'(?<profilePath>\\\\([^\\]+\\)+(?<profile>[^\']+))\' \(reason\: (?<reason>[^\)]+)\)") {
                # Ignored profile
                [pscustomobject]@{
                    Ignored = $true
                    Reason = $Matches.reason
                    ProfilePath = $Matches.profilePath
                    Profile = $Matches.profile
                    ComputerName = $Name
                }
            } elseif ($line -match "^(?<profilePath>\\\\([^\\]+\\)+(?<profile>\S+))") {
                # Included profile
                [pscustomobject]@{
                    Ignored = $false
                    Reason = $null
                    ProfilePath = $Matches.ProfilePath
                    Profile = $Matches.Profile
                    ComputerName = $Name
                }
            } elseif ($line -match "^Access denied to profile \'(?<profilePath>\\\\([^\\]+\\)+(?<profile>[^\']+))\'") {
                # Access denied
                [pscustomobject]@{
                    Ignored = $true
                    Reason = 'Access denied'
                    ProfilePath = $Matches.ProfilePath
                    Profile = $Matches.Profile
                    ComputerName = $Name
                }
            }
        }
    }
    End{}
}