Function Remove-InactiveUserProfiles {
    [cmdletbinding(
        DefaultParameterSetName = 'FromGet'
    )]
    param (
        [Parameter(
            ParameterSetName = 'FromSelf'
        )]
        [ValidateNotNullOrEmpty()]
        [string]$Name = $env:COMPUTERNAME
        ,
        [Parameter(
            ParameterSetName = 'FromSelf'
        )]
        [string[]]$include # Supports ? and *
        ,
        [Parameter(
            ParameterSetName = 'FromSelf'
        )]
        [string[]]$exclude # Supports ? and *
        ,
        [Parameter(
            ParameterSetName = 'FromSelf'
        )]
        [int]$olderThan
        ,
        [Parameter(
            ParameterSetName = 'FromGet',
            ValueFromPipeline = $true
        )]
        [PSObject]$ProfileObject
        ,
        [string]$delprof2 = "c:\windows\system32\delprof2.exe"
    )
    Begin {
        Write-Verbose $PSCmdlet.ParameterSetName
        $switches = & {
            "/u"
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
        if ($PSCmdlet.ParameterSetName -eq 'FromSelf'){
            if ($PSBoundParameters.ContainsKey('Name') -and ($Name -ne $env:COMPUTERNAME)) {
                if (Test-Connection $Name -Count 1 -Quiet) {
                    $computer = "/c:$Name"
                } else {
                    Throw "Cannot ping $Name"
                }
            }
            Write-Verbose "CMD: $delprof2 $computer $switches"
            $return = & $delprof2 $computer $switches
        } elseif ($PSCmdlet.ParameterSetName -eq 'FromGet'){
            if ( -not $profileObject.Ignored) {
                $computer = $null
                if ($null -ne $profileObject.Computer){
                    $computer = "/c:$($profileObject.Computer)"
                }
                Write-Verbose "CMD: $delprof2 $computer $switches /id:$($profileObject.Profile)"
                $return = & $delprof2 $computer $switches "/id:$($profileObject.Profile)"
            }
        }
        for ($x=0; $x -lt $return.Count; $x++){
            if ($return[$x] -match "^Ignoring profile \'(?<profilePath>\\\\([^\\]+\\)+(?<profile>[^\']+))\' \(reason\: (?<reason>[^\)]+)\)") {
                # Ignored profile
                If($PSCmdlet.ParameterSetName -ne 'FromGet'){
                    [pscustomobject]@{
                        Ignored = $true
                        Reason = $Matches.reason
                        ProfilePath = $Matches.profilePath
                        Profile = $Matches.profile
                        ComputerName = $Name
                        Deleted = $false
                    }
                }
            } elseif ($return[$x] -match "^Deleting profile \'(?<profilePath>\\\\([^\\]+\\)+(?<profile>[^\']+))\'") {
                # Profile to delete
                $firstMatch = $Matches
                $deleted = $false
                if ($return[$x+1] -match '\.\.\. done\.'){
                    $deleted = $true
                }
                [pscustomobject]@{
                    Ignored = $false
                    Reason = $null
                    ProfilePath = $firstMatch.ProfilePath
                    Profile = $firstMatch.Profile
                    ComputerName = $Name
                    Deleted = $deleted
                }
                $x++
            } elseif ($return[$x] -match "^Access denied to profile \'(?<profilePath>\\\\([^\\]+\\)+(?<profile>[^\']+))\'") {
                # Access denied
                if ($PSCmdlet.ParameterSetName -ne 'FromGet'){
                    [pscustomobject]@{
                        Ignored = $true
                        Reason = 'Access denied'
                        ProfilePath = $Matches.ProfilePath
                        Profile = $Matches.Profile
                        ComputerName = $Name
                        Deleted = $false
                    }
                }
            }
        }
    }
    End{}
}