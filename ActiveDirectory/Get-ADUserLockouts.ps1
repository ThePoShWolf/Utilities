Function Get-ADUserLockouts {
    [CmdletBinding(
        DefaultParameterSetName = 'All'
    )]
    param (
        [Parameter(
            ValueFromPipeline = $true,
            ParameterSetName = 'ByUser'
        )]
        [Microsoft.ActiveDirectory.Management.ADUser]$Identity
        ,
        [datetime]$StartTime
        ,
        [datetime]$EndTime
    )
    Begin{
        $filterHt = @{
            LogName = 'Security'
            ID = 4740
        }
        if ($PSBoundParameters.ContainsKey('StartTime')){
            $filterHt['StartTime'] = $StartTime
        }
        if ($PSBoundParameters.ContainsKey('EndTime')){
            $filterHt['EndTime'] = $EndTime
        }
        $PDCEmulator = (Get-ADDomain).PDCEmulator
        # Query the event log just once instead of for each user if using the pipeline
        $events = Get-WinEvent -ComputerName $PDCEmulator -FilterHashtable $filterHt
    }
    Process {
        if ($PSCmdlet.ParameterSetName -eq 'ByUser'){
            $user = Get-ADUser $Identity
            # Filter the events
            $output = $events | Where-Object {$_.Properties[0].Value -eq $user.SamAccountName}
        } else {
            $output = $events
        }
        foreach ($event in $output){
            [pscustomobject]@{
                UserName = $event.Properties[0].Value
                CallerComputer = $event.Properties[1].Value
                TimeStamp = $event.TimeCreated
            }
        }
    }
    End{}
}