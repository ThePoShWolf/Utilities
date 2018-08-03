Function Get-GPImpact
{
    [cmdletbinding()]
    Param
    (
        [Parameter(ParameterSetName="Name",Mandatory="True")]
        [string]$Name
        ,
        [Parameter(ParameterSetName="ID",Mandatory="True")]
        [string]$ID
    )
    Begin{}
    Process
    {
        Switch ($psCmdlet.ParameterSetName)
        {
            "Name" {
                Write-Verbose "using name"
                Try
                {
                    $gpo = Get-GPO -Name $Name
                }
                Catch
                {
                    Write-Error $error[0]
                    Return
                }
                $scope = Get-GPPermission -All -Name $Name | ?{$_.Permission -eq "GpoApply"}
                Write-Verbose "Scope: $($scope.trustee.name)"
            }
            "ID" {
                Write-Verbose "using id"
                Try
                {
                    $gpo = Get-GPO -Id $ID
                }
                Catch
                {
                    Write-Error $error[0]
                    Return
                }
                $scope = Get-GPPermission -All -Id $ID | ?{$_.Permission -eq "GpoApply"}
                Write-Verbose "Scope: $($scope.trustee.name)"
            }
        }
        $ous = Get-ADOrganizationalUnit -Filter * -Properties gplink | ?{$_.gplink -like "*$($gpo.id)*"}
        If((Get-ADDomain).LinkedGroupPolicyObjects -like "*$($gpo.id)*")
        {
            $ous += (Get-ADDomain).DistinguishedName
        }
        If($ous)
        {
            If($scope.Trustee.Name -contains "Authenticated Users")
            {
                Write-Verbose "Authenticated Users"
                ForEach($ou in $ous)
                {
                    Write-Verbose $ou.distinguishedname
                    Get-ADObject -SearchBase $ou.distinguishedname -Filter * | ?{@("user","computer") -contains $_.objectclass} | Select Name,ObjectClass
                }
            }
            ElseIf($scope.Trustee)
            {
                Write-Verbose "Custom Groups"
                $groupMembers = @()
                ForEach($group in ($scope | ?{$_.Trustee.SidType -eq "Group"}))
                {
                    Write-Verbose "Group: $($group.trustee.name)"
                    $groupMembers += Get-ADGroupMember $group.Trustee.Name
                }
                $groupMembers = $groupMembers | Select -Unique
                Foreach($ou in $ous)
                {
                    Write-Verbose $ou.DistinguishedName
                    Get-ADObject -SearchBase $ou.DistinguishedName -Filter * | ?{@("user","computer") -contains $_.objectclass} | ?{$groupMembers.Name -contains $_.name} | select Name,ObjectClass
                }
            }Else{
                Write-Error "GPO has no scope."
            }
        }Else{
            Write-Host "GPO has no links"
        }
    }
    End{}
}