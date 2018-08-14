<#

    Use text file with list of machines or list of machines to add or remove specified users in the local group specified (assumes user running has the rights to do this)

    @guyrleech (c) 2018

    Modification history:
#>

<#
.SYNOPSIS

Add or remove accounts from groups on Windows machines. 

.PARAMETER machines

A comma separated list of machines names to make the group changes on

.PARAMETER machinesFile

A text file containing one machine per line

.PARAMETER users

A comma separated list of user names add or remove from the group specified

.PARAMETER usersFile

A text file containing one user name per line

.PARAMETER group

The name of the group to make the changes to

.PARAMETER domain

The domain containing the user accounts. User names prefixed with domain\ will override the domain specified

.PARAMETER remove

Remove the specified users from the group specified rather than adding them

.EXAMPLE

& '.\Group membership modifier.ps1' -machines machine01,machines02 -users user1,user2 -group "Remote Desktop Users"

Add the specified users to the specified group on the specified machines

.EXAMPLE

& '.\Group membership modifier.ps1' -machinesFile c:\temp\machines.txt -usersFile c:\temp\users.txt -remove

Remove the users specified one per line in the file c:\temp\users.txt from the "Administrators" group on the machines specified one per line in c:\temp\machines.txt

.NOTES

The user running the script must have the rights to perform the group changes otherwise they will fail
#>

[CmdletBinding()]

Param
(
    [string[]]$machines = @() ,
    [string]$machinesFile ,
    [string[]]$users ,
    [string]$usersFile ,
    [string]$group = 'Administrators' ,
    [string]$domain = $env:USERDOMAIN ,
    [switch]$remove
)

if( ! [string]::IsNullOrEmpty( $machinesFile ) )
{
    $machines += Get-Content $machinesFile -ErrorAction Stop
}

if( ! [string]::IsNullOrEmpty( $usersFile ) )
{
    $users += Get-Content $usersFile -ErrorAction Stop
}

[int]$missingUsers = 0

[array]$adUsers = 
    @( ForEach( $user in $users )
    {
        if( $user -match '^[\- _a-z0-9]' )
        {
            [string]$domainName,[string]$userName = $user.Trim() -split '\\'

            if( [string]::IsNullOrEmpty( $userName ) )
            {
                $userName = $domainName
                $domainName = $domain
            }
            $thisUser = [ADSI]"WinNT://$domainName/$userName,user"
            if( ! $thisUser.Path )
            {
                Write-Error "Failed to find user $domainName\$userName"
                $missingUsers++
            }
            else
            {
                $thisUser
            }
        }
    })

if( $missingUsers )
{
    Write-Error "Failed to find $missingUsers user(s) - aborting"
    Exit 2
}

if( ! $adUsers.Count )
{
    Write-Error "No users specified - use -users or -usersFile - aborting"
    Exit 3
}

[string]$verb = $null
[string]$preposition = $null

if( $remove )
{
    $verb = 'Removing'
    $preposition = 'from'
}
else
{
    $verb = 'Adding'
    $preposition = 'to'
}

[int]$errors = 0

$machines | ForEach-Object `
{
    $computerName = $_.Trim()
    if( $computerName -match '^[a-z0-9]' )
    {
        $localGroup = [ADSI]"WinNT://$computerName/$group,group"
        ForEach( $adUser in $adUsers )
        {
            [string]$operation = "$verb $((($aduser.Path -split ':')[1] -split ',')[0] -replace '//' , '' -replace '/' , '\') $preposition `"$group`" on $computerName"
            Write-Verbose $operation
            try
            {
                if( $remove )
                {
                    $localGroup.Remove( $adUser.Path )
                }
                else
                {
                    $localGroup.Add( $adUser.Path )
                }
            }
            catch
            {
                Write-Error "Error $($operation.Substring(0,1).ToLower() + $operation.Substring(1))) - $($_.Exception.Message)"
                $errors++
            }
        }
    }
}

Write-Verbose "Finished with $errors errors"

Exit $errors
