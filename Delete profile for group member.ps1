<#
.SYNOPSIS

Delete profiles for members of an AD group which aren't in use

.PARAMETER group

The domain qualified group name for which any unloaded profiles found will be deleted

.EXAMPLE

& '.\Delete profile for group member.ps1' -group 'guyrleech\Delete Profiles'

Delete all local user profiles, if the "are you sure" prompt is confirmed, which are unloaded where the profile is for a member of the "Delete Profiles" AD group in domain "guyrleech"

.EXAMPLE

& '.\Delete profile for group member.ps1' -group 'guyrleech\Delete Profiles' -confirm:$false

Delete all local user profiles, without prompting for confirmation, which are unloaded where the profile is for a member of the "Delete Profiles" AD group in domain "guyrleech"

.NOTES

USE THIS SCRIPT AT YOUR OWN RISK. THE AUTHOR ACCEPTS NO RESPONSIBILITY FOR ANY UNINTENDED LOSS OR DAMAGE CAUSED BY USING THIS SCRIPT.

Modification History:

@guyrleech 21/05/20  Initial release
#>

[CmdletBinding(SupportsShouldProcess,ConfirmImpact='High')]

Param
(
    [Parameter(Mandatory,HelpMessage='Domain qualified group name whose members will have profiles deleted')]
    [string]$group
)

## check group is domain\groupname format
[string[]]$groupparts = $group -split '\\'
if( ! $groupparts -or $groupparts.Count -ne 2 )
{
    Throw "Group `"$group`" not in domain\group format"
}

$domain = $groupparts[0]
$group = $groupparts[1]

## Check group exists

if( ! ( $adGroup = ([ADSISearcher]"Name=$group").FindOne() ) )
{
    Throw "Unable to find group `"$domain\$group`""
}

[int]$totalUnloadedProfiles = 0

## if you use Get-CimInstance, there isn't a delete method
[array]$profilesDeleted = @( Get-WmiObject -ClassName Win32_UserProfile -Filter "Special = 'FALSE' and Loaded = 'FALSE'"| ForEach-Object `
{
    $totalUnloadedProfiles++
    $profile = $_
    [string[]]$usernameParts = ([System.Security.Principal.SecurityIdentifier]( $profile.sid )).Translate([System.Security.Principal.NTAccount]).Value -split '\\'
    if( $usernameParts -and $usernameParts.Count -eq 2 -and $usernameParts[0] -eq $domain )
    {
        Write-Verbose -Message "Checking group membership of $($usernameParts -join '\')"

        if( ( $searcher = [adsisearcher]"(samaccountname=$($usernameParts[1]))" ) `
            -and $searcher.FindOne().Properties.memberof -match "CN=$group,OU=" )
        {
            Write-Verbose -Message "Deleting profile in `"$($profile.LocalPath)`" for $($usernameParts -join '\\'), last used $(Get-Date -Date (([WMI] '').ConvertToDateTime($profile.LastUseTime)) -Format G), status $($profile.Status)"
            if( $PSCmdlet.ShouldProcess( "In `"$($profile.LocalPath)`" for $($usernameParts -join '\\')" , 'Delete Profile' ) )
            {
                $profile.psbase.Delete()
            }
            Add-Member -InputObject $profile -MemberType NoteProperty -Name Username -Value ($usernameParts -join '\\') -PassThru
        }
    }
})

if( ! $profilesDeleted -or ! $profilesDeleted.Count )
{
    Write-Warning -Message "No unloaded profiles found for group `"$group`" out of $totalUnloadedProfiles profiles checked"
}
else
{
    Write-Verbose -Message "Deleted $($profilesDeleted.Count) profile for $($profilesDeleted.username -join ' ')"
}
