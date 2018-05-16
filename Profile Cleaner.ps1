#required -version 3.0
<#
    Find all user profiles across machines selected from AD and show size, last access, last AD activity and allow removal

    Guy Leech, 2018

    Modification history:

    14/05/18  GL  Added missing help for -group and -ou arguments
#>

<#
.SYNOPSIS

Use WMI to enumerate all user profiles on Active Directory machine names matching a regular expression.
Those selected when OK is pressed will be deleted.

.DESCRIPTION

.PARAMETER name

Regular expression used to match machines retrieved from AD which will have their user profiles enumerated.
Can also be used to further restrict the machines returned by -ou or -group options

.PARAMETER group

An AD group whose members will be interrogated for user profiles.

.PARAMETER ou

An organisational unit, in either distinguished or canonical name format, whose members will be interrogated for user profiles.

.PARAMETER excludeUsers

A comma separated list of regular expressions to check user's owning profiles against and if they match any of them their profile will be ignored

.PARAMETER includeUsers

A comma separated list of regular expressions to check user's owning profiles against and only if they match one of these will the profile  be included

.PARAMETER excludeLocal

Exclude local accounts and their profiles. 

.PARAMETER csv

File name to write the data to in CSV format. If this is not specified then an on-screen grid view will display the results

.PARAMETER noDelete

Will not give OK/Cancel buttons in the grid view so profile deletions will not occur

.EXAMPLE

& '.\Profile Cleaner.ps1' -excludeLocal -machines ctx[pt]\d\d\d -excludeUsers [^A-Z]SVC-

Retrieve user profiles from machines named CTXPxxx and CTXTyyy (where xxx and yyy and numeric), don't include local accounts and exclude any which start with SVC- in the user name (e.g. service accounts)

& '.\Profile Cleaner.ps1' -OU CONTOSO.COM//Servers/Workplace/Citrix XenApp/Prod/Infrastructure Servers -includeUsers fredbloggs,johndoe

Retrieve user profiles from the specified OU, but only include users fredbloggs and johndoe (e.g. because they have left the company)

#>

[CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='High')]

Param
(
    [string]$name ,
    [string]$group ,
    [string]$OU ,
    [string[]]$excludeUsers  ,
    [string[]]$includeUsers ,
    [switch]$excludeLocal ,
    [string]$csv ,
    [switch]$noDelete ,
    [string[]]$columns = @( 'Machine Name','User Name','Profile Path','Profile Size (MB)','Loaded','Roaming','Last Used','Account Created','Last AD Logon','Account Enabled','Account Locked','Account Changed','Account Expires','Password Expired','Password Last Set','Last Bad Password' ) ,
    [string[]]$modules = @( 'ActiveDirectory' )
)

Function Calculate-FolderSize( [string]$machineName , [string]$folderName , [string]$sid , [string[]]$excludeUsers , [string[]]$includeUsers , [bool]$excludeLocal , [string]$domainName )
{
    ## can't do a Get-ChildItem -Recurse as can't seem to stop junction point traversal so do it manually
    Invoke-Command -ComputerName $machineName -ScriptBlock `
    { 
        [string]$username = if( $using:sid )
        {
            try
            {
                ([System.Security.Principal.SecurityIdentifier]($using:sid)).Translate([System.Security.Principal.NTAccount]).Value
            }
            catch
            {
                $null
            }
        }
        else
        {
            $null
        }
        if( $using:excludeLocal -and ! [string]::IsNullOrEmpty( $username ) -and ($username -split '\\')[0] -ne $using:domainName )
        {
            return $null,$null
        }
        ForEach( $includedUser in $using:includeUsers )
        {
            [bool]$found = $false
            if( $username -match $includedUser )
            {
                $found = $false
                break
            }
            if( ! $found )
            {
                return $null,$null
            }
        }
        ForEach( $excludedUser in $using:excludeUsers )
        {
            if( [string]::IsNullOrEmpty( $username ) -or $username -match $excludedUser )
            {
                return $null,$null
            }
        }
        $items = @( $using:folderName )
        [array]$files = While( $items )
        {
            $newitems = $items | Get-ChildItem -Force | Where-Object { ! ( $_.Attributes -band [System.IO.FileAttributes]::ReparsePoint ) }
            $newitems
            $items = $newitems | Where-Object { $_.Attributes -band [System.IO.FileAttributes]::Directory }
        }
        if( $files -and $files.Count )
        {
            [long]($files | Measure-Object -Property Length -Sum | Select -ExpandProperty Sum)
        }
        else
        {
            [long]0
        }
        $username
    }
}

ForEach( $module in $modules )
{
    Import-Module $module
}

[int]$ERROR_INVALID_PARAMETER = 87
[hashtable]$adUsers = @{}
[int]$count = 0
[hashtable]$profileObjects = @{}
[long]$totalSize = 0
$adDomain = Get-ADDomain

[hashtable]$searchProperties = @{}
[string]$command = $null

if( ! [string]::IsNullOrEmpty( $OU ) )
{
    ## see if canonical name (e.g. copied from AD Users & computers) and convert to distinguished
    if( $OU.IndexOf('/') -gt 0 )
    {
        try
        {
            ## http://www.itadmintools.com/2011/09/translate-active-directory-name-formats.html
            $NameTranslate = New-Object -ComObject NameTranslate 
            [System.__ComObject].InvokeMember( 'Init' , 'InvokeMethod', $null , $NameTranslate , @( 3 , $null ) ) ## ADS_NAME_INITTYPE_GC
            [System.__ComObject].InvokeMember( 'Set' , 'InvokeMethod', $null , $NameTranslate , (2 ,$OU)) ## CANONICALNAME
            $OU = [System.__ComObject].InvokeMember('Get' , 'InvokeMethod' , $null , $NameTranslate , 1) ## DISTINGUISHEDNAME
        }
        catch
        {
            Write-Error "Failed to translate OU `"$OU`" from canonical name to distinguished`n$_"
            exit $ERROR_INVALID_PARAMETER
        }
    }
    $searchProperties.Add( 'SearchBase' , $OU )
}
elseif( ! [string]::IsNullOrEmpty( $group ) )
{
    $command = 'Get-ADGroupMember -Recursive -Identity $group  | ? { $_.objectClass -eq ''Computer'' }'
}
elseif( [string]::IsNullOrEmpty( $name ) )
{
    Write-Error 'Must specify at least one of -group, -OU or -name to define what computers to interrogate'
    exit $ERROR_INVALID_PARAMETER
}

if( [string]::IsNullOrEmpty( $command ) )
{
    $command = 'Get-ADComputer -Filter * @searchProperties'
}

[array]$userProfiles = @( Invoke-Expression $command | Where-Object { $_.Name -match $name } | ForEach-Object `
{
    $count++
    [string]$machineName = $_.Name
    Write-Verbose "$count : $($_.Name)"

    $profiles = @( Get-WmiObject -Class win32_userprofile -ComputerName $machineName -ErrorAction SilentlyContinue )
    $profiles | ForEach-Object `
    {
        $profile = $_
        ## Get size of profile, last used - translate SID on remote machine in case a local account
        [long]$spaceUsed,[string]$username = Calculate-FolderSize -machineName $machineName -folderName $profile.LocalPath -sid $profile.sid -excludeUsers $excludeUsers -excludeLocal $excludeLocal -includeUsers $includeUsers -domain $adDomain.Name
        if( ! [string]::IsNullOrEmpty( $username ) ) ## username could be null if excluded by called function
        {
            [hashtable]$properties = [ordered]@{ 'Machine Name' = $machineName ; 'User Name' = $username ; 'Profile Path' = $profile.LocalPath ; 'Profile Size (MB)' = [math]::Round( $spaceUsed / 1MB ) -as [int] ;
                'Last Used' = [Management.ManagementDateTimeConverter]::ToDateTime( $profile.LastUseTime ) ; 'Roaming' = $profile.RoamingConfigured ; 'Loaded' = $profile.Loaded }
            $totalSize += $properties[ 'Profile Size (MB)' ]
            [string]$domainname,[string]$unqualifiedUserName = ( $username -split '\\' )
            ## if $unqualifiedUserNmame is null then not a domain\username so won't be in AD
            if( ! [string]::IsNullOrEmpty( $unqualifiedUserName ) -and $domainname -eq $adDomain.Name )
            {
                ## we stuff the profile object into a separate hash table so we can call its delete method later if required. If we put in the object's properties then Out-GridView would strip out since we can't display it
                $profileObjects.Add( ( $machineName + ':' + $username ) , $profile )
                $aduser = $adUsers[ $username ]
                if( ! $aduser )
                {
                    try
                    {
                        $aduser = Get-ADUser -Identity $unqualifiedUserName -Properties * -ErrorAction SilentlyContinue
                        $adUsers.Add( $username , $aduser )
                    }
                    catch {}
                }
                if( $aduser )
                {
                    $properties += @{ 'Account Enabled' = $aduser.Enabled ; 'Last AD Logon' = $aduser.LastLogonDate ; 'Account Created' = $aduser.Created ; 'Account Changed' = $aduser.Modified ;
                        'Password Expired' = $aduser.PasswordExpired ; 'Password Last Set' = $aduser.PasswordLastSet ; 'Account Expires' = $aduser.AccountExpirationDate ;
                        'Account Locked' = $aduser.LockedOut ; 'Last Bad Password' = $aduser.LastBadPasswordAttempt }
                }
            }
           
            [pscustomobject]$properties
            }
    }
})

if( $userProfiles -and $userProfiles.Count )
{
    if( [string]::IsNullOrEmpty( $csv ) )
    {
        [hashtable]$params = @{}
        if( ! $noDelete )
        {
            $params.Add( 'PassThru' , $true )
        }
        $selected = @( $userProfiles | Select $columns | Out-GridView -Title "$($userProfiles.Count) profiles found on $count machines using $totalSize MB" @params )

        if( $selected -and $selected.Count )
        {
            [long]$totalSizeDeleted = 0
            [int]$deleted = 0
            $selected | ForEach-Object `
            {
                $profile = $_
                $profileObject = $profileObjects[ ( $profile.'Machine Name' + ':' + $profile.'User Name' ) ]
                if( $profileObject )
                {
                    if( $PSCmdlet.ShouldProcess( "User $($profile.'User Name') from $($profile.'Machine Name')" , 'Delete Profile' ))
                    {
                        $profileObject.Delete()
                        if( $? )
                        {
                            $deleted++
                            $totalSizeDeleted += $profile.'Profile Size (MB)'
                        }
                    }
                }
                else
                {
                    Write-Warning "Failed to retrieve cached profile object for user $($profile.'User Name') on machine $($profile.'Machine Name')"
                }
            }
            Write-Output "Deleted $deleted profiles occupying $totalSizeDeleted MB"
        }
    }
    else
    {
        $userProfiles | Select $columns | Export-Csv -Path $csv -NoTypeInformation -NoClobber
    }
}
else
{
    Write-Warning "No user profiles found on the $count machines checked"
}
