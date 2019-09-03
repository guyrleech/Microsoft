<#
    Take action based on active network profile (Domain, Private, Public)

    @guyrleech 2019

    Modification History:
#>

<#
.SYNOPSIS

Check network connection profiles and if any are connected on a public network, or nothing is connected so offline, then set a registry value differently compared with private/domain network

.DESCRIPTION

The default values for the registry key and value will disable display of the user name on the lock screen when on public/no networks and enable when not

.PARAMETER key

The registry key containing the value which will be changed depending on the active network profiles

.PARAMETER valueName

The name of the registry value to be set depending on the active network profiles

.PARAMETER awayNetworks

The regular expression to match to denote a public network. If the network profile category does not match this then it is deemed a home network

.PARAMETER homeValue

The value to set in the registry value when the network profiles are connected at home (privately)

.PARAMETER awayValue

The value to set in the registry value when the network profiles are connected away from home (publicly)

.PARAMETER pollPeriod

The time in seconds to wait between checks. By default the script will only run once

.PARAMETER noAdminCheck

Do not check that the account running the script is an administrator

.PARAMETER logFile

The path to a logFile to append the results to

.PARAMETER noTraffic

The string to match when no traffic is flowing over the given interface (do not change)

.EXAMPLE

& '.\Network Profile Actioner.ps1' -Verbose -pollPeriod 60 -logFile c:\temp\iper.log

Check the network connection profiles every 60 seconds and make the required changes to the registry value, writing to a log file c:\temp\iper.log

.EXAMPLE

& '.\Network Profile Actioner.ps1'

Check the network connection profiles once and make the required changes to the registry value and then exit

#>

[CmdletBinding()]

Param
(
    [string]$key = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' ,
    [string]$valueName = 'DontDisplayLockedUserId' ,
    [string]$awayNetworks = 'Public' ,
    [int]$homeValue = 1 ,
    [int]$awayValue = 3 ,
    [int]$pollPeriod = 0 ,
    [switch]$noAdminCheck ,
    [string]$logFile ,
    [string]$noTraffic = 'NoTraffic'
)

if( $PSBoundParameters[ 'logFile' ] )
{
    Start-Transcript -Path $logFile -Append
}

Try
{
    if( $key -match '^HKLM:' -and ! $noAdminCheck )
    {
        $private:windowsPrincipal = New-Object System.Security.Principal.WindowsPrincipal([System.Security.Principal.WindowsIdentity]::GetCurrent())
        if( ! ( $windowsPrincipal.IsInRole( [System.Security.Principal.WindowsBuiltInRole]::Administrator )))
        {
            Throw "This script must be run with administrative privilege which $env:USERDOMAIN\$env:USERNAME does not have"
        }
    }

    [bool]$private:previousValue = $false
    [bool]$private:firstTime = $true

    Do
    {
        [bool]$private:atHome = $true
        [int]$private:awayNetworksChecked = 0

        ## Get active network profiles
        [array]$private:profiles = @( Get-NetConnectionProfile )
        
        $profiles | Where-Object { $_.NetworkCategory -match $awayNetworks } | ForEach-Object `
        {
            $awayNetworksChecked++
            if( $_.IPv4Connectivity -ne $noTraffic -or $_.IPv6Connectivity -ne $noTraffic )
            {
                $atHome = $false
                Write-Verbose -Message "$(Get-Date -Format G) : Interface `"$($_.InterfaceAlias)`" is connected on `"$($_.NetworkCategory)`" so setting as away"
            }
        }

        ## Check we have got some profiles which aren't public 
        if( $atHome )
        {
            if ( ! $profiles -or $profiles.Count -eq $awayNetworksChecked )
            {
                $atHome = $false
                Write-Verbose -Message "$(Get-Date -Format G) : $($profiles.Count) active network profiles found & all match `"$awayNetworks`" so setting as away"
            }
            elseif( $atHome )
            {
                Write-Verbose -Message "$(Get-Date -Format G) : No active network profiles found matching `"$awayNetworks`" so setting as home"
            }
        }
        ## else we have already set atHome as false because of connected network profiles

        ## set if first time in or profile has changed
        if( $firstTime -or $atHome -ne $previousValue )
        {
            if( ! ( Test-Path -Path $key -ErrorAction SilentlyContinue ) )
            {
                [void]( New-Item -Path $key -ItemType Key )
            }
            [int]$private:newValue = $(if( $atHome ) { $homeValue } else { $awayValue } )
            Write-Verbose -message "$(Get-Date -Format G) : Setting `"$valueName`" in `"$key`" to $newValue as at home is $atHome"
            Set-ItemProperty -Path $key -Name $valueName -Force -Value $newValue
            $firstTime = $false
            $previousValue = $atHome
        }

        if( $PSBoundParameters[ 'pollPeriod' ] )
        {
            Start-Sleep -Seconds $pollPeriod
        }
    } While( $pollPeriod -gt 0 )
}
Catch
{
    Throw $_
}
Finally
{
    if( $PSBoundParameters[ 'logFile' ] )
    {
        Stop-Transcript
    }
}
