#requires -version 3.0
<#
    Show all users on all nominated machines, grabbed from AD. Sessions selected in the resultant grid view will then be logged off if "OK" is pressed.

    Guy leech, 2018
#>

<#
.SYNOPSIS

Report the user sessions, as returned by quser, from the machines in Active Directory that match the regular expression specified via -name
These are then presented in a grid view and any selected when "OK" is pressed are then logged off

.DESCRIPTION

PowerShell's built-in conmfirmation mechanism is used before each logoff is performed to give granular control if required

.PARAMETER name

A regular expression used to match computers returned from Active Directory that will have their sessions enumerated

.PARAMETER user

An optional regular expression used to define which users are to be included. If not specified then all user sessions found are included

.PARAMETER current

Show only current sessions as returned by the quser command

.PARAMETER sinceBoot

Show sessions since the machine was last booted

.PARAMETER start

Show sessions starting from the time/date specified

.PARAMETER end

Show sessions starting up to the time/date specified

.PARAMETER csv

Name of csv file to write information to

.PARAMETER noProfile

Do not include user profile information

.PARAMETER jobTimeout

The number of seconds to wait for the quser command to complete before aborting the command

.EXAMPLE

& '.\Show users.ps1' -name '^CTX' -current

Show all users with current sessions on any machine present in Active Directory which starts with CTX, show them in a grid-view and then logoff those which have been selected when "OK" is pressed.

.EXAMPLE

& '.\Show users.ps1' -name '^CTX' -user '^adm' -sinceBoot

Show all users who have or have had sessions on any machine present in Active Directory which starts with CTX and where ther user name starts with ADM, show them in a grid-view and then logoff those which have been selected when "OK" is pressed.

.NOTES

Require the Active Directory PowerShell module to be available

#>

[CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='High')]

Param
(
    [Parameter(mandatory=$true,HelpMessage='Regular expression to match AD computers to interrogate')]
    [string]$name ,
    [string]$user ,
    [Parameter(ParameterSetName='Current')]
    [switch]$current ,
    [Parameter(ParameterSetName='Last')]
    [string]$last , 
    [Parameter(ParameterSetName='SinceBoot')]
    [switch]$sinceBoot ,
    [Parameter(ParameterSetName='Times')]
    [string]$start ,
    [Parameter(ParameterSetName='Times')]
    [string]$end ,
    [string]$csv ,
    [int]$jobTimeout = 15 ,
    [switch]$noprofile ,
    [string]$logName = 'Microsoft-Windows-User Profile Service/Operational' ,
    [string[]]$fieldNames = @( 'USERNAME','SESSIONNAME','ID','STATE','IDLE TIME','LOGON TIME' ) ,
    [string[]]$modules = @( 'ActiveDirectory' )
)

Function Calculate-FolderSize( [string]$machineName , [string]$folderName )
{
    ## can't do a Get-ChildItem -Recurse as can't seem to stop junction point traversal so do it manually
    Invoke-Command -ComputerName $machineName -ScriptBlock `
    { 
        $items = @( $using:folderName )
        [array]$files = While( $items )
        {
            $newitems = $items | Get-ChildItem -Force | Where-Object { ! ( $_.Attributes -band [System.IO.FileAttributes]::ReparsePoint ) }
            $newitems
            $items = $newitems | Where-Object { $_.Attributes -band [System.IO.FileAttributes]::Directory }
        }
        $files | Measure-Object -Property Length -Sum | Select -ExpandProperty Sum 
    }
}

ForEach( $module in $modules )
{
    Import-Module $module
}

[int]$count = 0
[int]$machinesWithUsers = 0
[datetime]$startDate = Get-Date
[datetime]$endDate = Get-Date

if( ! [string]::IsNullOrEmpty( $last ) )
{
    ## see what last character is as will tell us what units to work with
    [int]$multiplier = 0
    switch( $last[-1] )
    {
        "s" { $multiplier = 1 }
        "m" { $multiplier = 60 }
        "h" { $multiplier = 3600 }
        "d" { $multiplier = 86400 }
        "w" { $multiplier = 86400 * 7 }
        "y" { $multiplier = 86400 * 365 }
        default { Write-Error "Unknown multiplier `"$($last[-1])`"" ; return }
    }
    $endDate = Get-Date
    if( $last.Length -le 1 )
    {
        $startDate = $endDate.AddSeconds( -$multiplier )
    }
    else
    {
        $startDate = $endDate.AddSeconds( - ( ( $last.Substring( 0 ,$last.Length - 1 ) -as [int] ) * $multiplier ) )
    }
}
elseif( ! [string]::IsNullOrEmpty( $start ) )
{
    $startDate = [datetime]::Parse( $start )
    if( [string]::IsNullOrEmpty( $end ) )
    {
        $endDate = Get-Date
    }
    else
    {
        $endDate = [datetime]::Parse( $end )
    }
}

[array]$sessions = @( Get-ADComputer -Filter * | Where-Object { $_.Name -match $name } | ForEach-Object `
{
    $count++
    [string]$machineName = $_.Name
    Write-Verbose "$count : $($_.Name)"
    [datetime]$bootTime = (Get-Date).AddYears( 10 ) ## since we can't use null
    ## use WMI instead of CIM in case remote machine is older
    try
    {
        [string]$rawBootTime = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $machineName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty LastBootUpTime 
        if( ! [string]::IsNullOrEmpty( $rawBootTime ) )
        {
            $bootTime = [Management.ManagementDateTimeConverter]::ToDateTime( $rawBootTime )
        }
        else
        {
            Write-Warning "Failed to get boot time from $machineName : $($error[0].Exception.Message)"
        }

    }
    catch
    {
        Write-Warning "Failed to get boot time from $machineName  : $($_.Exception.Message)"
    }
    
    ## if couldn't get boot time then assume can't get events or run quser to save time
    if( $bootTime -and $bootTime -lt (Get-Date) )
    {
        if( $current )
        {
             ## Get users from machine - if we just run quser then get error for no users so this method make it squeaky clean/silent
            $pinfo = New-Object System.Diagnostics.ProcessStartInfo
            $pinfo.FileName = "quser.exe"
            $pinfo.Arguments = "/server:$machineName"
            $pinfo.RedirectStandardError = $true
            $pinfo.RedirectStandardOutput = $true
            $pinfo.UseShellExecute = $false
            $pinfo.WindowStyle = 'Hidden'
            $pinfo.CreateNoWindow = $true
            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $pinfo
            if( $process.Start() )
            {
                if( $process.WaitForExit( $jobTimeout * 1000 ) )
                {
                    ## Output of quser is fixed width but can't do simple parse as SESSIONNAME is empty when session is disconnected so we break it up based on header positions
                    [string[]]$allOutput = $process.StandardOutput.ReadToEnd() -split "`n"
                    [string]$header = $allOutput[0]
                    [bool]$counted = $false
                    [array]$profiles = $null
                    $allOutput | Select -Skip 1 | ForEach-Object `
                    {
                        [string]$line = $_
                        if( ! [string]::IsNullOrEmpty( $line ) )
                        {
                            if( ! $counted )
                            {
                                $machinesWithUsers++
                                $counted = $true
                                if( ! $noprofile )
                                {
                                    $profiles = @( Get-WmiObject -Class win32_userprofile -ComputerName $machineName -ErrorAction SilentlyContinue )
                                }
                            }
                            [hashtable]$properties = [ordered]@{ 'Machine' = $machineName ; 'Boot Time' = $( if( $bootTime -lt (Get-Date) ) { $bootTime } else { $null } ) }

                            For( [int]$index = 0 ; $index -lt $fieldNames.Count ; $index++ )
                            {
                                [int]$startColumn = $header.IndexOf($fieldNames[$index])
                                ## if last column then can't look at start of next field so use overall line length
                                [int]$endColumn = if( $index -eq $fieldNames.Count - 1 ) { $line.Length } else { $header.IndexOf( $fieldNames[ $index + 1 ] ) }
                                try
                                {
                                    $properties.Add( $fieldNames[ $index ] , ( $line.Substring( $startColumn , $endColumn - $startColumn ).Trim() ) )
                                }
                                catch
                                {
                                    throw $_
                                }
                            }
                            if( [string]::IsNullOrEmpty( $user ) -or $properties[ $fieldNames[ 0 ] ] -match $user )
                            {
                                if( ! $noprofile -and $profiles -and $profiles.Count )
                                {
                                    $sid = (New-Object System.Security.Principal.NTAccount($properties[$fieldNames[0]])).Translate([System.Security.Principal.SecurityIdentifier]).value
                                    $profile = $profiles | Where-Object { $_.sid -eq $sid } 
                                    if( $profile )
                                    {
                                        [long]$spaceUsed = Calculate-FolderSize -machineName $machineName -folderName $profile.LocalPath
                                        $properties += @{ 'Profile Path' = $profile.LocalPath ; 'Profile Size (MB)' = [math]::Round( $spaceUsed / 1MB ) -as [int] }
                                    }
                                    else
                                    {
                                        Write-Warning "Unable to get profile information for $($properties[$fieldNames[0]]) on $machineName"
                                    }
                                }
                                [pscustomobject]$properties
                            }
                            else
                            {
                                Write-Verbose "Ignoring $($properties[ $fieldNames[ 0 ] ]) on $machineName"
                            }
                            $properties = $null
                        }
                    }
                }
                else
                {
                    Write-Warning ( "Timeout of {0} seconds waiting for process to exit {1} {2}" -f $jobTimeout , $pinfo.FileName , $pinfo.Arguments )
                    $process.Kill()
                }
            }
            else
            {
                Write-Warning ( "Failed to start process {0} {1}" -f $pinfo.FileName , $pinfo.Arguments )
            }
        }
        else ## looking back historically
        {
            if( $sinceBoot )
            {
                $startDate = $bootTime
                $endDate = Get-Date
            }
            [array]$events = @( Get-WinEvent -ComputerName $machineName -FilterHashtable @{ LogName = $logName ; id = 1,4 ; StartTime=$startDate ; EndTime=$endDate } -ErrorAction SilentlyContinue -Oldest )
            if( $events -and $events.Count )
            {
                $machinesWithUsers++
                ## now we need to find logon events and corresponding logoff event for the same user - use index so can start search for logoff from current position
                For( [int]$index = 0 ; $index -lt $events.Count ; $index++ )
                {
                    if( $events[ $index ].Id -eq 1 )
                    {
                        $logonEvent = $events[ $index ]
                        [string]$userName = ([Security.Principal.SecurityIdentifier]($logonEvent.UserId)).Translate([Security.Principal.NTAccount]).Value
                        if( [string]::IsNullOrEmpty( $user ) -or $userName -match $user )
                        {                
                            [int]$sessionId = -1
                            if( $logonEvent.Message -match '(\d+)\.$' ) ## Recieved user logon notification on session 2.
                            {
                                $sessionId = $Matches[ 1 ]
                            }
                            if( $sessionId -le 0 )
                            {
                                Write-Warning "Failed to parse valid session id from text `"$($logonEvent.Message)`""
                            }
                            $logoffEvent = $null
                            [hashtable]$properties = [ordered]@{ 'Machine' = $machineName ; 'Boot Time' = $bootTime ; 'UserName' = $userName ; 'Id' = $sessionId ; 'Logon Time' = $logonEvent.TimeCreated }
                            if( $sessionId -gt 0 )
                            {
                                For( [int]$search = $index + 1 ; $search -lt $events.Count ; $search++ )
                                {
                                    if( $events[ $search ].Id -eq 4 -and $events[ $search ].UserId -eq $logonEvent.UserId )
                                    {
                                        if( $events[ $search ].Message -match '(\d+)\.$' ) ## Finished processing user logoff notification on session 2. 
                                        {
                                            [int]$loggedOffSessionId = $Matches[ 1 ]
                                            if( $loggedOffSessionId -eq $sessionId )
                                            {
                                                $properties.Add(  'Logoff Time' , $events[ $search ].TimeCreated )
                                                break
                                            }
                                        }
                                    }
                                }
                            }
                            [pscustomobject]$properties
                            $properties = $null
                        }
                    }
                }
            }
            else
            {
                Write-Warning "Got no logon or logoff events from $logName log on $machineName"
            }
        }
    }
 })

if( $sessions -and $sessions.Count )
{
    if( [string]::IsNullOrEmpty( $csv ) )
    {
        $selected = @( $sessions | Out-GridView -Title "$($sessions.Count) sessions found on $machinesWithUsers machines of $count checked" -PassThru )

        if( $selected -and $selected.Count )
        {
            $selected | ForEach-Object `
            {
                $session = $_
                if((  Get-Member -InputObject $session -Name 'Logoff Time' -ErrorAction SilentlyContinue ) -and ! [string]::IsNullOrEmpty( $session.'Logoff Time' ) )
                {
                    Write-Warning "Can't logoff user $($session.Username) from $($session.Machine) as already logged off at $($session.'Logoff Time')"
                }
                elseif( $PSCmdlet.ShouldProcess( "User $($session.Username) from $($session.Machine)" , 'Logoff' ))
                {
                    Start-Process -FilePath logoff.exe -ArgumentList "$($session.id) /server:$($session.Machine)" -Wait
                }
            }
        }
    }
    else
    {
        $sessions | Export-Csv -Path $csv -NoTypeInformation -NoClobber
    }
}
else
{
    Write-Warning "No sessions found on $count machines checked"
}
