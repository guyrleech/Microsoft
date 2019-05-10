#requires -version 3
<#
    Show process durations via security event logs when process creation/termination auditing is enabled

    @guyrleech 2019

    Modification History:

    10/05/2019  GRL   Added subject logon id to grid view output
#>

<#
.SYNOPSIS

Retrieve process start events from the security event log, try and find corresponding process exit and optionally also show start time relative to that user's logon and/or computer boot

.PARAMETER usernames

Only include processes for users which match this regular expression

.PARAMETER processName

Only include processes which match this regular expression

.PARAMETER start

Only retrieve processes started after this date/time

.PARAMETER end

Only retrieve processes started before this date/time

.PARAMETER last

Show processes started in the preceding period where 's' is seconds, 'm' is minutes, 'h' is hours, 'd' is days, 'w' is weeks and 'y' is years so 12h will retrieve all in the last 12 hours.

.PARAMETER logon

Include the logon time and the time since logon for the process creation for the logon session this process belongs to

.PARAMETER boot

Include the boot time and the time since boot for the process creation 

.PARAMETER outputFile

Write the results to the specified csv file

.PARAMETER noGridview

Output the results to the pipeline rather than a grid view

.PARAMETER excludeSystem

Do not include processes run by the system account

.EXAMPLE

& '.\Get Process Durations.ps1' -last 2d -logon -username billybob -boot

Find all process creations and corresponding terminations for the user billybob in the last 2 days, calculate the start time relative to logon for that user's session and relative to the the boot time and display in a grid view

.NOTES

Must have process creation and process termination auditing enabled.

If process command line auditing is enabled then the command line will be included. See https://docs.microsoft.com/en-us/windows-server/identity/ad-ds/manage/component-updates/command-line-process-auditing
#>

[CmdletBinding()]

Param
(
    [string]$username ,
    [string]$processName ,
    [string]$start ,
    [string]$end ,
    [string]$last ,
    [switch]$logon ,
    [switch]$boot ,
    [string]$outputFile ,
    [switch]$nogridview ,
    [switch]$excludeSystem 
)

[string[]]$startPropertiesMap = @(
    'SubjectUserSid' , 
    'SubjectUserName' ,
    'SubjectDomainName' ,
    'SubjectLogonId' ,
    'NewProcessId' ,
    'NewProcessName' ,
    'TokenElevationType' ,
    'ProcessId' ,
    'CommandLine' ,
    'TargetUserSid' ,
    'TargetUserName' ,
    'TargetDomainName' ,
    'TargetLogonId' ,
    'ParentProcessName'
)

Set-Variable -Name 'endSubjectUserSid' -Value 0 -Option ReadOnly -ErrorAction SilentlyContinue
Set-Variable -Name 'endSubjectUserName' -Value 1 -Option ReadOnly -ErrorAction SilentlyContinue
Set-Variable -Name 'endSubjectDomainName' -Value 2 -Option ReadOnly -ErrorAction SilentlyContinue
Set-Variable -Name 'endSubjectLogonId' -Value 3 -Option ReadOnly -ErrorAction SilentlyContinue
Set-Variable -Name 'endStatus' -Value 4 -Option ReadOnly -ErrorAction SilentlyContinue
Set-Variable -Name 'endProcessId' -Value 5 -Option ReadOnly -ErrorAction SilentlyContinue
Set-Variable -Name 'endProcessName' -Value 6 -Option ReadOnly -ErrorAction SilentlyContinue

[hashtable]$auditingGuids = @{
    'Process Creation'    = '{0CCE922C-69AE-11D9-BED3-505054503030}'
    'Process Termination' = '{0CCE922C-69AE-11D9-BED3-505054503030}' }

Function Get-AuditSetting
{
    [CmdletBinding()]
    Param
    (
        [string]$GUID
    )
    [string[]]$fields = ( auditpol.exe /get /subcategory:"$GUID" /r | Select-Object -Skip 1 ) -split ',' ## Don't use ConvertFrom-CSV as makes it harder to get the column we want
    if( $fields -and $fields.Count -ge 6 )
    {
        ## Machine Name,Policy Target,Subcategory,Subcategory GUID,Inclusion Setting,Exclusion Setting
        ## DESKTOP2,System,Process Termination,{0CCE922C-69AE-11D9-BED3-505054503030},No Auditing,
        $fields[5] ## get a blank field at the start
    }
    else
    {
        Write-Warning "Unable to determine audit setting"
    }
}

[string]$machineAccount = $env:COMPUTERNAME + '$'
[hashtable]$startEventFilter = @{
    'Logname' = 'Security'
    'Id' = 4688
}

if( $PSBoundParameters[ 'last' ] -and ( $PSBoundParameters[ 'start' ] -or $PSBoundParameters[ 'end' ] ) )
{
    Throw "Cannot use -last when -start or -end are also specified"
}

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
        default { Throw "Unknown multiplier `"$($last[-1])`"" }
    }
    $endDate = Get-Date
    if( $last.Length -le 1 )
    {
        $startDate = $endDate.AddSeconds( -$multiplier )
    }
    else
    {
        $startDate = $endDate.AddSeconds( - ( ( $last.Substring( 0 ,$last.Length - 1 ) -as [decimal] ) * $multiplier ) )
    }
    $startEventFilter.Add( 'StartTime' , $startDate )
}

if( $PSBoundParameters[ 'start' ] )
{
    $startEventFilter.Add( 'StartTime' , (Get-Date -Date $start ))
}

if( $PSBoundParameters[ 'end' ] )
{
    $startEventFilter.Add( 'EndTime' , (Get-Date -Date $end ))
}

[hashtable]$systemAccounts = @{}
Get-CimInstance -ClassName win32_SystemAccount | ForEach-Object `
{
    $systemAccounts.Add( $_.SID , $_.Name )
}

$bootTime = $null

if( $boot )
{
    ## get the boot time so we can add a column relative to it
    $bootTime = Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object -ExpandProperty LastBootupTime
}

[hashtable]$logons = @{}
## get logons so we can cross reference to the id of the logon
Get-WmiObject win32_logonsession -Filter "LogonType='10' or LogonType='12' or LogonType='2' or LogonType='11'" | ForEach-Object `
{
    $session = $_
    [array]$users = @( Get-WmiObject win32_loggedonuser -filter "Dependent = '\\\\.\\root\\cimv2:Win32_LogonSession.LogonId=`"$($session.LogonId)`"'" | ForEach-Object `
    {
        if( $_.Antecedent -match 'Domain="(.*)",Name="(.*)"$' ` )
        {
            [pscustomobject]@{ 'LogonTime' = (([WMI] '').ConvertToDateTime( $session.StartTime )) ; 'Domain' = $Matches[1] ; 'UserName' = $Matches[2]}
        }
        else
        {
            Write-Warning "Unexpected antecedent format `"$($_.Antecedent)`""
        }
    })
    if( $users -and $users.Count )
    {
        $logons.Add( $session.LogonId , $users )
    }
}

[hashtable]$stopEventFilter = $startEventFilter.Clone()
$stopEventFilter[ 'Id' ] = 4689
$eventError = $null
$error.Clear()

[array]$endEvents = @( Get-WinEvent -FilterHashtable $stopEventFilter -Oldest -ErrorAction SilentlyContinue )

Write-Verbose "Got $($endEvents.Count) process end events"

## Find all process starts then we'll look for the corresponding stops
[array]$processes = @( Get-WinEvent -FilterHashtable $startEventFilter -Oldest -ErrorAction SilentlyContinue -ErrorVariable 'eventError'  | ForEach-Object `
{
    $event = $_
    if( ( ! $username -or $event.Properties[ 1 ].Value -match $username ) -and ( ! $excludeSystem -or ( $event.Properties[ 1 ].Value -ne $machineAccount -and $event.Properties[ 1 ].Value -ne '-' )) -and ( ! $processName -or $event.Properties[5].value -match $processName ) )
    {
        [hashtable]$started = @{ 'Start' = $event.TimeCreated }
        For( [int]$index = 0 ; $index -lt [math]::Min( $startPropertiesMap.Count , $event.Properties.Count ) ; $index++ )
        {
            $started.Add( $startPropertiesMap[ $index ] , $event.Properties[ $index ].value )
        }
        if( $started[ 'SubjectUserName' ] -eq '-' )
        {
            $started.Set_Item( 'SubjectUserName' , $systemAccounts[ ($event.Properties[ 0 ].Value | Select-Object -ExpandProperty Value) ] )
            $started.Set_Item( 'SubjectDomainName' , $env:COMPUTERNAME )
        }
        ## now find corresponding termination event - don't use hashtable since could have duplicate pids
        $terminate = $endEvents | Where-Object { $_.Id -eq 4689 -and $_.TimeCreated -ge $event.TimeCreated -and $_.Properties[ $endProcessId ].Value -eq $started.NewProcessId  } | Select-Object -First 1

        if( ! $terminate ) ## probably still running
        {
            $existing = Get-Process -Id $started.NewProcessId -ErrorAction SilentlyContinue
            if( ! $existing )
            {
                Write-Warning "Cannot find process terminated event for pid $($started.NewProcessId) and not currently running"
            }
        }
        
        $started.Add( 'Exit Code' , $(if( $terminate ) { $terminate.Properties[ $endStatus ].value }))
        $started.Add( 'End' , $(if( $terminate ) { $terminate.TimeCreated }))
        $started.Add( 'Duration' , $(if( $terminate ) { (New-TimeSpan -Start $event.TimeCreated -End $terminate.TimeCreated | Select-Object -ExpandProperty TotalMilliSeconds) / 1000 }))

        if( $logon )
        {
            ## get the logon time so we can add a column relative to it
            $logonTime = $null
            $thisLogon = $logons[ $started.SubjectLogonId.ToString() ]

            if( $thisLogon )
            {
                $theLogon = $null
                if( $thisLogon -is [array] )
                {
                    ## Need to find this user
                    ForEach( $alogon in $thisLogon )
                    {
                        if( $started.SubjectDomainName -eq $alogon.Domain -and $started.SubjectUserName -eq $alogon.Username )
                        {
                            if( $theLogon )
                            {
                                Write-Warning "Multiple logons for same user $($started.SubjectDomainName)\$($started.SubjectUserName)"
                            }
                            $theLogon = $alogon
                        }
                    }
                }
                elseif( $started.SubjectDomainName -eq $thisLogon.Domain -and $started.SubjectUserName -eq $thisLogon.Username )
                {
                    $theLogon -eq $thisLogon
                }
                if( ! $theLogon )
                {
                    Write-Warning "Couldn't find logon for user $($started.SubjectDomainName)\$($started.SubjectUserName) for process $($started.NewProcessId) started @ $(Get-Date -Date $started.Start -Format G)"
                }
                $started.Add( 'Logon Time' , $(if( $theLogon ) { $theLogon.LogonTime } ) )
                $started.Add( 'After Logon (s)' , $(if( $theLogon ) { New-TimeSpan -Start $theLogon.LogonTime -End $started.Start | Select-Object -ExpandProperty TotalSeconds } ) )
            }
        }
        if( $bootTime )
        {
            $started.Add( 'After Boot (s)' , $(if( $theLogon ) { New-TimeSpan -Start $bootTime -End $started.Start | Select-Object -ExpandProperty TotalSeconds } ) )
        }
        [pscustomobject]$started
    }
})

if( $eventError )
{
    Write-Warning "Failed to get any events with ids $($startEventFilter[ 'Id' ] -join ',') from $($startEventFilter[ 'Logname' ]) event log"

    ## Look to see if there's an error otherwise check if required auditing is enabled
    if( $eventError -and $eventError.Count )
    {
        if( $eventError[0].Exception.Message -match 'No events were found that match the specified selection criteria' )
        {
            ForEach( $auditGuid in $auditingGuids.GetEnumerator() )
            {
                [string]$result = Get-AuditSetting -GUID $auditGuid.Value
                if( $result -notmatch 'Success' )
                {
                    Write-Warning "Auditing for $($auditGuid.Name) is `"$result`" so will not generate required events"
                }
            }
        }
        Throw $eventError[0]
    }
}

[string]$status = $null
if( $PSBoundParameters[ 'username' ] )
{
    $status += " matching user `"$username`""
}
if( $startEventFilter[ 'StartTime' ] )
{
    $status += " from $(Get-Date -Date $startEventFilter[ 'StartTime' ] -Format G)"
}
if( $startEventFilter[ 'EndTime' ] )
{
    $status += " to $(Get-Date -Date $startEventFilter[ 'EndTime' ] -Format G)"
}

if( ! $processes.Count )
{
    Write-Warning "No process start events found $status"
}
else
{
    $headings = [System.Collections.ArrayList]@( @{n='User Name';e={'{0}\{1}' -f $_.SubjectDomainName , $_.SubjectUserName}} , @{n='Process';e={$_.NewProcessName}} , @{n='PID';e={$_.NewProcessId}} , `
        'CommandLine' , @{n='Parent Process';e={$_.ParentProcessName}} , 'SubjectLogonId' , @{n='Parent PID';e={$_.ProcessId}} , @{n='Start';e={('{0}.{1}' -f (Get-Date -Date $_.start -Format G) , $_.start.Millisecond)}} , `
        @{n='End';e={('{0}.{1}' -f (Get-Date -Date $_.end -Format G) , $_.end.Millisecond)}} , 'Duration' , @{n='Exit Code';e={if( $_.'Exit Code' -ne $null ) { '0x{0:x}' -f $_.'Exit Code'}}} )
    if( $logon )
    {
        $headings += @( @{n='Logon Time';e={('{0}.{1}' -f (Get-Date -Date $_.'Logon Time' -Format G) , $_.'Logon Time'.Millisecond)}} , 'After Logon (s)' )
    }
    if( $bootTime )
    {
        $headings += @( @{n='Boot Time';e={Get-Date -Date $bootTime -Format G}} , 'After Boot (s)' )
    }
    if( $PSBoundParameters[ 'outputfile' ] )
    {
        $processes | Export-Csv -Path $outputFile -NoTypeInformation -NoClobber
    }
    if( $nogridview )
    {
        $processes
    }
    else
    {
        [array]$selected = @( $processes | Select-Object -Property $headings| Out-GridView -PassThru -Title "$($processes.Count) process starts $status" )
        if( $selected -and $selected.Count )
        {
            $selected | Set-Clipboard
        }
    }
}
