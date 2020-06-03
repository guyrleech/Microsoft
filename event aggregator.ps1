<#
    Get all events from all event logs in a given time window and output to grid view or csv

    @guyrleech 2019

    Modification History

    24/07/19   GRL   Added -last parameter, defaulted -end if not specified and verification of times/dates
    02/09/19   GRL   Added -passThru, -duration, -overwrite , -nogridview and pass thru for grid view
    17/09/19   GRL   Extra parameter validation
    12/12/19   GRL   Added -ids , -excludeProvider and -excludeIds parameters
    26/02/20   GRL   Changed -computer to take array
    27/02/20   GRL   Added -message and -ignore parameters
    03/06/20   GRL   Added -boot option
#>


<#
.SYNOPSIS

Retrieve all events from all 300+ event logs in a given time/date range and show in a sortable/filterable gridview or export to csv

.PARAMETER start

The start time/date to show events from. If no date is given then the current day is used.

.PARAMETER end

The end time/date to show events from. If no date is given then the current day is used.

.PARAMETER last

Show events logged in the preceding period where 's' is seconds, 'm' is minutes, 'h' is hours, 'd' is days, 'w' is weeks and 'y' is years so 12h will retrieve all in the last 12 hours.

.PARAMETER duration

Show events logged from the start specified via -start for the specified period where 's' is seconds, 'm' is minutes, 'h' is hours, 'd' is days, 'w' is weeks and 'y' is years so 2m will retrieve events for 2 minutes from the given start time

.PARAMETER message

Only include events where the message matches the regular expression specified

.PARAMETER ignore

Exclude events where the message matches the regular expression specified

.PARAMETER ids

only include events which have ids in this comma separated list of event ids

.PARAMETER excludeids

Exclude events which have ids in this comma separated list of event ids

.PARAMETER excludeProvider

Exclude events from any provider which matches this regular expression

.PARAMETER csv

The name of a non-existent csv file to write the results to. If not specified then the results will go to an on screne grid view

.PARAMETER badOnly

Only show critical, warning and error events. If not specified then all events will be shown.

.PARAMETER noGridView

Write the events found as objects to the pipeline

.PARAMETER overWrite

Overwrite the csv file if it exists already

.PARAMETER passThru

Selected items in the grid view when OK is clicked will be placed on the pipeline so can be put on clipboard for example

.PARAMETER eventLogs

A pattern matching the event logs to search. By default all event logs are searched.

.PARAMETER computer

One or more remote computers to query. If not specified then the local computer is used.

.EXAMPLE

& '.\event aggregator.ps1' -start 10:38 -end 10:45 -badOnly

Show all critical, warning and error events that occurred between 10:38 and 10:45 today in an on screen gridview

.EXAMPLE

& '.\event aggregator.ps1' -start "10:38 29/06/19" -end "10:45 29/06/19" -csv c:\badevents.csv

Export all events that occurred between 10:38 and 10:45 on the 29th June 2019 top the named csv file

.EXAMPLE

& '.\event aggregator.ps1' -last 5m -Computers fred,bloggs

Show allevents that occurred in the last 5 minutes on computers "fred" and "bloggs" in an on screen gridview

.EXAMPLE

& '.\event aggregator.ps1' -start 10:40 -duration 2m -eventLogs *shell-core/operational*

Show events from the specified event log that occurred between 10:40 and 10:42 today in an on screen gridview

.EXAMPLE

& '.\event aggregator.ps1' -start 10:40 -duration 2m -excludeids 405,3209 -excludeProvider 'kernel|security'

Show events from the specified event log that occurred between 10:40 and 10:42 today in an on screen gridview but exclude events with ids 405 or 3209 and exclude events which match the regular expression so where the provider matches 'kernel' or 'security'

#>

[CmdletBinding()]

Param
(
    [Parameter(ParameterSetName='BootTime' ,Mandatory,HelpMessage='Start from boot time')]
    [switch]$boot ,
    [Parameter(ParameterSetName='StartTime',Mandatory,HelpMessage='Start time/date for event search')]
	[string]$start ,
    [Parameter(ParameterSetName='StartTime')]
    [Parameter(ParameterSetName='BootTime')]
	[string]$end ,
    [Parameter(ParameterSetName='StartTime')]
    [Parameter(ParameterSetName='BootTime')]
	[string]$duration ,
    [Parameter(ParameterSetName='Last',Mandatory,HelpMessage='Search for events in last seconds/minutes/hours/days/weeks/years')]
    [string]$last ,
    [int[]]$ids ,
    [int[]]$excludeIds ,
    [string]$excludeProvider ,
    [string]$message ,
    [string]$ignore ,
    [string]$csv ,
    [switch]$badOnly ,
    [string]$eventLogs = '*' ,
    [switch]$noGridView ,
    [switch]$overWrite ,
    [switch]$passThru ,
    [string[]]$computer = @( 'localhost' )
)

Function Out-PassThru
{
    Process
    {
        $_
    }
}

[hashtable]$arguments = @{}
[string]$command = $null

if( $PSBoundParameters[ 'last' ] )
{
    if( $PSBoundParameters[ 'start' ] -or $PSBoundParameters[ 'end' ] -or $PSBoundParameters[ 'boot' ] )
    {
        Throw 'Cannot use -start, -boot or -end with -last'
    }

    ## see what last character is as will tell us what units to work with
    [int]$multiplier = 0
    switch( $last[-1] )
    {
        's' { $multiplier = 1 }
        'm' { $multiplier = 60 }
        'h' { $multiplier = 3600 }
        'd' { $multiplier = 86400 }
        'w' { $multiplier = 86400 * 7 }
        'y' { $multiplier = 86400 * 365 }
        default { Throw "Unknown multiplier `"$($last[-1])`"" }
    }
    Remove-Variable -Name 'End'
    [datetime]$script:end = Get-Date
    if( $last.Length -le 1 )
    {
        $secondsAgo = $multiplier
    }
    else
    {
        $secondsAgo = ( ( $last.Substring( 0 , $last.Length - 1 ) -as [decimal] ) * $multiplier )
    }
    
    Remove-Variable -Name 'Start'
    [datetime]$script:start = $end.AddSeconds( -$secondsAgo )

}
else
{
    $parsed = $null
    if( $PSBoundParameters[ 'boot' ] )
    {
        if( $lastbooted = Get-WmiObject -Class Win32_operatingsystem | Select-Object -ExpandProperty LastBootUpTime )
        {
            $parsed = ([WMI] '').ConvertToDateTime( $lastbooted )
        }
        else
        {
            Throw 'Failed to get last boot time via WMI'
        }
    }
    else
    {
        ## Check time formats as bad ones get stripped by query so search whole event log
        $parsed = New-Object -TypeName 'DateTime'
        if( ! [datetime]::TryParse( $start , [ref]$parsed ) )
        {
            Throw "Invalid start time/date `"$start`""
        }
    }

    if( $parsed )
    {
        Remove-Variable -Name 'Start'
        [datetime]$script:start = $parsed
    }

    if( $PSBoundParameters[ 'duration' ] )
    {
        if( $PSBoundParameters[ 'end' ] )
        {
            Throw 'Cannot use both -duration and -end'
        }
        [int]$multiplier = 0
        switch( $duration[-1] )
        {
            's' { $multiplier = 1 }
            'm' { $multiplier = 60 }
            'h' { $multiplier = 3600 }
            'd' { $multiplier = 86400 }
            'w' { $multiplier = 86400 * 7 }
            'y' { $multiplier = 86400 * 365 }
            default { Throw "Unknown multiplier `"$($duration[-1])`"" }
        }
        if( $duration.Length -le 1 )
        {
            $secondsDuration = $multiplier
        }
        else
        {
            $secondsDuration = ( ( $duration.Substring( 0 , $duration.Length - 1 ) -as [decimal] ) * $multiplier )
        }
        [datetime]$end = $parsed.AddSeconds( $secondsDuration )
    }
    elseif( ! $PSBoundParameters[ 'end' ] )
    {
        Remove-Variable -Name 'End'
        [datetime]$script:end = Get-Date
    }
    elseif( ! [datetime]::TryParse( $end , [ref]$parsed ) )
    {
        Throw "Invalid end time/date `"$start`""
    }
    else
    {
        Remove-Variable -Name 'End'
        [datetime]$script:end = $parsed
    }
}

if( $start -gt $end )
{
    Throw "Start $(Get-Date -Date $start -Format G) is after end $(Get-Date -Date $end -Format G)"
}

if( $start -gt (Get-Date) )
{
    Write-Warning "Start $(Get-Date -Date $start -Format G) is in the future by $([math]::Round(($start - (Get-Date)).TotalHours,1)) hours"
}

if( $PSBoundParameters[ 'csv' ] )
{
    $arguments += @{
        'NoTypeInformation' = $true 
        'NoClobber' = ( ! $overWrite )
        'Path' = $csv }
    $command = Get-Command -Name Export-Csv
}
elseif( ! $noGridView )
{
    $arguments.Add( 'Title' , "From $(Get-Date -Date $start -Format G) to $(Get-Date -Date $end -Format G)" )
    $arguments.Add( 'PassThru' , $passThru )
    $command = Get-Command -Name Out-GridView
}
else
{
    $command = Get-Command -name Out-PassThru
    $arguments = @{}
}

[hashtable]$eventFilter =  @{ starttime = $start ; endtime = $end }
if( $badOnly )
{
    $eventFilter.Add( 'Level' , @( 1 , 2 , 3 ) )
}

if( $PSBoundParameters[ 'ids' ] )
{
    $eventFilter.Add( 'ID' , $ids )
}

$results = New-Object -TypeName System.Collections.Generic.List[psobject]
[int]$counter = 0

[array]$results = @( $(ForEach( $thisComputer in $Computer )
{
    $counter++
    Write-Verbose -Message "$counter / $($computer.Count) : $thiscomputer"

    [hashtable]$computerArgument = @{}
    if( $thisComputer -ne 'localhost' -and $thisComputer -ne $env:COMPUTERNAME -and $thisComputer -ne '.' )
    {
        $computerArgument.Add( 'ComputerName' , $thisComputer ) ## not the most efficient way of doing this but it's better than having to do it manually!
    }

    Get-WinEvent -ListLog $eventLogs @ComputerArgument -Verbose:$false | Where-Object { $_.RecordCount } | . { Process { Get-WinEvent @ComputerArgument -ErrorAction SilentlyContinue -Verbose:$False -FilterHashtable ( @{ logname = $_.logname } + $eventFilter ) | `
        Where-Object { ( [string]::IsNullOrEmpty( $excludeProvider) -or $_.ProviderName -notmatch $excludeProvider ) -and ( ! $excludeIds -or ! $excludeIds.Count -or $_.Id -notin $excludeIds ) -and ( ! $message -or $_.message -match $message ) -and ( ! $ignore -or $_.message -notmatch $ignore ) }}}}
) | Sort-Object -Property TimeCreated | Select-Object -ExcludeProperty TimeCreated,?*Id,Version,Qualifiers,Level,Task,OpCode,Keywords,Bookmark,*Ids,Properties -Property @{n='Date';e={"$(Get-Date -Date $_.TimeCreated -Format d) $((Get-Date -Date $_.TimeCreated).ToString('HH:mm:ss.fff'))"}},* | . $command @arguments )

if( $command -ne 'Export-CSV' )
{
    $results
}

