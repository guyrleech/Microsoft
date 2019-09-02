<#
    Get all events from all event logs in a given time window and output to grid view or csv

    @guyrleech 2019

    Modification History

    24/07/19   GRL   Added -last parameter, defaulted -end if not specified and verification of times/dates
    02/09/19   GRL   Added -passThru, -duration, -overwrite , -nogridview and pass thru for grid view
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

A remote computer to query. If not specified then the local computer is used.

.EXAMPLE

& '.\event aggregator.ps1' -start 10:38 -end 10:45 -badOnly

Show all critical, warning and error events that occurred between 10:38 and 10:45 today in an on screen gridview

.EXAMPLE

& '.\event aggregator.ps1' -start "10:38 29/06/19" -end "10:45 29/06/19" -csv c:\badevents.csv

Export all events that occurred between 10:38 and 10:45 on the 29th June 2019 top the named csv file

.EXAMPLE

& '.\event aggregator.ps1' -last 5m -Computer fred

Show allevents that occurred in the last 5 minutes on computer "fred" in an on screen gridview

.EXAMPLE

& '.\event aggregator.ps1' -start 10:40 -duration 2m -eventLogs *shell-core/operational*

Show events from the specified event log that occurred between 10:40 and 10:42 today in an on screen gridview

#>

[CmdletBinding()]

Param
(
    [Parameter(ParameterSetName='StartTime',Mandatory=$true,HelpMessage='Start time/date for event search')]
	[string]$start ,
    [Parameter(ParameterSetName='StartTime',Mandatory=$false,HelpMessage='End time/date for event search')]
	[string]$end ,
    [Parameter(ParameterSetName='StartTime',Mandatory=$false,HelpMessage='Duration of the search')]
	[string]$duration ,
    [Parameter(ParameterSetName='Last',Mandatory=$true,HelpMessage='Search for events in last seconds/minutes/hours/days/weeks/years')]
    [string]$last ,
    [string]$csv ,
    [switch]$badOnly ,
    [string]$eventLogs = '*' ,
    [switch]$noGridView ,
    [switch]$overWrite ,
    [switch]$passThru ,
    [string]$computer
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
    if( $PSBoundParameters[ 'start' ] -or $PSBoundParameters[ 'end' ] )
    {
        Throw 'Cannot use -start or -end with -last'
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
    ## Check time formats as bad ones get stripped by query so search whole event log
    $parsed = New-Object -TypeName 'DateTime'
    if( ! [datetime]::TryParse( $start , [ref]$parsed ) )
    {
        Throw "Invalid start time/date `"$start`""
    }
    else
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

[hashtable]$computerArgument = @{}
if( $PSBoundParameters[ 'computer' ] -and $computer -ne 'localhost' -and $computer -ne $env:COMPUTERNAME )
{
    $computerArgument.Add( 'ComputerName' , $computer ) ## not the most efficient way of doing this!
}

[array]$results = @( Get-WinEvent -ListLog $eventLogs @ComputerArgument | Where-Object { $_.RecordCount } | . { Process { Get-WinEvent @ComputerArgument -ErrorAction SilentlyContinue -FilterHashtable ( @{ logname = $_.logname } + $eventFilter ) }} | Sort-Object -Property TimeCreated | Select-Object -ExcludeProperty TimeCreated,?*Id,Version,Qualifiers,Level,Task,OpCode,Keywords,Bookmark,*Ids,Properties -Property @{n='Date';e={"$(Get-Date -Date $_.TimeCreated -Format d) $((Get-Date -Date $_.TimeCreated).ToString('HH:mm:ss.fff'))"}},* | . $command @arguments )

if( $command -ne 'Export-Csv' )
{
    $results
}
