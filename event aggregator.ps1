<#
    Get all events from all event logs in a given time window and output to grid view or csv

    @guyrleech 2019

    Modification History

    24/07/19   GRL   Added -last parameter, defaulted -end if not specified and verification of times/dates
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

.PARAMETER csv

The name of a non-existent csv file to write the results to. If not specified then the results will go to an on screne grid view

.PARAMETER badOnly

Only show critical, warning and error events. If not specified then all events will be shown.

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

#>

[CmdletBinding()]

Param
(
    [Parameter(ParameterSetName='StartTime',Mandatory=$true,HelpMessage='Start time/date for event search')]
	[string]$start ,
    [Parameter(ParameterSetName='StartTime',Mandatory=$false,HelpMessage='End time/date for event search')]
	[string]$end ,
    [Parameter(ParameterSetName='Last',Mandatory=$true,HelpMessage='Search for events in last seconds/minutes/hours/days/weeks/years')]
    [string]$last ,
    [string]$csv ,
    [switch]$badOnly ,
    [string]$eventLogs = '*' ,
    [string]$computer
)

[hashtable]$arguments = @{}
[string]$command = $null

if( $PSBoundParameters[ 'csv' ] )
{
    $arguments += @{
        'NoTypeInformation' = $true 
        'NoClobber' = $true 
        'Path' = $csv }
    $command = Get-Command -Name Export-Csv
}
else
{
    $arguments.Add( 'Title' , "From $start to $end" )
    $command = Get-Command -Name Out-GridView
}

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
elseif( ! $PSBoundParameters[ 'end' ] )
{
    $end = Get-Date
}
else
{
    ## Check time formats as bad ones get stripped by query so search whole event log
    $parsed = New-Object -TypeName 'DateTime'
    if( ! [datetime]::TryParse( $start , [ref]$parsed ) )
    {
        Throw "Invalid start time/date `"$start`""
    }
    if( ! [datetime]::TryParse( $end , [ref]$parsed ) )
    {
        Throw "Invalid end time/date `"$start`""
    }
}

if( $start -gt $end )
{
    Throw "Start $(Get-Date -Date $start -Format G) is after end $(Get-Date -Date $end -Format G)"
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

Get-WinEvent -ListLog $eventLogs @ComputerArgument | Where-Object { $_.RecordCount } | . { Process { Get-WinEvent @ComputerArgument -ErrorAction SilentlyContinue -FilterHashtable ( @{ logname = $_.logname } + $eventFilter ) }} | Sort-Object -Property TimeCreated | Select-Object -ExcludeProperty TimeCreated,?*Id,Version,Qualifiers,Level,Task,OpCode,Keywords,Bookmark,*Ids,Properties -Property @{n='Date';e={"$(Get-Date -Date $_.TimeCreated -Format d) $((Get-Date -Date $_.TimeCreated).ToString('HH:mm:ss.fff'))"}},* | . $command @arguments
