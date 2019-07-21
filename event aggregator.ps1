<#
    Get all events from all event logs in a given time window and output to grid view or csv

    @guyrleech 2019
#>


<#
.SYNOPSIS

Retrieve all events from all 300+ event logs in a given time/date range and show in a sortable/filterable gridview or export to csv

.PARAMETER start

The start time/date to show events from. If no date is given then the current day is used.

.PARAMETER end

The end time/date to show events from. If no date is given then the current day is used.

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

#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$true)]
	[string]$start ,
    [Parameter(Mandatory=$true)]
	[string]$end ,
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

Get-WinEvent -ListLog $eventLogs @ComputerArgument | Where-Object { $_.RecordCount } | . { Process { Get-WinEvent @ComputerArgument -ErrorAction SilentlyContinue -FilterHashtable ( @{ logname = $_.logname } + $eventFilter ) }} | Sort-Object -Property TimeCreated | Select-Object -ExcludeProperty TimeCreated,?*Id,Version,Qualifiers,Level,Task,OpCode,Keywords,Bookmark,*Ids,Properties -Property @{n='Time';e={(Get-Date $_.TimeCreated).ToString('HH:mm:ss.fff')}},* | . $command @arguments
