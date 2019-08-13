#requires -version 3
<#
    Analyse IIS log files to show requests/min/sec, and min, max, average and median response times per time interval, usually seconds

    @guyrleech 2019

    Modification History:

    13/08/19   GRL  Initial publice release
#>

<#
.SYNOPSIS

Analyse IIS log files to show requests/min/sec, and min, max, average and median response times per time interval, usually seconds

.PARAMETER logFile

The IIS log to be examined

.PARAMETER url

Optional regular expression where only those URIs which match this will be analysed

.PARAMETER start

Optional start time expressed as hh:mm which will only show requests on or after this time

.PARAMETER end

Optional end time expressed as hh:mm which will only show requests on or before this time

.PARAMETER failureCode

Flag requests as in error if their return code is greater than or equal to this value

.EXAMPLE

& '.\Analyse IIS log files.ps1' -Verbose -logfile "C:\inetpub\logs\LogFiles\W3SVC1\u_ex190805.log" | Out-GridView

Analyse the specified log file and display in an on screen filterable/sortable grid view

.EXAMPLE

& '.\Analyse IIS log files.ps1' -logfile "C:\inetpub\logs\LogFiles\W3SVC1\u_ex190803.log" -url '/PersonalizationServer/' -start 0600 -end 0800

Analyse the specified log file and output statistics on all requests between 0600 and 0800 only which match the URL '/PersonalizationServer/'

#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$true,HelpMessage='IIS log file path')]
    [string]$logfile ,
    [string]$url ,
    [string]$start ,
    [string]$end ,
    [int]$failureCode = 500
)

[long]$fastest = [int]::MaxValue
[long]$slowest = 0
[long]$mostRequestsPerInterval = 0
[long]$mostErrorsPerInterval = 0 
[long]$overallTotal = 0
[long]$intervals = 0
[long]$totalRequests = 0
[long]$totalFailures = 0
[string]$mostErrorsTime = $null
[string]$mostRequestsTime = $null
[datetime]$startTime = Get-Date

[int]$fromHours = -1
[int]$fromMinutes = 0
[int]$endHours = -1
[int]$endMinutes = 0

if( $PSBoundParameters[ 'start' ] )
{
    $fromHours = $start.SubString( 0 , 2 )
    if( $start.Length -gt 2 )
    {
        $fromMinutes = $start.SubString( $start.Length - 2 , 2 )
    }
}

if( $PSBoundParameters[ 'end' ] )
{
    $endHours = $end.SubString( 0 , 2 )
    if( $end.Length -gt 2 )
    {
        $endMinutes = $end.SubString( $end.Length - 2 , 2 )
    }
}

if( $endHours -ge 0 -and $fromHours -ge 0 -and ( $endHours -lt $fromHours -or ( $endHours -eq $fromHours -and $endMinutes -lt $fromMinutes ) ) )
{
    Throw "End time is before start time"
}

Write-Verbose "Starting at $(Get-Date -Date $startTime -Format G)"

[array]$results = @( Get-Content -Path $logfile -ErrorAction Stop | Where-Object { $_ -match '^(\d|#Fields)' } | . { Process `
{ 
    $_ -replace '^#Fields: ' , '' ## Convert columns to headings
}} | ConvertFrom-Csv -Delimiter ' ' | Where-Object { $_.'cs-uri-stem' -match $url `
        -and ( $fromHours -lt 0 -or ( ($thisHour = [int]($_.Time.SubString(0,2))) -ge $fromHours -and ( $thisHour -ne $fromHours -or [int]($_.Time.SubString(3,2)) -ge $fromMinutes ) ) ) `
         -and ( $endHours -lt 0 -or ( ($thisHour = [int]($_.Time.SubString(0,2))) -le $endHours  -and ( $thisHour -ne $endHours  -or [int]($_.Time.SubString(3,2)) -le $endMinutes ) ) ) } | Group-Object -Property Time | . { Process `
{
    $intervals++
    $record = $_
    $totalRequests += $record.Group.Count
    [int]$min = [int]::MaxValue
    [int]$max = 0
    [long]$total = 0
    [int]$failures = 0
    $time = $null
    $date = $null

    [int[]]$times = @( 
        $record.Group.GetEnumerator() | . { Process `
        {
            if( ! $time )
            {
                $time = $_.Time
            }
            if( ! $date )
            {
                $date = $_.Date
            }
            [int]$timeTaken = [int]$_.'time-taken'
            
            if( $timeTaken -lt $min )
            {
                $min = $timeTaken
            }
            if( $timeTaken -gt $max )
            {
                $max = $timeTaken
            }
            $total += $timeTaken
            if( $_.'sc-status' -ge $failureCode )
            {
                $failures++
            }
            $timeTaken
        }} )
    $overallTotal += $total
    $totalFailures += $failures
    if( $min -lt $fastest )
    {
        $fastest = $min
    }
    if( $max -gt $slowest )
    {
        $slowest = $max
    }

    if( $record.Group.Count -gt $mostRequestsPerInterval )
    {
        $mostRequestsPerInterval = $record.Group.Count
        $mostRequestsTime = $time
    }
    if( $failures -gt $mostErrorsPerInterval )
    {
        $mostErrorsPerInterval = $failures
        $mostErrorsTime = $time
    }
    $result = [pscustomobject][ordered]@{
        'Date' = $date
        'Time' = $record.Name
        'Requests' = $record.Count
        'Fastest (ms)' = $min
        'Slowest (ms)' = $max
        'Average (ms)' = [int]($total / $record.Group.Count)
        'Median (ms)' = ( $times | Sort-Object | Select-Object -Skip ([math]::Floor( $times.Count / 2)) -First 1 )
        'Failures' = $failures
    }
    $result
}} )

[datetime]$endTime = Get-Date

Write-Verbose "Processesd $intervals intervals in $(($endTime - $startTime).TotalSeconds) seconds"
if( $totalRequests )
{
    Write-Verbose "$totalRequests total requests, fastest $fastest ms, slowest $slowest ms, average $([int]($overallTotal/$totalRequests)) ms, failures $totalFailures"
}
else
{
    Write-Warning "No requests found"
}
if( $mostErrorsPerInterval )
{
    Write-Verbose "Highest error rate per interval is $mostErrorsPerInterval at $mostErrorsTime"
}
else
{
    Write-Verbose "No error codes gerater than or equal to $failureCode found"
}
Write-Verbose "Most requests per interval is $mostRequestsPerInterval at $mostRequestsTime"

$results