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
    05/06/20   GRL   Added user name,sid, pid and tid
    20/07/20   GRL   Added -provider argument
    22/07/20   GRL   Fixed duplicate entries produced by duplicate events logs returned for providers
    25/09/20   GRL   Added minutes past hour functionality for troubleshooting recurring problems
    31/05/21   GRL   Added -credential parameter
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

.PARAMETER credential

Credential to use to remote to the specified computer(s). If $null is passed, a credential will be prompted for.

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

.PARAMETER minutesPastHour

Only include events after this number of minutes past the hour and for the number of minutes specified by -minutes.

.PARAMETER minutes

Only include events before this number of minutes past the hour when added to the minutes specified by -minutesPastHour.

.PARAMETER provider

The name, or pattern, for an event log source/provider to only return events from that

.PARAMETER eventLogs

A pattern matching the event logs to search. By default all event logs are searched.

.PARAMETER computer

One or more remote computers to query. If not specified then the local computer is used.

.EXAMPLE

& '.\event aggregator.ps1' -start 10:38 -end 10:45 -badOnly

Show all critical, warning and error events that occurred between 10:38 and 10:45 today in an on screen gridview

.EXAMPLE

& '.\event aggregator.ps1' -start 07:57 -minutesPastHour 57 -minutes 1.75

Show all  events that occurred since 0757 today but only those that occurred between 57:00 minutes and 58:45 past the hour in an on screen gridview

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
    [double]$minutesPastHour ,
    [double]$minutes ,
    [int[]]$ids ,
    [int[]]$excludeIds ,
    [string]$excludeProvider ,
    [string]$provider ,
    [string]$message ,
    [string]$ignore ,
    [string]$csv ,
    [switch]$badOnly ,
    [string]$eventLogs = '*' ,
    [switch]$noGridView ,
    [switch]$overWrite ,
    [switch]$passThru ,
    [PSCredential]$credential ,
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
    [string]$title = "From $(Get-Date -Date $start -Format G) to $(Get-Date -Date $end -Format G)"
    if( $PSBoundParameters[ 'provider' ])
    {
        $title += ", provider $provider"
    }
    if( $PSBoundParameters[ 'ids' ])
    {
        $title += ", ids $($ids -join ',')"
    }
    if( $PSBoundParameters[ 'badonly' ])
    {
        $title += ", bad only"
    }
    if( $PSBoundParameters[ 'message' ])
    {
        $title += ", message `"$message`""
    }
    $arguments.Add( 'Title' , $title )
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

if( $PSBoundParameters[ 'provider' ] )
{
    $eventFilter.Add( 'ProviderName' , $provider )
}

if( $PSBoundParameters[ 'ids' ] )
{
    $eventFilter.Add( 'ID' , $ids )
}

$results = New-Object -TypeName System.Collections.Generic.List[psobject]
[int]$counter = 0
[int]$secondsPastHour = -1
[int]$secondsPastHourEnd = -1

if( $PSBoundParameters[ 'minutesPastHour' ] )
{
    if( $minutesPastHour -lt 0 -or $minutesPastHour -ge 60 )
    {
        Throw "Minutes past hour value $minutesPastHour is invalid - must >= 0 and < 60"
    }
    if( ! $PSBoundParameters[ 'minutes' ] )
    {
        Throw 'Must specify the number of minutes to include via -minutes when using -minutesPastHour'
    }
    $secondsPastHour = $minutesPastHour * 60 ## allows fractional minutes
    $secondsPastHourEnd = $secondsPastHour + $minutes * 60
}
elseif( $PSBoundParameters[ 'minutes' ] )
{
    Throw 'Must specify the number of minutes past the hour via -minutesPastHour to start at when using -minutes'
}

[array]$results = @( $(ForEach( $thisComputer in $Computer )
{
    $counter++
    Write-Verbose -Message "$counter / $($computer.Count) : $thiscomputer"

    [bool]$continue = $true
    [hashtable]$computerArgument = @{}
    if( $thisComputer -ne 'localhost' -and $thisComputer -ne $env:COMPUTERNAME -and $thisComputer -ne '.' )
    {
        $computerArgument.Add( 'ComputerName' , $thisComputer ) ## not the most efficient way of doing this but it's better than having to do it manually!
    }
    if( $PSBoundParameters.ContainsKey( 'credential' ))
    {
        if( ! $credential )
        {
            $credential = Get-Credential -Message "For $($computer.Count) remote event logs"
        }
        if( $credential )
        {
            $computerArgument.Add( 'credential' , $credential )
        }
    }

    [string[]]$eventLogsToSearch = $eventLogs

    if( $eventFilter[ 'ProviderName' ] )
    {
        if( $providerDetails = Get-WinEvent -ListProvider $provider -ErrorAction SilentlyContinue @computerArgument )
        {
            $eventLogsToSearch = $providerDetails.LogLinks | Select-Object -ExpandProperty LogName | Sort-Object -Unique
            Write-Verbose -Message "Checking providers $($eventFilter.ProviderName) in $($eventLogsToSearch.Count) event logs $($eventLogsToSearch -join ' , ') on $thisComputer"
        }
        else
        {
            Write-Warning -Message "No event provider `"$provider`" on $thisComputer"
            $continue = $false
        }
    }

    if( $continue )
    {
        (Get-WinEvent -ListLog $eventLogsToSearch @ComputerArgument -Verbose:$false | Where-Object { $_.RecordCount } ).ForEach( { (Get-WinEvent @ComputerArgument -ErrorAction SilentlyContinue -Verbose:$False -FilterHashtable ( @{ logname = $_.logname } + $eventFilter )).Where(
        {            
            ([string]::IsNullOrEmpty( $excludeProvider) -or $_.ProviderName -notmatch $excludeProvider ) -and ( ! $excludeIds -or ! $excludeIds.Count -or $_.Id -notin $excludeIds ) -and ( ! $message -or $_.message -match $message ) -and ( ! $ignore -or $_.message -notmatch $ignore ) `
                -and ( $secondsPastHour -lt 0 -or (( $seconds = ( $_.TimeCreated.Minute * 60 + $_.TimeCreated.Seconds ) ) -ge $secondsPastHour -and $seconds -le $secondsPastHourEnd ))
        })})
    }
}) | Sort-Object -Property TimeCreated | Select-Object -ExcludeProperty TimeCreated,RecordId,ProviderId,*ActivityId,Version,Qualifiers,Level,Task,OpCode,Keywords,Bookmark,*Ids,Properties -Property @{n='Date';e={"$(Get-Date -Date $_.TimeCreated -Format d) $((Get-Date -Date $_.TimeCreated).ToString('HH:mm:ss.fff'))"}},*,@{n='User';e={if( $_.UserId ) { ([System.Security.Principal.SecurityIdentifier]($_.UserId)).Translate([System.Security.Principal.NTAccount]).Value }}} | . $command @arguments )

if( $command -ne 'Export-CSV' )
{
    $results
}

# SIG # Begin signature block
# MIINRQYJKoZIhvcNAQcCoIINNjCCDTICAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUyggCQU+tCaR4Bl9XcMCahhKr
# U26gggqHMIIFMDCCBBigAwIBAgIQBAkYG1/Vu2Z1U0O1b5VQCDANBgkqhkiG9w0B
# AQsFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVk
# IElEIFJvb3QgQ0EwHhcNMTMxMDIyMTIwMDAwWhcNMjgxMDIyMTIwMDAwWjByMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQg
# Q29kZSBTaWduaW5nIENBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA
# +NOzHH8OEa9ndwfTCzFJGc/Q+0WZsTrbRPV/5aid2zLXcep2nQUut4/6kkPApfmJ
# 1DcZ17aq8JyGpdglrA55KDp+6dFn08b7KSfH03sjlOSRI5aQd4L5oYQjZhJUM1B0
# sSgmuyRpwsJS8hRniolF1C2ho+mILCCVrhxKhwjfDPXiTWAYvqrEsq5wMWYzcT6s
# cKKrzn/pfMuSoeU7MRzP6vIK5Fe7SrXpdOYr/mzLfnQ5Ng2Q7+S1TqSp6moKq4Tz
# rGdOtcT3jNEgJSPrCGQ+UpbB8g8S9MWOD8Gi6CxR93O8vYWxYoNzQYIH5DiLanMg
# 0A9kczyen6Yzqf0Z3yWT0QIDAQABo4IBzTCCAckwEgYDVR0TAQH/BAgwBgEB/wIB
# ADAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwMweQYIKwYBBQUH
# AQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQwYI
# KwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFz
# c3VyZWRJRFJvb3RDQS5jcnQwgYEGA1UdHwR6MHgwOqA4oDaGNGh0dHA6Ly9jcmw0
# LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwOqA4oDaG
# NGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RD
# QS5jcmwwTwYDVR0gBEgwRjA4BgpghkgBhv1sAAIEMCowKAYIKwYBBQUHAgEWHGh0
# dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwCgYIYIZIAYb9bAMwHQYDVR0OBBYE
# FFrEuXsqCqOl6nEDwGD5LfZldQ5YMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6en
# IZ3zbcgPMA0GCSqGSIb3DQEBCwUAA4IBAQA+7A1aJLPzItEVyCx8JSl2qB1dHC06
# GsTvMGHXfgtg/cM9D8Svi/3vKt8gVTew4fbRknUPUbRupY5a4l4kgU4QpO4/cY5j
# DhNLrddfRHnzNhQGivecRk5c/5CxGwcOkRX7uq+1UcKNJK4kxscnKqEpKBo6cSgC
# PC6Ro8AlEeKcFEehemhor5unXCBc2XGxDI+7qPjFEmifz0DLQESlE/DmZAwlCEIy
# sjaKJAL+L3J+HNdJRZboWR3p+nRka7LrZkPas7CM1ekN3fYBIM6ZMWM9CBoYs4Gb
# T8aTEAb8B4H6i9r5gkn3Ym6hU/oSlBiFLpKR6mhsRDKyZqHnGKSaZFHvMIIFTzCC
# BDegAwIBAgIQBP3jqtvdtaueQfTZ1SF1TjANBgkqhkiG9w0BAQsFADByMQswCQYD
# VQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGln
# aWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQgQ29k
# ZSBTaWduaW5nIENBMB4XDTIwMDcyMDAwMDAwMFoXDTIzMDcyNTEyMDAwMFowgYsx
# CzAJBgNVBAYTAkdCMRIwEAYDVQQHEwlXYWtlZmllbGQxJjAkBgNVBAoTHVNlY3Vy
# ZSBQbGF0Zm9ybSBTb2x1dGlvbnMgTHRkMRgwFgYDVQQLEw9TY3JpcHRpbmdIZWF2
# ZW4xJjAkBgNVBAMTHVNlY3VyZSBQbGF0Zm9ybSBTb2x1dGlvbnMgTHRkMIIBIjAN
# BgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAr20nXdaAALva07XZykpRlijxfIPk
# TUQFAxQgXTW2G5Jc1YQfIYjIePC6oaD+3Zc2WN2Jrsc7bj5Qe5Nj4QHHHf3jopLy
# g8jXl7Emt1mlyzUrtygoQ1XpBBXnv70dvZibro6dXmK8/M37w5pEAj/69+AYM7IO
# Fz2CrTIrQjvwjELSOkZ2o+z+iqfax9Z1Tv82+yg9iDHnUxZWhaiEXk9BFRv9WYsz
# qTXQTEhv8fmUI2aZX48so4mJhNGu7Vp1TGeCik1G959Qk7sFh3yvRugjY0IIXBXu
# A+LRT00yjkgMe8XoDdaBoIn5y3ZrQ7bCVDjoTrcn/SqfHvhEEMj1a1f0zQIDAQAB
# o4IBxTCCAcEwHwYDVR0jBBgwFoAUWsS5eyoKo6XqcQPAYPkt9mV1DlgwHQYDVR0O
# BBYEFE16ovlqIk5uX2JQy6og0OCPrsnJMA4GA1UdDwEB/wQEAwIHgDATBgNVHSUE
# DDAKBggrBgEFBQcDAzB3BgNVHR8EcDBuMDWgM6Axhi9odHRwOi8vY3JsMy5kaWdp
# Y2VydC5jb20vc2hhMi1hc3N1cmVkLWNzLWcxLmNybDA1oDOgMYYvaHR0cDovL2Ny
# bDQuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1jcy1nMS5jcmwwTAYDVR0gBEUw
# QzA3BglghkgBhv1sAwEwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNl
# cnQuY29tL0NQUzAIBgZngQwBBAEwgYQGCCsGAQUFBwEBBHgwdjAkBggrBgEFBQcw
# AYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tME4GCCsGAQUFBzAChkJodHRwOi8v
# Y2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRTSEEyQXNzdXJlZElEQ29kZVNp
# Z25pbmdDQS5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAQEAU9zO
# 9UpTkPL8DNrcbIaf1w736CgWB5KRQsmp1mhXbGECUCCpOCzlYFCSeiwH9MT0je3W
# aYxWqIpUMvAI8ndFPVDp5RF+IJNifs+YuLBcSv1tilNY+kfa2OS20nFrbFfl9QbR
# 4oacz8sBhhOXrYeUOU4sTHSPQjd3lpyhhZGNd3COvc2csk55JG/h2hR2fK+m4p7z
# sszK+vfqEX9Ab/7gYMgSo65hhFMSWcvtNO325mAxHJYJ1k9XEUTmq828ZmfEeyMq
# K9FlN5ykYJMWp/vK8w4c6WXbYCBXWL43jnPyKT4tpiOjWOI6g18JMdUxCG41Hawp
# hH44QHzE1NPeC+1UjTGCAigwggIkAgEBMIGGMHIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAv
# BgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EC
# EAT946rb3bWrnkH02dUhdU4wCQYFKw4DAhoFAKB4MBgGCisGAQQBgjcCAQwxCjAI
# oAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIB
# CzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFGnNaSDwAQfk8s70WhTI
# 3jYAGPo6MA0GCSqGSIb3DQEBAQUABIIBAKuUy40fiBZSqtlMuZYPVlbkpf/yZLKA
# tqARX7N6bASK0EC5HjjVMlhhha8RfRh0cIK+mTMDGC5anxQB1FGoewrn3oac/ph1
# AbHJed5u2WCVpizG/N82LATvYgZImHLOKzodI9GC8Dx/ef7dbKmtza5dwcKXu0xK
# BrXhv0cAMyXXdAQDgCPehK9lTahAk2Uzu2/EP1OnHUG7idinAUV/uAluJfm9KWf6
# 759dqAmgRYwQrKp7YlyTmhFoej/Er+ZGNQBABW/03kqDzuvx6EVbJCq5NRgTwBC2
# F483zD6TmAmIYBHLaCO7X0EK5l3F4sPNDoQTd3mzgu/VxBM2ut6wF1k=
# SIG # End signature block
