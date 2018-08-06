<#
    Monitor process start/stop

    @guyrleech 2018

    Modification history:
#>

<#
.SYNOPSIS

Show details of processes started and stopped

.PARAMETER sessionId

Only processes in the session specified will be monitored. Note that process stop notifications do not have a valid session id.

.PARAMETER thisSession

Only processes in the current session will be monitored. Note that process stop notifications do not have a valid session id.

.PARAMETER newOnly

Only report on process activity from when the script was invoked

.PARAMETER waitFor

Stop monitoring when the given process starts or stops

.PARAMETER runFor

Monitor for this number of seconds

.EXAMPLE

& '.\Monitor process start stop.ps1' -runFor 120

Monitor all processes for 2 minutes

.EXAMPLE

& '.\Monitor process start stop.ps1' -newOnly -waitFor explorer.exe -runFor 300 | Out-GridView

Monitor new processes either for 5 minutes or until explorer.exe starts or stops and display the results in a grid view

.NOTES

Must be run as a user with administrative rights otherwise the event subscription will fail

#>

[CmdletBinding()]

Param
(
    [int]$sessionId = -1 ,
    [switch]$thisSession ,
    [switch]$newOnly ,
    [string]$waitFor ,
    [long]$runFor = 60
)

if( $thisSession )
{
    $sessionId = (Get-Process -Id $pid).SessionId
}

[string]$sourceIdentifierProcessStarted = 'Process Started' 
[string]$sourceIdentifierProcessStopped = 'Process Stopped' 

## in case script didn't terminate properly
Unregister-Event -SourceIdentifier $sourceIdentifierProcessStarted -ErrorAction SilentlyContinue
Unregister-Event -SourceIdentifier $sourceIdentifierProcessStopped -ErrorAction SilentlyContinue

Register-WmiEvent –class 'Win32_ProcessStopTrace'  –sourceIdentifier $sourceIdentifierProcessStopped  -ErrorAction Stop
Register-WmiEvent –class 'Win32_ProcessStartTrace' –sourceIdentifier $sourceIdentifierProcessStarted  -ErrorAction Stop

if( ! $? )
{
    Exit 1
}

[long]$runForMilliseconds = $runFor * 1000

[hashtable]$processes = @{}

[bool]$carryOn = $true

$timer = [Diagnostics.Stopwatch]::StartNew()

[datetime]$now = [System.DateTime]::Now

While( $carryOn -and $timer.ElapsedMilliseconds -le $runForMilliseconds )
{
    $event = Wait-Event -Timeout ( $runFor - $timer.ElapsedMilliseconds / 1000 )
    if( $event)
    {
        if( $event.SourceIdentifier -eq $sourceIdentifierProcessStarted -or $event.SourceIdentifier -eq $sourceIdentifierProcessStopped )
        {
            ## session id defined as "Session under which the process exists" - gets set to zero for process stopped!
            if( $sessionId -lt 0 -or $event.SourceIdentifier -eq $sourceIdentifierProcessStopped -or $event.SourceArgs.NewEvent.SessionId -eq $sessionId )
            {
                [datetime]$eventTime = [DateTime]::FromFileTime( $event.SourceArgs.NewEvent.TIME_CREATED )
                if( ! $newOnly -or $eventTime -ge $now )
                {
                    $result = [pscustomobject]@{ 
                        'Type' = $event.SourceIdentifier  
                        'Process Id' = $event.SourceArgs.NewEvent.ProcessID  
                        'Parent Process Id' = $event.SourceArgs.NewEvent.ParentProcessID  
                        'Process Name' = $event.SourceArgs.NewEvent.ProcessName  
                        'Session Id' = $event.SourceArgs.NewEvent.SessionId 
                        'Exit Code' = $null
                        'Command Line' = $null
                        'Owner' = $null
                        'Duration (ms)' = $null
                        'Time' = $eventTime }
                    if( $event.SourceIdentifier -eq $sourceIdentifierProcessStopped )
                    {
                        $result.'Exit Code' = $event.SourceArgs.NewEvent.ExitStatus
                        $process = $processes[ $event.SourceArgs.NewEvent.ProcessID ]
                        if( $process )
                        {
                            $result.'Duration (ms)'= [int]( ( $event.SourceArgs.NewEvent.TIME_CREATED - $process.TIME_CREATED ) / 1E4 )
                            ## these values get zeroed out
                            $result.'Parent Process Id' = $process.ParentProcessId
                            $result.'Session Id' = $process.SessionId
                            $processes.Remove( $event.SourceArgs.NewEvent.ProcessID )
                        }
                    }
                    else
                    {
                        $thisProcess = Get-WmiObject -Class Win32_process -Filter "ProcessId = '$($event.SourceArgs.NewEvent.ProcessID)'"
                        if( $thisProcess )
                        {
                            $result.'Command Line' = $thisProcess.CommandLine
                        }
                        if( $event.SourceArgs.NewEvent.Sid )
                        {
                            $sid = (New-Object System.Security.Principal.SecurityIdentifier($event.SourceArgs.NewEvent.Sid,0)).Value
                            if( $sid )
                            {
                                $user = ([System.Security.Principal.SecurityIdentifier]($sid)).Translate([System.Security.Principal.NTAccount]).Value
                                if( $user )
                                {
                                    $result.Owner = $user
                                }
                                else
                                {
                                    $result.Owner = $sid
                                }
                            }
                        }
                        $processes.Add( $event.SourceArgs.NewEvent.ProcessID , $event.SourceArgs.NewEvent )
                    }
                    $result
                    if( ! [string]::IsNullOrEmpty( $waitFor ) -and $waitFor -match $result.'Process Name' )
                    {
                        Write-Verbose "Waited for process occurred"
                        $carryOn = $false
                    }
                }
            }
            else
            {
                Write-Verbose ( "Ignoring {2} {0} pid {1} session {3}" -f $event.SourceArgs.NewEvent.ProcessName , $event.SourceArgs.NewEvent.ProcessId , $event.SourceIdentifier , $event.SourceArgs.NewEvent.SessionId )
            }
        }
        else
        {
            Write-Verbose "Ignoring event from source $($event.SourceIdentifier)"
        }
        $event | Remove-Event -ErrorAction SilentlyContinue
    }
}

Unregister-Event -SourceIdentifier $sourceIdentifierProcessStarted
Unregister-Event -SourceIdentifier $sourceIdentifierProcessStopped
