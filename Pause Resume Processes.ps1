#requires -version 3

<#
.SYNOPSIS

Pause/resume processes

.DESCRIPTION

Use to pause processes, and optionally empty their working sets, so they cannot be used and will consume no CPU. Use cases include:

  1. Stop unwanted software from running which would respawn if terminated
  2. Only allow certain applications to be used during given times but allow the user to continue when the app is running again after being paused
  3. Stop an over consuming process from impacting other processes but resume it when capacity is available to finish its work or troubleshoot it

.PARAMETER id

Comma separated list of process ids to operate on

.PARAMETER name

Comma separated list of process names to operate on

.PARAMETER resume

Processes will be resumed. If not specified they will be paused.

.PARAMETER allSessions

Operate on the named processes in all sessions. If not specified will only operate on processes in the session the script is running in.
Processes running in session 0 will not be operated on.

.PARAMETER pipename

The name of the named pipe to either wait for notification on or to send a notficiation over when -signal is specified

.PARAMETER quiet

Do not report errors

.PARAMETER signal

The signal to send down the named pipe specified via -pipename. When this is specified, this invocation of the script will not pause or resume any processes itself

.PARAMETER trim

Empty the working sets of processes operated on

.EXAMPLE

& ".\Pause Resume Processes.ps1" -Verbose -name DodgyApp -pipeName "Wait4DodgyApp" -allsessions -trim -logfile c:\scripts\pawser.log -append

Pause all instances of DodgyApp.exe in all sessions, empty their working sets and then wait for a message to come down the Wait4DodgyApp named pipe whereupon all paused processes will be resumed

.EXAMPLE

& ".\Pause Resume Processes.ps1" -pipeName "Wait4DodgyApp" -signal "All Done Thanks"

Send the message "All Done Thanks" down the named pipe called "Wait4DodgyApp" so that any instance of this script listening on that named pipe will resume processes that it paused

.EXAMPLE

& ".\Pause Resume Processes.ps1" -Verbose -name DodgyApp -resume -allsessions -logfile c:\scripts\pawser.log -append

Resume all instances of DodgyApp.exe in all sessions. An error message will be seen for any instances of dodgyapp.exe which weren't previosuly paused but the processes will not be affected in any way

.NOTES

The script uses the DebugActiveProcess() and DebugActiveProcessStop() API calls.

||||           ||||
VVVV IMPORTANT VVVV

Only the process that calls DebugActiveProcess() can successfuly invoke DebugActiveProcessStop() for the same process so if the pausing is invoked via a scheduled task or similar,
the -Pipename argument must be used if you want to resume the processes at a later time since otherwise the powershell.exe process invoking the pause will exit which will cause the
paused processes to be exited. To resume the processes, run the script again with -pipename specifying the same named pipe name and -signal

Named pipe code borrowed with thanks from https://stackoverflow.com/questions/36224688/pass-information-between-independently-running-powershell-scrips

USE THIS SCRIPT AT YOUR OWN RISK. THE AUTHOR ACCEPTS NO LIABILITY FOR PROBLEMS CAUSED BY USING THIS SCRIPT.

Modification Hitory:

    @guyrleech 26/04/20  Initial release
#>

[CmdletBinding()]

Param
(
    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [int[]]$id ,
    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [string[]]$name ,
    [switch]$resume ,
    [switch]$allSessions ,
    [string]$pipeName ,
    [switch]$quiet ,
    [string]$signal ,
    [switch]$trim ,
    [string]$logfile ,
    [switch]$append
)

Begin
{
    [string]$logging = $null

    if( $PSBoundParameters[ 'logfile' ] )
    {
        $loggin = Start-Transcript -Path $logfile -Append:$append
    }

    Function Invoke-PauseResumeProcess
    {
        [CmdletBinding()]

        Param
        (
            [int[]]$id ,
            [bool]$resume ,
            [bool]$trim
        )

        [int]$operatedOn = 0

        ForEach( $processId in $id )
        {
            if( ( $thisProcess = Get-Process -Id $processId -ErrorAction SilentlyContinue ) -and $thisProcess.SessionId -eq 0 )
            {
                Write-Warning -Message "Not operating on pid $id ($($thisProcess.Name)) as running in session 0"
            }
            else
            {
                Write-Verbose -Message "$(if( $resume ) { 'Resuming' } else { 'Pausing' }) pid $processId ($($thisProcess.Name)) in session $($thisProcess.SessionId)"

                [bool]$result = $( if( $resume )
                {
                    [Pinvoke.Win32.Process]::DebugActiveProcessStop( $processId ); $LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
                }
                else
                {
                    [Pinvoke.Win32.Process]::DebugActiveProcess( $processId ); $LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
                })

                if( ! $result )
                {
                    if( ! $quiet )
                    {
                        Write-Error -Message "Failed to $(if( $resume ) { 'resume' } else { 'pause' }) process id $processId - $lasterror"
                    }
                }
                else
                {
                    $operatedOn++
                    if( $trim )
                    {
                        if( $thisProcess -and $thisProcess.Handle )
                        {
                            Write-Verbose -Message "Emptying working set of pid $processId ($($thisProcess.Name)) in session $($thisProcess.SessionId)"

                            [bool]$trimmed = [Pinvoke.Win32.Process]::SetProcessWorkingSetSizeEx( $thisProcess.Handle , -1 , -1 , 0 ); $LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()

                            if( ! $trimmed )
                            {
                                Write-Warning -Message "Failed to empty working set of pid $processId - $LastError"
                            }
                        }
                        else
                        {
                            Write-Warning -Message "Failed to get handle to pid $processId to empty working set"
                        }
                    }
                }
            }
        }

        $operatedOn
    }

    Add-Type @"
    using System;

    using System.Runtime.InteropServices;

    namespace PInvoke.Win32
    {
        public static class Process
        {    
            [DllImport( "kernel32.dll",SetLastError = true )]
                public static extern bool DebugActiveProcess( UInt32 dwProcessId );
            [DllImport( "kernel32.dll",SetLastError = true )]
                public static extern bool DebugActiveProcessStop( UInt32 dwProcessId );
            [DllImport("kernel32.dll", SetLastError=true)]
                public static extern bool SetProcessWorkingSetSizeEx( IntPtr proc, int min, int max , int flags );
        }
    }
"@

    if( ! [string]::IsNullOrEmpty( $signal ) )
    {
        if( [string]::IsNullOrEmpty( $pipeName ) )
        {
            Throw 'Must specify -pipename when using -signal'
        }

        if( $namedPipe = New-Object -Typename IO.Pipes.NamedPipeServerStream( $pipeName , 'Out' ) )
        {
            $namedPipe.WaitForConnection()

            if( $writer = New-Object -Typename IO.StreamWriter( $namedPipe ) )
            {
                $writer.AutoFlush = $true
                $writer.WriteLine( $signal )
                $writer.Dispose()
                $writer = $null
            }
            $namedPipe.Dispose()
            $namedPipe = $null
        }
        Exit 0
    }
    
    [int]$operatedOn = 0
    [int]$processCount = 0

    if( $PSBoundParameters[ 'name' ] )
    {
        [int]$thisSessionId = Get-Process -Id $pid | Select-Object -ExpandProperty SessionId

        if( ! $thisSessionId -and ! $allSessions )
        {
            if( ! $quiet )
            {
                Throw "Failed to determine current session id from pid $pid"
            }
            else
            {
                Exit 1
            }
        }

        $id = @( Get-Process -Name $name -ErrorAction SilentlyContinue | Where-Object { $allSessions -or $_.SessionId -eq $thisSessionId } | Select-Object -ExpandProperty Id )

        if( ! $id -or ! $id.Count )
        {
            if( ! $quiet )
            {
                Throw "No processes found for process names $name"
            }
            else
            {
                Exit 2
            }
        }
    }
}

Process
{
    if( $id -and $id.Count )
    {
        $processCount += $id.Count
        $operatedOn += Invoke-PauseResumeProcess -id $id -resume $resume -trim $trim
    }
}

End
{
    Write-Verbose -Message "Successfully operated on $operatedOn processes out of $processCount"

    ## since only process that started debugging can stop it and debugged processes will exit if the debugger process exits, we give the option to wait to be signalled so we can stop debugging
    if( $operatedOn -and ! [string]::IsNullOrEmpty( $pipeName ) -and [string]::IsNullOrEmpty( $signal ) )
    {
        if( $resume )
        {
            Write-Warning -Message "No point waiting when -resume specified"
        }
        else
        {
            Try
            {
                if( $namedPipe = New-Object -TypeName IO.Pipes.NamedPipeClientStream( '.' , $pipeName,  'In' ) )
                {
                    Write-Verbose -Message "$(Get-Date -Format G): waiting on named pipe `"$pipeName`""
                    $namedPipe.Connect()

                    if( $reader = New-Object -TypeName IO.StreamReader( $namedPipe ) )
                    {
                        $readFromPipe = $reader.ReadLine()
                        Write-Verbose -Message "$(Get-Date -Format G): received `"$readFromPipe`" from named pipe `"$pipeName`""
                        $reader.Dispose()
                        $reader = $null
                    }
                    $namedPipe.Dispose()
                    $namedPipe = $null
                }
            }
            Catch
            {
                Write-Verbose -Message "In catch block : $_"
            }
            Finally
            {
                Write-Verbose -Message "Performing reversion before script exit"
                Invoke-PauseResumeProcess -id $id -resume (-not $resume)
            }
        }
    }

    If( ! [string]::IsNullOrEmpty( $logging ) )
    {
        Stop-Transcript
    }
}