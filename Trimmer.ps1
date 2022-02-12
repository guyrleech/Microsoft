#requires -version 3

<#
    Trim working sets of processes or set maximum working set sizes

    Guy Leech, 2018

    Modification History

    10/03/18  GL  Optimised code

    12/03/18  GL  Added reporting

    16/03/18  GL Fixed bug where -available was being passed as False when not specified
                 Workaround for invocation external to PowerShell for arrays being flattened
                 Made process ids parameter an array
                 Added process id and name filter to Get-Process cmdlet call for efficiency
                 Added include and exclude options for user names

    30/03/18  GL Added ability to wait for specific list of processes to start before continuing.
                 Added looping

    23/04/18  GL Added exiting from loop if pids specified no longer exist
                 Added -forceIt as equivalent to -confirm:$false for use via scheduled tasks

    12/02/22  GL Added output of whether hard workin set limits and -nogridview
#>

<#
.SYNOPSIS

Manipulate the working sets (memory usage) of processes or report their current memory usage and working set limits and types (hard or soft)

.DESCRIPTION

Can reduce the memory footprints of running processes to make more memory available and stop processes that leak memory from leaking

.PARAMETER Processes

A comma separated list of process names to use (without the .exe extension). By default all processes will be trimmed if the script has access to them.

.PARAMETER IncludeUsers

A comma separated of qualified names of process owners to include. Must be run as an admin for this to work. Specify domain or other qualifier, e.g. "NT AUTHORITY\SYSTEM'

.PARAMETER ExcludeUsers

A comma separated of qualified names of process owners to exclude. Must be run as an admin for this to work. Specify domain or other qualifier, e.g. "NT AUTHORITY\NETWORK SERVICE' or DOMAIN\Chris.Harvey

.PARAMETER Exclude

A comma separated list of process names to ignore (without the .exe extension).

.PARAMETER Above

Only trim the working set if the process' working set is currently above this value. Qualify with MB or GB as required. Default is to trim all processes

.PARAMETER WaitFor

A comma separated list of processes to wait for unless -alreadyStarted is specified and one of the processes is already running unless it is not in the current session and -thisSession specified

.PARAMETER AlreadyStarted

Used when -WaitFor specified such that waiting will not occur if any of the processes specified via -WaitFor are already running although only in the current session if -thisSession is specified

.PARAMETER PollPeriod

The time in seconds between checks for new processes that match the -WaitFor process list.

.PARAMETER MinWorkingSet

Set the minimum working set size to this value. Qualify with MB or GB as required. Default is to not set a minimum value.

.PARAMETER MaxWorkingSet

Set the maximum working set size to this value. Qualify with MB or GB as required. Default is to not set a maximum value.

.PARAMETER HardMin

When MinWorkingSet is specified, the limit will be enforced so the working set is never allowed to be less that the value. Default is a soft limit which is not enforced.

.PARAMETER HardMax

When MaxWorkingSet is specified, the limit will be enforced so the working set is never allowed to exceed the value. Default is a soft limit which can be exceeded.

.PARAMETER Loop

Loop infinitely

.PARAMETER forceIt

DO not prompt for confirmation before adjusting CPU priority

.PARAMETER Report

Produce a report of the current working set usage and limit types for processes in the selection. Will output to a grid view unless -outputFile is specified.

.PARAMETER OutputFile

Ue with -report to write the results to a csv format file. If the file already exists the operation will fail.

.PARAMETER ProcessIds

Only trim the specific process ids

.PARAMETER ThisSession

Will only trim working sets of processes in the same session as the sript is running in. The default is to trim in all sessions.

.PARAMETER SessionIds

Only trim processes running in the specified sessions which is a comma separated list of session ids. The default is to trim in all sessions.

.PARAMETER NotSessionId

Only trim processes not running in the specified sessions which is a comma separated list of session ids. The default is to trim in all sessions.

.PARAMETER Available

Specify as a percentage or an absolute value. Will only trim if the available memory is below the parameter specified. The default is to always trim.

.PARAMETER Savings

This will show a summary of the trimming at the end of processing. Note that working sets can grow once trimmed so the amount trimmed may be higher than the actual increase in available memory.

.PARAMETER Disconnected

This will only trim memory in sessions which are disconnected. The default is to target all sessions.

.PARAMETER Idle

If no user input has been received in the last x seconds, whre x is the parameter passed, then the session is considered idle and processes will be trimmed.

.PARAMETER nogridview

Put the results (use -report) onto the pipeline, not in a grid view

.PARAMETER Background

Only trim processes which are not the process responsible for the foreground window. Implies -ThisSession since cannot check windows in other sessions.

.PARAMETER Install

Create two scheduled tasks, one which will trim that user's session on disconnect or screen lock and the other runs at the frequency specified in seconds divided by two, checks if the user is idle for the specified number of seconds and trims if they are.
So if a parameter of 600 is passed, the task will run every 5 minutes and if the user has made no mouse or keyboard input for 10 minutes then their processes are trimmed.

.PARAMETER Uninstall

Removes the two scheduled tasks previously created for the user running the script.

.EXAMPLE

& .\Trimmer.ps1

This will trim all processes in all sessions to which the account running the script has access.

.EXAMPLE

& .\Trimmer.ps1 -ThisSession -Above 50MB

Only trim processes in the same session as the script and whose working set exceeds 50MB.

.EXAMPLE

& .\Trimmer.ps1 -MaxWorkingSet 100MB -HardMax -Above 50MB -Processes LeakyApp

Only trim processes called LeakyApp in any session whose working set exceeds 50MB and set the maximum working set size to 100MB which cannot be exceeded.
Will only apply to instances of LeakyApp which have already started, instances started after the script is run will not be subject to the restriction. 

.EXAMPLE

& .\Trimmer.ps1 -MaxWorkingSet 10MB -Processes Chrome

Trim Chrome processes to 10MB, rather than completely emptying their working set. If processes rapidly regain working sets after being trimmed, this can
cause page file thrashing so reducing the working set but not completely emptying them can still save memory but reduce the risk of page file thrashing.
Picking the figure to use for the working set is trial and error but typically one would use the value that it settles to a few minutes after trimming.

.EXAMPLE

& .\Trimmer.ps1 -Install 600 -Logoff

Create two scheduled tasks for this user which only run when the user is logged on. The first runs at session lock or disconnect and trims all processes in that session.
The second task runs every 300 seconds and if no mouse or keyboard input has been received in the last 600 seconds then all background processes in that session will be trimmed.
At logoff, the scheduled tasks will be removed.

.EXAMPLE

& .\Trimmer.ps1 Uninstall

Delete the two scheduled tasks for this user

.NOTES

If you trim too much and/or too frequently, you run the risk of reducing performance by overusing the page file.

Supports the "-whatif" parameter so you can see what processes it will trim without actually performing the trim.

If emptying the working set does cause too much paging, try using the -MaxWorkingSet parameter to apply a soft limit which will cause the process
to be trimmed down to that value but it can then grow larger if required.

Uses Windows API SetProcessWorkingSetSizeEx() - https://msdn.microsoft.com/en-us/library/windows/desktop/ms686237(v=vs.85).aspx
#>

[cmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='Low')]

Param
(
    [string]$logFile ,
    [int]$install ,
    [int]$idle ,  ## seconds
    [switch]$uninstall ,
    [string[]]$processes ,
    [string[]]$exclude , 
    [string[]]$includeUsers ,
    [string[]]$excludeUsers ,
    [string[]]$waitFor ,
    [switch]$alreadyStarted ,
    [switch]$report ,
    [switch]$nogridview ,
    [string]$outputFile ,
    [int]$above = 10MB ,
    [int]$minWorkingSet = -1 ,
    [int]$maxWorkingSet = -1 ,
    [int]$pollPeriod = 5 ,
    [switch]$hardMax ,
    [switch]$hardMin ,
    [switch]$newOnly ,
    [switch]$thisSession ,
    [string[]]$processIds ,
    [string[]]$sessionIds ,
    [string[]]$notSessionIds ,
    [string]$available ,
    [switch]$loop ,
    [switch]$savings ,
    [switch]$disconnected ,
    [switch]$background ,
    [switch]$scheduled ,
    [switch]$logoff ,
    [switch]$forceIt ,
    [string]$taskFolder = '\MemoryTrimming' 
)

[int]$minimumIdlePeriod = 120 ## where minimum reptition of a scheduled task must be at least 1 minute, thus idle time must be at least double that (https://msdn.microsoft.com/en-us/library/windows/desktop/aa382993(v=vs.85).aspx)

## Borrowed from http://stackoverflow.com/a/15846912 and adapted
Add-Type @'
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace PInvoke.Win32
{
  
    public static class Memory
    {
        [DllImport("kernel32.dll", SetLastError=true)]
        public static extern bool SetProcessWorkingSetSizeEx( IntPtr proc, int min, int max , int flags );
        [DllImport("kernel32.dll", SetLastError=true)]
        public static extern bool GetProcessWorkingSetSizeEx( IntPtr hProcess, ref int min, ref int max , ref int flags );
    }
    public static class UserInput
    {  
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();
        
        [DllImport("user32.dll")]
        public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
 
        [DllImport("user32.dll", SetLastError=false)]
        private static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

        [StructLayout(LayoutKind.Sequential)]
        private struct LASTINPUTINFO
        {
            public uint cbSize;
            public int dwTime;
        }
        public static DateTime LastInput
        {
            get
            {
                DateTime bootTime = DateTime.UtcNow.AddMilliseconds(-Environment.TickCount);
                DateTime lastInput = bootTime.AddMilliseconds(LastInputTicks);
                return lastInput;
            }
        }
        public static TimeSpan IdleTime
        {
            get
            {
                return DateTime.UtcNow.Subtract(LastInput);
            }
        }
        public static int LastInputTicks
        {
            get
            {
                LASTINPUTINFO lii = new LASTINPUTINFO();
                lii.cbSize = (uint)Marshal.SizeOf(typeof(LASTINPUTINFO));
                GetLastInputInfo(ref lii);
                return lii.dwTime;
            }
        }
    }
}
'@

Function Schedule-Task
{
    Param
    (
        [string]$taskFolder ,
        [string]$taskname , 
        [string]$script ,  ## if null then we are deleting
        [int]$idle ,
        [switch]$background ,
        [int]$above ,
        [switch]$savings ,
        [string]$available = $null,
        [string[]]$processes ,
        [string[]]$exclude ,
        [string]$logFile = $null
    )

    Write-Verbose "Schedule-Task( $taskFolder , $taskName , $script )"

    ## https://www.experts-exchange.com/articles/11591/VBScript-and-Task-Scheduler-2-0-Creating-Scheduled-Tasks.html

    Set-Variable TASK_LOGON_INTERACTIVE_TOKEN      3   #-Option Constant
    Set-Variable TASK_RUNLEVEL_LUA                 0   #-Option Constant
    Set-Variable TASK_TRIGGER_EVENT                0   #-Option Constant
    Set-Variable TASK_TRIGGER_TIME                 1
    Set-Variable TASK_TRIGGER_DAILY                2
    Set-Variable TASK_TRIGGER_IDLE                 6
    Set-Variable TASK_TRIGGER_SESSION_STATE_CHANGE 11  #-Option Constant
    Set-Variable TASK_STATE_SESSION_LOCK           7   #-Option Constant
    Set-Variable TASK_STATE_REMOTE_DISCONNECT      4
    Set-Variable TASK_ACTION_EXEC                  0   #-Option Constant
    Set-Variable TASK_CREATE_OR_UPDATE             6   #-Option Constant

    $objTaskService  = New-Object -ComObject "Schedule.Service" ##-Strict
    $objTaskService.Connect()

    $objRootFolder = $objTaskService.GetFolder("\")

    $objTaskFolders = $objRootFolder.GetFolders(0)

    [bool]$blnFoundTask = $false

    ForEach( $objTaskFolder In $objTaskFolders )
    {
	    If( $objTaskFolder.Path -eq $taskFolder )
        {
		    $blnFoundTask = $True
		    break
	    }
    }

    if( [string]::IsNullOrEmpty( $script ) )
    {
        ## Find task and delete
        if( $blnFoundTask )
        {
            [bool]$deleted = $false
            $objTaskFolder.GetTasks(0) | ?{ $_.Name -eq $taskname } | %{ $objTaskFolder.DeleteTask( $_.Name , 0 ) ; $deleted = $true }
            if( ! $deleted )
            {
                Write-Warning "Failed to find task `"$taskname`" so cannot remove it"
            }
        }
        else
        {
            Write-Warning "Unable to find task folder $taskFolder so cannot remove scheduled tasks"
        }
        return
    }
    elseif( ! $blnFoundTask )
    {
        $objTaskFolder = $objRootFolder.CreateFolder($taskFolder)
    }

    $objNewTaskDefinition = $objTaskService.NewTask(0) 

    $objNewTaskDefinition.Data = 'This is Guys task from PoSH'

    $objNewTaskDefinition.RegistrationInfo.Author = $objTaskService.ConnectedDomain  + "\" + $objTaskService.ConnectedUser
    $objNewTaskDefinition.RegistrationInfo.Date = ([datetime]::Now).ToString("yyyy-MM-dd'T'HH:mm:ss")
    $objNewTaskDefinition.RegistrationInfo.Description = 'Trim process memory'
    $objNewTaskDefinition.RegistrationInfo.Documentation = 'RTFM'
    $objNewTaskDefinition.RegistrationInfo.Source = 'PowerShell'
    $objNewTaskDefinition.RegistrationInfo.URI = 'http://guyrleech.wordpress.com'
    $objNewTaskDefinition.RegistrationInfo.Version = '1.0'

    $objNewTaskDefinition.Principal.Id = 'My ID'
    $objNewTaskDefinition.Principal.DisplayName = 'Principal Description'
    $objNewTaskDefinition.Principal.UserId = $objTaskService.ConnectedDomain  + "\" + $objTaskService.ConnectedUser
    $objNewTaskDefinition.Principal.LogonType = $TASK_LOGON_INTERACTIVE_TOKEN
    $objNewTaskDefinition.Principal.RunLevel = $TASK_RUNLEVEL_LUA

    $objTaskTriggers = $objNewTaskDefinition.Triggers
    
    $objTaskAction = $objNewTaskDefinition.Actions.Create($TASK_ACTION_EXEC)
    $objTaskAction.Id = 'Execute Action'
    ## powershell.exe even with windowstyle hidden still shows a window so we start via vbs in order for it to be truly hidden
$vbsscriptbody = @"
Dim objShell,strArgs , i
for i = 0 to WScript.Arguments.length - 1
    strArgs = strArgs & WScript.Arguments(i) & " "
next
Set objShell = CreateObject("WScript.Shell")
objShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -File ""$script"" -ThisSession -scheduled " & strArgs , 0

"@
   
    [string]$vbsscript = $script -replace '\.ps1$' , '.vbs'

    if( Test-Path $vbsscript )
    {
        [string]$content = ""
        $existingScript = Get-Content $vbsscript | %{ $content += $_ + "`r`n" }

        if( $content -ne $vbsscriptbody )
        {
            Write-Error "vbs script `"$vbsscript`" already exists but is different to the file we need to write"
        }
    }
    else
    {
        [io.file]::WriteAllText( $vbsscript , $vbsscriptbody ) ## ensure no newline as breaks comparison
        if( ! $? -or ! ( Test-Path $vbsscript ) )
        {
            Write-Error "Error creating vbs script `"$vbsscript`""
        }
    }
    
    $objTaskAction.WorkingDirectory = $env:TEMP
    $objTaskAction.Path = 'wscript.exe'
    $objTaskAction.Arguments = "//nologo `"$vbsscript`""

    if( $idle -gt 0 )
    {
        $objTaskAction.Arguments += " -Idle $install"
    }
    if( $background )
    {
        $objTaskAction.Arguments += ' -background'
    }
    if( $above -ge 0 )
    {
        $objTaskAction.Arguments += " -above $above"
    }
    if( $savings )
    {
        $objTaskAction.Arguments += " -savings"
    }
    if( ! [string]::IsNullOrEmpty( $available ) )
    {
        $objTaskAction.Arguments += " -available $available"
    }
    if( ! [string]::IsNullOrEmpty( $logFile ) )
    {
        $objTaskAction.Arguments += " -logfile `"$logfile`""
    }
    if( $processes -and $processes.Count )
    {
        $objTaskAction.Arguments += " -processes $processes"
    }
    if( $exclude -and $exclude.Count )
    {
        $objTaskAction.Arguments += " -exclude $exclude"
    }  
    if( $VerbosePreference -eq 'Continue' )
    {
        $objTaskAction.Arguments += " -verbose"
    }

    ## http://msdn.microsoft.com/en-us/library/windows/desktop/aa383480%28v=vs.85%29.aspx
    $objNewTaskDefinition.Settings.Enabled = $true
    $objNewTaskDefinition.Settings.Compatibility = 2 ## Win7/WS08R2
    $objNewTaskDefinition.Settings.Priority = 5 ## 0 High - 10 Low
    $objNewTaskDefinition.Settings.Hidden = $false
    
    ## Can't use idle trigger as means more than just no input from user so we run a standard, repeating scheduled task and check for no input in the script itself
    if( $idle -gt 0 )
    {
        $objTaskTrigger = $objTaskTriggers.Create($TASK_TRIGGER_DAILY)
        $objTaskTrigger.Enabled = $true
        $objTaskTrigger.DaysInterval = 1
        $objTaskTrigger.Repetition.Duration = 'P1D'
        $objTaskTrigger.Repetition.Interval = 'PT' + [math]::Round( $idle / 2 ) + 'S'
        $objTaskTrigger.Repetition.StopAtDurationEnd = $true
        $objTaskTrigger.StartBoundary = ([datetime]::Now).ToString('yyyy-MM-dd''T''HH:mm:ss')
    }
    else
    {
        $objTaskTrigger = $objTaskTriggers.Create($TASK_TRIGGER_SESSION_STATE_CHANGE)
        $objTaskTrigger.Enabled = $true
        $objTaskTrigger.Id = 'Session state change lock'
        $objTaskTrigger.StateChange = $TASK_STATE_SESSION_LOCK

        ## Format For Days = P#D where # is the number of days
        ## Format for Time = PT#[HMS] Where # is the duration and H for hours, M for minutes, S for seconds
        $objTaskTrigger.ExecutionTimeLimit = 'PT5M'
        $objTaskTrigger.Delay = 'PT5S'
        $objTaskTrigger.UserId = $objTaskService.ConnectedDomain  + '\' + $objTaskService.ConnectedUser

        ## http://msdn.microsoft.com/en-us/library/windows/desktop/aa382144%28v=vs.85%29.aspx

        $objTaskTrigger = $objTaskTriggers.Create($TASK_TRIGGER_SESSION_STATE_CHANGE)
        $objTaskTrigger.Enabled = $true
        $objTaskTrigger.Id = 'Session state change disconnect'
        $objTaskTrigger.StateChange = $TASK_STATE_REMOTE_DISCONNECT

        ## Format For Days = P#D where # is the number of days
        ## Format for Time = PT#[HMS] Where # is the duration and H for hours, M for minutes, S for seconds
        $objTaskTrigger.ExecutionTimeLimit = "PT5M"
        $objTaskTrigger.Delay = "PT5S"
        $objTaskTrigger.UserId = $objTaskService.ConnectedDomain  + '\' + $objTaskService.ConnectedUser
    }

    $objNewTaskDefinition.Settings.DisallowStartIfOnBatteries = $false
    $objNewTaskDefinition.Settings.AllowDemandStart = $true
    $objNewTaskDefinition.Settings.StartWhenAvailable = $true
    $objNewTaskDefinition.Settings.RestartInterval = 'PT10M'
    $objNewTaskDefinition.Settings.RestartCount = 2
    $objNewTaskDefinition.Settings.ExecutionTimeLimit = "PT1H"
    $objNewTaskDefinition.Settings.AllowHardTerminate = $true
    ## 0 = Run a second instance now (Parallel)
    ## 1 = Put the new instance in line behind the current running instance (Add To Queue)
    ## 2 = Ignore the new request
    $objNewTaskDefinition.Settings.MultipleInstances = 2

    try
    {
        $task = $objTaskFolder.RegisterTaskDefinition( $taskname , $objNewTaskDefinition , $TASK_CREATE_OR_UPDATE , $null , $null , $TASK_LOGON_INTERACTIVE_TOKEN )
    }
    catch
    {
        $task = $null
    }

    if( ! $task )
    {
        Write-Error ( "Failed to create scheduled task: {0}" -f $error[0] )
    }
}

if( ! [string]::IsNullOrEmpty( $logFile ) )
{
    Start-Transcript $logFile -Append
}

if( $install -gt 0 -or $uninstall )
{
    Write-Verbose ( "{0} requested" -f $( if( $uninstall ) { "Uninstall" } else {"Install"} )  )

    if( $uninstall -and $install -gt 0 )
    {
        Write-Error "Cannot specify -install and -uninstall together"
        return 1
    }
    elseif( $report )
    {
        Write-Error "Cannot specify -install or -uninstall with -report"
    }
    elseif( $uninstall )
    {
        $scriptName = $null
    }
    elseif( $install -lt $minimumIdlePeriod ) ## minimum repetition is 1 minute see https://msdn.microsoft.com/en-us/library/windows/desktop/aa382993(v=vs.85).aspx
    {
        Write-Error "Idle time is too low - minimum idle time is $minimumIdlePeriod seconds"
        return
    }
    else
    {
        $scriptName = & { $myInvocation.ScriptName }
    }

    [hashtable]$taskArguments =
    @{
        Taskfolder = $taskFolder 
        Script = $scriptName
        Above = $above 
        Savings = $savings
        Exclude = $exclude
        Processes = $processes
        Logfile = $logFile
        Available = $available
    }

    Schedule-Task -taskName "Trim on lock and disconnect for $($env:username)" @taskArguments
    Schedule-Task -taskName "Trim idle for $($env:username)" @taskArguments -idle $install -background $background

    ## if we have been asked to hook logoff, to uninstall, then we create a hidden window so we can capture events
    if( $logoff )
    {
        Add-Type –AssemblyName System.Windows.Forms 

        $form = New-Object Windows.Forms.Form
        $form.Size = New-Object System.Drawing.Size(0,0)
        $form.Location = New-Object System.Drawing.Point(-5000,-5000)
        $form.FormBorderStyle = 'FixedToolWindow'
        $form.StartPosition = 'manual'
        $form.ShowInTaskbar = $false
        $form.WindowState = 'Normal'
        $form.Visible = $false
	    $form.AutoSize = $false

        $form.add_FormClosing(
            {
               Write-Verbose "$(Get-Date) dialog closing" 
               Schedule-Task -taskName "Trim on lock and disconnect for $($env:username)" -taskFolder $taskFolder -Script $null 
               Schedule-Task -taskName "Trim idle for $($env:username)" -taskFolder $taskFolder -Script $null
            })

        $form.add_Load({ $form.Opacity = 0 })

        $form.add_Shown({ $form.Opacity = 100 })

        Write-Verbose "About to show dialog for logoff intercept - script will not exit until logoff"
        [void]$form.ShowDialog()
        ## We will only get here when the hidden dialogue exits which should only be logoff
    }

    if( ! [string]::IsNullOrEmpty( $logFile ) )
    {
        Stop-Transcript
    }

    Exit 0
}

[datetime]$monitoringStartTime = Get-Date
[int]$thisSessionId = (Get-Process -Id $pid).SessionId

## workaround for scheduled task not liking -confirm:$false being passed
if( $forceIt )
{
     $ConfirmPreference = 'None'
}

do
{
    if( $waitFor -and $waitFor.Count )
    {
        [bool]$found = $false
        $thisProcess = $null

        [datetime]$startedAfter = Get-Date
        if( $alreadyStarted ) ## we are not waiting for new instances so existing ones qualify too
        {
            $startedAfter = Get-Date -Date '01/01/1970' ## saves having to grab LastBootupTime
        }

        while( ! $found )
        {
            Write-Verbose "$(Get-Date): waiting for one of $($waitFor -join ',') to launch (only in session $thisSessionId is $thisSession)"
            ## wait for one of a set of specific processes to start - useful when you need to apply a hard working set limit to a known leaky process
            Get-Process -Name $waitFor -ErrorAction SilentlyContinue | Where-Object { $_.StartTime -gt $startedAfter } | ForEach-Object `
            {
                if( ! $found )
                {
                    $thisProcess = $_
                    ## we don't support all filtering options here
                    if( $thisSession )
                    {
                        $found = ( $thisSessionId -eq $thisProcess.SessionId )
                    }
                    else
                    {
                        $found = $true
                    }
                }
            }
            if( ! $found )
            {
                Start-Sleep -Seconds $pollPeriod
            }
        }
        Write-Verbose "$(Get-Date) : process $($thisProcess.Name) id $($thisProcess.Id) started at $($thisProcess.StartTime)"
    }
    
    if( $idle -gt 0 )
    {
        $idleTime = [PInvoke.Win32.UserInput]::IdleTime.TotalSeconds
        Write-Verbose "Idle time is $idleTime seconds"
        if( $idleTime -lt $idle )
        {
            Write-Verbose "Idle time is only $idleTime seconds, less than $idle"
            if( ! $scheduled -or ( $scheduled -and ! $background ) )
            {
                if( ! [string]::IsNullOrEmpty( $logFile ) )
                {
                    Stop-Transcript
                }
                return
            }
            else
            {
                Write-Verbose "Not idle but we are a scheduled task and trimming background processes so continue"
            }
        }
    }

    [long]$ActiveHandle = $null
    $activePid = [IntPtr]::Zero

    if( $background )
    {
        [long]$ActiveHandle = [PInvoke.Win32.UserInput]::GetForeGroundWindow( )
        if( ! $ActiveHandle )
        {
            Write-Error "Unable to find foreground window"
            return 1
        }
        else
        {
            $activeThreadId = [PInvoke.Win32.UserInput]::GetWindowThreadProcessId( $ActiveHandle , [ref] $activePid )
            if( $activePid -ne [IntPtr]::Zero )
            {
                Write-Verbose ( "Foreground window is pid {0} {1}" -f $activePid , (Get-Process -Id $activePid).Name )
            }
            else
            {
                Write-Error "Unable to get handle on process for foreground window $ActiveHandle"
                return 1
            }
        }
        $thisSession = $true ## can only check windows in this session
    }

    [int]$flags = 0
    if( $minWorkingSet -gt 0 )
    {
        if( $hardMin )
        {
            $flags = $flags -bor 1
        }
        else  ## soft
        {
            $flags = $flags -bor 2
        }
    }

    if( $maxWorkingSet -gt 0 )
    {
        if( $hardMax )
        {
            $flags = $flags -bor 4
        }
        else  ## soft
        {
            $flags = $flags -bor 8
        }
        if( $minWorkingSet -le 0 )
        {
            $minWorkingSet = 1 ## if a maximum is specified then we must specify a minimum too - this will default to the minimum
        }
    }

    [long]$availableMemory = (Get-Counter '\Memory\Available MBytes').CounterSamples[0].CookedValue * 1MB

    if( ! [string]::IsNullOrEmpty( $available ) )
    {
        ## Need to find out memory available and total
        [long]$totalMemory = ( Get-CimInstance -Class Win32_ComputerSystem -Property TotalPhysicalMemory ).TotalPhysicalMemory
        [int]$left = ( $availableMemory / $totalMemory ) * 100
        Write-Verbose ( "Available memory is {0}MB out of {1}MB total ({2}%)" -f ( $availableMemory / 1MB ) , [math]::Floor( $totalMemory / 1MB ) , [math]::Round( $left ) )

        [bool]$proceed = $false
        ## See if we are dealing with absolute or percentage
        if( $available[-1] -eq '%' )
        {
            [int]$percentage = $available -replace '%$'
            $proceed = $left -lt $percentage 
        }
        else ## absolute
        {
            [long]$threshold = Invoke-Expression $available
            $proceed = $availableMemory -lt $threshold
        }

        if( ! $proceed )
        {
            Write-Verbose "Not trimming as memory available is above specified threshold"
            if( ! [string]::IsNullOrEmpty( $logFile ) )
            {
                Stop-Transcript
            }
            Exit 0
        }
    }

    [long]$saved = 0
    [int]$trimmed = 0

    $params = @{}

    [int[]]$sessionsToTarget = @()
    $results = New-Object -TypeName System.Collections.ArrayList

    if( $disconnected ) 
    {
        ## no native session support so parse output of quser.exe
        ## Columns are 'USERNAME              SESSIONNAME        ID  STATE   IDLE TIME  LOGON TIME' but SESSIONNAME is empty for disconnected so all shifted left by one column (yuck!)

        $sessionsToTarget = @( (quser) -replace '\s{2,}', ',' | ConvertFrom-Csv | ForEach-Object ` 
        {
            $session = $_
            if( $session.Id -like "Disc*" )
            {
                $session.SessionName -as [int]
                Write-Verbose ( "Session {0} is disconnected for user {1} logon {2} idle {3}" -f $session.SessionName , $session.Username , $session.'Idle Time' , $session.State )
            }
        } )
    }

    ## Reform arrays as they will not be passed correctly if command not invoked natively in PowerShell, e.g. via cmd or scheduled task
    if( $processes )
    {
        if( $processes.Count -eq 1 -and $processes[0].IndexOf(',') -ge 0 )
        {
            $processes = $processes -split ','
        }
        $params.Add( 'Name' , $processes )
    }

    if( $processIds )
    {
        if( $processIds.Count -eq 1 -and $processIds[0].IndexOf(',') -ge 0 )
        {
            $processIds = $processIds -split ','
        }
        $params.Add( 'Id' , $processIds )
    }

    if( $includeUsers -or $excludeUsers )
    {
        $params.Add( 'IncludeUserName' , $true ) ## Needs admin rights
        if( $includeUsers.Count -eq 1 -and $includeUsers[0].IndexOf(',') -ge 0 )
        {
            $includeUsers = $includeUsers -split ','
        }
        if( $excludeUsers.Count -eq 1 -and $excludeUsers[0].IndexOf(',') -ge 0 )
        {
            $excludeUsers = $excludeUsers -split ','
        }
    }

    if( $exclude -and $exclude.Count -eq 1 -and $exclude[0].IndexOf(',') -ge 0 )
    {
        $exclude = $exclude -split ','
    }

    if( $sessionIds -and $sessionIds.Count -eq 1 -and $sessionIds[0].IndexOf(',') -ge 0 )
    {
        $sessionIds = $sessionIds -split ','
    }

    if( $notSessionIds -and $notSessionIds.Count -eq 1 -and $notSessionIds[0].IndexOf(',') -ge 0 )
    {
        $notSessionIds = $notSessionIds -split ','
    }

    [int]$adjusted = 0

    Get-Process @params -ErrorAction SilentlyContinue | ForEach-Object `
    {      
        $process = $_
        [bool]$doIt = $true
        if( $excludeUsers -and $excludeUsers.Count -And $excludeUsers -contains $process.UserName )
        {
            Write-Verbose ( "`tSkipping {0} pid {1} for user {2} as specifically excluded" -f $process.Name , $process.Id , $process.UserName )
            $doIt = $false
        }
        elseif( $doIt -and $includeUsers -and $includeUsers.Count -And $includeUsers -notcontains $process.UserName )
        {
            Write-Verbose ( "`tSkipping {0} pid {1} for user {2} as not included" -f $process.Name , $process.Id , $process.UserName )
            $doIt = $false
        }
        elseif( $doIt -and $exclude -and $exclude.Count -And $exclude -contains $process.Name )
        {
            Write-Verbose ( "`tSkipping {0} pid {1} as specifically excluded" -f $process.Name , $process.Id )
            $doIt = $false
        }
        elseif( $doIt -and $thisSession -And $process.SessionId -ne $thisSessionId )
        {
            Write-Verbose ( "`tSkipping {0} pid {1} as session {2} not {3}" -f $process.Name , $process.Id , $process.SessionId , $thisSessionId )
            $doIt = $false
        }
        elseif( $doIt -and $process.Id -eq $activePid -And $idle -eq 0 ) ## if idle then we'll trim anyway as not being used (will have quit already if not idle if idle parameter specified)
        {
            Write-Verbose ( "`tSkipping {0} pid {1} as it is the foreground window process" -f $process.Name , $process.Id )
            $doIt = $false
        }
        elseif( $doIt -and $sessionIds -and $sessionIds.Count -gt 0 -And $sessionIds -notcontains $process.SessionId.ToString() )
        {
            Write-Verbose ( "`tSkipping {0} pid {1} as session {2} not in list" -f $process.Name , $process.Id , $process.SessionId )
            $doIt = $false
        }
        elseif( $notsessionIds -and $notSessionIds.Count -gt 0 -And $notSessionIds -contains $process.SessionId.ToString() )
        {
            Write-Verbose ( "`tSkipping {0} pid {1} as session {2} is specifically excluded" -f $process.Name , $process.Id , $process.SessionId )
            $doIt = $false
        }
        elseif( $doIt -and $sessionsToTarget.Count -gt 0 -And $sessionsToTarget -notcontains $process.SessionId )
        {
            Write-Verbose ( "`tSkipping {0} pid {1} as session {2} which is not disconnected" -f $process.Name , $process.Id , $process.SessionId )
            $doIt = $false
        }
        elseif( $doIt -and $above -gt 0 -And $process.WS -le $above )
        {
            Write-Verbose ( "`tSkipping {0} pid {1} as working set only {2} MB" -f $process.Name , $process.Id , [Math]::Round( $process.WS / 1MB , 1 ) )
            $doIt = $false
        }
        elseif( $doIt -and $newOnly -and $process.StartTime -lt $monitoringStartTime )
        {
            Write-Verbose ( "`tSkipping {0} pid {1} as start time {2} prior to {3}" -f $process.Name , $process.Id , $process.StartTime , $monitoringStartTime )
            $doit = $false
        }

        if( $doIt )
        {
            $action = "Process {0} pid {1} session {2} working set {3} MB" -f $process.Name , $process.Id , $process.SessionId , [Math]::Floor( $process.WS / 1MB )

            if( $process.Handle )
            {
                if( $report )
                {
                    [int]$thisMinimumWorkingSet = -1 
                    [int]$thisMaximumWorkingSet = -1 
                    [int]$thisFlags = -1 ## Grammar alert! :-)
                    ## https://msdn.microsoft.com/en-us/library/windows/desktop/ms683227(v=vs.85).aspx
                    [bool]$result = [PInvoke.Win32.Memory]::GetProcessWorkingSetSizeEx( $process.Handle, [ref]$thisMinimumWorkingSet,[ref]$thisMaximumWorkingSet,[ref]$thisFlags);$LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
                    if( $result )
                    {
                        ## convert flags value - if not hard then will be soft so no point reporting that separately IMHO
                        [bool]$hardMinimumWorkingSet = $thisFlags -band 1 ## QUOTA_LIMITS_HARDWS_MIN_ENABLE
                        [bool]$hardMaximumWorkingSet = $thisFlags -band 4 ## QUOTA_LIMITS_HARDWS_MAX_ENABLE
                        $null = $results.Add( ([pscustomobject][ordered]@{ 'Name' = $process.Name ; 'PID' = $process.Id ; 'Handle Count' = $process.HandleCount ; 'Start Time' = $process.StartTime ;
                                'Hard Minimum Working Set Limit' = $hardMinimumWorkingSet ; 'Hard Maximum Working Set Limit' = $hardMaximumWorkingSet ;
                                'Working Set (MB)' = $process.WorkingSet64 / 1MB ;'Peak Working Set (MB)' = $process.PeakWorkingSet64 / 1MB ;
                                'Commit Size (MB)' = $process.PagedMemorySize / 1MB; 
                                'Paged Pool Memory Size (KB)' = $process.PagedSystemMemorySize64 / 1KB ; 'Non-paged Pool Memory Size (KB)' = $process.NonpagedSystemMemorySize64 / 1KB ;
                                'Minimum Working Set (KB)' = $thisMinimumWorkingSet / 1KB ; 'Maximum Working Set (KB)' = $thisMaximumWorkingSet / 1KB
                                'Hard Minimum Working Set' = $hardMinimumWorkingSet ; 'Hard Maximum Working Set' = $hardMaximumWorkingSet 
                                'Virtual Memory Size (GB)' = $process.VirtualMemorySize64 / 1GB; 'Peak Virtual Memory Size (GB)' = $process.PeakVirtualMemorySize64 / 1GB; }) )
                    }
                    else
                    {                   
                        Write-Warning ( "Failed to get working set info for {0} pid {1} - {2} " -f $process.Name , $process.Id , $LastError)
                    }
                }
                elseif( $pscmdlet.ShouldProcess( $action , 'Trim' ) ) ## Handle may be null if we don't have sufficient privileges to that process
                {
                    ## see https://msdn.microsoft.com/en-us/library/windows/desktop/ms686237(v=vs.85).aspx
                    [bool]$result = [PInvoke.Win32.Memory]::SetProcessWorkingSetSizeEx( $process.Handle,$minWorkingSet,$maxWorkingSet,$flags);$LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()

                    $adjusted++
                    if( ! $result )
                    {
                        Write-Warning ( "Failed to trim {0} pid {1} - {2} " -f $process.Name , $process.Id , $LastError)
                    }
                    elseif( $savings )
                    {
                        $now = Get-Process -Id $process.Id -ErrorAction SilentlyContinue
                        if( $now )
                        {
                            $saved += $process.WS - $now.WS
                            $trimmed++
                        }
                    }
                }
            }
            else
            {
                Write-Warning ( "No handle on process {0} pid {1} working set {2} MB so cannot access working set" -f $process.Name , $process.Id , [Math]::Floor( $process.WS / 1MB ) )
            }
        }
    }

    if( $report )
    {
        if( [string]::IsNullOrEmpty( $outputFile ) )
        {
            if( -Not $nogridview )
            {
                $selected = $results | Sort-Object Name | Out-GridView -PassThru -Title "Memory information from $($results.Count) processes at $(Get-Date -Format U)"
                if( $selected )
                {
                    $selected | clip.exe
                }
            }
            else
            {
                $results
            }
        }
        else
        {
            $results | Sort-Object Name | Export-Csv -Path $outputFile -NoTypeInformation -NoClobber
        }
    }

    if( $savings )
    {
        [long]$availableMemoryAfter = (Get-Counter '\Memory\Available MBytes').CounterSamples[0].CookedValue
        Write-Output ( "Trimmed {0}MB from {1} processes giving {2}MB extra available" -f [math]::Round( $saved / 1MB , 1 ) , $trimmed , ( $availableMemoryAfter - ( $availableMemory / 1MB ) ) )
    }
    if( $loop )
    {
        if( $processIds -and $processIds.Count -and ! $adjusted )
        {
            Write-Warning "None of the specified pids $($processIds -join ', ') were found or were not included or were excluded so exiting loop"
            $loop = $false
        }
        else
        {
            Write-Verbose "$(Get-Date) : sleeping for $pollPeriod seconds before looping"
            Start-Sleep -Seconds $pollPeriod
        }
    }
} while( $loop )

if( ! [string]::IsNullOrEmpty( $logFile ) )
{
    Stop-Transcript
}

# SIG # Begin signature block
# MIIZsAYJKoZIhvcNAQcCoIIZoTCCGZ0CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUkPd4kKxwEdoyJhqY8Z79bC+H
# PJ6gghS+MIIE/jCCA+agAwIBAgIQDUJK4L46iP9gQCHOFADw3TANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgVGltZXN0YW1waW5nIENBMB4XDTIxMDEwMTAwMDAwMFoXDTMxMDEw
# NjAwMDAwMFowSDELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMu
# MSAwHgYDVQQDExdEaWdpQ2VydCBUaW1lc3RhbXAgMjAyMTCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBAMLmYYRnxYr1DQikRcpja1HXOhFCvQp1dU2UtAxQ
# tSYQ/h3Ib5FrDJbnGlxI70Tlv5thzRWRYlq4/2cLnGP9NmqB+in43Stwhd4CGPN4
# bbx9+cdtCT2+anaH6Yq9+IRdHnbJ5MZ2djpT0dHTWjaPxqPhLxs6t2HWc+xObTOK
# fF1FLUuxUOZBOjdWhtyTI433UCXoZObd048vV7WHIOsOjizVI9r0TXhG4wODMSlK
# XAwxikqMiMX3MFr5FK8VX2xDSQn9JiNT9o1j6BqrW7EdMMKbaYK02/xWVLwfoYer
# vnpbCiAvSwnJlaeNsvrWY4tOpXIc7p96AXP4Gdb+DUmEvQECAwEAAaOCAbgwggG0
# MA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsG
# AQUFBwMIMEEGA1UdIAQ6MDgwNgYJYIZIAYb9bAcBMCkwJwYIKwYBBQUHAgEWG2h0
# dHA6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAfBgNVHSMEGDAWgBT0tuEgHf4prtLk
# YaWyoiWyyBc1bjAdBgNVHQ4EFgQUNkSGjqS6sGa+vCgtHUQ23eNqerwwcQYDVR0f
# BGowaDAyoDCgLoYsaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJl
# ZC10cy5jcmwwMqAwoC6GLGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9zaGEyLWFz
# c3VyZWQtdHMuY3JsMIGFBggrBgEFBQcBAQR5MHcwJAYIKwYBBQUHMAGGGGh0dHA6
# Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBPBggrBgEFBQcwAoZDaHR0cDovL2NhY2VydHMu
# ZGlnaWNlcnQuY29tL0RpZ2lDZXJ0U0hBMkFzc3VyZWRJRFRpbWVzdGFtcGluZ0NB
# LmNydDANBgkqhkiG9w0BAQsFAAOCAQEASBzctemaI7znGucgDo5nRv1CclF0CiNH
# o6uS0iXEcFm+FKDlJ4GlTRQVGQd58NEEw4bZO73+RAJmTe1ppA/2uHDPYuj1UUp4
# eTZ6J7fz51Kfk6ftQ55757TdQSKJ+4eiRgNO/PT+t2R3Y18jUmmDgvoaU+2QzI2h
# F3MN9PNlOXBL85zWenvaDLw9MtAby/Vh/HUIAHa8gQ74wOFcz8QRcucbZEnYIpp1
# FUL1LTI4gdr0YKK6tFL7XOBhJCVPst/JKahzQ1HavWPWH1ub9y4bTxMd90oNcX6X
# t/Q/hOvB46NJofrOp79Wz7pZdmGJX36ntI5nePk2mOHLKNpbh6aKLzCCBTAwggQY
# oAMCAQICEAQJGBtf1btmdVNDtW+VUAgwDQYJKoZIhvcNAQELBQAwZTELMAkGA1UE
# BhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2lj
# ZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4X
# DTEzMTAyMjEyMDAwMFoXDTI4MTAyMjEyMDAwMFowcjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEx
# MC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBD
# QTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAPjTsxx/DhGvZ3cH0wsx
# SRnP0PtFmbE620T1f+Wondsy13Hqdp0FLreP+pJDwKX5idQ3Gde2qvCchqXYJawO
# eSg6funRZ9PG+yknx9N7I5TkkSOWkHeC+aGEI2YSVDNQdLEoJrskacLCUvIUZ4qJ
# RdQtoaPpiCwgla4cSocI3wz14k1gGL6qxLKucDFmM3E+rHCiq85/6XzLkqHlOzEc
# z+ryCuRXu0q16XTmK/5sy350OTYNkO/ktU6kqepqCquE86xnTrXE94zRICUj6whk
# PlKWwfIPEvTFjg/BougsUfdzvL2FsWKDc0GCB+Q4i2pzINAPZHM8np+mM6n9Gd8l
# k9ECAwEAAaOCAc0wggHJMBIGA1UdEwEB/wQIMAYBAf8CAQAwDgYDVR0PAQH/BAQD
# AgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHkGCCsGAQUFBwEBBG0wazAkBggrBgEF
# BQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEMGCCsGAQUFBzAChjdodHRw
# Oi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0Eu
# Y3J0MIGBBgNVHR8EejB4MDqgOKA2hjRodHRwOi8vY3JsNC5kaWdpY2VydC5jb20v
# RGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMDqgOKA2hjRodHRwOi8vY3JsMy5k
# aWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsME8GA1UdIARI
# MEYwOAYKYIZIAYb9bAACBDAqMCgGCCsGAQUFBwIBFhxodHRwczovL3d3dy5kaWdp
# Y2VydC5jb20vQ1BTMAoGCGCGSAGG/WwDMB0GA1UdDgQWBBRaxLl7KgqjpepxA8Bg
# +S32ZXUOWDAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzANBgkqhkiG
# 9w0BAQsFAAOCAQEAPuwNWiSz8yLRFcgsfCUpdqgdXRwtOhrE7zBh134LYP3DPQ/E
# r4v97yrfIFU3sOH20ZJ1D1G0bqWOWuJeJIFOEKTuP3GOYw4TS63XX0R58zYUBor3
# nEZOXP+QsRsHDpEV+7qvtVHCjSSuJMbHJyqhKSgaOnEoAjwukaPAJRHinBRHoXpo
# aK+bp1wgXNlxsQyPu6j4xRJon89Ay0BEpRPw5mQMJQhCMrI2iiQC/i9yfhzXSUWW
# 6Fkd6fp0ZGuy62ZD2rOwjNXpDd32ASDOmTFjPQgaGLOBm0/GkxAG/AeB+ova+YJJ
# 92JuoVP6EpQYhS6SkepobEQysmah5xikmmRR7zCCBTEwggQZoAMCAQICEAqhJdbW
# Mht+QeQF2jaXwhUwDQYJKoZIhvcNAQELBQAwZTELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIG
# A1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4XDTE2MDEwNzEyMDAw
# MFoXDTMxMDEwNzEyMDAwMFowcjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lD
# ZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGln
# aUNlcnQgU0hBMiBBc3N1cmVkIElEIFRpbWVzdGFtcGluZyBDQTCCASIwDQYJKoZI
# hvcNAQEBBQADggEPADCCAQoCggEBAL3QMu5LzY9/3am6gpnFOVQoV7YjSsQOB0Uz
# URB90Pl9TWh+57ag9I2ziOSXv2MhkJi/E7xX08PhfgjWahQAOPcuHjvuzKb2Mln+
# X2U/4Jvr40ZHBhpVfgsnfsCi9aDg3iI/Dv9+lfvzo7oiPhisEeTwmQNtO4V8CdPu
# XciaC1TjqAlxa+DPIhAPdc9xck4Krd9AOly3UeGheRTGTSQjMF287DxgaqwvB8z9
# 8OpH2YhQXv1mblZhJymJhFHmgudGUP2UKiyn5HU+upgPhH+fMRTWrdXyZMt7HgXQ
# hBlyF/EXBu89zdZN7wZC/aJTKk+FHcQdPK/P2qwQ9d2srOlW/5MCAwEAAaOCAc4w
# ggHKMB0GA1UdDgQWBBT0tuEgHf4prtLkYaWyoiWyyBc1bjAfBgNVHSMEGDAWgBRF
# 66Kv9JLLgjEtUYunpyGd823IDzASBgNVHRMBAf8ECDAGAQH/AgEAMA4GA1UdDwEB
# /wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcDCDB5BggrBgEFBQcBAQRtMGswJAYI
# KwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3
# aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9v
# dENBLmNydDCBgQYDVR0fBHoweDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNlcnQu
# Y29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDA6oDigNoY0aHR0cDovL2Ny
# bDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDBQBgNV
# HSAESTBHMDgGCmCGSAGG/WwAAgQwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cu
# ZGlnaWNlcnQuY29tL0NQUzALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggEB
# AHGVEulRh1Zpze/d2nyqY3qzeM8GN0CE70uEv8rPAwL9xafDDiBCLK938ysfDCFa
# KrcFNB1qrpn4J6JmvwmqYN92pDqTD/iy0dh8GWLoXoIlHsS6HHssIeLWWywUNUME
# aLLbdQLgcseY1jxk5R9IEBhfiThhTWJGJIdjjJFSLK8pieV4H9YLFKWA1xJHcLN1
# 1ZOFk362kmf7U2GJqPVrlsD0WGkNfMgBsbkodbeZY4UijGHKeZR+WfyMD+NvtQEm
# tmyl7odRIeRYYJu6DC0rbaLEfrvEJStHAgh8Sa4TtuF8QkIoxhhWz0E0tmZdtnR7
# 9VYzIi8iNrJLokqV2PWmjlIwggVPMIIEN6ADAgECAhAE/eOq2921q55B9NnVIXVO
# MA0GCSqGSIb3DQEBCwUAMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lD
# ZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EwHhcNMjAwNzIwMDAw
# MDAwWhcNMjMwNzI1MTIwMDAwWjCBizELMAkGA1UEBhMCR0IxEjAQBgNVBAcTCVdh
# a2VmaWVsZDEmMCQGA1UEChMdU2VjdXJlIFBsYXRmb3JtIFNvbHV0aW9ucyBMdGQx
# GDAWBgNVBAsTD1NjcmlwdGluZ0hlYXZlbjEmMCQGA1UEAxMdU2VjdXJlIFBsYXRm
# b3JtIFNvbHV0aW9ucyBMdGQwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQCvbSdd1oAAu9rTtdnKSlGWKPF8g+RNRAUDFCBdNbYbklzVhB8hiMh48LqhoP7d
# lzZY3YmuxztuPlB7k2PhAccd/eOikvKDyNeXsSa3WaXLNSu3KChDVekEFee/vR29
# mJuujp1eYrz8zfvDmkQCP/r34Bgzsg4XPYKtMitCO/CMQtI6Rnaj7P6Kp9rH1nVO
# /zb7KD2IMedTFlaFqIReT0EVG/1ZizOpNdBMSG/x+ZQjZplfjyyjiYmE0a7tWnVM
# Z4KKTUb3n1CTuwWHfK9G6CNjQghcFe4D4tFPTTKOSAx7xegN1oGgifnLdmtDtsJU
# OOhOtyf9Kp8e+EQQyPVrV/TNAgMBAAGjggHFMIIBwTAfBgNVHSMEGDAWgBRaxLl7
# KgqjpepxA8Bg+S32ZXUOWDAdBgNVHQ4EFgQUTXqi+WoiTm5fYlDLqiDQ4I+uyckw
# DgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHcGA1UdHwRwMG4w
# NaAzoDGGL2h0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3VyZWQtY3Mt
# ZzEuY3JsMDWgM6Axhi9odHRwOi8vY3JsNC5kaWdpY2VydC5jb20vc2hhMi1hc3N1
# cmVkLWNzLWcxLmNybDBMBgNVHSAERTBDMDcGCWCGSAGG/WwDATAqMCgGCCsGAQUF
# BwIBFhxodHRwczovL3d3dy5kaWdpY2VydC5jb20vQ1BTMAgGBmeBDAEEATCBhAYI
# KwYBBQUHAQEEeDB2MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5j
# b20wTgYIKwYBBQUHMAKGQmh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydFNIQTJBc3N1cmVkSURDb2RlU2lnbmluZ0NBLmNydDAMBgNVHRMBAf8EAjAA
# MA0GCSqGSIb3DQEBCwUAA4IBAQBT3M71SlOQ8vwM2txshp/XDvfoKBYHkpFCyanW
# aFdsYQJQIKk4LOVgUJJ6LAf0xPSN7dZpjFaoilQy8Ajyd0U9UOnlEX4gk2J+z5i4
# sFxK/W2KU1j6R9rY5LbScWtsV+X1BtHihpzPywGGE5eth5Q5TixMdI9CN3eWnKGF
# kY13cI69zZyyTnkkb+HaFHZ8r6binvOyzMr69+oRf0Bv/uBgyBKjrmGEUxJZy+00
# 7fbmYDEclgnWT1cRROarzbxmZ8R7Iyor0WU3nKRgkxan+8rzDhzpZdtgIFdYvjeO
# c/IpPi2mI6NY4jqDXwkx1TEIbjUdrCmEfjhAfMTU094L7VSNMYIEXDCCBFgCAQEw
# gYYwcjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UE
# CxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1
# cmVkIElEIENvZGUgU2lnbmluZyBDQQIQBP3jqtvdtaueQfTZ1SF1TjAJBgUrDgMC
# GgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYK
# KwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG
# 9w0BCQQxFgQUC2hxLtH90lB2TYFoGr65UrFmfrUwDQYJKoZIhvcNAQEBBQAEggEA
# iTMGO16/t4uBMPDGCyc4nzygMG6IV6DxEJH7IxvOPSZppcNBMVrE/kkDWVtwiha2
# pYuuCdvsXmXoGfO+DHsiPBidKIrRFf/PhDSGJUdvzeC9Cr8kMozwB+r0xXGuZzHb
# c65lDiETZ1i0IJkLCANEhhtL4RgxoI3NUTNZrrpB+c8AGWU25H3ADK2fni1xlKlm
# Gogy/cqlIfeHg5rmTAZxa6FKQQUtw4EQdabt3xG189mmFvFz/Sh7KPt2ALR8yxoF
# hX1lozf4zotZUErJ/pUKn1DTJ8U23xGrsMa5vu1/qHBEHrQ4HbIYFFs1Z8XhV1Aa
# OIR5SPD8KIB8AAjxcBt05aGCAjAwggIsBgkqhkiG9w0BCQYxggIdMIICGQIBATCB
# hjByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQL
# ExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3Vy
# ZWQgSUQgVGltZXN0YW1waW5nIENBAhANQkrgvjqI/2BAIc4UAPDdMA0GCWCGSAFl
# AwQCAQUAoGkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUx
# DxcNMjIwMjEyMTQxNDAzWjAvBgkqhkiG9w0BCQQxIgQgbHeEhENzP/sA1wT2Twj9
# DGhGaccPuSMn9DDN28u1Z14wDQYJKoZIhvcNAQEBBQAEggEApPZU8hCI3vl0hgVI
# rWCZn1/ZpzGxeNA2WkRwR+kGyg6JRez2NJ368c4WKrXwD2fukKf0kqu4BxKSDVAC
# HCknSGvHfmZns1DNHtnDGnWVeMhwPc5frGFCe4cB/tXCjs//TZ0k8KDmcesCNJOM
# CbYz71rHb035UYWJwt2mHyKPxhF0yBjL4vJXa46A8A9ZZ836IoUxBnQf5h8sIxWc
# hx6NPqoOSqvc+P5rvuLk6r7InWS/ThQl4/XwYWX0ash5Y1aiA3v7wVJsMTBz4320
# GzL9aSL+ZOnDhJO8382LQmQOZiiuweSymbRBblJxQ+ZqouFTpDuMybGKf90F6qIo
# 0KH6WQ==
# SIG # End signature block
