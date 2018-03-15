#requires -version 3

<#
    Trim working sets of processes or set maximum working set sizes

    Guy Leech, 2018

    Modification History

    10/03/18  GL  Optimised code

    12/03/18  GL  Added reporting
#>

<#
.SYNOPSIS

Manipulate the working sets (memory usage) of processes or report their current memory usage and working set limits and types (hard or soft)

.DESCRIPTION

Can reduce the memory footprints of running processes to make more memory available and stop processes that leak memory from leaking

.PARAMETER Processes

A comma separated list of process names to use (without the .exe extension). By default all processes will be trimmed if the script has access to them.

.PARAMETER Exclude

A comma separated list of process names to ignore (without the .exe extension).

.PARAMETER Above

Only trim the working set if the process' working set is currently above this value. Qualify with MB or GB as required. Default is to trim all processes

.PARAMETER MinWorkingSet

Set the minimum working set size to this value. Qualify with MB or GB as required. Default is to not set a minimum value.

.PARAMETER MaxWorkingSet

Set the maximum working set size to this value. Qualify with MB or GB as required. Default is to not set a maximum value.

.PARAMETER HardMin

When MinWorkingSet is specified, the limit will be enforced so the working set is never allowed to be less that the value. Default is a soft limit which is not enforced.

.PARAMETER HardMax

When MaxWorkingSet is specified, the limit will be enforced so the working set is never allowed to exceed the value. Default is a soft limit which can be exceeded.

.PARAMETER Report

Produce a report of the current working set usage and limit types for processes in the selection. Will output to a grid view unless -outputFile is specified.

.PARAMETER OutputFile

Ue with -report to write the results to a csv format file. If the file already exists the operation will fail.

.PARAMETER ProcessId

Only trim this specific process

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
    [switch]$report ,
    [string]$outputFile ,
    [int]$above = 10MB ,
    [int]$minWorkingSet = -1 ,
    [int]$maxWorkingSet = -1 ,
    [switch]$hardMax ,
    [switch]$hardMin ,
    [switch]$thisSession ,
    [int]$processId ,
    [int[]]$sessionIds ,
    [int[]]$notSessionIds ,
    [string]$available ,
    [switch]$savings ,
    [switch]$disconnected ,
    [switch]$background ,
    [switch]$scheduled ,
    [switch]$logoff ,
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
        [string]$available ,
        [string[]]$processes ,
        [string[]]$exclude ,
        [string]$logFile
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
    }

    if( ! [ string]::IsNullOrEmpty( $available ) )
    {
        $taskArguments.Add( 'Available' , $available )
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

[int]$thisSessionId = (Get-Process -Id $pid).SessionId
[long]$saved = 0
[int]$trimmed = 0

$params = @{}
if( $processId -gt 0 )
{
    $params.Add( 'Id' , $processId )
}

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

Get-Process @params | ForEach-Object `
{
    $process = $_
    [bool]$doIt = $true
    if( $processes -and $processes.Count -And $processes -notcontains $process.Name )
    {
        Write-Verbose ( "`tSkipping {0} pid {1} as not specified process" -f $process.Name , $process.Id )
        $doIt = $false
    }   
    elseif( $exclude -and $exclude.Count -And $exclude -contains $process.Name )
    {
        Write-Verbose ( "`tSkipping {0} pid {1} as specifically excluded" -f $process.Name , $process.Id )
        $doIt = $false
    }
    elseif( $thisSession -And $process.SessionId -ne $thisSessionId )
    {
        Write-Verbose ( "`tSkipping {0} pid {1} as session {2} not {3}" -f $process.Name , $process.Id , $process.SessionId , $thisSessionId )
        $doIt = $false
    }
    elseif( $process.Id -eq $activePid -And $idle -eq 0 ) ## if idle then we'll trim anyway as not being used (will have quit already if not idle if idle parameter specified)
    {
        Write-Verbose ( "`tSkipping {0} pid {1} as it is the foreground window process" -f $process.Name , $process.Id )
        $doIt = $false
    }
    elseif( $sessionIds -and $sessionIds.Count -gt 0 -And $sessionIds -notcontains $process.SessionId )
    {
        Write-Verbose ( "`tSkipping {0} pid {1} as session {2} not in list" -f $process.Name , $process.Id , $process.SessionId )
        $doIt = $false
    }
    elseif( $notsessionIds -and $notSessionIds.Count -gt 0 -And $notSessionIds -contains $process.SessionId )
    {
        Write-Verbose ( "`tSkipping {0} pid {1} as session {2}" -f $process.Name , $process.Id , $process.SessionId )
        $doIt = $false
    }
    elseif( $sessionsToTarget.Count -gt 0 -And $sessionsToTarget -notcontains $process.SessionId )
    {
        Write-Verbose ( "`tSkipping {0} pid {1} as session {2} which is not disconnected" -f $process.Name , $process.Id , $process.SessionId )
        $doIt = $false
    }
    elseif( $above -gt 0 -And $process.WS -le $above )
    {
        Write-Verbose ( "`tSkipping {0} pid {1} as working set only {2} MB" -f $process.Name , $process.Id , [Math]::Round( $process.WS / 1MB , 1 ) )
        $doIt = $false
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
                            'Minimum Working Set (KB)' = $thisMinimumWorkingSet / 1KB ; 'Maximum Working Set (KB)' = $thisMaximumWorkingSet / 1KB;
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
        $selected = $results | Sort-Object Name | Out-GridView -PassThru -Title "Memory information from $($results.Count) processes at $(Get-Date -Format U)"
        if( $selected )
        {
            $selected | clip.exe
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

if( ! [string]::IsNullOrEmpty( $logFile ) )
{
    Stop-Transcript
}
