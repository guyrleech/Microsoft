<#
    Alert if there is no process of the given name(s) and offer to start or just start anyway
    Can install/uninstall from the registry (per user or machine) so runs automatically at logon.
    Can be used to launch and monitor key apps like Outlook and Skype for Business (Lync)

    @guyrleech 2018
#>

<#
.SYNOPSIS

Check if any of a list of processes are not running and either just alert, prompt whether to run or just run based on parameters

.PARAMETER processes

Comma separated list of processes to check are running. Optionally add arguments which are after the character specified by the -delimiter parameter.
If specified without a full path then it must be executable from the shell, e.g. vi Start->Run

.PARAMETER launch

Prompt to launch any process which is found not to be running unless -auto is passed which will not prompt but start the missing process

.PARAMETER auto

Automatically launch any missing process without prompting

.PARAMETER ignoreFirst

Ignore the first instance of a missing process. Useful if something else is going to start the process

.PARAMETER install

Install the script into the registry so it runs with the specified arguments at logon

.PARAMETER uninstall

Remove the script from the registry so it doesn't run at logon

.PARAMETER allUsers

When used with -install or -uninstall will perform the action for all users. Default is for the user running the script only. Requires administrative privileges.

.PARAMETER checkPeriod

How often, in seconds, to check that processes are running. If zero then only check once rather than looping indefinitely.

.PARAMETER startDelay

Delay, in seconds, before the script starts checking for missing processes. Useful if something else is going to start the process

.PARAMETER delimiter

The character used to separate the process from any optional arguments that need passing when the process is invoked if it is missing

.PARAMETER valueName

The name of the value to create/delete when -install/-uninstall are used

.EXAMPLE

'.\Check and start processes.ps1'  -processes "outlook.exe","lync.exe:/fromrunkey" -startDelay 75 -auto

After an initial 75 second delay, check every 5 minutes that processes outlook.exe and lync.exe are running and if not automatially start them, passing "/fromrunkey" to lync.exe

'.\Check and start processes.ps1'  -processes "outlook.exe","lync.exe:/fromrunkey" -startDelay 75 -auto -install

As previous example but install into the user's run key in the registry so the script runs at logon

#>

[CmdletBinding()]

Param
(
    [string[]]$processes ,
    [switch]$launch ,
    [switch]$auto ,
    [switch]$ignoreFirst ,
    [switch]$install ,
    [switch]$uninstall ,
    [switch]$allUsers ,
    [int]$checkPeriod = 300 ,
    [int]$startDelay = 0 ,
    [string]$delimiter = ';' ,
    [string]$valueName = 'Guy''s Process Checker'
)

if( $install -and $uninstall )
{
    Throw 'Can''t install and uninstall in the same invocation'
}

[string]$runKey = $( if( $allUsers ) { 'HKLM' } else { 'HKCU' } ) + ':\Software\Microsoft\Windows\CurrentVersion\Run'

if( $install )
{
    [string]$valueData = ("powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File `"{0}`" -processes `"{1}`" -checkPeriod {2} -startDelay {3} -delimiter `"{4}`"" -f (& { $myInvocation.ScriptName }) , ($processes -join '","') ,  $checkPeriod , $startDelay , $delimiter )
    if( $launch )
    {
        $valueData += ' -launch'
    }
    if( $auto )
    {
        $valueData += ' -auto'
    }
    if( $ignoreFirst )
    {
        $valueData += ' -ignoreFirst'
    }
    Set-ItemProperty -Path $runKey -Name $valueName -Value $valueData -ErrorAction Stop
    Exit $?
}
elseif( $uninstall )
{
    Remove-ItemProperty -Path $runKey -Name $valueName -ErrorAction Stop
    Exit $?
}

## Workaround for array passed as single element which happens when script not invoked from within PowerShell
if( $processes[0].IndexOf( ',' ) -ge 0 )
{
    $processes = $processes[0] -split ','
}

if( ! $processes -or ! $processes.Count )
{
    Throw "Must specify one or more processes via -processes"
}

[int]$thisSessionId = (Get-Process -Id $pid -ErrorAction Stop).SessionId
[string]$scriptName = Split-Path -Path (& { $myInvocation.ScriptName }) -Leaf
[bool]$first = $true

## in case launched at logon before the monitored processes are launched
Write-Verbose ( "{0} : sleeping for {1} seconds" -f (Get-Date -Format G) , $startDelay )
Start-Sleep -Seconds $startDelay

[void](Add-type -assembly 'Microsoft.VisualBasic')

Write-Verbose ( "{0} : started monitoring {1} processes" -f (Get-Date -Format G) , $processes.Count )

do
{
    ForEach( $process in $processes )
    {
        ## Process may have full path and/or arguments so must strip out for Get-Process
        if( ! ( Get-Process -Name ([io.path]::GetFileNameWithoutExtension( ($process -split $delimiter)[0] )) -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $thisSessionId } ) )
        {
            if( ! $ignoreFirst -or ! $first )
            {
                [string]$theProcess,[string]$arguments = $process -split $delimiter
                [string]$message = ( "{0} : {1} is not running" -f (Get-Date -Format G) , $theProcess )
                Write-Verbose $message
                if( $launch -or $auto )
                {
                    [string]$answer = 'No'
                    if( ! $auto )
                    {
                        $message += '. Start it?'
                        $answer = [Microsoft.VisualBasic.Interaction]::MsgBox( $message , 'YesNo,SystemModal,Exclamation' , $scriptName )
                    }
                    if( $auto -or $answer -eq 'Yes' )
                    {
                        $pinfo = New-Object System.Diagnostics.ProcessStartInfo
                        $pinfo.FileName = $theProcess
                        $pinfo.Arguments = $arguments
                        $pinfo.RedirectStandardError = $false
                        $pinfo.RedirectStandardOutput = $false
                        $pinfo.UseShellExecute = $true
                        $pinfo.WindowStyle = 'Normal'
                        $pinfo.CreateNoWindow = $false
                        $process = New-Object System.Diagnostics.Process
                        $process.StartInfo = $pinfo
                        $launched = $null
                        try
                        {
                            $launched = $process.Start()
                        }
                        catch
                        {
                            $launched = $null
                            Write-Verbose $_
                        }
                        if( ! $launched )
                        {
                            [void][Microsoft.VisualBasic.Interaction]::MsgBox( "Failed to launch $($pinfo.FileName) $($pinfo.Arguments) - $($Error[0].Exception.Message)" , 'OKOnly,SystemModal,Exclamation' , $scriptName )
                        }
                        else
                        {
                            Write-Verbose "`"$($pinfo.FileName)`" launched ok"
                        }
                    }
                }
                else
                {
                    [void][Microsoft.VisualBasic.Interaction]::MsgBox( $message , 'OKOnly,SystemModal,Exclamation' , $scriptName )
                }
            }
        }
    }
    Start-Sleep -Seconds $checkPeriod
    $first = $false
} While( $checkPeriod -gt 0 )
