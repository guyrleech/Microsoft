#requires -version 3
<#
    Modification History

    @guyrleech 23/03/2020  Initial Release
#>

<#
.SYNOPSIS

Find window for given process and set as foreground window or perform another operation on it

.DESCRIPTION

Pass processes to operate via pid or name, with optional command line argument matching for the latter.
By default will restore the processes to their original size and position but can also be used to minimise, maximise, hide, etc

.PARAMETER id

Comma separated list of process ids to operate on

.PARAMETER name

Comma separated list of process names to operate on

.PARAMETER arguments

Regular expression which will only match processes specified by -name whose command line arguments match the regex

.PARAMETER command

The window operation to perform. The default is SW_RESTORE which activates and displays the window

.EXAMPLE 

& '.\Set Foreground Window.ps1' -name notepad.exe -command SW_MINIMIZE

Find all notepad.exe processes and minimise those windows

.EXAMPLE 

& '.\Set Foreground Window.ps1' -name Teams.exe -arguments Avanite 

Find all Teams.exe processes which have the word "Avanite" in their command line arguments and restore those windows

.NOTES

Using SW_HIDE will mean that this script cannot be used subsequently to restore the window as Get-Process then does not return the MainWindowHandle property. These processes can be restored using task manager.

https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-showwindowasync

#>

[CmdletBinding()]

Param
(
    [Parameter(ParameterSetName='Pid',Mandatory=$true)]
    [int[]]$id ,
    [Parameter(ParameterSetName='Name',Mandatory=$true)]
    [string[]]$name ,
    [Parameter(ParameterSetName='Name',Mandatory=$false)]
    [string]$arguments ,
    [ValidateSet('SW_RESTORE','SW_FORCEMINIMIZE','SW_HIDE','SW_MAXIMIZE','SW_MINIMIZE','SW_SHOW','SW_SHOWDEFAULT','SW_SHOWMAXIMIZED','SW_SHOWMINIMIZED','SW_SHOWMINNOACTIVE','SW_SHOWNA','SW_SHOWNOACTIVATE','SW_SHOWNORMAL')]
    [string]$command = 'SW_RESTORE'
)

## https://docs.microsoft.com/en-gb/windows/win32/api/winuser/nf-winuser-showwindow
[hashtable]$cmdShowTranslation = @{
    'SW_RESTORE' = 9
    'SW_FORCEMINIMIZE' = 11
    'SW_HIDE' = 0
    'SW_MAXIMIZE' = 3
    'SW_MINIMIZE' = 6
    'SW_SHOW' = 5
    'SW_SHOWDEFAULT' = 10
    'SW_SHOWMAXIMIZED' = 3
    'SW_SHOWMINIMIZED' = 2
    'SW_SHOWMINNOACTIVE' = 7
    'SW_SHOWNA' = 8
    'SW_SHOWNOACTIVATE' = 4
    'SW_SHOWNORMAL' = 1
}

$operation = $cmdShowTranslation[ $command ]

if( $operation -eq $null )
{
    Throw "Unknown command `"$command`", see https://docs.microsoft.com/en-gb/windows/win32/api/winuser/nf-winuser-showwindow"
}

$pinvokeCode = @'
        [DllImport("user32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow); 
'@

if( ! ([System.Management.Automation.PSTypeName]'Win32.User32').Type )
{
    Add-Type -MemberDefinition $pinvokeCode -Name 'User32' -Namespace 'Win32' -UsingNamespace System.Text -ErrorAction Stop
}

[hashtable]$getprocessParameters = @{ 'ErrorAction' = 'Stop' }

if( $PSBoundParameters[ 'id' ] )
{
    $getprocessParameters.Add( 'Id' , $id )
}
elseif( $PSBoundParameters[ 'arguments' ] )
{
    [array]$ids = @( Get-CimInstance -ClassName win32_process -Filter "Name = '$name'" | Where-Object CommandLine -match $arguments | Select-Object -ExpandProperty ProcessId )
    Write-Verbose -Message "Got $($ids.Count) matching $name processes: $($ids -join ' , ')"
    if( ! $ids -or ! $ids.Count )
    {
        Throw "Found no $name processes with command lines matching regex `"$arguments`""
    }
    else
    {
        $getprocessParameters.Add( 'Id' , $ids )
    }
}
else
{
    $getprocessParameters.Add( 'Name' , $name -replace '\.exe$' )
}

[int]$processedCount = 0

Get-Process @getprocessParameters | ForEach-Object `
{
    $process = $_
    $processedCount++
    if( ( [intptr]$windowHandle = $process.MainWindowHandle ) -and $windowHandle -ne [intptr]::Zero )
    {
        [bool]$setForegroundWindow = [win32.user32]::ShowWindowAsync( $windowHandle , [int]$operation ) ; $lastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
        if( ! $setForegroundWindow )
        {
            Write-Warning -Message "Failed to set window to foreground for process $($process.Name) (pid $($process.Id)): $lastError"
        }
        else
        {
            Write-Verbose -Message "Operation $operation on $($process.Name) (pid $($process.Id)) succeeded"
        }
    }
    else
    {
        Write-Warning -Message "No main window handle for process $($process.Name) (pid $($process.Id))"
    }
}

Write-Verbose -Message "Processed $processedCount processes"

if( ! $processedCount )
{
    Write-Warning -Message "Failed to find any main window handles"
}
