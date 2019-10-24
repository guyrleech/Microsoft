#requires -version 3
<#
    Show processes with a given module loaded

    @guyrleech 2019
#>

<#
.SYNOPSIS

Examine loaded modules all or specific processes by name or pid and show those that match a specified string/regex

.DESCRIPTION

Designed to help spot processes hooked by 3rd party software like Citrix, Ivanti, Lakeside, etc. Shows module versions so can also be used to play spot the difference between processes.

.PARAMETER moduleName

The name of the module or regex to match. Also specify -regex when the module name is a regex

.PARAMETER processName

The name of a process or pattern to only show matching modules for those processes

.PARAMETER processId

The pid of a process or pattern to only show matching modules for that process

.PARAMETER regex

Specifies that the module name passed via -moduleName is a regular expression so special characters like \ will not be escaped

.PARAMETER noUsername

Do not retrieve the username. Use this when not running the script elevated

.PARAMETER quiet

Do not display warning messages, some of which may be benign such as insufficient access to a process so no process handle available

.PARAMETER noOtherBitness

Do not check processes that are not the same bitness (32 or 64 bit) as the PowerShell process running the script if on x64 computer

.PARAMETER convertToCsv

Used internally when calling the other bitness powershell executable. Pipe the results through Export-CSV if you want to produce a csv file

.EXAMPLE

& '.\Find loaded modules.ps1' -processName notepad -Verbose -modulename msvcrt.dll

Show what modules called msvcrt.dll are loaded into all notepad processes running as 32 or 64 bit

.EXAMPLE

& '.\Find loaded modules.ps1' -modulename MfApHook -quiet

Show what modules with MfApHook in their name are loaded into all processes running as 32 or 64 bit whilst suppressing any warning messages

.EXAMPLE

& '.\Find loaded modules.ps1' -modulename .*

Show all modules loaded into all processes running as 32 or 64 bit. Pipe through Out-GridView, Export-CSV, Format-Table, etc to save/visualise

#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$true,HelpMessage='Module name to search for')]
    [string]$moduleName ,
    [string]$processName ,
    [int]$processId ,
    [switch]$quiet ,
    [switch]$regex ,
    [switch]$noUsername ,
    [switch]$noOtherBitness ,
    [switch]$convertToCsv
)

[string]$escapedModuleName = $(if( $regex ) { $moduleName } else { [regex]::Escape( $moduleName ) })

[hashtable]$processArguments = @{}

if( ! $noUsername )
{
    $processArguments.Add( 'IncludeUserName' , $true )
}

if( $PSBoundParameters[ 'processname' ] )
{
    $processArguments.Add( 'Name' , $processName )
}

if( $PSBoundParameters[ 'processId' ] )
{
    $processArguments.Add( 'Id' , $processId )
}

Add-Type @'
using System;
using System.Runtime.InteropServices;

public static class Kernel32
{
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool IsWow64Process(
        [In] IntPtr hProcess ,
        [Out,MarshalAs(UnmanagedType.Bool)] out bool wow64Process );
}
'@

[bool]$is32bitProcess = $false
[bool]$powerShellis32bit = [IntPtr]::Size -eq 4
[int]$unchecked = 0
[int]$otherBitness = 0

Write-Verbose "PowerShell is 32 bit - $powerShellis32bit"

[System.Collections.ArrayList]$results = @( Get-Process @processArguments | . { Process `
{
    $process = $_
    $processProperties = $null

    if( $process.Id -gt 4 ) ## don't check idle and system processes
    {
        if( ! $process.Handle )
        {
            $unchecked++
            if( ! $quiet )
            {
                Write-Warning "No handle on $($process.Name) (PID $($process.Id)) to check if 32 or 64 bit"
            }
        }
        elseif( ! [kernel32]::IsWow64Process( $process.Handle , [ref]$is32bitProcess ) )
        {
            $unchecked++
            if( ! $quiet )
            {
                Write-Warning -Message "Failed to determine if process $($process.Name) (PID $($process.Id)) is 32 or 64 bit"
            }
        }
        else
        {
            if( ( $powerShellis32bit -and ! $is32bitProcess ) -or ( ! $powerShellis32bit -and $is32bitProcess ) )
            {
                $otherBitness++
            }
        }

        ForEach( $module in $process.Modules )
        {
            if( $module.FileName -match $escapedModuleName )
            {
                if( ! $processProperties )
                {
                    $processProperties = $process | Get-ItemProperty -ErrorAction SilentlyContinue | Select-Object -ExpandProperty VersionInfo
                }
                [pscustomobject][ordered]@{
                    'Process' = $process.Name
                    'PID' = $process.Id
                    'Bitness' = $(if( $is32bitProcess ) { '32' } else { '64' })
                    'Path' = $processProperties | Select-Object -ExpandProperty FileName -ErrorAction SilentlyContinue
                    'User' = $process.UserName
                    'Started' = $process.StartTime
                    'Product' = $processProperties | Select-Object -ExpandProperty ProductName -ErrorAction SilentlyContinue
                    'Version' = $processProperties | Select-Object -ExpandProperty ProductVersion -ErrorAction SilentlyContinue
                    'Module Name' = $module.ModuleName
                    'Module Path' = $module.FileName
                    'Module Version' = $module.FileVersion
                    'Module Product' = $module.FileVersionInfo.ProductName
                    'Module Company' = $module.FileVersionInfo.CompanyName
                }
            }
        }
    }
}})

if( ! $noOtherBitness -and $otherBitness -and [System.Environment]::Is64BitOperatingSystem )
{
    Write-Verbose "Checking other bitness (we are 32 bit is $is32bitProcess) because $otherBitness other bitness processes found"

    ## get parent process so can flick to other bitness if necessary to get those process details too
    [string]$exePath = (Get-Process -Id $pid).Path
    if( $exePath )
    {
        [string]$otherExePath = $( if( $powerShellis32bit )
            {
                $exePath -replace ( '\\syswow64\\' ) , '\\sysnative\\'
            }
            else
            {
                $exePath -replace ( "^{0}" -f [regex]::Escape( [System.Environment]::GetFolderPath( [System.Environment+SpecialFolder]::System ))) , [System.Environment]::GetFolderPath( [System.Environment+SpecialFolder]::SystemX86 )
            })
        if( $otherExePath -ne $exePath )
        {
            if( $host.Name -ne 'ConsoleHost' )
            {
                $otherExePath = $otherExePath -replace '\\powershell_ise\.exe$' , '\powershell.exe'
            }
            ## Don't test path to exe since sysnative isn't a real path and will fail
            ## If result is being redirected or piped then we need to truncate our new command at that point
            [string]$cmdLine = (($script:MyInvocation.Line.Replace( "`"" , "'" ) ).Trim() -replace '\|.*$' , '' -replace , '\>.*$' , '' -replace ';.*$' , '' -replace "^&" , "$otherExePath -File " ) + ' -NoOtherBitness -ConvertToCsv'
            Write-Verbose "Running $cmdLine"
            [array]$otherResults = Invoke-Expression -Command $cmdLine
            if( $otherResults -and $otherResults.Count )
            {
                ## filter out warning, verbose
                $results += $otherResults | Where-Object { $_ -match '^"' } | ConvertFrom-Csv
            }
        }
        else
        {
            if( ! $quiet )
            {
                Write-Warning -Message "Unable to determine other bitness of powershell"
            }
        }
    }
    elseif( ! $quiet )
    {
        Write-Warning -Message "Unable to determine path of executable so cannot rerun for other bitness"
    }
}

if( $unchecked -and ! $quiet )
{
    Write-Warning -Message "Unable to check $unchecked processes as either no permissions or not $(if( $powerShellis32bit ) { '32' } else { '64' }) bit"
}

if( $results -and $results.Count )
{
    Write-Verbose -Message "Got $($results.Count) matching modules"
    if( $convertToCsv ) ## so we can convert back to objects when running the opposite bitness PowerShell from within this script
    {
        $results | ConvertTo-Csv -NoTypeInformation
    }
    else
    {
        $results
    }
}
elseif( ! $quiet )
{
    Write-Warning -Message "No matches found"
}
