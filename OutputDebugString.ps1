
<#
.SYNOPSIS

Send a string to OutputDebugString() so it can be picked up by SysInternals dbgview

.PARAMETER message

The message text to write to the debug channel. %environment% variables will be expanded as well as $env: PowerShell ones

.EXAMPLE

.\OutputDebugString.ps1 -message "The script is running as $env:username"

Write the string "The script is running as $env:username", with %username% expanded, to the debug channel

#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$true,HelpMessage="Message to send to debug output")]
    [string]$message
)

## https://stackoverflow.com/questions/60363812/how-to-write-to-debug-stream-with-powershell-core-so-its-captured-by-sysinterna
$WinAPI = @'
    public class WinAPI
    {
        [System.Runtime.InteropServices.DllImport("kernel32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        public static extern void OutputDebugString(string message);
    }
'@

Add-Type $WinAPI -Language CSharp

[WinAPI]::OutputDebugString( [Environment]::ExpandEnvironmentVariables( $message ) )
