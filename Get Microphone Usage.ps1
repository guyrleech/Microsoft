#requires -version 3

<#
.SYNOPSIS

Show Microphone, Webcam, etc usage from registry keys where it gets recorded

.DESCRIPTION

Will only show last used and what process/app and what, if anything, is currently using it. Only works on Windows 10

.PARAMETER registryKey

The registry key to query looking for the chosen device's usage history

.PARAMETER device

The device to query, eg microphone or webcam

.PARAMETER global

Query HKLM rather than HKCU

.EXAMPLE

& '.\Get Microphone Usage.ps1'

Show last usage for all microphone devices for the user running the script.

.EXAMPLE

& '.\Get Microphone Usage.ps1' -device webcam

Show last usage for all webcam devices for the user running the script.

.NOTES

Modification History:

    @guyrleech 21/09/20  Initial release
#>

[CmdletBinding()]

Param
(
    [string]$registryKey = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore' ,
    [string]$device = 'microphone' ,
    [switch]$global
)

if( $global )
{
    $registryKey = $registryKey -replace '^HKCU:' , 'HKLM:'
}

[string]$deviceRegistryPath = Join-Path -Path $registryKey -ChildPath $device

if( ! ( Test-Path -Path $deviceRegistryPath -PathType Container -ErrorAction SilentlyContinue ) )
{
    Throw "Path $deviceRegistryPath does not exist"
}

Get-ChildItem -Path $deviceRegistryPath -Force -Recurse -ErrorAction SilentlyContinue | ForEach-Object `
{
    $key = $_
    if( $started = $key.GetValue( 'LastUsedTimeStart' ) )
    {
        [bool]$inuse = $false
        [datetime]$startedTime = [datetime]::FromFileTime( $started )
        $stoppedTime = $null
        if( $null -ne ( [uint64]$stopped = $key.GetValue( 'LastUsedTimeStop' ) ) )
        {
            if( $stopped -eq 0 )
            {
                $inuse = $true
            }
            else
            {
                $stoppedTime = [datetime]::FromFileTime( $stopped )
            }
        }
        [pscustomobject]@{
            'App' =  $(if( $key.PSChildName.IndexOf( '#' ) -gt 0 ) { $key.PSChildName -replace '#' , '\' } else { $key.PSChildName -replace '_[a-z0-9]+$' } )
            'Started' = $startedTime
            'Stopped' = $stoppedTime
            'In Use' = $inuse
            'Duration (s)' = $(if( $stoppedTime ) { [int]($stoppedTime - $startedTime).TotalSeconds })
        }
    }
}
