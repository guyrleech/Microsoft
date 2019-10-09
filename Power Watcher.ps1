#requires -version 3
<#
    Modification History:

    24/09/19  GRL  Initial public release

    @guyrleech 2019
#>

<#
.SYNOPSIS

Wait for a power event, show grid view of available schemes and set the one chosen when OK is clicked

.DESCRIPTION

Designed to help set the most suitable power scheme when using an external power bank for a laptop as the laptop sees it as still being powered by an external power source
so does not implement any power saving.

.PARAMETER wmiClass

The WMI class to query - do not change this

.PARAMETER sourceIdentifier

The string to use to identify events from this watcher. A default is provided so there is no need to use this parameter

.PARAMETER waitPeriod

How long to sleep for in seconds before timing out the event wait. There is no need to set this as the default is to wait indefinitely

.EXAMPLE

& '.\Power Watcher.ps1'

Wait for a power change event, show a grid view of the available power schemes and if one is selected and OK pressed, set that as the active power scheme. Then repeat, as in wait for another power event.

.NOTES

Use powercfg.exe to clone a power scheme via its GUID, rename it and then change the settings so all "plugged in" settings are set as per battery saving.
Get the GUID for the existing schemes via powercfg.exe /list

powercfg /Duplicatescheme GUID
powercfg /Changename GUID "On Powerbank"

#>

[CmdletBinding()]

Param
(
    [string]$wmiClass = 'Win32_PowerManagementEvent' ,
    [string]$sourceIdentifier ,
    [int]$waitPeriod = -1
)

if( ! $PSBoundParameters[ 'sourceIdentifier' ] )
{
    $sourceIdentifier = ($wmiClass -split '_')[-1]
}

Unregister-Event -SourceIdentifier $sourceIdentifier -ErrorAction SilentlyContinue

Register-CimIndicationEvent -Query "Select * From $wmiClass" -SourceIdentifier $sourceIdentifier

[bool]$carryOn = $true

$null = [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

While( $carryOn  )
{
    $event = Wait-Event -Timeout $waitPeriod
    if( $event)
    {
        if( $event.SourceIdentifier -eq $sourceIdentifier )
        {
            ## https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-powermanagementevent
            [string]$eventType = $event.SourceEventArgs.NewEvent.ToString()
            [string]$text = "created $(Get-Date -Date $event.TimeGenerated -Format G) : $eventType : "
            switch( $eventType )
            {
                "Win32_PowerManagementEvent" 
                {
                    $batteryStatus = Get-CimInstance -ClassName BatteryStatus -Namespace root\WMI
                    [string]$powerSource = $(if( $batteryStatus.PowerOnline ) { 'external power' } else { 'battery' })
                    $text += "Event type $($event.SourceEventArgs.NewEvent.EventType), power source $powerSource"
                    if( $event.SourceEventArgs.NewEvent.EventType -eq 10 )
                    {
                        ## Get power schemes so we can prompt
                        [array]$powerSchemes = powercfg.exe /list | ForEach-Object `
                        {
                            ## 'Power Scheme GUID: 381b4222-f694-41f0-9685-ff5bb260df2e  (Balanced)'
                            if( $_ -match '.*([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}).*\((.*)\)(.*)' )
                            {
                                [pscustomobject]@{
                                    'Scheme' = $Matches[2]
                                    'GUID' = $Matches[1]
                                    'Active' = $(if( $Matches[3] -match '\*' ) { 'Yes' } else { 'No' } ) }
                            }
                        }
                        $selected = $powerSchemes | Out-GridView -Title "$(Get-Date -Format G) : select power plan, now on $powerSource" -PassThru
                        if( $selected )
                        {
                            if( $selected -is [array] )
                            {
                                $null = [System.Windows.Forms.MessageBox]::Show( 'Only select one power scheme' , ( Split-Path -Path (& { $myInvocation.ScriptName }) -Leaf) , 0 , [System.Windows.Forms.MessageBoxIcon]::Error )
                            }
                            elseif( $selected.Active -ne 'Yes' )
                            {
                                $result = Start-Process -FilePath powercfg.exe -ArgumentList "/setactive $($selected.GUID)" -PassThru -Wait -WindowStyle Hidden
                                if( ! $result -or $result.ExitCode )
                                {
                                    $null = [System.Windows.Forms.MessageBox]::Show( "Failed to set active power scheme to `"$($selected.Scheme)`"" , ( Split-Path -Path (& { $myInvocation.ScriptName }) -Leaf) , 0 , [System.Windows.Forms.MessageBoxIcon]::Error )
                                }
                            }
                            else
                            {
                                Write-Warning "Selected power scheme `"$($selected.Scheme)`" is already active so not changing"
                            }
                        }
                    }
                }
                default
                {
                    $text += " unexpected event type!" 
                }
            }
            Write-Verbose "$(Get-Date -Format G): $text"
        }
        $event | Remove-Event -ErrorAction SilentlyContinue
    }
}
