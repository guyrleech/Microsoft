<#
.SYNOPSIS
    Pull out logon task events from shell-core to show any delays in each

.PARAMETER daysback
    The number of days back (can be non-integer) to search in the event log

.PARAMETER user
    Only user names matching this regex will be analysed
    
.PARAMETER task
    Only show the logon tasks which match this regex

.PARAMETER mostRecent
    Only show the most recent instance of the logon task for the user
    
.EXAMPLE
    & '.\Shell Core Phases.ps1' -days 7 -mostRecent

    Show the most recent shell phase duration only for all users for the last 7 days
    
.EXAMPLE
    & '.\Shell Core Phases.ps1' -task PreShellTasks -user fred

    Show the most recent "PreShellTasks" shell phase duration only for user fred for the last 1 days

.NOTES
    Modification History:

    2023/03/03 @guyrleech  Initial Version
#>

[CmdletBinding()]

Param
(
    [decimal]$daysBack = 1 ,
    [string]$user ,
    [string]$task ,
    [switch]$mostRecent
)

[int]$counter = 0

Get-WinEvent -ErrorAction SilentlyContinue -Verbose:$false -FilterHashtable @{ ProviderName = 'Microsoft-Windows-Shell-Core' ; id = 62170,62171 ; StartTime = [datetime]::Now.AddDays( -$daysBack )} -Oldest | Select *,@{n='LogonTask';e={$_.Properties[1].Value}},@{n='TaskState';e={if( $_.Id -eq 62170 ) { 'Start' } else { 'Finish' }}}|select timecreated,id,logontask,Taskstate,message,userid|Group-Object -Property userid | ForEach-Object `
{
    $eventsForUser = $_
    $counter++
    [string]$username = ([System.Security.Principal.SecurityIdentifier]( $eventsForUser.Name )).Translate([System.Security.Principal.NTAccount]).Value 
    Write-Verbose -Message "$counter : $username"

    if( $username -match $user )
    {
        $eventsForUser.Group | Group-Object -Property LogonTask | ForEach-Object `
        {
            $eventsForTask = $_
            For( [int]$index = 0 ; $index -lt $_.Group.Count ; $index += 2 ) ## Too simple? What if a phase start/stop missed ?
            {
                if( $eventsForTask.Group[$index].LogonTask -match $task )
                {
                    [pscustomobject]@{
                        LogonTask = $eventsForTask.Group[$index].LogonTask
                        UserName = $username
                        Start = $eventsForTask.Group[$index].TimeCreated
                        End = $(if( $index -lt $eventsForTask.Group.Count -1 ) { $eventsForTask.Group[$index + 1].TimeCreated })
                        'Duration (ms)' = $(if( $index -lt $eventsForTask.Group.Count -1 ) { ( $eventsForTask.Group[$index + 1].TimeCreated - $_.Group[$index].TimeCreated).TotalMilliSeconds -as [int] })
                    }
                    if( $mostRecent )
                    {
                        break
                    }
                }
            }
        }
    }
}

