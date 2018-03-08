#Requires -version 3.0
<#
    Change priorities of processes which over consume CPU

    Use this script at your own risk.

    Guy Leech, 2018
#>

<#
.SYNOPSIS

Change the priority class of processes which consume over a specified amount of CPU during the monitoring period to aid fair sharing of CPU

.DESCRIPTION

When one or more threads are ready to run, the OS will place the higher priority thread on a CPU first. 
Therefore reducing the priority class of processes which are deemed to be consuming too much CPU means 
that other processes will get preferential access to the CPU.

When a process stops over-consuming then its priority will be returned to what it was before it was reduced by the script.

.PARAMETER frequency

How often, in seconds, to check the CPU consumption of processes

.PARAMETER threshold

The percentage of time, in the time period defined by -frequency, which the process was consuming CPU, over which its CPU priority class is lowered

.PARAMETER includeNames

A comma separated list of process names that will have their processes instances and potentially punished

.PARAMETER excludeNames

A comma separated list of process names that will not have their instances examined or punished

.PARAMETER includeUsers

A comma separated list of usernames that will have their processes examined and potentially punished. The script must be run elevated

.PARAMETER excludeUsers

A comma separated list of usernames that will not have their processes examined or punished. The script must be run elevated

.PARAMETER excludeSessions

A comma separated list of session IDs which will not have their processes examined or punished

.PARAMETER selfOnly

Only examine processes in the current session

.EXAMPLE

& '\Change CPU priorities.ps1' -confirm:$false

Monitor all accessible processes, except those in session 0 like services, and change their priority as needed without prompting

.EXAMPLE

& '\Change CPU priorities.ps1' -confirm:$false -includeNames msiexec -confirm:$false -frequency 30 -threshold 33 -verbose

Monitor all msiexec.exe processes every 30 seconds and if any of those have consumed over 33% CPU in those 30 seconds then
their CPU priority will be decreased by one class. When any affected processes reduce their consumption to below 33% then
their priority will be returned to what it was before it was changed by this script. Verbose output will be shown.

.NOTES

The PowerShell process running the script will have its priority class set to High to give it the best chance of being able to run on a busy system.
This priority will be reset when the script exits, e.g. via ctrl-C

There is a very small risk that a process could terminate and a new one be created with the same process id when the script is sleeping but this is
vyer unlikely and the script may cope with it anyway.
#>

[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='High')]

Param
(
    [ValidateScript({$_ -gt 0 })]
    [int]$frequency = 60 ,
    [ValidateScript({$_ -gt 0 -and $_ -le 100 })]
    [decimal]$threshold = 50 ,
    [string[]]$includeNames = @() ,
    [string[]]$excludeNames = @(),
    [string[]]$includeUsers = @() ,
    [string[]]$excludeUsers = @() ,
    [int[]]$excludeSessions = @( 0 ) ,
    [switch]$selfOnly
)

[hashtable]$getProcessParams = @{}

if( $includeUsers.Count -or $excludeUsers.Count )
{
    if( $PSVersionTable.PSVersion.Major -ge 3 )
    {     
        $myWindowsPrincipal = New-Object System.Security.Principal.WindowsPrincipal([System.Security.Principal.WindowsIdentity]::GetCurrent())

        # Get the security principal for the Administrator role
        $adminRole = [System.Security.Principal.WindowsBuiltInRole]::Administrator
 
        # Check to see if we are currently running "as Administrator"
        if ( $myWindowsPrincipal.IsInRole($adminRole))
        {
            $getProcessParams.Add( 'IncludeUserName' , $true )
        }
        else
        {
            Write-Warning "Unable to get user names since not running elevated"
            return
        }
    }
    else
    {
        Write-Warning "Unable to get user names as requires PowerShell 3.0 or higher and this is $($PSVersionTable.psversion.ToString())"
        return
    }
}

[hashtable]$processes = @{}
[int]$pulse = 0
[decimal]$excessive = 100 / $frequency
## Priority reducing table
[hashtable]$lowerPriority = 
@{
    'Idle' = 'Idle'
    'BelowNormal' = 'Idle'
    'Normal' = 'BelowNormal'
    'AboveNormal' = 'Normal'
    'High' = 'AboveNormal'
    'Realtime' = 'High'
}

[int]$ownSession = -1
if( $selfOnly )
{
    $ownSession = (Get-Process -Id $pid).SessionId
}

Function Revert-Processes( [hashtable]$processes )
{
    $processes.GetEnumerator() | ForEach-Object `
    {
        if( $_.Value.AdjustedPriority )
        {
            $priority = $_.Value.OriginalPriority
            Get-Process -Id $_.Value.Id -ErrorAction SilentlyContinue | ForEach-Object { $_.PriorityClass = $priority }
            $_.Value.AdjustedPriority = $false
        }
    }
}

[string]$originalPriority = $null

Get-Process -Id $pid | Select -First 1 | ForEach-Object { $originalPriority = $_.PriorityClass ; $_.PriorityClass = 'High' }

## Put main loop in a try block so can revert process priorties back at exit via finally block, e.g. if ctrl-c pressed
try
{
    While( $true )
    {
        [int]$excluded = 0
        Get-Process @getProcessParams | ForEach-Object `
        {
            $thisProcess = $_
            if( $thisProcess.HasExited )
            {
                $processes.Remove( $thisProcess.Id )
            }
            elseif( $thisProcess.Id -ne $PID ) ## don't adjust ourself
            {
                [bool]$included = $true

                if( $includeUsers.Count )
                {
                    $included = $includeUsers -contains $thisProcess.UserName
                }
                elseif( $excludeUsers.Count )
                {
                    $included = $excludeUsers -notcontains $thisProcess.UserName
                }
                if( $included -and $includeNames.Count )
                {
                    $included = $includeNames -contains $thisProcess.Name
                }
                elseif( $included -and $excludeNames.Count )
                {
                    $included = $excludeNames -notcontains $thisProcess.Name
                }
                if( $included -and $excludeSessions.Count )
                {
                    $included = $excludeSessions -notcontains $thisProcess.SessionId
                }
                if( $included -and $ownSession -ge 0 )
                {
                    $included = ( $thisProcess.SessionId -eq $ownSession )
                }

                if( $included )
                {
                    ## If we have the process already then look how much CPU it has used since last time sampled
                    $existingProcess = $processes[ $thisProcess.Id ]
                    if( $existingProcess )
                    {
                        [int]$cpuConsumptionPercent = ( $thisProcess.TotalProcessorTime.TotalSeconds - $existingProcess.TotalCPUSeconds ) * $excessive
                        if( $cpuConsumptionPercent -gt $threshold )
                        {
                            Write-Verbose "$($thisProcess.Id) : $($thisProcess.Name) : Consumed $cpuConsumptionPercent % CPU (had $($thisProcess.TotalProcessorTime.TotalSeconds - $existingProcess.TotalCPUSeconds) secs), priority class $($thisProcess.PriorityClass)"
                            if( $existingProcess.OriginalPriority )
                            {
                                $newPriority = $lowerPriority[ $existingProcess.OriginalPriority.ToString() ] ## set to same reduced priority rather than keep decreasing it
                                if( $PSCmdlet.ShouldProcess( "$($thisProcess.Name) ($($thisProcess.id))" , "Change priority to $newPriority" ))
                                {
                                    $thisProcess.PriorityClass = $newPriority
                                    $existingProcess.AdjustedPriority = $true
                                }
                            }
                            else
                            {
                                Write-Warning "Unable to get current priority for $($thisProcess.Name) ($($thisProcess.Id)) so unable to reduce it - probably a permissions issue"
                            }
                        }
                        elseif( $existingProcess.AdjustedPriority )
                        {
                            Write-Verbose  "$($thisProcess.Id) : $($thisProcess.Name) : Consumed $cpuConsumptionPercent % CPU, priority class $($thisProcess.PriorityClass) but now below threshold of $threshold so setting back to $($existingProcess.OriginalPriority)"
                            $existingProcess.AdjustedPriority = $false
                            $thisProcess.PriorityClass = $existingProcess.OriginalPriority 
                        }
                        ## update main table with changed stats
                        $existingProcess.TotalCPUSeconds = $thisProcess.TotalProcessorTime.TotalSeconds
                        ## mark this process as alive so we can make a pass to remove dead processes
                        $existingProcess.Pulse = $pulse
                    }
                    else ## process did not exist in our hash table so add it with extra properties
                    {
                        try
                        {
                            Add-Member -InputObject $thisProcess -NotePropertyMembers `
                            @{
                                ## Record CPU consumption as TotalProcessTime is read-only so we can't update the existing hash table entry on next pass
                                TotalCPUSeconds = $thisProcess.TotalProcessorTime.TotalSeconds;
                                Pulse = $pulse ;
                                OriginalPriority = $thisProcess.PriorityClass
                                AdjustedPriority = $false
                            }
                            $processes.Add( $thisProcess.Id , $thisProcess )
                        }
                        catch
                        {
                            Write-Verbose "Failed to add process id $($thisProcess.Id) $($thisprocess.Name): $($_.Exception.Message)"
                        }
                    }
                }
                else
                {
                    $excluded++
                }
            }
        }
    
        ## remove processes which are no longer alive so have missed the change in the pulse value - have to use a clone of the hashtable since we can't change it when enumerating over it
        if( $processes.Count )
        {
            [hashtable]$clonedProcesses = $processes.Clone()
            $clonedProcesses.GetEnumerator() | ForEach-Object `
            {
                if( $_.Value.HasExited -or $_.Value.Pulse -ne $pulse )
                {
                    $processes.Remove( $_.Key )
                }
            }
            Clear-Variable clonedProcesses
        }

        Write-Verbose "$(Get-Date -Format u) : sleeping for $frequency seconds with $($processes.Count) processes monitored & $excluded excluded"
        Start-Sleep -Seconds $frequency
        $pulse = ! $pulse
    }
}
Finally
{
    Revert-Processes $processes
    (Get-Process -Id $pid).PriorityClass = $originalPriority
}
