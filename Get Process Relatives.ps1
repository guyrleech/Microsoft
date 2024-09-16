#requires -Version 5

<#
.SYNOPSIS
    Get parent and child processes details and recurse

.DESCRIPTION
    Uses win32_process. Level 0 processes are those specified via parameter, positive levels are parent processes & negative levels are child processes
    Child processes are not guaranteed to be directly below their parent - check process id and parent process id

.PARAMETER name
    A regular expression to match the name(s) of the process(es) to retrieve.

.PARAMETER id
    The ID(s) of the process(es) to retrieve.

.PARAMETER indentMultiplier
    The multiplier for the indentation level. Default is 1.

.PARAMETER indenter
    The character(s) used for indentation. Default is a space.

.PARAMETER unknownProcessName
    The placeholder name for unknown processes. Default is '<UNKNOWN>'.

.PARAMETER properties
    The properties to retrieve for each process

.PARAMETER quiet
    Suppresses warning output if specified.

.PARAMETER norecurse
    Prevents recursion through processes if specified.

.PARAMETER noIndent
    Disables creatring indented name if specified.

.PARAMETER noChildren
    Excludes child processes from the output if specified.

.PARAMETER noOwner
    Excludes the owner from the output if specified which speeds up the script.
    
.PARAMETER sessionId
    Process all processes passed via -id or -name regardless of session if * is passed (default)
    Only process processes passed via -id or -name if they are in the same session as the script if -1 is passed
    Only process processes passed via -id or -name if they are in the same session as the value passed if it is a positive integer

.EXAMPLE
   & '.\Get Process Relatives.ps1' -id 12345

   Get parent and child processes of the running process with pid 12345

.EXAMPLE
   & '.\Get Process Relatives.ps1' -name notepad.exe,winword.exe -properties *

   Get parent and child processes of all running processes of notepad and winword, outputting all win32_process properties & added ones

.EXAMPLE
   & '.\Get Process Relatives.ps1' -name powershell.exe -sessionid -1

   Get parent and child processes of powershll.exe processes running in the same session as the script

.NOTES
    Modification History:

    2024/09/13  @guyrleech  Script born
    2024/09/16  @guyrleech  First release
#>

[CmdletBinding(DefaultParameterSetName='name')]

Param
(
    [Parameter(ParameterSetName='name',Mandatory=$true)]
    [string[]]$name ,
    [Parameter(ParameterSetName='id',Mandatory=$true)]
    [int[]]$id ,
    [string]$sessionId = '*' ,
    [int]$indentMultiplier = 1 ,
    [string]$indenter = ' ' ,
    [string]$unknownProcessName = '<UNKNOWN>' ,
    [string[]]$properties = @( 'IndentedName' , 'ProcessId' , 'ParentProcessId' , 'Sessionid' , '-' , 'Owner' , 'CreationDate' , 'Level' , 'CommandLine' , 'Service' ) ,
    [switch]$quiet ,
    [switch]$norecurse ,
    [switch]$noIndent ,
    [switch]$noChildren ,
    [switch]$noOwner
)

Function Get-DirectRelativeProcessDetails
{
    Param
    (
        [int]$id ,
        [int]$level = 0 ,
        [datetime]$created ,
        [bool]$children = $false ,
        [switch]$recurse ,
        [switch]$quiet ,
        [switch]$firstCall
    )
    Write-Verbose -Message "Get-DirectRelativeProcessDetails pid $id level $level"
    $processDetail = $null
    ## array is of win32_process objects where we order & search on process id
    [int]$processDetailIndex = $script:processes.BinarySearch( [pscustomobject]@{ ProcessId = $id } , $comparer )
    if( $processDetailIndex -ge 0 )
    {
        $processDetail = $script:processes[ $processDetailIndex ]
    }
    ## else not found

    ## guard against pid re-use (do not need to check pid created after child process since could not exist before with same pid although can't guarantee that pid hasn't been reused since unless we check process auditing/sysmon)
    if( $null -ne $processDetail -and ( $null -eq $created -or ( -not $children -and $processDetail.CreationDate -le $created ) -or $children ) )
    {
        ## * means any session, -1 means session script is running in any other positive value is session id it process must be running in
        if( $sessionId -ne '*' -and $firstCall )
        {
            if( $script:sessionIdAsInt -lt 0 ) ## session for script only
            {
                if( $processDetail.SessionId -ne $script:thisSessionId )
                {
                    $processDetail = $null
                }
            }
            elseif( $script:sessionIdAsInt -ne $processDetail.SessionId ) ## session id passed so check process is in this session
            {
                $processDetail = $null
            }
        }
        if( $null -ne $processDetail -and $null -ne $processDetail.ParentProcessId -and $processDetail.ParentProcessId -gt 0 )
        {
            if( $recurse )
            {
                if( $children )
                {
                    $script:processes | Where-Object ParentProcessId -eq $id -PipelineVariable childProcess | ForEach-Object `
                    {
                        Get-DirectRelativeProcessDetails -id $childProcess.ProcessId -level ($level - 1) -recurse -children $true -created $processDetail.CreationDate -quiet:$quiet
                    }
                }
                if( $firstCall -or -not $children ) ## getting parents
                {
                    Get-DirectRelativeProcessDetails -id $processDetail.ParentProcessId -level ($level + 1) -children $false -recurse  -created $processDetail.CreationDate -quiet:$quiet
                }
            }

            ## don't just look up svchost.exe as could be a service with it's own exe
            [string]$service = ($script:runningServices[ $processDetail.ProcessId ]| Select-Object -ExpandProperty Name) -join '/'

            $owner = $null
            if( -Not $noOwner )
            {
                if( -Not $processDetail.PSObject.Properties[ 'Owner' ] )
                {
                    $ownerDetail = Invoke-CimMethod -InputObject $processDetail -MethodName GetOwner -ErrorAction SilentlyContinue
                    if( $null -ne $ownerDetail -and $ownerDetail.ReturnValue -eq 0 )
                    {
                        $owner = "$($ownerDetail.Domain)\$($ownerDetail.User)"
                    }

                    Add-Member -InputObject $processDetail -MemberType NoteProperty -Name Owner -Value $owner
                }
                else
                {
                    $owner = $processDetail.owner
                }
            }
            
            ## clone the process detail since may be used by another process being analysed and could be at a different level in that
            ## clone() method not available in PS 7.x
            $clone = [CimInstance]::new( $processDetail )

            Add-Member -InputObject $clone -PassThru -NotePropertyMembers @{ ## return
                Owner   = $owner
                Service = $service
                Level   = $level
                '-'     = $(if( $firstCall ) { '*' } else {''})
            }
        }
        ## else no parent or excluded based on session id
    }
    elseif( $firstCall ) ## only warn on first call
    {
        if( -not $quiet )
        {
            Write-Warning "No process found for id $id"
        }
    }
    elseif( -not $quiet )
    {
        ## TODO search process auditing/sysmon ?
        [pscustomobject]@{
            Name = $unknownProcessName
            ProcessId = $id
            Level = $level
        }
    }
}

## main

[int]$script:thisSessionId = (Get-Process -id $pid).SessionId
$script:sessionIdAsInt = $sessionId -as [int]

class NameComparer : System.Collections.Generic.IComparer[PSCustomObject]
{
    [int] Compare( [PSCustomObject]$x , [PSCustomObject]$y )
    {
        ## cannot simply return difference directly since Compare must return int but uint32 could be bigger
        [int64]$difference = $x.ProcessId - $y.ProcessId
        if( $difference -eq 0 )
        {
            return 0
        }
        elseif( $difference -lt 0 )
        {
            return -1
        }
        else
        {
            return 1
        }
    }
}

## use sorted array so can find quicker
$comparer = [NameComparer]::new()
## get all processes so quicker to find parents and children regardless of session id as only filter on session id of processes specified by paramter, not parent/child
$script:processes = [System.Collections.Generic.List[PSCustomObject]]( Get-CimInstance -ClassName win32_process )
$script:processes.Sort( $comparer )

Write-Verbose -Message "Got $($script:processes.Count) processes"

## if names passed as parameter then get pids for them
if( $null -ne $name -and $name.Count -gt 0 )
{
    $id = @( ForEach( $processName in $name ) 
    {
        $script:processes | Where-Object Name -Match $processName | Select-Object -ExpandProperty ProcessId
    })

    if( $id.Count -eq 0 )
    {
        Throw "No processes found for $name"
    }
    Write-Verbose -Message "Got $($id.Count) pids for process $name"
}

## get all services so we can quickly look them up
[hashtable]$script:runningServices = @{}

Get-CimInstance -ClassName win32_service -filter 'ProcessId > 0' -PipelineVariable service | ForEach-Object `
{
    ## could be multiple so store as array
    $existing = $runningServices[ $service.ProcessId ]
    if( $null -eq $existing )
    {
        $runningServices.Add( $service.ProcessId , ( [System.Collections.Generic.List[object]]$service ))
    }
    else ## already have this pid
    {
        $existing.Add( $service )
    }
}

Write-Verbose -Message "Got $($script:runningServices.Count) running service pids"

[array]$results = @( ForEach( $processId in $id )
{
    [array]$result = @( Get-DirectRelativeProcessDetails -id $processId -recurse:(-Not $norecurse) -quiet:$quiet -children (-Not $noChildren) -firstCall | Sort-Object -Property Level -Descending )
    ## now we know how many levels we can indent so the topmost process has no ident - no point waiting for all results as some may not have still existing parents so don't know what level in relation to other processes
    if( -not $noIndent -and $null -ne $result -and $result.Count -gt 1 )
    {
        $levelRange = $result | Measure-Object -Maximum -Minimum -Property Level
        ForEach( $item in $result )
        {
            Add-Member -InputObject $item -MemberType NoteProperty -Name IndentedName -Value ("$($indenter * ($levelRange.Maximum - $item.level) * $indentMultiplier)$($item.name)")
        }
    }
    else ## not indenting
    {
        $properties[ 0 ] = 'Name'
    }
    $result
})

$results | Select-Object -Property $properties
