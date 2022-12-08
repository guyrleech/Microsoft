
<#
.SYNOPSIS

Use WMI to query computers to find out, since boot, when any interactive desktop connections logged on

.PARAMETER computers

A comma separated list of computers to query

.PARAMETER user

A specific user to query otherwise all users who have logged on are returned

.PARAMETER before

Only sessions logged on on or before this date/time are returned

.PARAMETER after

Only sessions logged on on or after this date/time are returned

.EXAMPLE

& '\get remote user logon times.ps1' 

Show all remote desktop logons since boot on the current computer

.EXAMPLE

& '\get remote user logon times.ps1' -Computers computer01,computer02 -user guy.leech -After "01/08/2018"

Show all remote desktop logons for user guy.leech on computers computer01 and computer02 since 1st August 2018

.NOTES

Pipe the output through cmdlets like Out-GridView, Export-CSV or Format-Table

Modification History:

    2018/08/14  @guyrleech  Initial Release
    2022/12/08  @guyrleech  Fixed flattening of array bug

#>

[CmdletBinding()]

Param
(
    [string]$user = $env:USERNAME ,
    [string]$shellProcess = 'Explorer' ,
    [datetime]$after ,
    [datetime]$before ,
    [int]$sessionid = ( Get-Process -Id $pid | select -ExpandProperty SessionId ) ,
    [switch]$noPrompt
)

$timeNow = [datetime]::Now

$culture = Get-Culture
[string]$preciseTimeFormat = '{0}.FFFFFFF' -f $culture.DateTimeFormat.SortableDateTimePattern 

$logons = @( Get-WmiObject win32_logonsession  | ForEach-Object `
{
    $session = $_
    Get-WmiObject win32_loggedonuser -filter "Dependent = '\\\\.\\root\\cimv2:Win32_LogonSession.LogonId=`"$($session.logonid)`"'" | ForEach-Object `
    {
        [datetime]$logonTime = (([WMI] '').ConvertToDateTime( $session.StartTime ))
        if( ( ! $PSBoundParameters[ 'after'] -or $logonTime -ge $after ) `
            -and ( ! $PSBoundParameters[ 'before' ] -or $logonTime -le $before ) `
                -and $_.Antecedent -match 'Domain="(.*)",Name="(.*)"$' `
                    -and ( [string]::IsNullOrEmpty( $user ) -or ( $matches -and $matches.Count -eq 3 -and $user -eq $matches[2] ) ) )
        {
            [pscustomobject]@{ 'User' = $Matches[1] + '\' + $Matches[2] ; 'Logon Time' = $logonTime ; 'Logon Id' = $session.LogonId ; 'Logon Type' = $session.LogonType ; 'Logon Time Precise' = (Get-Date $logonTime -Format $preciseTimeFormat ) ;'Authentication Package' = $session.AuthenticationPackage }
        }
    }
} | Sort-Object -Property 'Logon Time' -Descending )

$logons | Format-Table -AutoSize

if( $logons -and $logons.Count )
{
    if( ! [string]::IsNullOrEmpty( $shellProcess ) )
    {
        [array]$processes = @( Get-Process -Name $shellprocess -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $sessionid } )
        if( ! $processes -or ! $processes.Count )
        {
            Write-Warning "Found no instances of process $shellProcess for session $sessionid"
        }
        else
        {
            [array]$shells = @( ForEach( $process in $processes )
            {
                [pscustomobject]@{ 'Started' = $process.StartTime ; 'Seconds after logon' = (New-TimeSpan -End $process.StartTime -Start $logons[0].'Logon Time').TotalSeconds ; 'Seconds to now' = (New-TimeSpan -End $timeNow -Start $process.StartTime).TotalSeconds }
            })
        }

        "$((New-TimeSpan -End $timeNow -Start $logons[0].'Logon Time').TotalSeconds ) seconds elapsed from logon to now on $env:COMPUTERNAME for $user"

        $logons[0] | Format-Table -AutoSize

        "$shellProcess instances for session $sessionid"

        $shells | Format-Table -AutoSize
    }

    quser $user
}
else
{
    Write-Warning "No logons found for $user in session $sessionid"
}

if( ! $noPrompt )
{
    $null = Read-Host 'Hit enter to exit'
}
