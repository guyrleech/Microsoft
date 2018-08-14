<#
    Get remote logon times via WMI

    @guyrleech (c) 2018
#>

<#
.SYNOPSIS

Use WMI to query computers to find out, since boot, when any remote desktop connections logged on

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
#>

[CmdletBinding()]

Param
(
    [string[]]$computers = @( $env:COMPUTERNAME ) ,
    [string]$user ,
    [datetime]$after ,
    [datetime]$before
)

$culture = Get-Culture
[string]$preciseTimeFormat = '{0}.FFFFFFF' -f $culture.DateTimeFormat.SortableDateTimePattern 

ForEach( $computer in $computers )
{
    Get-WmiObject win32_logonsession -ComputerName $computer -Filter "LogonType='10'" | ForEach-Object `
    {
        $session = $_
        Get-WmiObject win32_loggedonuser -ComputerName $computer -filter "Dependent = '\\\\.\\root\\cimv2:Win32_LogonSession.LogonId=`"$($session.logonid)`"'" | ForEach-Object `
        {
            [datetime]$logonTime = (([WMI] '').ConvertToDateTime( $session.StartTime ))
            if( ( ! $PSBoundParameters[ 'after'] -or $logonTime -ge $after ) `
                -and ( ! $PSBoundParameters[ 'before' ] -or $logonTime -le $before ) `
                    -and $_.Antecedent -match 'Domain="(.*)",Name="(.*)"$' `
                        -and ( [string]::IsNullOrEmpty( $user ) -or ( $matches -and $matches.Count -eq 3 -and $user -eq $matches[2] ) ) )
            {
                [pscustomobject]@{ 'Computer' = $computer ; 'User' = ($matches[1] + '\' + $matches[2] ) ; 'Logon Time' = $logonTime ; 'Logon Time Precise' = (Get-Date $logonTime -Format $preciseTimeFormat ) ;'Authentication Package' = $session.AuthenticationPackage }
            }
        }
    }
}
