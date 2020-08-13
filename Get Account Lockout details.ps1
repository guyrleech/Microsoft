#requires -version 3

<#
.SYNOPSIS

Look for account lock out events on all domain controllers

.PARAMETER dayago

The number of days back to start the event log search from

.PARAMETER username

Username to report lockouts for otherwise will report for all users

.PARAMETER silent

Do not output any status messages

.EXAMPLE

& '.\Get Account Lockout details.ps1' -username ritasue -silent -daysAgo 7

Look for account lockout events that have occurred in the last 7 days and report any for the user ritasue.
If none are found there will be no output.

.NOTES

Must have account lockout auditing enable otherwise the events do not get generated

@guyrleech 24/07/2020
#>

[CmdletBinding()]


Param
(
    [double]$daysAgo = 1 ,
    [string]$username ,
    [switch]$silent
)

if( ! ( $domain = [directoryServices.ActiveDirectory.Domain]::GetCurrentDomain() ) )
{
    Throw "Failed to get the current domain"
}

[array]$domainControllers = @( $domain.FindAllDomainControllers() )

if( ! $domainControllers -or ! $domainControllers.Count )
{
    Throw "Failed to get any domain controllers for domain $($domain.Name)"
}

Write-Verbose -Message "Got $($domainControllers.Count) domain controllers in domain $($domain.Name)"

[datetime]$startDate = (Get-Date).AddDays( -$daysAgo )

## Invoke command will run some in parallel and Get-WinEvent will only take a single machine
[array]$events = @( Invoke-Command -ComputerName $domainControllers -ScriptBlock { Get-WinEvent -FilterHashtable @{ LogName = 'Security' ; Id = 4740 ; StartTime = $using:startDate } -ErrorAction SilentlyContinue | Select TimeCreated,Properties } )

if( $events -and $events.Count )
{
    Write-Verbose -Message "Found $($events.Count) events on $($domainControllers.Count) domain controllers"

    <#
        Properties array:

        TargetUserName johndoe 
        TargetDomainName GLS16MCS01 
        TargetSid S-1-5-21-1721611859-3364803896-2099701507-2124 
        SubjectUserSid S-1-5-18 
        SubjectUserName GRL-DC03$ 
        SubjectDomainName GUYRLEECH 
        SubjectLogonId 0x3e7 
    #>
    [array]$filtered = @( $events | Where-Object { [string]::IsNullOrEmpty( $username ) -or $_.properties[0].value -eq $username } | Select-Object -Property TimeCreated,@{n='User name';e={$_.Properties[0].value}},@{n='Computer name';e={$_.Properties[1].value}},@{n='Domain Controller';e={($_.PSComputerName -split '\.')[0]}} | Sort-Object -Property TimeCreated )
    if( $filtered -and $filtered.Count )
    {
        $filtered
    }
    elseif( ! $silent )
    {
        [string]$message = "Found no lock out events on $($domainControllers.Count) domain controllers since $(Get-Date -Date $startDate -Format G)"
        if( ! [string]::IsNullOrEmpty( $username ) )
        {
            $message += " for user $username"
        }
        Write-Output -InputObject $message
    }
}
elseif( ! $silent )
{
    Write-Output -InputObject "No lock out events found on $($domainControllers.Count) domain controllers since $(Get-Date -Date $startDate -Format G)"
}
