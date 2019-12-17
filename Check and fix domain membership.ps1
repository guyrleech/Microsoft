#requires -version 3
<#
    Check domain membership is ok and if not fix via stored credentials

    @guyrleech 16/12/19
#>

<#
.SYNOPSIS

Check domain membership is ok and if not fix via stored credentials

.PARAMETER username

Domain account with permissions to fix the secure channel

.PARAMETER credential

A credential object with permissions to fix the secure channel

.PARAMETER password

A secure string representation of the password for the account used to fix the secure channel

.PARAMETER passwordFile

Path to a file containing a secure string representation of the password for the account used to fix the secure channel

.PARAMETER server

Specific domain controller to target, otherwise one is picked

.PARAMETER retries

The number of times to retry repairing the secure channel

.PARAMETER logfile

Path to a file which will have a log from the script appended to it

.PARAMETER sleepForMilliSeconds

The number of milliseconds to sleep for between retries

.PARAMETER encrypt

Will prompt for a password which will be output such that it can be placed in a file for the -passwordFile option or passed directly to -password

.PARAMETER simulateFail

Simulates the secure channel test failing such that the fixing can be tested

.EXAMPLE

& '.\Check and fix domain membership.ps1' -encrypt

Will prompt for a password such that it can be placed in a file for the -passwordFile option or passed directly to -password

.EXAMPLE

& '.\Check and fix domain membership.ps1' -username domain\someuser -passwordFile c:\scripts\crud -retries 4

If the secure channel is broken, up to four attempts will be made to fix the secure channel using the domain\someuser account and the secure string password stored in c:\scripts\crud
which has been generated previosuly via the -encrypt option

.NOTES

Does not require a reboot once it has fixed the secure channel

Run at startup and/or periodically via a scheduled task, running under a local administrator account
#>

[CmdletBinding()]

Param
(
    [string]$username ,
    [Parameter(Mandatory=$true,ParameterSetName='Credential',HelpMessage='Credential object for domain join account')]
    [System.Management.Automation.PSCredential]$credential ,
    [Parameter(Mandatory=$true,ParameterSetName='Password',HelpMessage='Secure string password for domain join account')]
    [string]$password ,
    [Parameter(Mandatory=$true,ParameterSetName='PasswordFile',HelpMessage='File containing secure string password for domain join account')]
    [string]$passwordFile ,
    [string]$server ,
    [int]$retries = 5 ,
    [string]$logFile ,
    [int]$sleepForMilliSeconds = 1000 ,
    [Parameter(Mandatory=$true,ParameterSetName='Encrypt',HelpMessage='Enter and encrypt a password')]
    [switch]$encrypt ,
    [switch]$simulateFail 
)

If( $encrypt )
{
    $credential = Get-Credential -Message "Enter password for encryption" -UserName 'N\A'
    If( $credential )
    {
        $credential.Password|ConvertFrom-SecureString
        Exit 0
    }
    Else
    {
        Exit 1
    }
}

If( $PSBoundParameters[ 'logfile' ] )
{
    Start-Transcript -Path $logfile -Append
}

Try
{
    [hashtable]$parameters = @{ 'ErrorAction' = 'SilentlyContinue' }

    If( $PSBoundParameters[ 'server' ] )
    {
        $parameters.Add( 'Server' , $server )
    }

    If( $simulateFail -or ! ( Test-ComputerSecureChannel @parameters ) )
    {
        Write-Verbose -Message 'Test-ComputerSecureChannel returned false'
    
        If( ! $PSBoundParameters[ 'credential' ] )
        {
            If( ! $PSBoundParameters[ 'username' ] )
            {
                $username = '{0}\{1}' -f $env:USERDOMAIN , $env:USERNAME
            }

            If( $PSBoundParameters[ 'passwordFile' ] )
            {
                $password = Get-Content -Path $passwordFile
            }

            [System.Security.SecureString]$secureString = ConvertTo-SecureString -String $password -ErrorAction Stop
            $credential = New-Object -TypeName System.Management.Automation.PSCredential( $username , $secureString )
            If( ! $credential )
            {
                Throw "Failed to create credential for $username"
            }
        }

        $parameters += @{ 'Repair' = $true ; 'Credential' = $credential }

        [bool]$result = $false
        While( $retries -gt 0 -and ! $result )
        {
            $result = Test-ComputerSecureChannel @parameters
            If( ! $result -and $sleepForMilliSeconds -gt 0 )
            {
                Start-Sleep -Milliseconds $sleepForMilliSeconds
            }
            $retries--
        }

        If( $result )
        {
            Write-Output -InputObject 'Domain membership repaired ok'
        }
        Else
        {
            Write-Error -Message 'Failed to repair domain membership'
        }
    }
}
Catch
{
    Throw $_
}
Finally
{
    If( $PSBoundParameters[ 'logfile' ] )
    {
        Stop-Transcript
    }
}
