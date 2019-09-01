#requires -version 3.0
<#
    Find AD accounts with passwords or accounts expiring within the specified number of days or are locked out or disabled and optionally send an email containing the information
    
    This script is provided as is and with absolutely no warranty such that the author cannot be held responsible for any untoward behaviour or failure of the script.

    @guyrleech 2019

    Modification History:

    01/09/19  GRL  Fixed bug with -password and -mailpassword
#>

<#
.SYNOPSIS

Find AD accounts with passwords or accounts expiring within the specified number of days or are locked out or disabled and optionally send an email containing the information

.PARAMETER server

The specific domain controller to connect to

.PARAMETER accountName

Only AD accounts matching this pattern will be returned.

.PARAMETER expireWithinDays

Only report accounts which have passwords or accounts which are set to expire within this number of days

.PARAMETER ignore

A comma separated list of accounts to ignore

.PARAMETER encryptPassword

Encrypt the password passed by the -mailpassword option so it can be passed to -mailHashedPassword. The encrypted password is specific to the user and machine where they are encrypted.
Pipe through clip.exe or Set-ClipBoard (scb) to place in the Windows clipboard

.PARAMETER mailServer

The SMTP mail server to use

.PARAMETER proxyMailServer

If email relaying is only allowed from specific computers, try and remote the Send-Email cmdlet via the server specific via this argument

.PARAMETER noSSL

Do not use SSL to communicate with the mail server

.PARAMETER subject

The subject of the email sent with the expiring account list

.PARAMETER from

The email address to send the email from. Some mail servers must have a valid email address specified

.PARAMETER recipients

A comma separated list of the email addresses to send emails to

.PARAMETER mailUsername

The username to authenticate with at the mail server

.PARAMETER mailPassword

The clear text password for the -mailUsername argument. If the %_MVal12% environment variable is set then its contents are used for the password

.PARAMETER mailHashedPassword

The hashed mail password returned from a previous call to the script with -encryptPassword. Will only work on the same machine and for the same user that ran the script to encrypt the password

.PARAMETER mailPort

The port to use to communicate with the mail server

.PARAMETER nogridview

If not emailing then output the results to the pipeline rather than displaying in a gridview

.PARAMETER logFile

The full path to a log file to append the output of the script to

.EXAMPLE

& '.\Check SQL account expiry.ps1' -accountName svc* -recipients guyl@hell.com -mailServer smtp.hell.com -expireWithinDays 10

This will connect to a domain controller and email a list of all AD accounts starting with "svc" which expire within the next 10 days to the given recipient.

.EXAMPLE

& '.\Check SQL account expiry.ps1' -accountName svc* -ignore SVC_dummy -server dc04 -expireWithinDays 10

This will connect to domain controller dc04 and displaly a list of all AD accounts starting with "svc", except svc_dummy, which expire within the next 10 days in an on screen grid view

.EXAMPLE

& '.\Check AD account expiry.ps1' -encryptPassword -mailpassword thepassword

This will encrypt the given password and output its encrypted form so that it can be passed as the argument to the -mailHashedPassword option to avoid having to specify the password on the command line.
The encrypted password will only work for the same user that encrypted it and on the same machine. Email credentials are only required when pass through authentication does not work, e.g. ISP SMTP servers

.NOTES

Place in a scheduled task where the action should be set to start a program, powershell.exe, with the arguments starting with '-file "C:\Scripts\Check AD account expiry.ps1"' and then having the rest of the required arguments after this.
Requires the ActiveDirectory PowerShell module

#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$true,HelpMessage='Pattern to match account names')]
    [string]$accountName  ,
    [int]$expireWithinDays = 7 ,
    [string]$server ,
    [string[]]$ignore ,
    [string]$hashedPassword ,
    [switch]$encryptPassword ,
    [string]$mailServer ,
    [string]$proxyMailServer = 'localhost' ,
    [switch]$noSSL ,
    [string]$subject = "AD Accounts with passwords expiring in the next $expireWithinDays days" ,
    [string]$from = "$($env:computername)@$($env:userdnsdomain)" ,
    [string[]]$recipients ,
    [string]$mailUsername ,
    [string]$mailPassword ,
    [string]$mailHashedPassword ,
    [int]$mailport ,
    [switch]$nogridview ,
    [string]$logFile
)

if( $encryptPassword )
{
    if( ! $PSBoundParameters[ 'mailpassword' ] -and ! ( $mailPassword = $env:_Mval12 ) )
    {
        Throw 'Must specify the mail username''s password when encrypting via -mailpassword or _Mval12 environment variable'
    }
    
    ConvertTo-SecureString -AsPlainText -String $mailPassword -Force | ConvertFrom-SecureString
    Exit 0
}

try
{
    if( ! [string]::IsNullOrEmpty( $logFile ) )
    {
        Start-Transcript -Path $logFile -Append
    }
    
    $queryError = $null

    Import-Module -Name ActiveDirectory -ErrorVariable queryError -Verbose:$false
    if( $ignore -and $ignore.Count -eq 1 -and $ignore[0].IndexOf(',') -ge 0 ) ## fix scheduled task not passing array correctly
    {
        $ignore = $ignore -split ','
    }
    [datetime]$dateThreshold = (Get-Date).AddDays( $expireWithinDays )
    [hashtable]$serverParam = @{}
    if( $PSBoundParameters[ 'server' ] )
    {
        $serverParam.Add( 'Server' , $server )
    }
    [array]$results = @( Get-ADUser @serverParam -ErrorVariable queryError -Filter "name -like '$accountName'" -Properties AccountExpirationDate,AccountLockoutTime,PasswordExpired,PasswordLastSet,badPwdCount,BadLogonCount,PasswordNeverExpires,LastLogonDate,LockedOut,Enabled,msDS-UserPasswordExpiryTimeComputed,LastBadPasswordAttempt,TrustedForDelegation,Created| . { Process `
    {
        $account = $_
        Write-Verbose -Message "Checking $($account.Name)"
        if( $account.Name -notin $ignore `
            -and ( ( $account.PasswordNeverExpires -eq $false -and $account.'msDS-UserPasswordExpiryTimeComputed' -and [datetime]::FromFileTime( $account.'msDS-UserPasswordExpiryTimeComputed' ) -le $dateThreshold ) `
                -or ( $account.AccountExpirationDate -ne $null -and $account.AccountExpirationDate -le $dateThreshold ) -or $account.Enabled -eq $false -or $account.LockedOut -eq $true -or $account.PasswordExpired -eq $true ) )
        {
            Write-Verbose -Message "`tadded"
            [pscustomobject][ordered]@{
                'Name' = $account.Name
                'OU' = ($account.DistinguishedName -split ',' , 2)[-1]
                'Enabled' = $account.Enabled
                'Created' = $account.Created
                'Account Expires' = $account.AccountExpirationDate
                'Password Expires' = $(if( $account.'msDS-UserPasswordExpiryTimeComputed' -and ! $account.PasswordNeverExpires ) { [datetime]::FromFileTime( $account.'msDS-UserPasswordExpiryTimeComputed' ) })
                'Password Expired' = $account.PasswordExpired
                'Password Never Expires' = $account.PasswordNeverExpires
                'Password Last Set' = $account.PasswordLastSet
                'Locked Out' = $account.LockedOut
                'Lockout time' = $account.AccountLockoutTime
                'Last Bad Password' = $account.LastBadPasswordAttempt
                'Bad Password Count' = $account.badPwdCount
                'Bad Logon Count' = $account.BadLogonCount
                'Trusted for Delegation' = $account.TrustedForDelegation
            }
        }
    }})

    if( $results -and $results.Count )
    {    
        [hashtable]$mailParams = $null

        if( ( ! [string]::IsNullOrEmpty( $proxyMailServer )  -or ! [string]::IsNullOrEmpty( $mailServer ) ) -and $recipients.Count )
        {
            if( $recipients -and $recipients.Count -eq 1 -and $recipients[0].IndexOf(',') -ge 0 ) ## fix scheduled task not passing array correctly
            {
                $recipients = $recipients -split ','
            }

            $mailParams = @{
                    'To' =  $recipients
                    'SmtpServer' = $mailServer
                    'From' =  $from
                    'UseSsl' = ( ! $noSSL ) }
            if( $PSBoundParameters[ 'mailport' ] )
            {
                $mailParams.Add( 'Port' , $mailport )
            }
            if( $PSBoundParameters[ 'mailUsername' ] )
            {
                $thePassword = $null
                if( ! $PSBoundParameters[ 'mailPassword' ] )
                {
                    if( $PSBoundParameters[ 'mailHashedPassword' ] )
                    {
                        Write-Verbose "Using hashed password of length $($mailHashedPassword.Length)"
                        $thePassword = $mailHashedPassword | ConvertTo-SecureString
                    }
                    elseif( Get-ChildItem -Path env:_MVal12 -ErrorAction SilentlyContinue )
                    {
                        $thePassword = ConvertTo-SecureString -AsPlainText -String $env:_MVal12 -Force
                    }
                }
                else
                {
                    $thePassword = ConvertTo-SecureString -AsPlainText -String $mailPassword -Force
                }
        
                if( $thePassword )
                {
                    $mailParams.Add( 'Credential' , ( New-Object System.Management.Automation.PSCredential( $mailUsername , $thePassword )))
                }
                else    
                {
                    Write-Error "Must specify mail account password via -mailPassword, -mailHashedPassword or _MVal12 environment variable"
                }
            }
        }
    
        if( ( ! [string]::IsNullOrEmpty( $proxyMailServer )  -or ! [string]::IsNullOrEmpty( $mailServer ) ) -and $recipients.Count )
        {
            ## Ok, so it's hard coded! Puts a border on our table
            [string]$style = '<style>BODY{font-family: Arial; font-size: 10pt;}'
            $style += "TABLE{border: 1px solid black; border-collapse: collapse;}"
            $style += "TH{border: 1px solid black; background: #dddddd; padding: 5px; }"
            $style += "TD{border: 1px solid black; padding: 5px; }"
            $style += "</style>"

            Write-Verbose "Emailing $($results.Count) results to $($recipients -join ',') via $mailServer"

            [string]$htmlBody = $results | ConvertTo-Html -Head $style

            $mailParams.Add( 'Body' , $htmlBody )
            $mailParams.Add( 'BodyAsHtml' , $true )
            $mailParams.Add( 'Subject' ,  $subject )
         
            if( $PSBoundParameters[ 'proxyMailServer' ] )
            {
                Invoke-Command -ComputerName $proxyMailServer -ScriptBlock { [hashtable]$mailParams = $using:mailParams ; Send-MailMessage @mailParams }
            }
            else
            {
                Send-MailMessage @mailParams 
            }
        }
        elseif( ! $nogridview )
        {
            [array]$selected = @( $results | Out-GridView -PassThru )
            if( $selected -and $selected.Count )
            {
                $selected | Set-Clipboard
            }
        }
        else
        {
            $results
        }
    }
    else
    {
        Write-Output -InputObject "No accounts found expiring before $(Get-Date -Date $dateThreshold -Format G) or currently disabled or locked out"
    }
}
catch
{
    Throw $_
}
finally
{
    if( ! [string]::IsNullOrEmpty( $logFile ) )
    {
        Stop-Transcript
    }
}