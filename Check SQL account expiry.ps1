#requires -version 3.0
<#
    Find SQL accounts with passwords expiring within the specified number of days and optionally send an email containing the information
    Can also be used to send an email alert if the specified SQL server cannot be connected to

    This script is provided as is and with absolutely no warranty such that the author cannot be held responsible for any untoward behaviour or failure of the script.

    @guyrleech 2019
#>

<#
.SYNOPSIS

Query SQL accounts on a SQL server and get those which expire within a given number of days

.PARAMETER sqlserver

The sqlserver\instance to connect to

.PARAMETER accountName

Only SQL accounts matching this regular expression will be returned. If not specified then all accounts found are returned.

.PARAMETER expireWithinDays

Only report accounts which have passwords which are set to expire within this number of days

.PARAMETER username

The SQL username to use to connect to the server rather than using the account that is running the script. Must also specify the password

.PARAMETER password

The password for the -username argument. If the %SQLval1% environment variable is set then its contents are used for the password

.PARAMETER hashedPassword

The hashed password returned from a previous call to the script with -encryptPassword. Will only work on the same machine and for the same user that ran the script to encrypt the password

.PARAMETER mailOnFail

Send an email if the connection to SQL fails

.PARAMETER includeDisabled

Includes accounts which have been disabled and therefore will fail to connect to SQL

.PARAMETER encryptPassword

Encrypt the password passed by the -password option so it can be passed to -hashedPassword or -mailHashedPassword. The encrypted password is specific to the user and machine where they are encrypted.
Pipe through clip.exe or Set-ClipBoard to place in the Windows clipboard

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

The password for the -mailUsername argument. If the %_MVal12% environment variable is set then its contents are used for the password

.PARAMETER hashedMailPassword

The hashed mail password returned from a previous call to the script with -encryptPassword. Will only work on the same machine and for the same user that ran the script to encrypt the password

.PARAMETER port

The port to use to communicate with the mail server

.PARAMETER nogridview

If not emailing then output the results to the pipeline rather than displaying in a gridview

.PARAMETER logFile

The full path to a log file to append the output of the script to

.EXAMPLE

'.\Check SQL account expiry.ps1' -sqlServer SQL01\instance01 -recipients guyl@hell.com -mailServer smtp.hell.com -mailOnFail -expireWithinDays 10

This will connect to the specified SQL server instance as the user running the script and email a list of all SQL accounts which expire within the next 10 days to the given recipient.
If the script fails to connect to the SQL server then an email will also be sent containing details of the error.

.EXAMPLE

'.\Check SQL account expiry.ps1' -sqlServer SQL01\instance01 -accountName bob -includeDisabled

This will connect to the specified SQL server instance as the user running the script and display all SQL accounts with "bob" in the name which either expire within the next 7 days or are disabled

.EXAMPLE

'.\Check SQL account expiry.ps1' -encryptPassword -password thepassword

This will encrypt the given password and output its encrypted form so that it can be passed as the argument to the -hashedPassword option to avoid having to specify the password on the command line.
The encrypted password will only work for the same user that encrypted it and on the same machine.

.NOTES

Place in a scheduled task where the action should be set to start a program, namely powershell.exe, with the arguments starting with '-file "C:\Scripts\Check SQL account expiry.ps1"' and then having the rest of the required arguments after this.

#>

[CmdletBinding()]

Param
(
    [Parameter(mandatory=$true, ParameterSetName='Query')]
    [string]$sqlServer ,
    [string]$accountName  ,
    [int]$expireWithinDays = 7 ,
    [string]$username ,
    [string]$password ,
    [string]$hashedPassword ,
    [switch]$mailOnFail ,
    [switch]$includeDisabled ,
    [Parameter(mandatory=$true, ParameterSetName='Encrypt')]
    [switch]$encryptPassword ,
    [string]$mailServer ,
    [string]$proxyMailServer = 'localhost' ,
    [switch]$noSSL ,
    [string]$subject = "SQL Accounts with passwords expiring in the next $expireWithinDays days on server $sqlServer" ,
    [string]$from = "$($env:computername)@$($env:userdnsdomain)" ,
    [string[]]$recipients ,
    [string]$mailUsername ,
    [string]$mailPassword ,
    [string]$mailHashedPassword ,
    [int]$port ,
    [switch]$nogridview ,
    [string]$logFile
)

## if dates aren't set, so year 1900, just output a dash
Function Get-RealDate
{
    Param
    (
        [datetime]$date
    )

    if( $date.Year -lt 1980 -and $date.Year -ge 100 )
    {
        '-'
    }
    else
    {
        Get-Date -Date $date -Format G
    }
}

## https://docs.microsoft.com/en-us/sql/t-sql/functions/loginproperty-transact-sql?view=sql-server-2017
$sqlQuery = @'
    SELECT  name AS 'AccountName'
	    ,LOGINPROPERTY(name, 'BadPasswordCount') AS 'BadPasswordCount'
	    ,LOGINPROPERTY(name, 'BadPasswordTime') AS 'BadPasswordTime'
	    ,LOGINPROPERTY(name, 'DaysUntilExpiration') AS 'DaysUntilExpiration'
	    ,LOGINPROPERTY(name, 'DefaultDatabase') AS 'DefaultDatabase'
	    ,LOGINPROPERTY(name, 'DefaultLanguage') AS 'DefaultLanguage'
	    ,LOGINPROPERTY(name, 'HistoryLength') AS 'HistoryLength'
	    ,LOGINPROPERTY(name, 'IsExpired') AS 'IsExpired'
	    ,LOGINPROPERTY(name, 'IsLocked') AS 'IsLocked'
	    ,LOGINPROPERTY(name, 'IsMustChange') AS 'IsMustChange'
	    ,LOGINPROPERTY(name, 'LockoutTime') AS 'LockoutTime'
	    ,LOGINPROPERTY(name, 'PasswordLastSetTime') AS 'PasswordLastSetTime'
	    ,is_expiration_checked, is_disabled , create_date , modify_date
    FROM    sys.sql_logins
    WHERE   is_policy_checked = 1
'@

if( $encryptPassword )
{
    if( ! $PSBoundParameters[ 'password' ] -and ! ( $password = $env:SQLval1 ) )
    {
        Throw 'Must specify the password when encrypting via -password or SQLval1 environment variable'
    }
    
    ConvertTo-SecureString -AsPlainText -String $password -Force | ConvertFrom-SecureString
    Exit 0
}

try
{
    if( ! [string]::IsNullOrEmpty( $logFile ) )
    {
        Start-Transcript -Path $logFile -Append
    }

    [hashtable]$mailParams = $null

    if( ( ! [string]::IsNullOrEmpty( $proxyMailServer )  -or ! [string]::IsNullOrEmpty( $mailServer ) ) -and $recipients.Count )
    {
        if( $recipients -and $recipients.Count -eq 1 -and $recipients[0].IndexOf(',') -ge 0 ) ## fix scheduled task not passing array correctly
        {
            $recipients = $recipients -split ','
        }

        ## Set mail parameters in case we have to send an email that we can't connect to SQL
        $mailParams = @{
                'To' =  $recipients
                'SmtpServer' = $mailServer
                'From' =  $from
                'UseSsl' = ( ! $noSSL ) }
        if( $PSBoundParameters[ 'port' ] )
        {
            $mailParams.Add( 'Port' , $port )
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

    #region SQL

    $connectionString = "Data Source=$sqlServer;"

    if( ! [string]::IsNullOrEmpty( $username ) )
    {
        ## will only work for SQL auth, Windows must be done via RunAs
        $connectionString += "Integrated Security=no;"
        $connectionString += "uid=$username;"
        if( ! $PSBoundParameters[ 'password' ] )
        {
            if( ! ( $password = $env:SQLval1 ) )
            {
                if( $PSBoundParameters[ 'hashedPassword' ] )
                {
                    $password = [Runtime.interopServices.marshal]::PtrToStringAuto( [Runtime.Interopservices.Marshal]::SecurestringToBstr( ( $hashedPassword|ConvertTo-SecureString ) ) )
                }
                else
                {
                    Throw 'Must specify password'
                }
            }
        }
        $connectionString += "pwd=$password;"
    }
    else
    {
        $connectionString += "Integrated Security=SSPI;"
    }

    $conn = New-Object System.Data.SqlClient.SqlConnection
    $conn.ConnectionString = $connectionString

    try
    {
        $conn.Open()
    }
    catch
    {
        Write-Error "Failed to connect with `"$connectionString`" : $($_.Exception.Message)"
        if( $mailOnFail -and $mailParams )
        {
            [string]$subject = "Failed to connect to $sqlServer as user "
            $subject += $( if( ! [string]::IsNullOrEmpty( $username ) )
            {
                $username
            }
            else
            {
                $env:USERNAME
            })
            $mailParams.Add( 'Subject' , $subject )
            $mailParams.Add( 'Body' , $_.Exception.Message )
            $mailParams.Add( 'BodyAsHtml' , $false )
          
            if( $PSBoundParameters[ 'proxyMailServer' ] )
            {
                Invoke-Command -ComputerName $proxyMailServer -ScriptBlock { [hashtable]$mailParams = $using:mailParams ; Send-MailMessage @mailParams }
            }
            else
            {
                Send-MailMessage @mailParams 
            }
        }

        Exit 1
    }

    $cmd = New-Object System.Data.SqlClient.SqlCommand
    $cmd.connection = $conn

    ## Now query the database
    $cmd.CommandText = $sqlQuery

    [datetime]$startTime = Get-Date

    $sqlreader = $cmd.ExecuteReader()

    [datetime]$endTime = Get-Date

    Write-Verbose "Got $($sqlreader.FieldCount) columns from query in $(($endTime - $startTime).TotalSeconds) seconds"

    $datatable = New-Object System.Data.DataTable
    $datatable.Load( $sqlreader )

    $sqlreader.Close()
    $conn.Close()

    #endregion

    Write-Verbose "Retrieved $($datatable.Rows.Count) rows"

    [array]$results = $null

    if( $datatable.Rows -and $datatable.Rows.Count -gt 0 )
    {
        $results = @( $datatable | ForEach-Object `
        {
            $item = $_
            if( ! $accountName -or $item.AccountName -match $accountName )
            {
                if( ( $item.is_expiration_checked -and $item.DaysUntilExpiration -le $expireWithinDays ) `
                    -or $item.IsExpired -or $item.IsLocked -or $item.IsMustChange -or ( $includeDisabled -and $item.is_disabled ) )
                {
                    $item
                }
            }
        })
    }

    if( $results -and $results.Count )
    {   
        if( ( ! [string]::IsNullOrEmpty( $proxyMailServer )  -or ! [string]::IsNullOrEmpty( $mailServer ) ) -and $recipients.Count )
        {
            ## Ok, so it's hard coded! Puts a border on our table
            [string]$style = '<style>BODY{font-family: Arial; font-size: 10pt;}'
            $style += "TABLE{border: 1px solid black; border-collapse: collapse;}"
            $style += "TH{border: 1px solid black; background: #dddddd; padding: 5px; }"
            $style += "TD{border: 1px solid black; padding: 5px; }"
            $style += "</style>"

            Write-Verbose "Emailing to $($recipients -join ',') via $mailServer"

            [string]$htmlBody = $results | Sort-Object -Property 'DaysUntilExpiration' | Select-Object -Property 'AccountName',@{n='Expiry Date';e={if( $_.DaysUntilExpiration ) { Get-Date -Date (Get-Date).AddDays( $_.DaysUntilExpiration ) -Format d } elseif( $_.DaysUntilExpiration ) { 'EXPIRED' }}},
                'DaysUntilExpiration','IsExpired','IsLocked','IsMustChange',@{n='Last Lockout Time';e={Get-RealDate -date $_.LockoutTime}},@{n='Password Last Set';e={Get-RealDate -date $_.PasswordLastSetTime}},@{n='Last Bad Password';e={Get-RealDate -date $_.BadPasswordTime}},
                    'is_expiration_checked','is_disabled' | ConvertTo-Html -Head $style

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
        Write-Warning 'No results returned'
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
# SIG # Begin signature block
# MIINKgYJKoZIhvcNAQcCoIINGzCCDRcCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUGxQkUHCIdi8PPdoMNdwu2MwG
# AlmgggpsMIIFMDCCBBigAwIBAgIQBAkYG1/Vu2Z1U0O1b5VQCDANBgkqhkiG9w0B
# AQsFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVk
# IElEIFJvb3QgQ0EwHhcNMTMxMDIyMTIwMDAwWhcNMjgxMDIyMTIwMDAwWjByMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQg
# Q29kZSBTaWduaW5nIENBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA
# +NOzHH8OEa9ndwfTCzFJGc/Q+0WZsTrbRPV/5aid2zLXcep2nQUut4/6kkPApfmJ
# 1DcZ17aq8JyGpdglrA55KDp+6dFn08b7KSfH03sjlOSRI5aQd4L5oYQjZhJUM1B0
# sSgmuyRpwsJS8hRniolF1C2ho+mILCCVrhxKhwjfDPXiTWAYvqrEsq5wMWYzcT6s
# cKKrzn/pfMuSoeU7MRzP6vIK5Fe7SrXpdOYr/mzLfnQ5Ng2Q7+S1TqSp6moKq4Tz
# rGdOtcT3jNEgJSPrCGQ+UpbB8g8S9MWOD8Gi6CxR93O8vYWxYoNzQYIH5DiLanMg
# 0A9kczyen6Yzqf0Z3yWT0QIDAQABo4IBzTCCAckwEgYDVR0TAQH/BAgwBgEB/wIB
# ADAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwMweQYIKwYBBQUH
# AQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQwYI
# KwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFz
# c3VyZWRJRFJvb3RDQS5jcnQwgYEGA1UdHwR6MHgwOqA4oDaGNGh0dHA6Ly9jcmw0
# LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwOqA4oDaG
# NGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RD
# QS5jcmwwTwYDVR0gBEgwRjA4BgpghkgBhv1sAAIEMCowKAYIKwYBBQUHAgEWHGh0
# dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwCgYIYIZIAYb9bAMwHQYDVR0OBBYE
# FFrEuXsqCqOl6nEDwGD5LfZldQ5YMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6en
# IZ3zbcgPMA0GCSqGSIb3DQEBCwUAA4IBAQA+7A1aJLPzItEVyCx8JSl2qB1dHC06
# GsTvMGHXfgtg/cM9D8Svi/3vKt8gVTew4fbRknUPUbRupY5a4l4kgU4QpO4/cY5j
# DhNLrddfRHnzNhQGivecRk5c/5CxGwcOkRX7uq+1UcKNJK4kxscnKqEpKBo6cSgC
# PC6Ro8AlEeKcFEehemhor5unXCBc2XGxDI+7qPjFEmifz0DLQESlE/DmZAwlCEIy
# sjaKJAL+L3J+HNdJRZboWR3p+nRka7LrZkPas7CM1ekN3fYBIM6ZMWM9CBoYs4Gb
# T8aTEAb8B4H6i9r5gkn3Ym6hU/oSlBiFLpKR6mhsRDKyZqHnGKSaZFHvMIIFNDCC
# BBygAwIBAgIQBUgpPCsQ1AN/x7Nr/GVKWTANBgkqhkiG9w0BAQsFADByMQswCQYD
# VQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGln
# aWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQgQ29k
# ZSBTaWduaW5nIENBMB4XDTE5MDQwOTAwMDAwMFoXDTIwMDQxMzEyMDAwMFowcTEL
# MAkGA1UEBhMCR0IxEjAQBgNVBAcTCVdha2VmaWVsZDEmMCQGA1UEChMdU2VjdXJl
# IFBsYXRmb3JtIFNvbHV0aW9ucyBMdGQxJjAkBgNVBAMTHVNlY3VyZSBQbGF0Zm9y
# bSBTb2x1dGlvbnMgTHRkMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA
# 6TIrjnIU6zyvuEl3U4MgxLi98idcMWzVeWifcQl9EUN4MxBisMI9hNrnv5jvheOk
# suPpeyIIcZ1mDpuzSgWndchERk1uRdIl0MNruQg81R1d6h3PTGGSm8/cRDjBroZJ
# 0kFawFsBTxkMjHX9alOEXQT3xPr5WpsKeuRybNTKfpcbO25deK7wgO4mvU/U545n
# p92AVHX07ONtXIAF8PXmT1a25PfqKAnt81xdfhOVVI54Ru2mT8lxODSASk9JCBjN
# i5RyIowxDYAWx4UiIfdNgatC7foiDlVGxirK+5+e/jpEor1Bj98546ibVrUpLuuV
# 1nqkdxpCQAV/RFL6+kc36QIDAQABo4IBxTCCAcEwHwYDVR0jBBgwFoAUWsS5eyoK
# o6XqcQPAYPkt9mV1DlgwHQYDVR0OBBYEFHyupN5JxjpSZ2qGWJKVdjgDE6G6MA4G
# A1UdDwEB/wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzB3BgNVHR8EcDBuMDWg
# M6Axhi9odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vc2hhMi1hc3N1cmVkLWNzLWcx
# LmNybDA1oDOgMYYvaHR0cDovL2NybDQuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJl
# ZC1jcy1nMS5jcmwwTAYDVR0gBEUwQzA3BglghkgBhv1sAwEwKjAoBggrBgEFBQcC
# ARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAIBgZngQwBBAEwgYQGCCsG
# AQUFBwEBBHgwdjAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29t
# ME4GCCsGAQUFBzAChkJodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRTSEEyQXNzdXJlZElEQ29kZVNpZ25pbmdDQS5jcnQwDAYDVR0TAQH/BAIwADAN
# BgkqhkiG9w0BAQsFAAOCAQEAroNPk8xHRUMXkqvWN0sVjMNoh0iYFVhbz6Tylfuq
# UxjOLLW6WxTc1OtvLVOBUJHQabGgv5BadDD1jzrlEVauE42Fvd1rjXpAbQ2pTCK6
# JoRaMfon3wD3dNC5g4zqXWj5pM6IzyxcpH7k6u9qdC1hgrgpXoEvgkDuNrhmY5dO
# pJxjiv9uW1icUA8LWCfGuesCZMe5JYP7iGMUxw/ANC3WqhUImEr534fsb5rY28dX
# GvAyQyzdA8Lu62vDrUdU+PN3ccudpumRxM9q4r/a7TvUiEOvOJKSyT4jiJX2F7UU
# qk4t8ipKsoezPuKWfQIhQ2LcLT6GCGsIIVuJLuVJC/MOWTGCAigwggIkAgEBMIGG
# MHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsT
# EHd3dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJl
# ZCBJRCBDb2RlIFNpZ25pbmcgQ0ECEAVIKTwrENQDf8eza/xlSlkwCQYFKw4DAhoF
# AKB4MBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisG
# AQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcN
# AQkEMRYEFCjKHirqxCW9y+kpaMvrKhAnJcU6MA0GCSqGSIb3DQEBAQUABIIBAIAY
# huDkdSvSY1QijOSrIx+biauxG028gGXQOpbZ7D572Hs3geQgJotDfftA892ud9zl
# rXbeRoVuCrUXsD4P0ugPm25jhoc+S1J3vE9v00oyDEsGLKwTgSgX9CK0g7UeBZzq
# hea+fO1h075+BWL3CyFDUBrXFzm6LHCvaBsPS2wBDxr/2viPQdOhdbfh0Gg/4rET
# Lmi99KbYE40UAeqfp4l23LBk/UsQmnT3HmI3Al+pYhf646H7+T8FbUzyJhdeLXHd
# Ugh4BCtMnXq9IJAm0g6cseHMP1AMHUQv/+XGJ3qmAOTV/qjc4c8tvknv0Xcn1zW5
# VzI9dfDC1jDzzNocyH8=
# SIG # End signature block
