#requires -version 3

<#
.SYNOPSIS

Check and optionally change printer spooler service on domain controllers

.PARAMETER domainControllers

List of domain controllers to check. If not specified, will try and get list of domain controllers for the domain

.PARAMETER disable

Stop the service and disable it

.PARAMETER enable

Set service to automatic start and start it

.PARAMETER force

Force the enable/disable operation even when service is already in desired state

.PARAMETER serviceName

Name of the service to check/operate on. Do not change

.EXAMPLE 

& '.\Check DC spooler.ps1' -disable -verbose

Disable the spooler service on all domain controllers in the domain, prompting for confirmation before operating

.EXAMPLE 

& '.\Check DC spooler.ps1' -enable domainControllers grl-dc01,grl-dc02 -confirm:$false -force

Enable the spooler service on the two machines listed without asking for confirmation and regardless of the service's current state

.NOTES

https://msrc.microsoft.com/update-guide/en-US/vulnerability/CVE-2021-1675

Modification History

@guyrleech 02/07/2021   Initial release.
                        Code around error because Server 2012R2 not providing StartupTye property from Get-Service

#>

<#
Copyright (c) 2021 Guy Leech

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, 
including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

[CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='High')]

Param
(
    [string[]]$domainControllers ,
    [switch]$disable ,
    [switch]$enable ,
    [switch]$force ,
    [string]$serviceName = 'Spooler'
)

if( $enable -and $disable )
{
    Throw "Cannot enable and disable in the same script invocation"
}

[string]$domainName = $env:USERDOMAIN

if( ! $PSBoundParameters[ 'domainControllers' ] )
{
    if( ! ( $domain = [directoryServices.ActiveDirectory.Domain]::GetCurrentDomain() ) )
    {
        Throw "Failed to get the current domain"
    }

    $domainName = $domain.Name
    $domainControllers = @( $domain.FindAllDomainControllers() )
}

if( ! $domainControllers -or ! $domainControllers.Count )
{
    Throw "Failed to get any domain controllers for domain $domainName"
}

Write-Verbose -Message "Got $($domainControllers.Count) domain controllers in domain $domainName"
$domainControllers | Write-Verbose

## Invoke command will run some in parallel and Get-WinEvent will only take a single machine
[array]$spoolers = @( Invoke-Command -ComputerName $domainControllers -ScriptBlock { Get-Service -Name $using:serviceName} )

if( $spoolers -and $spoolers.Count )
{
    [array]$disabled = @()
    $running  = @( $spoolers.Where( { $_.Status    -eq 'Running' } ) )
    $disabled = @( $spoolers.Where( { $_.PSObject.Properties[ 'StartType' ] -and $_.StartType -eq 'Disabled' } ) )
    Write-Verbose -Message "Found $($spoolers.Count) $serviceName services on $($domainControllers.Count) domain controllers with $($running.Count) running & $($disabled.Count) disabled"
    
    [string]$operation = $null
    $collection = $null
    if( $disable )
    {
        $operation = 'Stop and disable'
        $collection = $running
    }
    elseif( $enable )
    {
        $operation = 'Enable and start'
        $collection = $disabled
    }

    [int]$errors = 0

    if( $null -ne $collection -and ( ( $collection.Count -gt 0 -or $force ) -and ( $enable -or $disable ) ) )
    {
        if( $force )
        {
            $collection = $spoolers
        }
        ForEach( $spooler in $collection )
        {
            if( $PSCmdlet.ShouldProcess( $spooler.PSComputername , "$operation $serviceName" ) )
            {
                if( $disable )
                {
                    if( ! ( $result = Invoke-Command -ComputerName $spooler.PSComputerName -ScriptBlock { Stop-Service $using:spooler.Name -PassThru } ) -or $result.Status -ne 'Stopped' )
                    {
                        if( $result )
                        {
                            Write-Error -Exception "$serviceName failed to stop on $($spooler.PSComputerName) - status is $($result.Status)"
                        }
                        $errors++
                        Add-Member -InputObject $spooler -MemberType NoteProperty -Name OperationFailure -Value $true
                    }
                    if( ! ( $result = Set-Service -ComputerName $spooler.PSComputerName -Name $spooler.Name -StartupType Disabled -PassThru ) -or $result.StartType -ne 'Disabled' )
                    {
                        if( $result )
                        {
                            Write-Error -Exception "Failed to change startup of $serviceName on $($spooler.PSComputerName) - startup is $($result.StartType)"
                        }
                        $errors++
                        Add-Member -InputObject $spooler -MemberType NoteProperty -Name OperationFailure -Value $true -Force
                    }
                }
                elseif( $enable )
                {
                    if( ! ( $result = Set-Service -ComputerName $spooler.PSComputerName -Name $spooler.Name -StartupType Automatic -PassThru ) -or $result.StartType -ne 'Automatic' )
                    {
                        if( $result )
                        {
                            Write-Error -Exception "Failed to change startup of $serviceName on $($spooler.PSComputerName) - startup is $($result.StartType)"
                        }
                        $errors++
                        Add-Member -InputObject $spooler -MemberType NoteProperty -Name OperationFailure -Value $true -Force
                    }
                    if( ! ( $result = Invoke-Command -ComputerName $spooler.PSComputerName -ScriptBlock { Start-Service $using:spooler.Name -PassThru } ) -or $result.Status -ne 'Running' )
                    {
                        if( $result )
                        {
                            Write-Error -Exception "$serviceName failed to start on $($spooler.PSComputerName) - status is $($result.Status)"
                        }
                        $errors++
                        Add-Member -InputObject $spooler -MemberType NoteProperty -Name OperationFailure -Value $true -Force
                    }
                }
            }
        }
        if( $errors -gt 0 )
        {
            Write-Output -InputObject "Encountered $errors errors on the following machines:"
            $collection.Where( { $_.PSObject.Properties[ 'OperationFailure' ] } ) | Select-Object -ExpandProperty PSComputerName
        }
        else
        {
            Write-Verbose -Message "Operations appeared to complete ok on $($collection.Count) machines"
        }
    }
}
else
{
    Write-Error -Message "No $serviceName services found on $($domainControllers.Count) domain controllers"
}

# SIG # Begin signature block
# MIINRQYJKoZIhvcNAQcCoIINNjCCDTICAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUz1qK0ThiEkJ9EtmVNiHa5Ztr
# RTCgggqHMIIFMDCCBBigAwIBAgIQBAkYG1/Vu2Z1U0O1b5VQCDANBgkqhkiG9w0B
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
# T8aTEAb8B4H6i9r5gkn3Ym6hU/oSlBiFLpKR6mhsRDKyZqHnGKSaZFHvMIIFTzCC
# BDegAwIBAgIQBP3jqtvdtaueQfTZ1SF1TjANBgkqhkiG9w0BAQsFADByMQswCQYD
# VQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGln
# aWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQgQ29k
# ZSBTaWduaW5nIENBMB4XDTIwMDcyMDAwMDAwMFoXDTIzMDcyNTEyMDAwMFowgYsx
# CzAJBgNVBAYTAkdCMRIwEAYDVQQHEwlXYWtlZmllbGQxJjAkBgNVBAoTHVNlY3Vy
# ZSBQbGF0Zm9ybSBTb2x1dGlvbnMgTHRkMRgwFgYDVQQLEw9TY3JpcHRpbmdIZWF2
# ZW4xJjAkBgNVBAMTHVNlY3VyZSBQbGF0Zm9ybSBTb2x1dGlvbnMgTHRkMIIBIjAN
# BgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAr20nXdaAALva07XZykpRlijxfIPk
# TUQFAxQgXTW2G5Jc1YQfIYjIePC6oaD+3Zc2WN2Jrsc7bj5Qe5Nj4QHHHf3jopLy
# g8jXl7Emt1mlyzUrtygoQ1XpBBXnv70dvZibro6dXmK8/M37w5pEAj/69+AYM7IO
# Fz2CrTIrQjvwjELSOkZ2o+z+iqfax9Z1Tv82+yg9iDHnUxZWhaiEXk9BFRv9WYsz
# qTXQTEhv8fmUI2aZX48so4mJhNGu7Vp1TGeCik1G959Qk7sFh3yvRugjY0IIXBXu
# A+LRT00yjkgMe8XoDdaBoIn5y3ZrQ7bCVDjoTrcn/SqfHvhEEMj1a1f0zQIDAQAB
# o4IBxTCCAcEwHwYDVR0jBBgwFoAUWsS5eyoKo6XqcQPAYPkt9mV1DlgwHQYDVR0O
# BBYEFE16ovlqIk5uX2JQy6og0OCPrsnJMA4GA1UdDwEB/wQEAwIHgDATBgNVHSUE
# DDAKBggrBgEFBQcDAzB3BgNVHR8EcDBuMDWgM6Axhi9odHRwOi8vY3JsMy5kaWdp
# Y2VydC5jb20vc2hhMi1hc3N1cmVkLWNzLWcxLmNybDA1oDOgMYYvaHR0cDovL2Ny
# bDQuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1jcy1nMS5jcmwwTAYDVR0gBEUw
# QzA3BglghkgBhv1sAwEwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNl
# cnQuY29tL0NQUzAIBgZngQwBBAEwgYQGCCsGAQUFBwEBBHgwdjAkBggrBgEFBQcw
# AYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tME4GCCsGAQUFBzAChkJodHRwOi8v
# Y2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRTSEEyQXNzdXJlZElEQ29kZVNp
# Z25pbmdDQS5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAQEAU9zO
# 9UpTkPL8DNrcbIaf1w736CgWB5KRQsmp1mhXbGECUCCpOCzlYFCSeiwH9MT0je3W
# aYxWqIpUMvAI8ndFPVDp5RF+IJNifs+YuLBcSv1tilNY+kfa2OS20nFrbFfl9QbR
# 4oacz8sBhhOXrYeUOU4sTHSPQjd3lpyhhZGNd3COvc2csk55JG/h2hR2fK+m4p7z
# sszK+vfqEX9Ab/7gYMgSo65hhFMSWcvtNO325mAxHJYJ1k9XEUTmq828ZmfEeyMq
# K9FlN5ykYJMWp/vK8w4c6WXbYCBXWL43jnPyKT4tpiOjWOI6g18JMdUxCG41Hawp
# hH44QHzE1NPeC+1UjTGCAigwggIkAgEBMIGGMHIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAv
# BgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EC
# EAT946rb3bWrnkH02dUhdU4wCQYFKw4DAhoFAKB4MBgGCisGAQQBgjcCAQwxCjAI
# oAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIB
# CzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFO/XIn6Q/2wjRP8iBBgC
# MmxQp9eyMA0GCSqGSIb3DQEBAQUABIIBAFF5oZKQV2aYYjmr9gAqLhDvHJTzrWx4
# OoTi1TY5XGmYxuYtarPn+prsqdKDU3bgZvRYPYA5NP8vg0emhr58WGfomZokUNjQ
# qUwbJBb0x4Qg/Vn9wFCgssseEDYBUmMUOoHJbIaSmds0qgJCB9SjdzTc6QPsF7ki
# 6qyjpCJTKc7nFgw22BcnUs2PdRrxNtR2h8+RbAF7X3htHH0T44LSzowgS2+CD0/f
# PCG0JeCrQmGF6xljRx4sVHWXOdAIFibZGzXN7mkPmx1RNhl4d0agv8OodZjkfSXl
# 858mJXWPGhWeWzZ+8yLWSbm/Eb91owf9pxb2EC8SP+6M0iYT83D175o=
# SIG # End signature block
