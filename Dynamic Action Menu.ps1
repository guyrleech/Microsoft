#required -version 3

<#
.SYNOPSIS
    Create a form with buttons for actions coming from a json config file

.PARAMETER configFile

    JSON configuration file to read actions from (Name,Command,Arguments,Options). Will read from registry if not specified

.PARAMETER title

    Title of the dialog

.PARAMETER buttonHeight

    Height of the buttons in pixels

.PARAMETER buttonWidthFactor

    Multiplier to produce button width from longest string in the buttons

.PARAMETER font

    Font family to use for the text

.PARAMETER fontSize

    Font size to use

.PARAMETER buttonGap

    Gap in pixels between buttons
    
.PARAMETER buttonOffset

    Offset in pixels from left hand side of form to where button starts 

.PARAMETER loop

    Always show the menu, do not allow it to be closed and if it is, show it again
    
.PARAMETER logoff

    Logoff when the main window is closed

.EXAMPLE

 & '.\Dynamic Action Menu.ps1'

 Read the location of the JSON configuration file from the registry and show a window with the actions defined in the config file

.NOTES

    Modification History:

    @guyrleech  03/07/2021  Initial release
#>

<#
Copyright © 2021 Guy Leech

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, 
including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

[CmdletBinding()]

Param
(
    [string]$configFile ,
    [string]$title = 'Self Service Utility' ,
    [double]$buttonHeight = 65 ,
    [double]$buttonWidthFactor = 20 ,
    [double]$buttonOffset = 25 ,
    [int]$buttonGap = 35 ,
    [string]$font = 'Segoe UI' ,
    [float]$fontSize = 12 ,
    [switch]$logoff ,
    [switch]$loop
)

Function Invoke-Action
{
    Param
    (
        $control ,
        [string]$name ,
        [string]$command ,
        [string]$arguments ,
        [string[]]$options
    )

    ## TODO add log file

    Write-Verbose "Invoke-Action -name $name -command $command -arguments $arguments -options $options"

    [hashtable]$processArguments = @{
        'PassThru' = $true
        'FilePath' = $command 
    }

    if( ! [string]::IsNullOrEmpty( $arguments ) )
    {
        $processArguments.Add( 'ArgumentList' , [Environment]::ExpandEnvironmentVariables( $arguments ) )
    }
        
    if( $options -contains 'Hide' )
    {
        $processArguments.Add( 'WindowStyle' , 'Hidden' )
        $processArguments.Add( 'NoNewWindow' , $true )
    }

    $process = $null
    $process = Start-Process @processArguments

    if( ! $process )
    {        
        [void][Windows.MessageBox]::Show( "Failed to run $($processArguments.FilePath)" , 'Action Error' , 'Ok' ,'Exclamation' )
    }
    else
    {
        Write-Verbose -Message "$(Get-Date -Format G): pid $($process.Id) - $($process.Name) `"$($processArguments[ 'ArgumentList'] )`""

        if( $options -contains 'Wait' )
        {
            $process.WaitForExit()
            Write-Verbose -Message "$(Get-Date -Format G -Date $process.ExitTime): pid $($process.Id) - $($process.Name) exited with status $($process.ExitCode)"
        }
    }
}

if( ! $PSBoundParameters[ 'configFile' ] )
{
    if( $configFile = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Guy Leech" -Name 'Dynamic Action Menu Configuration File' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty 'Dynamic Action Menu Configuration File') )
    {
        $configFile = [System.Environment]::ExpandEnvironmentVariables( $configFile )
    }
    Write-Verbose -Message "Configuration file from registry is `"$configFile`""
}

Add-Type -AssemblyName PresentationFramework,System.Windows.Forms

if( [string]::IsNullOrEmpty( $configFile ))
{
    [void][Windows.MessageBox]::Show( "No configuration file" , 'Menu Error' , 'Ok' ,'Exclamation' )
    Throw "Failed to get a configuration file name"
}

if( ! ($actions = Get-Content -Path $configFile | ConvertFrom-Json ) -or ! $actions.Count )
{
    [void][Windows.MessageBox]::Show( "Failed to read configuration file" , 'Menu Error' , 'Ok' ,'Exclamation' )
    Throw "Failed to get any actions from `"$configFile`""
}

$form = New-Object -TypeName Windows.Forms.Form

[int]$formWidth = 0
[int]$formHeight = 0
[int]$x = $buttonOffset
[int]$y = $buttonGap
[int]$longestString = 0

ForEach( $action in $actions )
{  
    if( $action.Name.Length -gt $longestString )
    {
        $longestString = $action.Name.Length
    }
}

[double]$buttonWidth = $longestString * $buttonWidthFactor

$buttonFont = New-Object -TypeName System.Drawing.Font( $font , $fontSize )

$form.Font = $buttonFont

ForEach( $action in $actions )
{  
    $button = New-Object -TypeName Windows.Forms.Button
    $button.Text = $action.Name
    $button.Font = $buttonFont
    
    $button.Location = New-Object -TypeName Drawing.Point( $x , $y )
    $y += $buttonHeight + $buttonGap
    $button.Width = $buttonWidth
    $button.Height = $buttonHeight
    
    [scriptblock]$clickAction = [scriptblock]::Create( "Invoke-Action -control `$_ -name `"$($action.Name)`" -command `"$($action.command)`" -arguments `"$(($action|Select-Object -ExpandProperty arguments -ErrorAction SilentlyContinue) -replace '"' , '`"`"')`" -options `"$(($action|Select-Object -ExpandProperty options -ErrorAction SilentlyContinue) -split ',')`"" )
    $button.add_click( $clickAction )
    $form.Controls.Add( $button )
}

$formHeight = $y + $buttonGap
$formWidth = ($buttonOffset * 2) + $buttonwidth

$form.Size = New-Object System.Drawing.Size($formWidth,$formHeight)
$form.AutoSize = $true
$form.Text = $title

if( $loop )
{
    $form.ControlBox = $false
    $form.add_Closing({
        $_.Cancel = $true
    })
}

do
{
    $answer = $form.ShowDialog()
}
while( $loop )

if( $logoff )
{
    logoff.exe
}

# SIG # Begin signature block
# MIINRQYJKoZIhvcNAQcCoIINNjCCDTICAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU2gd0YklZF5oqvNbtICdl2eZO
# 9e2gggqHMIIFMDCCBBigAwIBAgIQBAkYG1/Vu2Z1U0O1b5VQCDANBgkqhkiG9w0B
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
# CzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFPyERLQ5/OllddzTcS4h
# 1RfAM19KMA0GCSqGSIb3DQEBAQUABIIBAKi9cHCeiccjzAAboEbh5J8Dh5waRUGH
# 9xOBCX65npojN1QtCeSSPEpCNBu3CCBsfZL1GJZsF2YhDpb5Ywq4lerTx0J2GsZc
# Zg3J/wzS8/nRjlGxythTihGe6aSebGIth8IDePAl9Wj7zctQ9ngg5VjM4tEippJ8
# /rhPN6ShW1HPH1dKpCtlwRarn0PnuxcTuG/7pZPKa0csqYY/MIyG+M0xQBOVFkUb
# V2LGFGjZGqievE9Ebr1RbxQh0kuiXOJ8K1l3X7yCvV49nSqCQbtayE0judk4Uj9E
# /Qmr08JoEcemX08rjxRr1k/fymfFy3vMH2Me0HFAun3KDGwT33fzxcE=
# SIG # End signature block
