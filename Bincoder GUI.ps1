#requires -version 5

<#
.SYNOPSIS

Base64 encode/decode clipboard contents to/from specified file either via GUI or command line options

.DESCRIPTION

Useful to transfer binary files to/from a remote Citrix or RDSH session when client drive mapping, cloud file services, email attachments, etc are blocked

.PARAMETER filename

The name of the file to encode the contents of if -decode not specified or the name of the file to write the decoded content to if -decode is used

.PARAMETER decode

Take the base64 encoded contents of the clipboard, decode it and write to the file specified in -filename otherwise base64 encode the contents of the file specified by -filename

.PARAMETER nochecksum

Do not compute and display a checksum of the file

.EXAMPLE

& '.\Bincoder GUI.ps1'

Run the script so a GUI is displayed to allow selection of the file and then press the "encode" button on one system and then the "decode" button on the other system

.EXAMPLE

& '.\Bincoder GUI.ps1' -filename c:\temp\tools.zip -encode

Base64 encode the contents of the file c:\temp\tools.zip and place on the windows clipboard

.EXAMPLE

& '.\Bincoder GUI.ps1' -filename c:\temp\tools.zip -decode

Take the contents of the windows clipboard, from a prior invocation of this script without -encode, decode the base64 text and write to the file c:\temp\tools.zip

.NOTES

Modification history:

  30/11/19   @guyrleech  First public release

#>

[CmdletBinding()]

Param
(
    [string]$filename  ,
    [switch]$decode ,
    [switch]$noChecksum
)

[string]$mainwindowXAML = @'
<Window x:Class="MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Bincoder_GUI"
        mc:Ignorable="d"
        Title="Base64 Encoder/Decoder via Clipboard" Height="235.385" Width="616.538">
    <Grid Margin="0,0,22.5,88.269">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="37*"/>
            <ColumnDefinition Width="96*"/>
            <ColumnDefinition Width="454*"/>
        </Grid.ColumnDefinitions>
        <TextBox x:Name="txtFileName" HorizontalAlignment="Left" Height="36.346" Margin="27.346,24.923,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="429.231" Grid.Column="1" Grid.ColumnSpan="2"/>
        <Label Content="File" HorizontalAlignment="Left" Height="36.346" Margin="23.846,24.923,0,0" VerticalAlignment="Top" Width="53.5" Grid.ColumnSpan="2"/>
        <Button x:Name="btnEncode" Content="_Encode" HorizontalAlignment="Left" Height="36.347" Margin="27.346,71.203,0,0" VerticalAlignment="Top" Width="166.154" Grid.Column="1" Grid.ColumnSpan="2"/>
        <Button x:Name="btnDecode" Content="_Decode" HorizontalAlignment="Left" Height="36.347" Margin="137.922,71.203,0,0" VerticalAlignment="Top" Width="166.154" Grid.Column="2"/>
        <Button x:Name="btnFileChooser" Content="..." HorizontalAlignment="Left" Height="36.346" Margin="365.846,24.923,0,0" VerticalAlignment="Top" Width="57.308" Grid.Column="2"/>

    </Grid>
</Window>
'@

Function Load-GUI
{
    Param
    (
        [Parameter(Mandatory=$true)]
        $inputXaml
    )

    $form = $null
    $inputXML = $inputXaml -replace 'mc:Ignorable="d"' , '' -replace 'x:N' ,'N'  -replace '^<Win.*' , '<Window'
 
    [xml]$xaml = $inputXML

    if( $xaml )
    {
        $reader = New-Object -TypeName Xml.XmlNodeReader -ArgumentList $xaml

        try
        {
            $form = [Windows.Markup.XamlReader]::Load( $reader )
        }
        catch
        {
            Throw "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .NET is installed.`n$_"
        }
 
        $xaml.SelectNodes( '//*[@Name]' ) | ForEach-Object `
        {
            Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -Scope Global
        }
    }
    else
    {
        Throw "Failed to convert input XAML to WPF XML"
    }

    $form
}

Function Invoke-Operation
{
    [CmdletBinding()]

    Param
    (
        [string]$filename ,
        [switch]$decode ,
        $form
    )

    if( $decode )
    {
        ## if filename is not absolute then make it so
        if( $filename -notmatch '^\\\\' -and $filename -notmatch '^[a-z]:\\' )
        {
            $filename = Join-Path -Path (Get-Location -PSProvider FileSystem | Select-Object -ExpandProperty Path) -ChildPath $filename
            Write-Verbose -Message "File name is now `"$filename`""
        }
        if( ! ( Test-Path -Path $filename -ErrorAction SilentlyContinue ) `
            -or ( $form -and [Windows.MessageBox]::Show( "`"$filename`" already exists. Overwrite?" , 'Decoding Error' , 'YesNo' ,'Exclamation' ) -eq 'Yes' ) )
        {
            if( $form )
            {
                $oldCursor = $form.Cursor
                $form.Cursor = [Windows.Input.Cursors]::Wait
            }
            
            [string]$clipboard = Get-Clipboard -Format Text -TextFormatType Text -Raw
            [byte[]]$transmogrified = [System.Convert]::FromBase64String( $clipboard )

            [string]$message = $null 

            if( $transmogrified.Count )
            {
                $fileStream = New-Object System.IO.FileStream( $filename , [System.IO.FileMode]::Create , [System.IO.FileAccess]::Write )
                if( $fileStream )
                {
                    $fileStream.Write( $transmogrified , 0 , $transmogrified.Count )
                    $fileStream.Close()
                    if( ! $noChecksum )
                    {
                        $hash = Get-FileHash -Path $filename 
                        if( $hash )
                        {
                            $message = "`n$($hash.Algorithm) checksum $($hash.hash)"
                        }
                    }
                }
                elseif( $form )
                {
                    [void][Windows.MessageBox]::Show( "Failed to create file `"$filename`"" , 'Decoding Error' , 'Ok' , 'Error' )
                }
            }
            else
            {
                if( $form )
                {
                    [void][Windows.MessageBox]::Show( "Failed to decode clipboard contents ($([int]($clipboard.Length / 1KB))KB)" , 'Decoding Error' , 'Ok' , 'Error' )
                }
            }
            
            $transmogrified = $null
            $clipboard = $null

            $fileProperties = Get-ItemProperty -Path $filename
            [int]$bytesWritten = 0 
            if( $fileProperties )
            {
                $bytesWritten = $fileProperties.Length
            }
            
            $message = "Wrote $([int]($bytesWritten / 1KB))KB to $filename" + $message

            Write-Verbose -Message $message

            if( $form )
            {
                $form.Cursor = $oldCursor
                [void][Windows.MessageBox]::Show( $message , 'Decoding Complete' , 'Ok' , $(if( $bytesWritten ) { 'Information' } else { 'Error' } ) )
            }
        }
    }
    else ## encode
    {
        [byte[]]$data = [System.IO.File]::ReadAllBytes( $filename )
        if( $data -and $data.Count )
        {
            [System.Convert]::ToBase64String( $data ) | Set-Clipboard
            [string]$message = "$([int]($data.Count / 1KB))KB from `"$filename`" encoded and put on clipboard"
            if( ! $noChecksum )
            {
                $hash = Get-FileHash -Path $filename
                $message += "`n$($hash.Algorithm) checksum $($hash.hash)"
            }
            Write-Verbose -Message $message
            if( $form )
            {
                [void][Windows.MessageBox]::Show( $message , 'Encoding Complete' , 'Ok' ,'Information' )
            }
        }
        else
        {
            if( $form )
            {
                [void][Windows.MessageBox]::Show( "No data read from `"$($WPFtxtFileName.Text)`"" , 'Encoding Error' , 'Ok' ,'Exclamation' )
            }
            else
            {
                Write-Error -Message "No data read from `"$($WPFtxtFileName.Text)`""
            }
        }
        $data = $null
    }
}

Add-Type -AssemblyName System.Windows.Forms

$mainForm = $null

if( $PSBoundParameters[ 'filename' ] )
{
    Invoke-Operation -filename $filename -decode:$decode
}
else
{
    Add-Type -AssemblyName Presentationframework
    $mainForm = Load-GUI -inputXaml $mainwindowXAML

    if( ! $mainForm )
    {
        return
    }

    $mainForm.Title += " on $env:COMPUTERNAME"

    $WPFbtnEncode.Add_Click({
        $_.Handled = $true
        if( [string]::IsNullOrEmpty( $WPFtxtFileName.Text ) )
        {
            [void][Windows.MessageBox]::Show( "Must specify source file name" , 'Encoding Error' , 'Ok' ,'Exclamation' )
        }
        else
        {
            Invoke-Operation -filename $WPFtxtFileName.Text -form $mainForm
        }
    })

    $WPFbtnDecode.Add_Click({
        $_.Handled = $true
        if( [string]::IsNullOrEmpty( $WPFtxtFileName.Text ) )
        {
            [void][Windows.MessageBox]::Show( "Must specify destination file name" , 'Decoding Error' , 'Ok' ,'Exclamation' )
        }
        else
        {
            Invoke-Operation -filename $WPFtxtFileName.Text -decode -form $mainForm
        }
    })

    $WPFbtnFileChooser.Add_Click({
        $_.Handled = $true
        $fileBrowser = New-Object -TypeName System.Windows.Forms.OpenFileDialog
        if( $fileBrowser )
        {
            $fileBrowser.InitialDirectory = Get-Location -PSProvider FileSystem | Select-Object -ExpandProperty Path
            $file = $fileBrowser.ShowDialog()
            if( $file -eq 'OK' )
            {
                $WPFtxtFileName.Text = $fileBrowser.FileName
            }
        }
    })

    $result = $mainForm.ShowDialog()
}

# SIG # Begin signature block
# MIINKgYJKoZIhvcNAQcCoIINGzCCDRcCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUTHMa4JASthihvhvU5zzs3FmM
# qF6gggpsMIIFMDCCBBigAwIBAgIQBAkYG1/Vu2Z1U0O1b5VQCDANBgkqhkiG9w0B
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
# AQkEMRYEFFQ5efTRxqWZXm2+12bllvZGFmL8MA0GCSqGSIb3DQEBAQUABIIBAN3a
# WMY8pdRK40EUCkLVnNseN/70cn/gMzk/xlO1EleWFZGGBkxtK3xhsNwNAj1YYQ7P
# rLss8dZhe9YbC42w4tcP8kKejBZGtitVHFAaax/5wfFkP0WeeS901gyYe4olNZLQ
# aJOk1pY8ktW7vV8IcuhKJcwynUNbW418w3d/t925WL57fUiQykzIPYVxnv6tHUwg
# 8/5y7uJL+FRZlZidxmOJePw0GCWYVv6CJ5fjPiQ/l1Qm69ff98Y5RzyxLuHAbiNG
# mxDWU1wUuD3B+l/zkeAKmPV7ElZ56m3zNxYUIM+LS3b+oFAHGkGk2D5MfSdR9tnp
# TrM25LlPl7kRNUf1aZs=
# SIG # End signature block
