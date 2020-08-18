
<#
.SYNOPSIS

Take an XML template exported from perfmon and add the given list of machines to it by copying the existing local or \\computer counters with each of the specified computers for all counters

.PARAMETER source

Path to the source xml file, exported from perfmon

.PARAMETER dest

Path to the destination xml file, to be imported into perfmon. Use -force to overwrite if it already exists

.PARAMETER computers

Comma separated list of computers to place into the new XML file for each existing counter occurrence

.PARAMETER force

Overwrite the destination file if it already exists

.EXAMPLE

& '.\Add computers to perfmon xml.ps1' -source 'C:\temp\perfmon.xml' -dest C:\temp\perfmon.changed.xml -computers grl-dc03,grl-jump01 -Verbose

Read source file C:\temp\perfmon.xml and for each counter \\computer\, add a new counter instance for computers grl-dc03 and grl-jump01.
Write the result to C:\temp\perfmon.changed.xml which can be imported into perfmon as a new Data Collector Set

.NOTES

@guyrleech 10/08/2020

#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory,HelpMessage="Path to source xml file")]
    [string]$source ,
    [Parameter(Mandatory,HelpMessage="Path to destination xml file")]
    [string]$dest ,
    [Parameter(Mandatory,HelpMessage="Computers to add to new xml file")]
    [string[]]$computers ,
    [switch]$force
)

[xml]$perf = Get-Content -Path $source

if( ! $perf )
{
    Throw "Failed to read XML from `"$source`""
}

if( ! $force -and ( Test-Path -Path $dest -ErrorAction SilentlyContinue ) )
{
    Throw "Destination file `"$dest`" already exists - user -force to overwrite"
}

[array]$existingCounters = @( $perf.DataCollectorSet.PerformanceCounterDataCollector.Counter )
[array]$existingCounterDisplayNames = @( $perf.DataCollectorSet.PerformanceCounterDataCollector.CounterDisplayName )

if( ! $existingCounters -or ! $existingCounters.Count )
{
    Throw "No counters found in `"$source`""
}
elseif( $existingCounters.Count -ne $existingCounterDisplayNames.Count )
{
    Throw "Count mismatch between counters ($($existingCounters.Count)) and counter display names ($($existingCounterDisplayNames))"
}

Write-Verbose -Message "Got $($existingCounters.Count) existing counters and $($existingCounterDisplayNames.Count) counter display names"

## sanity check the existing counters to make sure there is only 1 machine name if they start with \\
[array]$existingComputers = @( ($perf.DataCollectorSet.PerformanceCounterDataCollector.Counter) | Where-Object { $_ -match '^\\\\([^\\]+)\\' } | ForEach-Object { $Matches[1] }|Sort-Object -Unique )

if( $existingComputers -and $existingComputers.Count -gt 1 )
{
    Throw "Found $($existingComputers.Count) different computers in `"$source`" ($($existingComputers -join ','))"
}
## else if none found then is all local counters which we can deal with

[int]$additions = 0

ForEach( $computer in $computers )
{
    ForEach( $counter in $existingCounters )
    {
        $newCounter = $perf.CreateElement( 'Counter' )
        ## need to deal with local and remote counters
        if( $counter -match '^\\\\[^\\]+\\(.*)$' ) ## \\grl-jump02\PhysicalDisk(*)\% Disk Time
        {
            $newCounter.InnerXml = '\\{0}\{1}' -f $computer, $Matches[ 1 ]
        }
        else
        {
            $newCounter.InnerXml = "\\$computer\$counter"
        }
        if( $newCounter.InnerXml -eq $counter )
        {
            Write-Warning -Message "Failed to change `"$counter`" counter to include $computer"
        }
        else
        {
            if( $perf.DataCollectorSet.PerformanceCounterDataCollector.AppendChild( $newCounter ) )
            {
                $additions++
            }
        }
    }
    ForEach( $counterDisplayName in $existingCounterDisplayNames )
    {
        $newCounterDisplayName = $perf.CreateElement( 'CounterDisplayName' )
        ## need to deal with local and remote counters
        if( $counterDisplayName -match '^\\\\[^\\]+\\(.*)$' ) ## \\grl-jump02\PhysicalDisk(*)\% Disk Time
        {
            $newCounterDisplayName.InnerXml = '\\{0}\{1}' -f $computer, $Matches[ 1 ]
        }
        else
        {
            $newCounterDisplayName.InnerXml = "\\$computer\$counterDisplayName"
        }
        if( $newCounterDisplayName.InnerXml -eq $counterDisplayName )
        {
            Write-Warning -Message "Failed to change `"$counterDisplayName`" display name to include $computer"
        }
        else
        {
            if( $perf.DataCollectorSet.PerformanceCounterDataCollector.AppendChild( $newCounterDisplayName ) )
            {
                $additions++
            }
        }
    }
}

Write-Verbose -Message "$additions additions made to file"

if( $additions -gt 0 )
{
    $perf.Save( $dest )
}
else
{
    Write-Warning -Message "Added no new data so not writing file"
}

# SIG # Begin signature block
# MIINRQYJKoZIhvcNAQcCoIINNjCCDTICAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU5m0f1n+0OLWQvBGjY+G4EmgB
# IKOgggqHMIIFMDCCBBigAwIBAgIQBAkYG1/Vu2Z1U0O1b5VQCDANBgkqhkiG9w0B
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
# CzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFKwmbLZVBntTxUXiZ6sx
# SfuYUYQGMA0GCSqGSIb3DQEBAQUABIIBAHBpLKjH4bcm9Aa5r+wRrlKNg9Xwq0FV
# t4EreeYUQw4RXjq+IvPya+oUMhrOJGB4dmxGgb3wTsDGFANBV6KqfzLlU7ql6k0j
# dMVjtRalczhF5spn/NLSJ1hgRhMxKByffdAMqn1d2yQEZjSHlwSZgMYMP42yN+Kk
# ssa8OtZJh8KaS14QggEDkqN3ckPrmG28e2LlRGaY6b0lo5/E3ncCWpK4f47TaShQ
# eQxMGRXORtVM5sgd4UOmrXSat6HSGG3Jzcg9FErhyJxX3zjGCIqbjR/89nSiqSgH
# Z4YPR/LbJD62k9XzK3C+VlT3XM57bfPQKgrmHL1fC+de20TVo/Tku8I=
# SIG # End signature block
