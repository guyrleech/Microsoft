#requires -version 3

<#
.SYNOPSIS
    Run multiple script blocks in parallel to demomstrate how to use runspaces

.PARAMETER maxThreads
    The maximum number of threads to use. Where this is less than the number of jobs, thread re-use will occur

.PARAMETER numberJobs
    The number of jobs to run
    
.PARAMETER maximumSleepSeconds
    The maximum number of seconds to generate a random sleep interval for
    
.PARAMETER showErrors
    Show any errors that occurred in the runspaces

.EXAMPLE
    & '.\Multiple runspaces.ps1' -numberJobs 10 -maxThreads 4 -maximumSleepSeconds 20

    Run 10 parallel runspaces using a pool of 4 threads and sleep, randomly, for up to 20 seconds in each runspace script block

.NOTES
    Modification History:

        2018/11/08  @guyrleech  Initial version
        2022/12/14  @guyrleech  Updated for demo at @psugbde
#>

<#
Copyright © 2022 Guy Leech

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, 
including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

[CmdletBinding()]

Param
(
    [uint16]$maxThreads = 10 ,
    [uint16]$numberJobs = 20 ,
    [uint16]$maximumSleepSeconds = 10 ,
    [switch]$showErrors
)

## local function that we will pass into each runspace so it can be called from there
Function Invoke-Panic
{
    Param
    (
        [int]$jobnumber = -1
    )

    Write-Warning -Message "Invoking panic in job $jobnumber !"
}

## local function that we will not pass into runspaces but call it from there to show doesn't get it
Function Invoke-Calm
{
    Param
    (
        [string]$message
    )

    Write-Output -InputObject $message
}

$seedy = Get-Random -SetSeed ( [datetime]::Now.ToFileTime() % [int32]::MaxValue)

## base parameters passed to the script block run in each runspace which could be different for each run space
[hashtable]$Parameters = @{
    Param1 = 'Guy'
    Param2 = 'Leech'
    MaximumSleepSeconds = $maximumSleepSeconds
}

$SessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()

## when using with GUIs, may need to set ApartmentState otherwise get exception making GUI - "The calling thread must be STA, because many UI components require this"

## make a thread safe dictionary for sharing of variables so we can write to them in parallel without corruption
$sharedVariables = [hashtable]::Synchronized(@{ 'Stuff' = (New-Object -TypeName System.Collections.Generic.List[string]) }) 

$SessionState.Variables.Add( (New-Object -Typename System.Management.Automation.Runspaces.SessionStateVariableEntry  -ArgumentList 'sharedVariables' , $sharedVariables , 'Shared Variables hashtable' ) ) 

## need to import any functions we need from this script
ForEach( $function in @( 'Invoke-Panic' ) )
{
    $Definition = $null
    $Definition = Get-Content -Path Function:\$function -ErrorAction Continue
    if( $Definition )
    {
        $SessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList $function , $Definition
        $sessionState.Commands.Add( $SessionStateFunction )
    }
}

$RunspacePool = [runspacefactory]::CreateRunspacePool(
    1, #Minimum Runspaces
    $maxThreads ,
    $SessionState ,
    $host
)

## ensure all initialisation/configuration is done before we open the run space pool
$RunspacePool.Open()

$jobs = New-Object -TypeName System.Collections.Generic.List[psobject]

[datetime]$startTime = [datetime]::Now

Write-Verbose -Message "$(Get-Date -Date $startTime -Format G): starting firing script blocks at runspaces"

1..$numberJobs | ForEach `
{
    [int]$numberJob = $_
    $PowerShell = [PowerShell]::Create() 
    $PowerShell.RunspacePool = $RunspacePool

    [void]$PowerShell.AddScript({
        Param
        (
            $Param1,
            $Param2,
            $jobNumber ,
            $maximumSleepSeconds
        )

        [int]$ThreadID = [appdomain]::GetCurrentThreadId()

        Write-Verbose "ThreadID: Beginning $ThreadID for job $jobNumber" -Verbose

        [int]$sleep = Get-Random -Minimum 1 -Maximum $maximumSleepSeconds

        ## this is the return from the script
        [pscustomobject]@{
            Param1    = $param1
            Param2    = $param2
            ThreadID  = $ThreadID
            ProcessID = $PID
            SleepTime = $sleep
        }

        Start-Sleep -Seconds $sleep

        Invoke-Panic -jobnumber $jobNumber ## demonstrate use of parent function in runspace script block which has previously been added to the session state

        Invoke-Calm -message 'All good here ;-)' ## show error in stream output later
        
        [string]$stuff = "Sleep time for job $jobnumber was $sleep"

        [System.Diagnostics.Trace]::WriteLine( $stuff ) ## demonstrate writing debug that can be picked  up by dbgview, etc

        $sharedVariables.Stuff.Add( $stuff )

        Write-Verbose "ThreadID: Ending $ThreadID" -Verbose
    })

    ## make a copy of the static part of the parameters dictionary so we can add our dynamic/variable parameter to it which will be unique for this runspace
    [hashtable]$theseParameters = $Parameters.Clone()
    $theseParameters.Add( 'jobNumber' , $numberJob )
    [void]$PowerShell.AddParameters( $theseParameters )

    ## start runspace asynchronously so that it does not block this loop
    $handle = $PowerShell.BeginInvoke()

    $job = [pscustomobject]@{ 
        PowerShell = $PowerShell
        Handle = $handle
    }

    $jobs.Add( $job )
}

Write-Verbose -Message "$(Get-Date -Format G): finished firing script blocks at runspaces, waiting for results"

[array]$results = @( ForEach( $job in $jobs )
{
    ## wait for this runspace to finish but by the time it finishes, others may have finished too which will mean no waiting for them when the loop gets around to them
    $returned = $job.powershell.EndInvoke( $job.handle )
    $returned

    ## show output streams, etc
    if( $showErrors -and $job.PowerShell.HadErrors )
    {
        $job.PowerShell.Streams.Error | Write-Error
    }

    $job.PowerShell.Dispose()
})

[datetime]$endTime = [datetime]::Now

Write-Verbose -Message "$(Get-Date -Date $endTime -Format G): finished waiting for results"

$jobs.clear()

## see shared variables

Write-Output -InputObject "Shared variable 'Stuff'"
$sharedVariables.Stuff

## examine results - see same process id and maybe different or same thread ids

Write-Output -InputObject "Objects returned from all runspaces"
$results | Format-Table

## show total time spent sleeping and actual elapsed time to create, run and wait for all runspaces to finish

Write-Output -InputObject "Sum of sleep times is $(($results | Measure-Object -Property SleepTime -Sum).Sum) seconds"

Write-Output -InputObject "Elapsed time was $(($endTime - $startTime).TotalSeconds) seconds"

[array]$uniqueThreads = @( $results | Group-Object -Property ThreadId )

Write-Output -InputObject "Unique threads were $($uniqueThreads.Count)"

Write-Output -InputObject "Currently have $((Get-Runspace | Measure-Object).Count) runspaces available"

## can do this earlier but left here so we can examine in the demo

$RunspacePool.Close()
$RunspacePool.Dispose()
$RunspacePool = $null

Write-Output -InputObject "Now have $((Get-Runspace | Measure-Object).Count) runspaces available"

# SIG # Begin signature block
# MIIjlQYJKoZIhvcNAQcCoIIjhjCCI4ICAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDnPxZHgoM8Bigf
# bTtwwZOyUMe7mS/QRJEIpu0hch2/4KCCHY4wggUwMIIEGKADAgECAhAECRgbX9W7
# ZnVTQ7VvlVAIMA0GCSqGSIb3DQEBCwUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNV
# BAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0xMzEwMjIxMjAwMDBa
# Fw0yODEwMjIxMjAwMDBaMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lD
# ZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQD407Mcfw4Rr2d3B9MLMUkZz9D7RZmxOttE9X/l
# qJ3bMtdx6nadBS63j/qSQ8Cl+YnUNxnXtqrwnIal2CWsDnkoOn7p0WfTxvspJ8fT
# eyOU5JEjlpB3gvmhhCNmElQzUHSxKCa7JGnCwlLyFGeKiUXULaGj6YgsIJWuHEqH
# CN8M9eJNYBi+qsSyrnAxZjNxPqxwoqvOf+l8y5Kh5TsxHM/q8grkV7tKtel05iv+
# bMt+dDk2DZDv5LVOpKnqagqrhPOsZ061xPeM0SAlI+sIZD5SlsHyDxL0xY4PwaLo
# LFH3c7y9hbFig3NBggfkOItqcyDQD2RzPJ6fpjOp/RnfJZPRAgMBAAGjggHNMIIB
# yTASBgNVHRMBAf8ECDAGAQH/AgEAMA4GA1UdDwEB/wQEAwIBhjATBgNVHSUEDDAK
# BggrBgEFBQcDAzB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9v
# Y3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDCBgQYDVR0fBHow
# eDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJl
# ZElEUm9vdENBLmNybDA6oDigNoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0Rp
# Z2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDBPBgNVHSAESDBGMDgGCmCGSAGG/WwA
# AgQwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAK
# BghghkgBhv1sAzAdBgNVHQ4EFgQUWsS5eyoKo6XqcQPAYPkt9mV1DlgwHwYDVR0j
# BBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8wDQYJKoZIhvcNAQELBQADggEBAD7s
# DVoks/Mi0RXILHwlKXaoHV0cLToaxO8wYdd+C2D9wz0PxK+L/e8q3yBVN7Dh9tGS
# dQ9RtG6ljlriXiSBThCk7j9xjmMOE0ut119EefM2FAaK95xGTlz/kLEbBw6RFfu6
# r7VRwo0kriTGxycqoSkoGjpxKAI8LpGjwCUR4pwUR6F6aGivm6dcIFzZcbEMj7uo
# +MUSaJ/PQMtARKUT8OZkDCUIQjKyNookAv4vcn4c10lFluhZHen6dGRrsutmQ9qz
# sIzV6Q3d9gEgzpkxYz0IGhizgZtPxpMQBvwHgfqL2vmCSfdibqFT+hKUGIUukpHq
# aGxEMrJmoecYpJpkUe8wggVPMIIEN6ADAgECAhAE/eOq2921q55B9NnVIXVOMA0G
# CSqGSIb3DQEBCwUAMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJ
# bmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lDZXJ0
# IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EwHhcNMjAwNzIwMDAwMDAw
# WhcNMjMwNzI1MTIwMDAwWjCBizELMAkGA1UEBhMCR0IxEjAQBgNVBAcTCVdha2Vm
# aWVsZDEmMCQGA1UEChMdU2VjdXJlIFBsYXRmb3JtIFNvbHV0aW9ucyBMdGQxGDAW
# BgNVBAsTD1NjcmlwdGluZ0hlYXZlbjEmMCQGA1UEAxMdU2VjdXJlIFBsYXRmb3Jt
# IFNvbHV0aW9ucyBMdGQwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCv
# bSdd1oAAu9rTtdnKSlGWKPF8g+RNRAUDFCBdNbYbklzVhB8hiMh48LqhoP7dlzZY
# 3YmuxztuPlB7k2PhAccd/eOikvKDyNeXsSa3WaXLNSu3KChDVekEFee/vR29mJuu
# jp1eYrz8zfvDmkQCP/r34Bgzsg4XPYKtMitCO/CMQtI6Rnaj7P6Kp9rH1nVO/zb7
# KD2IMedTFlaFqIReT0EVG/1ZizOpNdBMSG/x+ZQjZplfjyyjiYmE0a7tWnVMZ4KK
# TUb3n1CTuwWHfK9G6CNjQghcFe4D4tFPTTKOSAx7xegN1oGgifnLdmtDtsJUOOhO
# tyf9Kp8e+EQQyPVrV/TNAgMBAAGjggHFMIIBwTAfBgNVHSMEGDAWgBRaxLl7Kgqj
# pepxA8Bg+S32ZXUOWDAdBgNVHQ4EFgQUTXqi+WoiTm5fYlDLqiDQ4I+uyckwDgYD
# VR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHcGA1UdHwRwMG4wNaAz
# oDGGL2h0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3VyZWQtY3MtZzEu
# Y3JsMDWgM6Axhi9odHRwOi8vY3JsNC5kaWdpY2VydC5jb20vc2hhMi1hc3N1cmVk
# LWNzLWcxLmNybDBMBgNVHSAERTBDMDcGCWCGSAGG/WwDATAqMCgGCCsGAQUFBwIB
# FhxodHRwczovL3d3dy5kaWdpY2VydC5jb20vQ1BTMAgGBmeBDAEEATCBhAYIKwYB
# BQUHAQEEeDB2MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20w
# TgYIKwYBBQUHMAKGQmh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2Vy
# dFNIQTJBc3N1cmVkSURDb2RlU2lnbmluZ0NBLmNydDAMBgNVHRMBAf8EAjAAMA0G
# CSqGSIb3DQEBCwUAA4IBAQBT3M71SlOQ8vwM2txshp/XDvfoKBYHkpFCyanWaFds
# YQJQIKk4LOVgUJJ6LAf0xPSN7dZpjFaoilQy8Ajyd0U9UOnlEX4gk2J+z5i4sFxK
# /W2KU1j6R9rY5LbScWtsV+X1BtHihpzPywGGE5eth5Q5TixMdI9CN3eWnKGFkY13
# cI69zZyyTnkkb+HaFHZ8r6binvOyzMr69+oRf0Bv/uBgyBKjrmGEUxJZy+007fbm
# YDEclgnWT1cRROarzbxmZ8R7Iyor0WU3nKRgkxan+8rzDhzpZdtgIFdYvjeOc/Ip
# Pi2mI6NY4jqDXwkx1TEIbjUdrCmEfjhAfMTU094L7VSNMIIFjTCCBHWgAwIBAgIQ
# DpsYjvnQLefv21DiCEAYWjANBgkqhkiG9w0BAQwFADBlMQswCQYDVQQGEwJVUzEV
# MBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29t
# MSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3QgQ0EwHhcNMjIwODAx
# MDAwMDAwWhcNMzExMTA5MjM1OTU5WjBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMM
# RGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQD
# ExhEaWdpQ2VydCBUcnVzdGVkIFJvb3QgRzQwggIiMA0GCSqGSIb3DQEBAQUAA4IC
# DwAwggIKAoICAQC/5pBzaN675F1KPDAiMGkz7MKnJS7JIT3yithZwuEppz1Yq3aa
# za57G4QNxDAf8xukOBbrVsaXbR2rsnnyyhHS5F/WBTxSD1Ifxp4VpX6+n6lXFllV
# cq9ok3DCsrp1mWpzMpTREEQQLt+C8weE5nQ7bXHiLQwb7iDVySAdYyktzuxeTsiT
# +CFhmzTrBcZe7FsavOvJz82sNEBfsXpm7nfISKhmV1efVFiODCu3T6cw2Vbuyntd
# 463JT17lNecxy9qTXtyOj4DatpGYQJB5w3jHtrHEtWoYOAMQjdjUN6QuBX2I9YI+
# EJFwq1WCQTLX2wRzKm6RAXwhTNS8rhsDdV14Ztk6MUSaM0C/CNdaSaTC5qmgZ92k
# J7yhTzm1EVgX9yRcRo9k98FpiHaYdj1ZXUJ2h4mXaXpI8OCiEhtmmnTK3kse5w5j
# rubU75KSOp493ADkRSWJtppEGSt+wJS00mFt6zPZxd9LBADMfRyVw4/3IbKyEbe7
# f/LVjHAsQWCqsWMYRJUadmJ+9oCw++hkpjPRiQfhvbfmQ6QYuKZ3AeEPlAwhHbJU
# KSWJbOUOUlFHdL4mrLZBdd56rF+NP8m800ERElvlEFDrMcXKchYiCd98THU/Y+wh
# X8QgUWtvsauGi0/C1kVfnSD8oR7FwI+isX4KJpn15GkvmB0t9dmpsh3lGwIDAQAB
# o4IBOjCCATYwDwYDVR0TAQH/BAUwAwEB/zAdBgNVHQ4EFgQU7NfjgtJxXWRM3y5n
# P+e6mK4cD08wHwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8wDgYDVR0P
# AQH/BAQDAgGGMHkGCCsGAQUFBwEBBG0wazAkBggrBgEFBQcwAYYYaHR0cDovL29j
# c3AuZGlnaWNlcnQuY29tMEMGCCsGAQUFBzAChjdodHRwOi8vY2FjZXJ0cy5kaWdp
# Y2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3J0MEUGA1UdHwQ+MDww
# OqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJ
# RFJvb3RDQS5jcmwwEQYDVR0gBAowCDAGBgRVHSAAMA0GCSqGSIb3DQEBDAUAA4IB
# AQBwoL9DXFXnOF+go3QbPbYW1/e/Vwe9mqyhhyzshV6pGrsi+IcaaVQi7aSId229
# GhT0E0p6Ly23OO/0/4C5+KH38nLeJLxSA8hO0Cre+i1Wz/n096wwepqLsl7Uz9FD
# RJtDIeuWcqFItJnLnU+nBgMTdydE1Od/6Fmo8L8vC6bp8jQ87PcDx4eo0kxAGTVG
# amlUsLihVo7spNU96LHc/RzY9HdaXFSMb++hUD38dglohJ9vytsgjTVgHAIDyyCw
# rFigDkBjxZgiwbJZ9VVrzyerbHbObyMt9H5xaiNrIv8SuFQtJ37YOtnwtoeW/VvR
# XKwYw02fc7cBqZ9Xql4o4rmUMIIGrjCCBJagAwIBAgIQBzY3tyRUfNhHrP0oZipe
# WzANBgkqhkiG9w0BAQsFADBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNl
# cnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdp
# Q2VydCBUcnVzdGVkIFJvb3QgRzQwHhcNMjIwMzIzMDAwMDAwWhcNMzcwMzIyMjM1
# OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xOzA5
# BgNVBAMTMkRpZ2lDZXJ0IFRydXN0ZWQgRzQgUlNBNDA5NiBTSEEyNTYgVGltZVN0
# YW1waW5nIENBMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAxoY1Bkmz
# wT1ySVFVxyUDxPKRN6mXUaHW0oPRnkyibaCwzIP5WvYRoUQVQl+kiPNo+n3znIkL
# f50fng8zH1ATCyZzlm34V6gCff1DtITaEfFzsbPuK4CEiiIY3+vaPcQXf6sZKz5C
# 3GeO6lE98NZW1OcoLevTsbV15x8GZY2UKdPZ7Gnf2ZCHRgB720RBidx8ald68Dd5
# n12sy+iEZLRS8nZH92GDGd1ftFQLIWhuNyG7QKxfst5Kfc71ORJn7w6lY2zkpsUd
# zTYNXNXmG6jBZHRAp8ByxbpOH7G1WE15/tePc5OsLDnipUjW8LAxE6lXKZYnLvWH
# po9OdhVVJnCYJn+gGkcgQ+NDY4B7dW4nJZCYOjgRs/b2nuY7W+yB3iIU2YIqx5K/
# oN7jPqJz+ucfWmyU8lKVEStYdEAoq3NDzt9KoRxrOMUp88qqlnNCaJ+2RrOdOqPV
# A+C/8KI8ykLcGEh/FDTP0kyr75s9/g64ZCr6dSgkQe1CvwWcZklSUPRR8zZJTYsg
# 0ixXNXkrqPNFYLwjjVj33GHek/45wPmyMKVM1+mYSlg+0wOI/rOP015LdhJRk8mM
# DDtbiiKowSYI+RQQEgN9XyO7ZONj4KbhPvbCdLI/Hgl27KtdRnXiYKNYCQEoAA6E
# VO7O6V3IXjASvUaetdN2udIOa5kM0jO0zbECAwEAAaOCAV0wggFZMBIGA1UdEwEB
# /wQIMAYBAf8CAQAwHQYDVR0OBBYEFLoW2W1NhS9zKXaaL3WMaiCPnshvMB8GA1Ud
# IwQYMBaAFOzX44LScV1kTN8uZz/nupiuHA9PMA4GA1UdDwEB/wQEAwIBhjATBgNV
# HSUEDDAKBggrBgEFBQcDCDB3BggrBgEFBQcBAQRrMGkwJAYIKwYBBQUHMAGGGGh0
# dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBBBggrBgEFBQcwAoY1aHR0cDovL2NhY2Vy
# dHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZFJvb3RHNC5jcnQwQwYDVR0f
# BDwwOjA4oDagNIYyaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1
# c3RlZFJvb3RHNC5jcmwwIAYDVR0gBBkwFzAIBgZngQwBBAIwCwYJYIZIAYb9bAcB
# MA0GCSqGSIb3DQEBCwUAA4ICAQB9WY7Ak7ZvmKlEIgF+ZtbYIULhsBguEE0TzzBT
# zr8Y+8dQXeJLKftwig2qKWn8acHPHQfpPmDI2AvlXFvXbYf6hCAlNDFnzbYSlm/E
# UExiHQwIgqgWvalWzxVzjQEiJc6VaT9Hd/tydBTX/6tPiix6q4XNQ1/tYLaqT5Fm
# niye4Iqs5f2MvGQmh2ySvZ180HAKfO+ovHVPulr3qRCyXen/KFSJ8NWKcXZl2szw
# cqMj+sAngkSumScbqyQeJsG33irr9p6xeZmBo1aGqwpFyd/EjaDnmPv7pp1yr8TH
# wcFqcdnGE4AJxLafzYeHJLtPo0m5d2aR8XKc6UsCUqc3fpNTrDsdCEkPlM05et3/
# JWOZJyw9P2un8WbDQc1PtkCbISFA0LcTJM3cHXg65J6t5TRxktcma+Q4c6umAU+9
# Pzt4rUyt+8SVe+0KXzM5h0F4ejjpnOHdI/0dKNPH+ejxmF/7K9h+8kaddSweJywm
# 228Vex4Ziza4k9Tm8heZWcpw8De/mADfIBZPJ/tgZxahZrrdVcA6KYawmKAr7ZVB
# tzrVFZgxtGIJDwq9gdkT/r+k0fNX2bwE+oLeMt8EifAAzV3C+dAjfwAL5HYCJtnw
# ZXZCpimHCUcr5n8apIUP/JiW9lVUKx+A+sDyDivl1vupL0QVSucTDh3bNzgaoSv2
# 7dZ8/DCCBsAwggSooAMCAQICEAxNaXJLlPo8Kko9KQeAPVowDQYJKoZIhvcNAQEL
# BQAwYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMTswOQYD
# VQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQwOTYgU0hBMjU2IFRpbWVTdGFt
# cGluZyBDQTAeFw0yMjA5MjEwMDAwMDBaFw0zMzExMjEyMzU5NTlaMEYxCzAJBgNV
# BAYTAlVTMREwDwYDVQQKEwhEaWdpQ2VydDEkMCIGA1UEAxMbRGlnaUNlcnQgVGlt
# ZXN0YW1wIDIwMjIgLSAyMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA
# z+ylJjrGqfJru43BDZrboegUhXQzGias0BxVHh42bbySVQxh9J0Jdz0Vlggva2Sk
# /QaDFteRkjgcMQKW+3KxlzpVrzPsYYrppijbkGNcvYlT4DotjIdCriak5Lt4eLl6
# FuFWxsC6ZFO7KhbnUEi7iGkMiMbxvuAvfTuxylONQIMe58tySSgeTIAehVbnhe3y
# YbyqOgd99qtu5Wbd4lz1L+2N1E2VhGjjgMtqedHSEJFGKes+JvK0jM1MuWbIu6pQ
# OA3ljJRdGVq/9XtAbm8WqJqclUeGhXk+DF5mjBoKJL6cqtKctvdPbnjEKD+jHA9Q
# Bje6CNk1prUe2nhYHTno+EyREJZ+TeHdwq2lfvgtGx/sK0YYoxn2Off1wU9xLokD
# EaJLu5i/+k/kezbvBkTkVf826uV8MefzwlLE5hZ7Wn6lJXPbwGqZIS1j5Vn1TS+Q
# Hye30qsU5Thmh1EIa/tTQznQZPpWz+D0CuYUbWR4u5j9lMNzIfMvwi4g14Gs0/EH
# 1OG92V1LbjGUKYvmQaRllMBY5eUuKZCmt2Fk+tkgbBhRYLqmgQ8JJVPxvzvpqwcO
# agc5YhnJ1oV/E9mNec9ixezhe7nMZxMHmsF47caIyLBuMnnHC1mDjcbu9Sx8e47L
# ZInxscS451NeX1XSfRkpWQNO+l3qRXMchH7XzuLUOncCAwEAAaOCAYswggGHMA4G
# A1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsGAQUF
# BwMIMCAGA1UdIAQZMBcwCAYGZ4EMAQQCMAsGCWCGSAGG/WwHATAfBgNVHSMEGDAW
# gBS6FtltTYUvcyl2mi91jGogj57IbzAdBgNVHQ4EFgQUYore0GH8jzEU7ZcLzT0q
# lBTfUpwwWgYDVR0fBFMwUTBPoE2gS4ZJaHR0cDovL2NybDMuZGlnaWNlcnQuY29t
# L0RpZ2lDZXJ0VHJ1c3RlZEc0UlNBNDA5NlNIQTI1NlRpbWVTdGFtcGluZ0NBLmNy
# bDCBkAYIKwYBBQUHAQEEgYMwgYAwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRp
# Z2ljZXJ0LmNvbTBYBggrBgEFBQcwAoZMaHR0cDovL2NhY2VydHMuZGlnaWNlcnQu
# Y29tL0RpZ2lDZXJ0VHJ1c3RlZEc0UlNBNDA5NlNIQTI1NlRpbWVTdGFtcGluZ0NB
# LmNydDANBgkqhkiG9w0BAQsFAAOCAgEAVaoqGvNG83hXNzD8deNP1oUj8fz5lTmb
# Jeb3coqYw3fUZPwV+zbCSVEseIhjVQlGOQD8adTKmyn7oz/AyQCbEx2wmIncePLN
# fIXNU52vYuJhZqMUKkWHSphCK1D8G7WeCDAJ+uQt1wmJefkJ5ojOfRu4aqKbwVNg
# CeijuJ3XrR8cuOyYQfD2DoD75P/fnRCn6wC6X0qPGjpStOq/CUkVNTZZmg9U0rIb
# f35eCa12VIp0bcrSBWcrduv/mLImlTgZiEQU5QpZomvnIj5EIdI/HMCb7XxIstiS
# DJFPPGaUr10CU+ue4p7k0x+GAWScAMLpWnR1DT3heYi/HAGXyRkjgNc2Wl+WFrFj
# DMZGQDvOXTXUWT5Dmhiuw8nLw/ubE19qtcfg8wXDWd8nYiveQclTuf80EGf2JjKY
# e/5cQpSBlIKdrAqLxksVStOYkEVgM4DgI974A6T2RUflzrgDQkfoQTZxd639ouiX
# dE4u2h4djFrIHprVwvDGIqhPm73YHJpRxC+a9l+nJ5e6li6FV8Bg53hWf2rvwpWa
# SxECyIKcyRoFfLpxtU56mWz06J7UWpjIn7+NuxhcQ/XQKujiYu54BNu90ftbCqhw
# fvCXhHjjCANdRyxjqCU4lwHSPzra5eX25pvcfizM/xdMTQCi2NYBDriL7ubgclWJ
# LCcZYfZ3AYwxggVdMIIFWQIBATCBhjByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMM
# RGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQD
# EyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQgQ29kZSBTaWduaW5nIENBAhAE/eOq
# 2921q55B9NnVIXVOMA0GCWCGSAFlAwQCAQUAoIGEMBgGCisGAQQBgjcCAQwxCjAI
# oAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIB
# CzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIMf3FVStH21powLHxSLB
# JLzLSlSgB1aOsVhG/CDqukWxMA0GCSqGSIb3DQEBAQUABIIBADD1pNQEsjAwEiPr
# ytQWyrjnkQcy6iKX4J+pXkms0kDBoPo8oeDp2gq/CtFssNGufFN9NUczo6seOFV6
# 54tTGu/osgaw/1JJbyJYd5ZrzgFoaP1aiuHF3bxoldp5Yl+dQ17QjrsjFFwQ+ks3
# 7sTeL/Hub0IU4PO1MYSDVaUfq+/1Mm0YUZGIedYDcSqCrRw/KTlR8wCojl2wQSuh
# INPAqTrxVDNHvwo36EJKfYOzIJrBgWMlGICZ91fKm33cl9dcypfYmz4vn8HC95iP
# Eizp8Eq0csJZeR2FHsLI4LGdV3rCT9jkHcQpBBZWjg39Nbv1Se0H+fcyKZvpJ6KB
# tfR++6OhggMgMIIDHAYJKoZIhvcNAQkGMYIDDTCCAwkCAQEwdzBjMQswCQYDVQQG
# EwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0
# IFRydXN0ZWQgRzQgUlNBNDA5NiBTSEEyNTYgVGltZVN0YW1waW5nIENBAhAMTWly
# S5T6PCpKPSkHgD1aMA0GCWCGSAFlAwQCAQUAoGkwGAYJKoZIhvcNAQkDMQsGCSqG
# SIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMjIxMjE0MTk0ODE1WjAvBgkqhkiG9w0B
# CQQxIgQgw4cgpLalEnjw2FIj8HnGZo1kY+Ad32z9lJg3IpGhkmUwDQYJKoZIhvcN
# AQEBBQAEggIAHqrw4cbisdT0df6CPG5daklGT8deHV2XG2zDv4yc/5BdKjUBjYKl
# J+OCNR71Jqe9vnLTjVdSKEazPwaNE882v90aY9g78MpxmTi3wGNsysLm7pB0Zcbx
# crSjK62GHgUpPfye4fokQ4ljLrMryJfbycSpFCFDVC/ww4KQEWdgDpVi0OxFbGXC
# +CJ6N48zZHcCKpu4WJuR8ZiVuU+s/+rVAYdqnY4cRaQo7G9r3OZvzdreq1LitUr7
# cJzF2Nm904MZzJHdtpcC3025gUqdTeP3ce/Ryjl5XIP9+gF0+pm5oRWqtFp2rSL8
# kSX+XLYKVqLYF2atoh3hpairk6b+sP1HCbFQPHeFpazYy2X712EkBua/pW9++OMo
# 2CUzf8CtkrTrqTL56iMW3gPL7ECWbxGCZoxOs11ODdq59TIQCdxaR4m916DP7wym
# NONxn7RHxmtSsB7lb5aCbLuG69xqRoJlwRWsM1i31XFBDmo743+bZRAMd7WCYDfu
# xyqZxKvTh8vJ60OD1XY52/D3Acmy/2O6ILKKoiyh2JgBUaNvEZFNEDI9YtNIkcyC
# GAnbuPuFttEcaVOKUy80kWjfwC8gl70yizlix50vJSlfFWH92XgJ9rXbnEhLCzJf
# V646dm9ILpkmkl3bhnCGODv399T7aXnV147uDykgJ8H+RaXWWltsuC4=
# SIG # End signature block
