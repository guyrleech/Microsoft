<#
.SYNOPSIS
    Resize and/or reposition the window for a specified process or launch a process with specified size and/or position

.PARAMETER processName
    The names of processes to operate on

.PARAMETER processId
    The PID of the process to operate on

.PARAMETER x
    The x coordinate to position the window at

.PARAMETER y
    The y coordinate to position the window at

.PARAMETER style
    The style of the window

.PARAMETER mostRecent
    Only operate on the most recently launched process where multiple process names match

.PARAMETER alwaysOntop
    Make the window always on top

.PARAMETER timeoutMilliseconds
    Time in milliseconds to wait from when the process is started to when the window adjustments are made

.PARAMETER launch
    Launch the process specified rather than operating on an existing instance

.PARAMETER arguments
    Arguments to pass to the process when using -launch

.EXAMPLE
    & '.\Set window.ps1' -processId $pid -alwaysOnTop $true

    Set the PowerShell window running the script to always be on top of other windows
    
.EXAMPLE
    & '.\Set window.ps1' -processId $pid -alwaysOnTop $false

    Set the PowerShell window running the script to not be on top of other windows
    
.EXAMPLE
    & '.\Set window.ps1' -processname notepad -mostrecent -width 640 -height 480

    Find the window for the most recently launched notepad process and set its dimentsions to 640 by 480
    
.EXAMPLE
    & '.\Set window.ps1' -processname notepad -launch -width 1000 -height 500 -x 50 -y 50 -arguments "fred bloggs.txt"

    Launch notepad on the file "fred bloggs.txt" and set the windows dimentsions to 1000 by 500 and position it at 50,50

.EXAMPLE
    & '.\Set window.ps1' -processname notepad -launch -arguments "fred bloggs.txt" -style MAXIMIZE
    
    Launch notepad on the file "fred bloggs.txt" and make the window maximized
    
.EXAMPLE
    & '.\Set window.ps1' -processId 12345
    
    Find the window for the process with pid 12345 and show its current size, position and attributes

.NOTES

    Modification History:

    2022/07/25  @guyrleech  First public release
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
    [Parameter(Mandatory=$True,ValuefromPipeline=$True,ParameterSetName='ByName')]
    [Alias("name")]
    [array]$processName ,
    [Parameter(Mandatory=$True,ValuefromPipeline=$False,ParameterSetName='ById')]
    [int]$processId ,
    [int]$width ,
    [int]$height ,
    [int]$x ,
    [int]$y ,
    [ValidateSet('FORCEMINIMIZE', 'HIDE', 'MAXIMIZE', 'MINIMIZE', 'RESTORE', 'SHOW', 'SHOWDEFAULT', 'SHOWMAXIMIZED', 'SHOWMINIMIZED', 'SHOWMINNOACTIVE', 'SHOWNA', 'SHOWNOACTIVATE', 'SHOWNORMAL')]
    [string]$style ,
    [switch]$mostRecent ,
    [bool]$alwaysOnTop ,
    [int]$timeoutMilliseconds ,
    [switch]$launch ,
    [string[]]$arguments
)

Begin
{
    Add-Type @'
        using System;
        using System.Runtime.InteropServices;
    
        [StructLayout(LayoutKind.Sequential)]

        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        public static class user32
        {
            [DllImport("user32.dll", SetLastError=true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool SetForegroundWindow(IntPtr hWnd);

            [DllImport("user32.dll", SetLastError=true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow); 
            
            [DllImport("user32.dll", SetLastError=true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow); 
            
            [DllImport("user32.dll", SetLastError=true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool IsIconic(IntPtr hWnd); 
        
            [DllImport("user32.dll", SetLastError=true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool IsZoomed(IntPtr hWnd); 
        
            [DllImport("user32.dll", SetLastError=true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool IsWindowVisible(IntPtr hWnd); 
        
            [DllImport("user32.dll", SetLastError=true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool IsWindowUnicode(IntPtr hWnd); 
        
            [DllImport("user32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
   
            [DllImport("user32.dll", SetLastError=true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool SetWindowPos( IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        }
'@ -ErrorAction Stop
}

Process
{
    $process = $null

    if( $processId -gt 0 )
    {
        $process = Get-Process -Id $processId -ErrorAction Stop
    }
    elseif( -Not [string]::IsNullOrEmpty( $processName ) )
    {
        if( $launch )
        {
            if( $processName.Count -ne 1 )
            {
                Throw "Can only launch a single process"
            }

            [hashtable]$launchParameters = @{
                'FilePath' = $processName[0]
                'PassThru' = $true
                'ErrorAction' = 'stop'
            }
            if( $PSBoundParameters[ 'arguments' ] )
            {
                $launchParameters.Add( 'ArgumentList' , $arguments )
            }
            if( $PSBoundParameters[ 'style' ] )
            {
                [string]$startStyle = switch -Regex ($style)
                {
                    'MAX'  { 'Maximized' }
                    'MIN'  { 'Minimized' }
                    'HIDE' { 'Hidden' ; $launchParameters.Add( 'NoNewWindow' , $true ) }
                }
                if( $null -ne $startStyle )
                {
                    $launchParameters.Add( 'WindowStyle' , $startStyle )
                }
            }
            $process = Start-Process @launchParameters
            if( $process )
            {
                Write-Verbose -Message "Launched $($launchParameters.FilePath) pid is $($process.Id)"
                ## we need to wait for it to have a window
                if( $PSBoundParameters.ContainsKey( '   ' ) )
                {
                    $null = $process.WaitForInputIdle( $timeoutMilliseconds )
                }
                else
                {
                    $null = $process.WaitForInputIdle()
                }
            }
            ## else launch failure will be reported above and code below will not run due to $process being null
        }
        else ## looking for existing process
        {
            if( $processName -and $processName -is [string] )
            {
                $processes = Get-Process -Name $processName -ErrorAction Stop 
            }
            else 
            {
                $processes = Get-Process -Name $processName
            }
            if( $processes -is [array] -and $processes.Count -gt 1 )
            {
                if( -Not $mostRecent )
                {
                    ## TODO should we have an option to iterate on all matching processes ?
                    Throw "Found $($processes.Count) processes for $processName and -mostrecent not specified"
                }
                else
                {
                    $process = $processes | Sort-Object -Property StartTime -Descending | Select-Object -First 1
                }
            }
            else
            {
                $process = $processes
            }
        }
    }

    if( $process )
    {
        ## TODO we could do various things to the process such as change CPU priority, adjust working sets, etc
        if( ( $windowHandle = $process.MainWindowHandle ) -and $windowHandle -ne [IntPtr]::Zero )
        {
            ## https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowpos

            [uint32]$flags = 0x4040   ## SWP_ASYNCWINDOWPOS | SWP_SHOWWINDOW
            [IntPtr]$insertAfter = -2 ## -1 = HWND_TOPMOST , -2 = HWND_NOTOPMOST 
            [int]$changes = 0
            if( $PSBoundParameters.ContainsKey( 'alwaysOnTop' ))
            {
                if( $alwaysOnTop )
                {
                    $insertAfter = -1
                }
                $changes++
            }
            if( $style -ieq 'hide' )
            {
                $flags = $flags -bor 0x0080 ## SWP_HIDEWINDOW
                $changes++
            }
            if( $width -gt 0 -or $height -gt 0 )
            {
                $changes++
            }
            else
            {
                $flags = $flags -bor 0x0001 ## SWP_NOSIZE
            }
            if( $PSBoundParameters.ContainsKey( 'x' ) -or $PSBoundParameters.ContainsKey( 'y' ) )
            {
                $changes++
            }
            else
            {
                $flags = $flags -bor 0x0002 ## SWP_NOMOVE
            }

            [bool]$result = $false

            if( $changes -gt 0 -or $PSBoundParameters.ContainsKey( 'style' ) )
            {
                if( $launch -and $style -match 'MAX' )  ## if we launched the process maximised then we can't change it here otherwise it will not stay maximised
                {
                    $result = $true
                }
                elseif( $changes -gt 0 )
                {
                    Write-Verbose -Message ( "insertAfter {0} flags {1:x}" -f $insertAfter , $flags )
        
                    $result = [user32]::SetWindowPos( $windowHandle, $insertAfter , $x , $y , $width , $height , $flags ) ; $lastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
                }
                if( -Not [string]::IsNullOrEmpty( $style ) )
                {
                    $WindowStates = @{
                        'FORCEMINIMIZE'   = 11
                        'HIDE'            = 0
                        'MAXIMIZE'        = 3
                        'MINIMIZE'        = 6
                        'RESTORE'         = 9
                        'SHOW'            = 5
                        'SHOWDEFAULT'     = 10
                        'SHOWMAXIMIZED'   = 3
                        'SHOWMINIMIZED'   = 2
                        'SHOWMINNOACTIVE' = 7
                        'SHOWNA'          = 8
                        'SHOWNOACTIVATE'  = 4
                        'SHOWNORMAL'      = 1
                    }
                    $cmdShow = $WindowStates[ $style ]
                    if( $null -ne $cmdShow )
                    {
                        Write-Verbose -Message "Setting style to $style ($cmdShow)"
                        [bool]$styleResult = $false
                        $styleResult = [user32]::ShowWindowAsync( $windowHandle, [int]$cmdShow ) ; $lastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
                        if( -Not $styleResult )
                        {
                            Write-Warning -Message "Failed ShowWindowAsync - $lastError"
                        }
                        elseif( $changes -eq 0 )
                        {
                            $result = $true ## so doesn't throw an error from the SetWindowPos since it didn't happen
                        }
                    }
                    else
                    {
                        Write-Warning -Message "Cannot map style `"$style`""
                    }
                }

                if( -Not $result) 
                {
                     Throw "Failed SetWindowPos - $lastError"
                }
            }

            ## no changes so report current window size and position
            $windowRect = New-Object -TypeName RECT
            $result = [user32]::GetWindowRect( $windowHandle, [ref]$windowRect ) ; $lastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
            if( -Not $result) 
            {
                Write-Warning -Message "Failed GetWindowRect - $lastError"
            }
            else
            {
                [pscustomobject]@{
                    'pid'   = $process.Id
                    'title' = $process.MainWindowTitle
                    'x' = $windowRect.Left
                    'y' = $windowRect.Top
                    'width'  = $windowRect.Right - $windowRect.Left
                    'height' = $windowRect.Bottom - $windowRect.Top
                    'minimised' = [user32]::IsIconic( $windowHandle )
                    'maximised' = [user32]::IsZoomed( $windowHandle )
                    'unicode'   = [user32]::IsWindowUnicode( $windowHandle )
                    'visible'   = [user32]::IsWindowVisible( $windowHandle )
                }
            }
        }
        else
        {
            Throw "Process $($process.Name) ($($process.Id)) has no window handle"
        }
    }
    else
    {
        Throw 'Unable to find process'
    }
}

End
{
}

# SIG # Begin signature block
# MIIjmgYJKoZIhvcNAQcCoIIjizCCI4cCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUCrvrDy4yF3iTnS10K4syeOeN
# 1pCggh24MIIFMDCCBBigAwIBAgIQBAkYG1/Vu2Z1U0O1b5VQCDANBgkqhkiG9w0B
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
# hH44QHzE1NPeC+1UjTCCBbEwggSZoAMCAQICEAEkCvseOAuKFvFLcZ3008AwDQYJ
# KoZIhvcNAQEMBQAwZTELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IElu
# YzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQg
# QXNzdXJlZCBJRCBSb290IENBMB4XDTIyMDYwOTAwMDAwMFoXDTMxMTEwOTIzNTk1
# OVowYjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UE
# CxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBS
# b290IEc0MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+Rd
# SjwwIjBpM+zCpyUuySE98orYWcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20d
# q7J58soR0uRf1gU8Ug9SH8aeFaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7f
# gvMHhOZ0O21x4i0MG+4g1ckgHWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRA
# X7F6Zu53yEioZldXn1RYjgwrt0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raR
# mECQecN4x7axxLVqGDgDEI3Y1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzU
# vK4bA3VdeGbZOjFEmjNAvwjXWkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2
# mHY9WV1CdoeJl2l6SPDgohIbZpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkr
# fsCUtNJhbesz2cXfSwQAzH0clcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaA
# sPvoZKYz0YkH4b235kOkGLimdwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxf
# jT/JvNNBERJb5RBQ6zHFynIWIgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEe
# xcCPorF+CiaZ9eRpL5gdLfXZqbId5RsCAwEAAaOCAV4wggFaMA8GA1UdEwEB/wQF
# MAMBAf8wHQYDVR0OBBYEFOzX44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaA
# FEXroq/0ksuCMS1Ri6enIZ3zbcgPMA4GA1UdDwEB/wQEAwIBhjATBgNVHSUEDDAK
# BggrBgEFBQcDCDB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9v
# Y3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDBFBgNVHR8EPjA8
# MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVk
# SURSb290Q0EuY3JsMCAGA1UdIAQZMBcwCAYGZ4EMAQQCMAsGCWCGSAGG/WwHATAN
# BgkqhkiG9w0BAQwFAAOCAQEAmhYCpQHvgfsNtFiyeK2oIxnZczfaYJ5R18v4L0C5
# ox98QE4zPpA854kBdYXoYnsdVuBxut5exje8eVxiAE34SXpRTQYy88XSAConIOqJ
# LhU54Cw++HV8LIJBYTUPI9DtNZXSiJUpQ8vgplgQfFOOn0XJIDcUwO0Zun53OdJU
# lsemEd80M/Z1UkJLHJ2NltWVbEcSFCRfJkH6Gka93rDlkUcDrBgIy8vbZol/K5xl
# v743Tr4t851Kw8zMR17IlZWt0cu7KgYg+T9y6jbrRXKSeil7FAM8+03WSHF6EBGK
# CHTNbBsEXNKKlQN2UVBT1i73SkbDrhAscUywh7YnN0RgRDCCBq4wggSWoAMCAQIC
# EAc2N7ckVHzYR6z9KGYqXlswDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMx
# FTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNv
# bTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTIyMDMyMzAw
# MDAwMFoXDTM3MDMyMjIzNTk1OVowYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRp
# Z2lDZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBUcnVzdGVkIEc0IFJTQTQw
# OTYgU0hBMjU2IFRpbWVTdGFtcGluZyBDQTCCAiIwDQYJKoZIhvcNAQEBBQADggIP
# ADCCAgoCggIBAMaGNQZJs8E9cklRVcclA8TykTepl1Gh1tKD0Z5Mom2gsMyD+Vr2
# EaFEFUJfpIjzaPp985yJC3+dH54PMx9QEwsmc5Zt+FeoAn39Q7SE2hHxc7Gz7iuA
# hIoiGN/r2j3EF3+rGSs+QtxnjupRPfDWVtTnKC3r07G1decfBmWNlCnT2exp39mQ
# h0YAe9tEQYncfGpXevA3eZ9drMvohGS0UvJ2R/dhgxndX7RUCyFobjchu0CsX7Le
# Sn3O9TkSZ+8OpWNs5KbFHc02DVzV5huowWR0QKfAcsW6Th+xtVhNef7Xj3OTrCw5
# 4qVI1vCwMROpVymWJy71h6aPTnYVVSZwmCZ/oBpHIEPjQ2OAe3VuJyWQmDo4EbP2
# 9p7mO1vsgd4iFNmCKseSv6De4z6ic/rnH1pslPJSlRErWHRAKKtzQ87fSqEcazjF
# KfPKqpZzQmiftkaznTqj1QPgv/CiPMpC3BhIfxQ0z9JMq++bPf4OuGQq+nUoJEHt
# Qr8FnGZJUlD0UfM2SU2LINIsVzV5K6jzRWC8I41Y99xh3pP+OcD5sjClTNfpmEpY
# PtMDiP6zj9NeS3YSUZPJjAw7W4oiqMEmCPkUEBIDfV8ju2TjY+Cm4T72wnSyPx4J
# duyrXUZ14mCjWAkBKAAOhFTuzuldyF4wEr1GnrXTdrnSDmuZDNIztM2xAgMBAAGj
# ggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBS6FtltTYUvcyl2
# mi91jGogj57IbzAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNV
# HQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBp
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUH
# MAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRS
# b290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EM
# AQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAfVmOwJO2b5ipRCIB
# fmbW2CFC4bAYLhBNE88wU86/GPvHUF3iSyn7cIoNqilp/GnBzx0H6T5gyNgL5Vxb
# 122H+oQgJTQxZ822EpZvxFBMYh0MCIKoFr2pVs8Vc40BIiXOlWk/R3f7cnQU1/+r
# T4osequFzUNf7WC2qk+RZp4snuCKrOX9jLxkJodskr2dfNBwCnzvqLx1T7pa96kQ
# sl3p/yhUifDVinF2ZdrM8HKjI/rAJ4JErpknG6skHibBt94q6/aesXmZgaNWhqsK
# RcnfxI2g55j7+6adcq/Ex8HBanHZxhOACcS2n82HhyS7T6NJuXdmkfFynOlLAlKn
# N36TU6w7HQhJD5TNOXrd/yVjmScsPT9rp/Fmw0HNT7ZAmyEhQNC3EyTN3B14OuSe
# reU0cZLXJmvkOHOrpgFPvT87eK1MrfvElXvtCl8zOYdBeHo46Zzh3SP9HSjTx/no
# 8Zhf+yvYfvJGnXUsHicsJttvFXseGYs2uJPU5vIXmVnKcPA3v5gA3yAWTyf7YGcW
# oWa63VXAOimGsJigK+2VQbc61RWYMbRiCQ8KvYHZE/6/pNHzV9m8BPqC3jLfBInw
# AM1dwvnQI38AC+R2AibZ8GV2QqYphwlHK+Z/GqSFD/yYlvZVVCsfgPrA8g4r5db7
# qS9EFUrnEw4d2zc4GqEr9u3WfPwwggbGMIIErqADAgECAhAKekqInsmZQpAGYzhN
# hpedMA0GCSqGSIb3DQEBCwUAMGMxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQgVHJ1c3RlZCBHNCBSU0E0MDk2
# IFNIQTI1NiBUaW1lU3RhbXBpbmcgQ0EwHhcNMjIwMzI5MDAwMDAwWhcNMzMwMzE0
# MjM1OTU5WjBMMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4x
# JDAiBgNVBAMTG0RpZ2lDZXJ0IFRpbWVzdGFtcCAyMDIyIC0gMjCCAiIwDQYJKoZI
# hvcNAQEBBQADggIPADCCAgoCggIBALkqliOmXLxf1knwFYIY9DPuzFxs4+AlLtIx
# 5DxArvurxON4XX5cNur1JY1Do4HrOGP5PIhp3jzSMFENMQe6Rm7po0tI6IlBfw2y
# 1vmE8Zg+C78KhBJxbKFiJgHTzsNs/aw7ftwqHKm9MMYW2Nq867Lxg9GfzQnFuUFq
# RUIjQVr4YNNlLD5+Xr2Wp/D8sfT0KM9CeR87x5MHaGjlRDRSXw9Q3tRZLER0wDJH
# GVvimC6P0Mo//8ZnzzyTlU6E6XYYmJkRFMUrDKAz200kheiClOEvA+5/hQLJhuHV
# GBS3BEXz4Di9or16cZjsFef9LuzSmwCKrB2NO4Bo/tBZmCbO4O2ufyguwp7gC0vI
# CNEyu4P6IzzZ/9KMu/dDI9/nw1oFYn5wLOUrsj1j6siugSBrQ4nIfl+wGt0ZvZ90
# QQqvuY4J03ShL7BUdsGQT5TshmH/2xEvkgMwzjC3iw9dRLNDHSNQzZHXL537/M2x
# wafEDsTvQD4ZOgLUMalpoEn5deGb6GjkagyP6+SxIXuGZ1h+fx/oK+QUshbWgaHK
# 2jCQa+5vdcCwNiayCDv/vb5/bBMY38ZtpHlJrYt/YYcFaPfUcONCleieu5tLsuK2
# QT3nr6caKMmtYbCgQRgZTu1Hm2GV7T4LYVrqPnqYklHNP8lE54CLKUJy93my3YTq
# J+7+fXprAgMBAAGjggGLMIIBhzAOBgNVHQ8BAf8EBAMCB4AwDAYDVR0TAQH/BAIw
# ADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAgBgNVHSAEGTAXMAgGBmeBDAEEAjAL
# BglghkgBhv1sBwEwHwYDVR0jBBgwFoAUuhbZbU2FL3MpdpovdYxqII+eyG8wHQYD
# VR0OBBYEFI1kt4kh/lZYRIRhp+pvHDaP3a8NMFoGA1UdHwRTMFEwT6BNoEuGSWh0
# dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFJTQTQwOTZT
# SEEyNTZUaW1lU3RhbXBpbmdDQS5jcmwwgZAGCCsGAQUFBwEBBIGDMIGAMCQGCCsG
# AQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wWAYIKwYBBQUHMAKGTGh0
# dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFJTQTQw
# OTZTSEEyNTZUaW1lU3RhbXBpbmdDQS5jcnQwDQYJKoZIhvcNAQELBQADggIBAA0t
# I3Sm0fX46kuZPwHk9gzkrxad2bOMl4IpnENvAS2rOLVwEb+EGYs/XeWGT76TOt4q
# OVo5TtiEWaW8G5iq6Gzv0UhpGThbz4k5HXBw2U7fIyJs1d/2WcuhwupMdsqh3KEr
# lribVakaa33R9QIJT4LWpXOIxJiA3+5JlbezzMWn7g7h7x44ip/vEckxSli23zh8
# y/pc9+RTv24KfH7X3pjVKWWJD6KcwGX0ASJlx+pedKZbNZJQfPQXpodkTz5GiRZj
# IGvL8nvQNeNKcEiptucdYL0EIhUlcAZyqUQ7aUcR0+7px6A+TxC5MDbk86ppCaiL
# fmSiZZQR+24y8fW7OK3NwJMR1TJ4Sks3KkzzXNy2hcC7cDBVeNaY/lRtf3GpSBp4
# 3UZ3Lht6wDOK+EoojBKoc88t+dMj8p4Z4A2UKKDr2xpRoJWCjihrpM6ddt6pc6pI
# allDrl/q+A8GQp3fBmiW/iqgdFtjZt5rLLh4qk1wbfAs8QcVfjW05rUMopml1xVr
# NQ6F1uAszOAMJLh8UgsemXzvyMjFjFhpr6s94c/MfRWuFL+Kcd/Kl7HYR+ocheBF
# ThIcFClYzG/Tf8u+wQ5KbyCcrtlzMlkI5y2SoRoR/jKYpl0rl+CL05zMbbUNrkdj
# OEcXW28T2moQbh9Jt0RbtAgKh1pZBHYRoad3AhMcMYIFTDCCBUgCAQEwgYYwcjEL
# MAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3
# LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElE
# IENvZGUgU2lnbmluZyBDQQIQBP3jqtvdtaueQfTZ1SF1TjAJBgUrDgMCGgUAoHgw
# GAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGC
# NwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQx
# FgQUs1/lnb4jsRQHZoLbzlbkI/6UEfAwDQYJKoZIhvcNAQEBBQAEggEAoFgjzM72
# 6B4lC6ph9TWM16aSfXWHf8PHp3wINWeHly4XTkyC8DFkeSI4YvwxWUuhjixRefve
# JcrB/Q15laQ//YY3ncq4et8D1/OFs35IREsxrcAZiWxvUKtMg+1GGiMBQjuAuPvN
# lxuVDjmQeZ/fzsBT+HddeiwkQMm+zjdLAytGoO90JLzS5hffwNKcddmsJi5N1I0a
# t9AsMksMaPta3bKmwhzxMvgYVWlEI9kQ//xmRabAYvPVs9vEPVvqZcaqQh3KDYTZ
# bgNI8Mv5SNK4qgmGJd18r1s8IJA+UG6SbTmuSWmcgjdizvIp7yh7FmMtXHKRxG3V
# sNocmYCMuFv6GaGCAyAwggMcBgkqhkiG9w0BCQYxggMNMIIDCQIBATB3MGMxCzAJ
# BgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjE7MDkGA1UEAxMyRGln
# aUNlcnQgVHJ1c3RlZCBHNCBSU0E0MDk2IFNIQTI1NiBUaW1lU3RhbXBpbmcgQ0EC
# EAp6SoieyZlCkAZjOE2Gl50wDQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMx
# CwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yMjA3MjUyMjQ3MDRaMC8GCSqG
# SIb3DQEJBDEiBCCBWibpdzY+JcNPesPKiqgPo6RA0QH1uzK1Du55Q7WyejANBgkq
# hkiG9w0BAQEFAASCAgBaN4QpWPVCohJMGJP52QmWsXqw4QYqZLdbp0ff0P1fIxEo
# tPhaLAiKdUAwR6j1gLEPJIRMQVpEGkS7pAaWJXnaYq1HNW9PdVFn3vgSzKGtpsrK
# wJ0KHP42PBi6e0RqigNOJjtwzUYexEoft1LXy34EFj16QYY8bJg2YCi+tlDmNcEa
# RjzLZbzrnnj551QUw1YlOY75gEE7wwi7EB2Mj7wudj41J4EHF+DYepvV60wHkFxY
# jApChwG+bfrC5nlK/jfvB9D/BirfV8SZkt2UGNzgRkNdLzQ14yHveOjZEfslgwCX
# +X/TaUoLisTmaIZ9ErQP/t6Ki/wwVRjspoImmQ6W3SHq7uS/br/c9wh+jCBNGBDb
# 4k7BHAxByudXObLXoEMyZEsbs40WLYQctyVjCATSIYo39+BypvuQAhSZTU90wxRQ
# 9r74zFgyuwOkcCQ4Q5hzxU51V5J10r0mLxpgesbmyH2ecPpTTLtnRz1RhOSTRORq
# Gz3XOUdfLPxvi5YGG39PidhUoKxvkd2G/7nGqa0TwsXwq6RnRjZ7d0DBDSCeg+ud
# xSObBag3FP1GFz+sA8kTcdVHDwE90kRcevw+m6koR8rk+2Cz2UMMZgdPxW7ULFln
# AvpT+UhGUevxeU7wDkmV4+Ylm4TAqB01Ht2iod0pvz7MLimZZ06izsm3lHiXgQ==
# SIG # End signature block
