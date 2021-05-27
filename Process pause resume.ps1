<#
.SYNOPSIS

    Pause/resume processes by pausing/resuming all their threads

.DESCRIPTION

    Note that if a thread is already paused when the script tries to pause it, its pause count will be increased so that when this script resumes it, it will still be paused since something else paused it, not this script (possible the program itself).

.PARAMETER id

    Process Ids to operate on

.PARAMETER name

    Process names (without .exe) to operate on

.PARAMETER sessionIds

    Only operate on processes in the give session ids

.PARAMETER resume

    Resume all threads. Without this all threas will be paused

.PARAMETER allSessions

    Operate on processes in all sessions. Without this only the current session is connsidered unless -sessionIds is used

.PARAMETER quiet

    No extraneous output
    
.PARAMETER windowControl

    Minimise the window associated with the process if pausing and restore the window if resuming

.PARAMETER trim

    Trim the working set of processes which have been paused

.PARAMETER logfile

    Write to the specified log file

.PARAMETER append

    Append to ann existing log file. Without this it will be overwritten

.EXAMPLE

    & '.\Process pause resume.ps1' -name badboy -trim

    Pause all the threads in all badboy.exe processes in the current session and trim their working sets.
    
.EXAMPLE

    & '.\Process pause resume.ps1' -name badboy -resume

    Resume all the threads in all badboy.exe processes in the current session.

.NOTES

    Return value of SuspendThread/ResumeThread is the previous suspend count so if SuspendThread returns > 0, thread already suspended & if ResumeThread returns >1, thread will still be suspended

    Modification History:

    @guyrleech 2021/02/02  Initial version
    @guyrleech 2021/05/27  Initial public release
    @guyrleech 2021/05/27  Added ability to minimise/restore window
#>

[CmdletBinding()]

Param
(
    [Parameter(ValueFromPipelineByPropertyName=$true,ParameterSetName='Id',Mandatory = $true)]
    [int[]]$id ,
    [Parameter(ValueFromPipelineByPropertyName=$true,ParameterSetName='Name',Mandatory = $true)]
    [string[]]$name ,
    [Parameter(ValueFromPipelineByPropertyName=$true,ParameterSetName='Session',Mandatory = $true)]
    [int[]]$sessionIds ,
    [switch]$resume ,
    [switch]$allSessions ,
    [switch]$quiet ,
    [switch]$windowControl ,
    [switch]$trim ,
    [string]$logfile ,
    [int]$maximumMinimisedRetries = 10 ,
    [switch]$append
)

Begin
{
    [string]$logging = $null

    if( $PSBoundParameters[ 'logfile' ] )
    {
        $logging = Start-Transcript -Path $logfile -Append:$append
    }
    
    Add-Type @'
        using System;
        using System.Runtime.InteropServices;

        public static class Kernel32
        {
            [DllImport( "kernel32.dll",SetLastError = true )]
            public static extern IntPtr OpenThread( 
                UInt32 dwDesiredAccess, 
                bool bInheritHandle, 
                UInt32 dwThreadId );

            [DllImport("kernel32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool CloseHandle(
                [In] IntPtr hHandle );
            
            // https://docs.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-suspendthread

            [DllImport("kernel32.dll", SetLastError = true)]
            public static extern int SuspendThread(
                [In] IntPtr hThread );

            [DllImport("kernel32.dll", SetLastError = true)]
            public static extern int Wow64SuspendThread(
                [In] IntPtr hThread );

            // https://docs.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-resumethread

            [DllImport("kernel32.dll", SetLastError = true)]
            public static extern int ResumeThread(
                [In] IntPtr hThread );

            [DllImport("kernel32.dll", SetLastError = true)]
            public static extern bool IsWow64Process(
                [In] IntPtr hProcess ,
                [Out,MarshalAs(UnmanagedType.Bool)] out bool wow64Process );
            
            [DllImport("kernel32.dll", SetLastError = true)]
                public static extern bool SetProcessWorkingSetSizeEx( IntPtr proc, int min, int max , int flags );

            public enum ThreadAccess
            {
                THREAD_SUSPEND_RESUME = 0x2 ,
                THREAD_QUERY_INFORMATION = 0x40 ,
                THREAD_QUERY_LIMITED_INFORMATION = 0x800,
            };
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
            public static extern bool SetWindowPos( IntPtr hWnd, IntPtr hWndInsertAfter, int  X, int  Y, int  cx, int  cy,  uint uFlags);
        }
'@ -ErrorAction Stop

    Function Invoke-PauseResumeProcess
    {
        [CmdletBinding()]

        Param
        (
            [int[]]$id ,
            [bool]$resume ,
            [bool]$trim ,
            [bool]$windowControl 
        )
        
        [int]$operatedOn = 0
        [int]$threadsNotChanging = 0

        ForEach( $processId in $id )
        {
            if( $processId -eq $pid )
            {
                Write-Warning -Message "Not operating on self (pid $pid)"
            }
            elseif( $thisProcess = Get-Process -Id $processId -ErrorAction SilentlyContinue )
            {
                ## if pausing then must change window state before hand
                
                if( $windowControl -and ! $resume -and $thisProcess.MainWindowHandle -ne [IntPtr]::Zero )
                {
                    ## sync call so minimised before we pause threads
                    if( ! [user32]::ShowWindow( $thisProcess.MainWindowHandle , 11 ) ) ## SW_FORCEMINIMIZE
                    {
                        if( ! $quiet )
                        {
                            Write-Warning -Message "Failed to minimise window for pid $($thisProcess.Id)"
                        }
                    }
                    else
                    {
                        [int]$retries = 0 
                        While( ! [user32]::IsIconic( $thisProcess.MainWindowHandle ) -and $retries++ -le $maximumMinimisedRetries )
                        {
                            Start-Sleep -Milliseconds 250
                        }
                    }
                }

                [int]$is32bitProcess = -1
                [int]$threadsOperatedOn = 0

                Write-Verbose -Message "$(if( $resume ) { 'Resuming' } else { 'Pausing' }) $($thisProcess.Threads.Count) threads in pid $processId ($($thisProcess.Name)) in session $($thisProcess.SessionId)"
           
                if( ! $thisProcess.Handle -or ! [kernel32]::IsWow64Process( $thisProcess.Handle , [ref]$is32bitProcess ) )
                {
                    Write-Warning -Message "Failed to determine if process $processId is 32 or 64 bit"
                }
                ForEach( $thread in $thisProcess.Threads )
                {
                    ## if thread already suspended, we will suspend again so that threads which were suspended before we interfered with them get an extra suspend so upon resume are still suspended
                    [IntPtr]$threadHandle = [kernel32]::OpenThread(  [Kernel32+ThreadAccess]::THREAD_SUSPEND_RESUME , $false , $thread.Id ) ; $lastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
                    if( $threadHandle -ne [IntPtr]::Zero)
                    {
                        [string]$warning = $null
                        [int]$result = -1
                        if( $resume )
                        {
                            $result = [kernel32]::ResumeThread( $threadHandle ) ;  $lastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
                            
                            if( $result -eq 0 )
                            {
                                $warning = "Thread id $($thread.Id) in process id $processId was not paused"
                            }
                            elseif( $result -gt 1 )
                            {
                                $warning = "Thread id $($thread.Id) in process id $processId is still paused (count $result)"
                            }
                        }
                        else ## suspend
                        {
                            if( $is32bitProcess )
                            {
                                $result = [kernel32]::Wow64SuspendThread( $threadHandle ) ; $lastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
                            }
                            else
                            {
                                $result = [kernel32]::SuspendThread( $threadHandle ) ;  $lastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
                            }
                            
                            if( $result -gt 0 )
                            {
                                $warning = "Thread id $($thread.Id) in process id $processId was already paused (count $result)"
                            }
                        }

                        if( $warning )
                        {
                            $threadsNotChanging++
                            if( ! $quiet )
                            {
                                Write-Warning -Message $warning
                            }
                        }
        
                        if( $result -ge 0 )
                        {
                            $threadsOperatedOn++
                            Write-Verbose -Message "`tReturned $result for tid $($thread.id) pid $($thisProcess.Id)"
                        }
                        else
                        {
                            if( ! $quiet )
                            {
                                Write-Warning -Message "Failed to $(if( $resume ) { 'resume' } else { 'pause' }) thread id $($thread.Id) in process id $processId - $lasterror"
                            }
                        }
                        $null = [kernel32]::CloseHandle( $threadHandle )
                        $threadHandle = [IntPtr]::Zero
                    }
                    else
                    {
                        Write-Warning -Message "Failed to open thread id $($thread.Id) in process $processId - $lasterror"
                    }
                }
                
                if( $windowControl -and $resume -and $thisProcess.MainWindowHandle -ne [IntPtr]::Zero )
                {
                    if( ! [user32]::ShowWindowAsync( $thisProcess.MainWindowHandle , 9 ) -and ! $quiet ) ## SW_RESTORE
                    {
                        Write-Warning -Message "Failed to restore window for pid $($thisProcess.Id)"
                    }
                }

                if( $threadsOperatedOn )
                {
                    $operatedOn++
                }

                Write-Verbose -Message "Successfully operated on $threadsOperatedOn / $($thisProcess.Threads.Count) threads of pid $processId although $threadsNotChanging not changed state"

                if( ! $resume -and $trim )
                {
                    if( $thisProcess.Handle )
                    {
                        Write-Verbose -Message "Emptying working set of $([math]::Round( $thisProcess.WorkingSet64 / 1MB , 1 ))MB in pid $processId ($($thisProcess.Name)) in session $($thisProcess.SessionId)"

                        [bool]$trimmed = [kernel32]::SetProcessWorkingSetSizeEx( $thisProcess.Handle , -1 , -1 , 0 ); $LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()

                        if( ! $trimmed )
                        {
                            Write-Warning -Message "Failed to empty working set of pid $processId - $LastError"
                        }
                    }
                    else
                    {
                        Write-Warning -Message "Failed to get handle to pid $processId to empty working set"
                    }
                }
            }
            else
            {
                Write-Warning -Message "Failed to find process pid $processId"
            }
        }

        $operatedOn
    }

    [int]$operatedOn = 0
    [int]$processCount = 0

    if( $PSBoundParameters[ 'name' ] )
    {
        [int]$thisSessionId = Get-Process -Id $pid | Select-Object -ExpandProperty SessionId

        if( ! $thisSessionId -and ! $allSessions )
        {
            if( ! $quiet )
            {
                Throw "Failed to determine current session id from pid $pid"
            }
            else
            {
                Exit 1
            }
        }

        if( ! ( $id = @( Get-Process -Name $name -ErrorAction SilentlyContinue | Where-Object { ( $allSessions -or $_.SessionId -eq $thisSessionId ) -and $_.Id -ne $pid } | Select-Object -ExpandProperty Id ) ) -or ! $id.Count )
        {
            if( ! $quiet )
            {
                Throw "No processes found for process names $name"
            }
            else
            {
                Exit 2
            }
        }
    }
    elseif( $PSBoundParameters[ 'sessionids' ] )
    {   
        if( ! ( $id = @( Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -in $sessionIds -and $_.Id -ne $pid } | Select-Object -ExpandProperty Id ) ) -or ! $id.Count )
        {
            if( ! $quiet )
            {
                Throw "No processes found for session ids $($sessionIds -join ',')"
            }
            else
            {
                Exit 2
            }
        }
    }
}

Process
{
    if( $id -and $id.Count )
    {
        $processCount += $id.Count
        $operatedOn += Invoke-PauseResumeProcess -id $id -resume $resume -trim $trim -windowControl $windowControl
    }
}

End
{
    If( ! [string]::IsNullOrEmpty( $logging ) )
    {
        Stop-Transcript
    }
    $operatedOn ## return
}

# SIG # Begin signature block
# MIINRQYJKoZIhvcNAQcCoIINNjCCDTICAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUE+MUsLfOFe1jlwAUW0kWyapv
# fLagggqHMIIFMDCCBBigAwIBAgIQBAkYG1/Vu2Z1U0O1b5VQCDANBgkqhkiG9w0B
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
# CzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFIiBJO5wtD8BQa4fQvnf
# CjkpOXbAMA0GCSqGSIb3DQEBAQUABIIBAFtj5EeBqemvERw8gvNAWYGGRjGNJhq2
# Q3hUYwyYkOsv5tNxoKv6AgL7QI9Nx9tg2+DEBzIPBwmA0IO9PUECvHIAzu0JNTjP
# U90Zg8OxQY+vymycGux2btK4qXpr+CeAhw/wwI75Dti5Zv8FrWaQeEf7aoaCdUoH
# 22uBu3P4PoMBp7qweOu3e5RZACX8o7oRKkbcooD7HyYPZpcJXQrDUykbYL6GzrWR
# LIlZ2hq0o37KBM4LS2bBM/EUZNm9E+m8+Z1imnLijAumXszRBhwgs6jpwl1bWiS+
# GU24gmZAVqu6X+fW4b5gYLyUBsBnUDc0EOXKVNRzlmX7ZXbvKHXDjK0=
# SIG # End signature block
