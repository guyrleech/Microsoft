<#
.SYNOPSIS
    Use Windows Update COM interface to show update history

.DESCRIPTION
    This script uses the Microsoft.Update.Session COM object to query the Windows Update history 
    and display information about installed updates. It provides filtering options to include/exclude 
    updates based on regex patterns, date ranges, and update status. By default, it shows only 
    successfully installed updates (excluding Defender updates) but can be configured to show all updates.

.PARAMETER includeRegex
    Regular expression pattern to filter updates by title. Only updates matching this pattern will be included.
    Example: "KB5.*" to show only KB updates starting with 5

.PARAMETER excludeRegex
    Regular expression pattern to exclude updates by title. Updates matching this pattern will be filtered out.
    Example: "Preview|Beta" to exclude preview and beta updates

.PARAMETER since
    Show only updates installed on or after this date. 
    Accepts standard DateTime formats like "2024-01-01" or "01/01/2024"
    Default: Shows all updates (DateTime.MinValue)

.PARAMETER sortAscending
    Sort the results by date in ascending order (oldest first).
    Default: Results are sorted in descending order (newest first)

.PARAMETER allUpdates
    Show all updates regardless of status or type.
    Default: Shows only successfully installed updates, excluding Defender updates

.PARAMETER KBonly
    Show only updates that are KB (Knowledge Base) articles.
    Default: Shows all updates
    
.PARAMETER computerName
    List of 1 or more computers to query using PowerShell remoting

.PARAMETER passThru
    Return the update objects instead of formatting them as a table.
    Useful for further processing or exporting to other formats.

.PARAMETER defenderUpdate
    The title pattern used to identify Microsoft Defender updates for filtering.
    Default: "Security Intelligence Update for Microsoft Defender Antivirus"

.NOTES
    Modification History:

    2025/08/18  @guyrleech  Script born out of AI
    2025/08/19  @guyrleech  Added -computername
#>

[CmdletBinding()]

Param
(
    [string]$includeRegex  ,
    [string]$excludeRegex ,
    [datetime]$since = [datetime]::MinValue ,
    [string[]]$computerName = @( '.' ),
    [switch]$sortAscending ,
    [switch]$allUpdates ,
    [switch]$KBonly ,
    [switch]$passThru ,
    [string[]]$displayProperties = @( 'Date','kbnumber','result','title','description' ) ,
    [string]$defenderUpdateRegex = 'Security Intelligence Update for Microsoft Defender Antivirus' ,
    [string]$notAvailable = 'N/A'
)

[array]$updates = @( foreach( $computer in $computerName )
{
    Write-Verbose "Checking $computer"
    [hashtable]$parameters = @{
        scriptBlock = {
            $Session = New-Object -ComObject Microsoft.Update.Session
            $Searcher = $Session.CreateUpdateSearcher()
            $Searcher.QueryHistory(0, $Searcher.GetTotalHistoryCount())
        }
    }
    if( $computer -ine '.'-and $computer -ine 'localhost' -and $computer -ine $env:COMPUTERNAME )
    {
        $parameters.Add( 'ComputerName' , $computer )
        if( $displayProperties -notcontains 'Computer' )
        {
            $displayProperties = , 'Computer' + $displayProperties 
        }
    }
    $history = @()
    $history = Invoke-Command @parameters
    
    Write-Verbose "Got $($history.Count) updates from $computer"
    [int]$totalUpdates = 0

    foreach ($Update in $History)
    {
        $totalUpdates++
        $result = [PSCustomObject]@{
            Date        = $Update.Date
            Title       = $Update.Title
            Description = $Update.Description
            KBNumber    = if ($Update.Title -match '\bKB\d+') { $Matches[0] } else { $notAvailable }
            Operation   = switch ($Update.Operation) {
                            1 { 'Installation' }
                            2 { 'Uninstallation' }
                            default { 'Unknown' }
                          }
            Result      = switch ($Update.ResultCode) {
                            0 { 'Not started' }
                            1 { 'In progress' }
                            2 { 'Succeeded' }
                            3 { 'Succeeded with errors' }
                            4 { 'Failed' }
                            5 { 'Aborted' }
                            default { 'Unknown' }
                          }
        }
        if( $KBonly -and $result.KBNumber -eq $notAvailable )
        {
            continue
        }
        elseif( -Not [string]::IsNullOrEmpty( $includeRegex ) -and $result.title -notmatch $includeRegex )
        {
            continue
        }
        elseif( -Not [string]::IsNullOrEmpty( $excludeRegex ) -and $result.title -match $excludeRegex )
        {
            continue
        }
        elseif( $null -ne $result.Date -and $result.Date -lt $since )
        {
            continue
        }
        elseif( -Not $allUpdates -and ( $result.Result -ine 'Succeeded' -or $result.Operation -ine 'Installation' -or $result.Title -match $defenderUpdateRegex ) )
        {
            continue
        }

        if( $parameters.ContainsKey( 'ComputerName' ) )
        {
            Add-Member -InputObject $result -MemberType NoteProperty -Name Computer -Value $computer -PassThru
        }
        else ## local
        {
            $result ## output
        }
    }
} ) | Sort-Object Date -Descending:$(-Not $sortAscending) 

Write-Verbose "Got $($Updates.Count) updates out of $totalUpdates total"

if( $passThru )
{
    $Updates 
}
else
{
    $Updates | Select-Object -Property $displayProperties | Format-Table -AutoSize
}

# SIG # Begin signature block
# MIIktwYJKoZIhvcNAQcCoIIkqDCCJKQCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAkdN1eXEFq+z8t
# OjTPc4bzafd7O5Oih7BM6E9kllU0AqCCH2AwggV9MIIDZaADAgECAhAB1rN1Nl8g
# zZEd1y/l+ZNkMA0GCSqGSIb3DQEBCwUAMFoxCzAJBgNVBAYTAkxWMRkwFwYDVQQK
# ExBFblZlcnMgR3JvdXAgU0lBMTAwLgYDVQQDEydHb0dldFNTTCBHNCBDUyBSU0E0
# MDk2IFNIQTI1NiAyMDIyIENBLTEwHhcNMjUwNzIxMDAwMDAwWhcNMjYwNzIwMjM1
# OTU5WjBxMQswCQYDVQQGEwJHQjESMBAGA1UEBxMJV2FrZWZpZWxkMSYwJAYDVQQK
# Ex1TZWN1cmUgUGxhdGZvcm0gU29sdXRpb25zIEx0ZDEmMCQGA1UEAxMdU2VjdXJl
# IFBsYXRmb3JtIFNvbHV0aW9ucyBMdGQwdjAQBgcqhkjOPQIBBgUrgQQAIgNiAARE
# VushBxmaLDZJys/h4fGHMe+gEacCcTcalje+NTkKlUboku0+BdNDPxotbsh0aHHv
# HhwndrrL7f/pD45f5VUVKK5F3rQY7bZjZ6gxwGa/BzuFZsRhO12MTMC7zawyaQCj
# ggHUMIIB0DAfBgNVHSMEGDAWgBTJ/BDvUMjLa3+9CETvOmKT7VtemjAdBgNVHQ4E
# FgQU+7W5w1B8mWyXaJebkcnMJkcSB+cwPgYDVR0gBDcwNTAzBgZngQwBBAEwKTAn
# BggrBgEFBQcCARYbaHR0cDovL3d3dy5kaWdpY2VydC5jb20vQ1BTMA4GA1UdDwEB
# /wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzCBlwYDVR0fBIGPMIGMMESgQqBA
# hj5odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vR29HZXRTU0xHNENTUlNBNDA5NlNI
# QTI1NjIwMjJDQS0xLmNybDBEoEKgQIY+aHR0cDovL2NybDQuZGlnaWNlcnQuY29t
# L0dvR2V0U1NMRzRDU1JTQTQwOTZTSEEyNTYyMDIyQ0EtMS5jcmwwgYMGCCsGAQUF
# BwEBBHcwdTAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tME0G
# CCsGAQUFBzAChkFodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vR29HZXRTU0xH
# NENTUlNBNDA5NlNIQTI1NjIwMjJDQS0xLmNydDAJBgNVHRMEAjAAMA0GCSqGSIb3
# DQEBCwUAA4ICAQALHZsdOuMeT/e1fsdQfhIz/2wS19UWlG1lxXieYOmPAju0DA5I
# ZheTgMtWMkUm96gWNtixny+q5nX8ckzuuD47esI2bM4G9RcVVKN0vdLZHv6QXZE5
# Ht2qTX8E1bqfejtDcGY0aqOjVYeOi/o98BsR98ItkjWNP3xE2oKEx6xYyZBL6d/z
# HB2ySd7hdk4VfmH9rTRftAsAn5L9s6m2ILRK8QRrkUJY9RxXswvQy2gNzccg+eYw
# y5gvLnzp4kdsTleV8SyZZQ2Tcp+HHPGxekB1NIM55vlCb9ocYw5j7noae3/PF+u/
# Zt/E+copm1c+MDju2bz1EelqXxuVsICRV9ikpJ7QEU+LwUiT7Ne+mgBmQ3IIyb8d
# QwR0xu1E/sKoWZjPRha6JLe65RaoBnXOX6fWQglPx467qjTUQpLxKGKMdQjS+LGJ
# uI2/BMWBHJfdhz/3GR9XVaDOWLhk+ChkjoXBgF2uFEXSiv4LNgigQ1R9RiojukEz
# mSRe5LK0UPdJch5I/HXg1lJFPORx05Ila0uSMusisgrPNvl5fEuf+DGYl2ywHsZQ
# pkeKT5wHQ5QJEocTKwGPfiGe9drO5DoMos5AXL5xnrPh/aQB4XKrttZTFy0+YXrU
# WSYa9v2rp7cAwuDsmd9gLoQzfW7jagbHfxvmLQ0CJTO9Y4BsJZML4bjimDCCBY0w
# ggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTELMAkG
# A1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRp
# Z2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENB
# MB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMCVVMx
# FTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNv
# bTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE98orY
# WcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9SH8ae
# FaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g1ckg
# HWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RYjgwr
# t0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgDEI3Y
# 1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNAvwjX
# WkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDgohIb
# Zpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQAzH0c
# lcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOkGLim
# dwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHFynIW
# IgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gdLfXZ
# qbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFOzX
# 44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3z
# bcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDBF
# BgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkqhkiG
# 9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7IviH
# GmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/59Pes
# MHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0POz3
# A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISfb8rb
# II01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhULSd+
# 2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBqEwggSJoAMCAQICEAeEPa0B
# wRXCdO5BpygiRnkwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8G
# A1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTIyMDYyMzAwMDAwMFoX
# DTMyMDYyMjIzNTk1OVowWjELMAkGA1UEBhMCTFYxGTAXBgNVBAoTEEVuVmVycyBH
# cm91cCBTSUExMDAuBgNVBAMTJ0dvR2V0U1NMIEc0IENTIFJTQTQwOTYgU0hBMjU2
# IDIwMjIgQ0EtMTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAK0e9Aey
# Q2aKomd3JZUKpfgW1inkV8ks71KHQG7Be68F41i3bXF/yH+ksn/tjN4pw2r3PjMS
# F5yq0PTGFu9IHySwEB9YExk7Q0t85PSPtbI24Puu/5kXNr6bEhDv2zV0KLBQzAai
# dqgMruapl8OkoQTFHpTIoGHdpq1PvdTYibH/H59hOZAWr43wWsuzoWHpgQZYlOCz
# HLDV8AKEJ+C0RxmR21yAruq6qyQe1bo8n2XlU0ntPdZenOew47GvPerHQLNaPArz
# 7cq/ZfqmJa93xhF8A7JxKtPj88zRwcsVznGz/ib96TgUKdhJ6bjd2gV+0HFkFCQc
# qg9bcG3pUiuFUjC87uGtSOFyEkllh1KV3dsA0O+Inn9Og/sQ2/3tT0y0oW+YLl3N
# 3WngfEVHSnnBaZhBtb7LdEWbSenof2bnsxQw2nTKyZ0mNvR/v51Utfc4QFRvof6v
# UtEtlP/EQ5O7A7EaDZLjDbkoYv1IIFRbieQGWM8d4lOhT5Me3Q/xlB/gH0gWUbG0
# srSDe44CfBIWAq2Y2OGROXxxosBDBuAQg0KquFzRvqTZnz5DCBUAvSci7Mz1yQo/
# zG0hxaDzkvf+XRxOQdB176MoYMbYJ3EEe1+UshS3CoPskDa4LFaUjNE2DBPSQfZ6
# fT5BXwrOJIZqo7avTqTBuirb/HLh73wIR9JjAgMBAAGjggFZMIIBVTASBgNVHRMB
# Af8ECDAGAQH/AgEAMB0GA1UdDgQWBBTJ/BDvUMjLa3+9CETvOmKT7VtemjAfBgNV
# HSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNVHQ8BAf8EBAMCAYYwEwYD
# VR0lBAwwCgYIKwYBBQUHAwMwdwYIKwYBBQUHAQEEazBpMCQGCCsGAQUFBzABhhho
# dHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUHMAKGNWh0dHA6Ly9jYWNl
# cnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3J0MEMGA1Ud
# HwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRy
# dXN0ZWRSb290RzQuY3JsMBwGA1UdIAQVMBMwBwYFZ4EMAQMwCAYGZ4EMAQQBMA0G
# CSqGSIb3DQEBCwUAA4ICAQAL2wrXsh2YpMJRq0Szv7J7CEmcni3KsvA0Sd+Xoesb
# w+bsdnRv7klz4YaolPyRFzuaKG5Wt2xg0d2Jy54Mv2KaG0K6wj+tSaPCG1+Wn5eA
# uSaAsauawSGjVv6UaJGnssL/XR2LxIA6KUOQePlnHXjbFG8GutaP1451IYfBi5sC
# wR3oIOiPBxpXP2l9czNg7eRzQ7o9eDVORyCRiUJQC4Me6T+xnHrxbQVWPU/aIyH5
# VSr2UvWm6ugDJ2h5ZVSH/5wxd6p+CkGqUBb6vyZrkXrIovSy1VAfy9hvURLSYlIg
# /Ih+QiMLVXSltlLenRBawppp8Sivx8t8s19LGe1Uj+6C63T7p6SW49ZGkRcf4kCI
# 11GNsttlyEJER7f9eXRiXCQD55BUl8zQDteK4W1j+Y6krYA34Tbm3mZBgWGl3Flf
# ktMMpaAOdv4DpCcS3iI3K6Rwtom5Lwg+A/PQTYAtku3dGqz6VeJ8r8bCc0hZA1t/
# sWYsMH2mHyLx2+xHWGyNzYo8RbhsCxu8tzPyE3XMTX9Bs5X3a8MagWPBk6LaShD5
# TISHR7yMMUcAol5Mj6nwQ8T+aq+iMsUCe3fXJcgDa2O3QRG2yOSEE2ZljpIQ5+ci
# g7C/Kpq8s/JrVS3f/Zw4Uu41DxCXodpmw1ASuefOf5kQBpROQ9lp5XKgciQ4MQIv
# OTCCBrQwggScoAMCAQICEA3HrFcF/yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAw
# YjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQ
# d3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290
# IEc0MB4XDTI1MDUwNzAwMDAwMFoXDTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMC
# VVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBU
# cnVzdGVkIEc0IFRpbWVTdGFtcGluZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCC
# AiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBALR4MdMKmEFyvjxGwBysdduj
# Rmh0tFEXnU2tjQ2UtZmWgyxU7UNqEY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S
# 9SLrC6Kbltqn7SWCWgzbNfiR+2fkHUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+
# 42DFUF0mR/vtLa4+gKPsYfwEu7EEbkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg6
# 2IVwxKSpO0XaF9DPfNBKS7Zazch8NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21
# Qomb+zzQWKhxKTVVgtmUPAW35xUUFREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8
# y9IaaGBpPNXKFifinT7zL2gdFpBP9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQ
# NfVmUB5KlCX3ZA4x5HHKS+rqBvKWxdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gao
# u30yZ46t4Y9F20HHfIY4/6vHespYMQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6g
# qztiT96Fv/9bH7mQyogxG9QEPHrPV6/7umw052AkyiLA6tQbZl1KhBtTasySkuJD
# psZGKdlsjg4u70EwgWbVRSX1Wd4+zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D
# 8bpfm4CLKczsG7ZrIGNTAgMBAAGjggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEA
# MB0GA1UdDgQWBBTvb1NK6eQGfHrK4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC
# 0nFdZEzfLmc/57qYrhwPTzAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYB
# BQUHAwgwdwYIKwYBBQUHAQEEazBpMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5k
# aWdpY2VydC5jb20wQQYIKwYBBQUHMAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSG
# Mmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQu
# Y3JsMCAGA1UdIAQZMBcwCAYGZ4EMAQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0B
# AQsFAAOCAgEAF877FoAc/gc9EXZxML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6F
# TGNpoV2V4wzSUGvI9NAzaoQk97frPBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mC
# efSG+tXqGpYZ3essBS3q8nL2UwM+NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57m
# QfQXwcAEGCvRR2qKtntujB71WPYAgwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9
# ydOal95CHfmTnM4I+ZI2rVQfjXQA1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dB
# wp9nEC8EAqoxW6q17r0z0noDjs6+BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdq
# fMTCW/NmKLJ9M+MtucVGyOxiDf06VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2
# puE6FndlENSmE+9JGYxOGLS/D284NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAO
# k5eCkhSxZON3rGlHqhpB/8MluDezooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL
# 0Q4ssd8xHZnIn/7GELH3IdvG2XlM9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBun
# vAZapsiI5YKdvlarEvf8EA+8hcpSM9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE
# 1aADAgECAhAKgO8YS43xBYLRxHanlXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNV
# BAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNl
# cnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBD
# QTEwHhcNMjUwNjA0MDAwMDAwWhcNMzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJV
# UzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNI
# QTI1NiBSU0E0MDk2IFRpbWVzdGFtcCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkq
# hkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHf
# yjfMGUIwYzKomd8U1nH7C8Dr0cVMF3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPx
# NyFPJIDZHhAqlUPt281mHrBbZHqRK71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpk
# BaMUNg7MOLxI6E9RaUueHTQKWXymOtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFv
# ZSjKs3SKO1QNUdFd2adw44wDcKgH+JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1zn
# OM8odbkqoK+lJ25LCHBSai25CFyD23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8f
# cpK40uhktzUd/Yk0xUvhDU6lvJukx7jphx40DQt82yepyekl4i0r8OEps/FNO4ah
# fvAk12hE5FVs9HVVWcO5J4dVmVzix4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUD
# y9Z2hSgctaepZTd0ILIUbWuhKuAeNIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9
# w6CtjuuVHJOVoIJ/DtpJRE7Ce7vMRHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTn
# nkrT3pXWETTJkhd76CIDBbTRofOsNyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKa
# cJ+A9/z7eacCAwEAAaOCAZUwggGRMAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7
# /PIx7f391/ORcWMZUEPPYYzoMB8GA1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ
# 7MtOMA4GA1UdDwEB/wQEAwIHgDAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYI
# KwYBBQUHAQEEgYgwgYUwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0
# LmNvbTBdBggrBgEFBQcwAoZRaHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0Rp
# Z2lDZXJ0VHJ1c3RlZEc0VGltZVN0YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEu
# Y3J0MF8GA1UdHwRYMFYwVKBSoFCGTmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9E
# aWdpQ2VydFRydXN0ZWRHNFRpbWVTdGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0Ex
# LmNybDAgBgNVHSAEGTAXMAgGBmeBDAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcN
# AQELBQADggIBAGUqrfEcJwS5rmBB7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF
# 0RkP2AGr181o2YWPoSHz9iZEN/FPsLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKq
# dT8wv2UV+Kbz/3ImZlJ7YXwBD9R0oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbU
# UO75ZSpbh1oipOhcUT8lD8QAGB9lctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTe
# HihsQyfFg5fxUFEp7W42fNBVN4ueLaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG
# 7aEQJmmrJTV3Qhtfparz+BW60OiMEgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NB
# qycz0BZwhB9WOfOu/CIJnzkQTwtSSpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6
# +iX8MmB10nfldPF9SVD7weCC3yXZi/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaA
# yBjFBtXVLcKtapnMG3VH3EmAp/jsJ3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyP
# ehwJVxwC+UpX2MSey2ueIu9THFVkT+um1vshETaWyQo8gmBto/m3acaP9QsuLj3F
# NwFlTxq25+T4QwX9xa6ILs84ZPvmpovq90K8eWyG2N01c4IhSOxqt81nMYIErTCC
# BKkCAQEwbjBaMQswCQYDVQQGEwJMVjEZMBcGA1UEChMQRW5WZXJzIEdyb3VwIFNJ
# QTEwMC4GA1UEAxMnR29HZXRTU0wgRzQgQ1MgUlNBNDA5NiBTSEEyNTYgMjAyMiBD
# QS0xAhAB1rN1Nl8gzZEd1y/l+ZNkMA0GCWCGSAFlAwQCAQUAoIGEMBgGCisGAQQB
# gjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYK
# KwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEID7nKOpX
# isU27j9xn9kyp8OXFPC+95/gQIfXV9eCAAaXMAsGByqGSM49AgEFAARnMGUCMDrf
# UW6JAOE/kiWGMnz4HBW8GiTPFB5quwhhoXklTvQAe/3v9k7C4tst9h2ySIY48QIx
# ANAdutXCrCIjQLqW5048FIGHtuW1CQotxciCcm4w6xcJ/H0dhnsjKTmsKMkHN7pK
# LqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVT
# MRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1
# c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA
# 7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJ
# KoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNTA4MTgyMzM3NDFaMC8GCSqGSIb3
# DQEJBDEiBCBNnOOJokuA8z0lugY3DAdmgvY9LRNxR/xxh+PVhy2hojANBgkqhkiG
# 9w0BAQEFAASCAgAUTO1tX56DSVAvvUrRUnxOmugao6HiA5/G/LyYcxY7HvOBlhYy
# LXDSmeOunSZMd3ZjArDL1qDolWrhAYti8MkDCJ4kDSBLuXGKatN3DolF3694DNB3
# wPh0xren6SU0/A/L1F8h1aFP/cmLnz6ejjLoE3SZ5aeYxUVHIeZNNj2hciIkuZTA
# zq6VOXVa1gkYdqDCLuI3823K41KOA5fYrBHXxzKKapIna0H0BcQQB84jHt4mXjEL
# 0B9/Dhi9NKn8rKSXaOuSasCK5zmtEXfTqccD8xaRHmvpPCXR04c6eyqisFDCd1tx
# IBUOsglvixU9+ekTp0/RnlVNzR5SMemnwxUvRtNLoqEsOHYAuco7+bh+PEyDje2+
# 24lZNowzwPHern3O+e1ZpownHYbgDXOT/Yw/n47GUdQk/CJSuCEvPUPIhkIyxU+z
# 09fNflOHLYvuH/6SmCAGPYyeuc40MvocBK1DC8NsEYX999sT8nK1hWUPuye9xxbR
# cPB2hZDyrhtrbMGfOHNkRlWUQBekjCvHN/vbzWVsPsIdiOdk4rx/71J5QsgXeKJU
# 6kDcK9i2qv759MuFw7biFbsaRAh++A7QLRSwOSXIylBlUYA29fDRhQYa7qf5qBBz
# 6d8JqDBRkc8zcdTIZFVMb1ef20FA8IoCXS4ncPv3Xi6Y2Ef//VQWXCi9hw==
# SIG # End signature block
