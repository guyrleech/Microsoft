#requires -RunAsAdministrator

<#
.SYNOPSIS
    Show available windows updates and optionally install

.PARAMETER query
    The query to be used to search for windows updates

.PARAMETER force
    Force the restart of the computer if one is required

.NOTES
    https://learn.microsoft.com/en-us/windows/win32/api/wuapi/nf-wuapi-iupdatesearcher-search

    Modification History:

    2020/11/20  @guyrleech  Script born
#>

<#
Copyright © 2025 Guy Leech

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, 
including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

[CmdletBinding( SupportsShouldProcess = $True , ConfirmImpact = 'High' )]

Param
(
    [string]$query = "IsInstalled = 0 and Type = 'Software'" ,
    [switch]$force
)

## https://learn.microsoft.com/en-us/windows/win32/api/wuapi/ne-wuapi-operationresultcode
[string[]]$resultCodes = @(
  'NotStarted' ## 0,
  'In Progress' ## 1,
  'Succeeded' ## 2,
  'Succeeded With Errors' ## 3,
  'Failed' ## 4,
  'Aborted' ## 5
)

$updateSession = New-Object -comobject "Microsoft.Update.Session"

$updateSearcher = $updateSession.CreateupdateSearcher()
$updateSearcher.Online = $true
Write-Verbose -Message "$([datetime]::Now.ToString('G')): searching for updates with `"$query`""
$searchResult = $null
$searchResult = $updateSearcher.Search( $query )

if( $null -eq $searchResult -or $null -eq $searchResult.Updates -or $searchResult.Updates.Count -eq 0 )
{
    Write-Verbose "$([datetime]::Now.ToString('G')): No updates"
}
else
{
    try
    {
        Write-Verbose "$([datetime]::Now.ToString('G')): $($searchResult.updates.Count) updates available"
        $searchResult.Updates | Select-Object -Property Title, @{n='KBArticleID';e={$_.KBArticleIDs}}, MsrcSeverity | Format-Table | Out-String | Write-Verbose
        if( $PSCmdlet.ShouldProcess( $env:COMPUTERNAME , "Download & Install $($searchResult.updates.Count) updates" ) )
        {
            $updatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl

            foreach ($update in $searchResult.Updates)
            {
                [void]$updatesToInstall.Add($update)
            }

            Write-Verbose -Message "$([datetime]::Now.ToString('G')): downloading updates ..."
            $downloader = $null
            $downloader = $updateSession.CreateUpdateDownloader()
            $downloader.Updates = $updatesToInstall
            [void]$downloader.Download()
            
            Write-Verbose -Message "$([datetime]::Now.ToString('G')): installing updates ..."
            $installer = $updateSession.CreateUpdateInstaller()
            $installer.Updates = $updatesToInstall
            $installationResult = $null
            $installationResult = $installer.Install()
            
            Write-Verbose -Message "$([datetime]::Now.ToString('G')): installation finished, result is $($installationResult.ResultCode)"
            [string]$result = 'Unknown'
            if( $installationResult.ResultCode -ge 0 -and $installationResult.ResultCode -lt $resultCodes.Count )
            {
                $result = $resultCodes[ $installationResult.ResultCode ]
            }
            Add-Member -InputObject $installationResult -MemberType NoteProperty -Name Result -Value $result -PassThru ## output
            if( $installationResult.RebootRequired )
            {
                if( $PSCmdlet.ShouldProcess( $env:COMPUTERNAME , 'Reboot' ) )
                {
                    Write-Verbose -Message "$([datetime]::Now.ToString('G')): rebooting"
                    Restart-Computer -Force:$force -Confirm:$false
                }
                else
                {
                    Write-Warning -Message "Updates will not all apply until after a reboot"
                }
            }
        }
    }
    catch
    {
        throw
    }
}
