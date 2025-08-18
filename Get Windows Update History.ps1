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

.PARAMETER passThru
    Return the update objects instead of formatting them as a table.
    Useful for further processing or exporting to other formats.

.PARAMETER defenderUpdate
    The title pattern used to identify Microsoft Defender updates for filtering.
    Default: "Security Intelligence Update for Microsoft Defender Antivirus"

.NOTES
    Modification History:

    2025/08/18  @guyrleech  Script born out of AI
#>

[CmdletBinding()]

Param
(
    [string]$includeRegex  ,
    [string]$excludeRegex ,
    [datetime]$since = [datetime]::MinValue ,
    [switch]$sortAscending ,
    [switch]$allUpdates ,
    [switch]$KBonly ,
    [switch]$passThru ,
    [string]$defenderUpdateRegex = 'Security Intelligence Update for Microsoft Defender Antivirus' ,
    [string]$notAvailable = 'N/A'
)
$Session = New-Object -ComObject Microsoft.Update.Session
$Searcher = $Session.CreateUpdateSearcher()
$History = $Searcher.QueryHistory(0, $Searcher.GetTotalHistoryCount())

[int]$totalUpdates = 0

# Process each update
$updates = @( foreach ($Update in $History)
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

   $result ## output
}) | Sort-Object Date -Descending:$(-Not $sortAscending) 

Write-Verbose "Got $($Updates.Count) updates out of $totalUpdates total"

if( $passThru )
{
    $Updates 
}
else
{
    $Updates | Select-Object -Property Date,kbnumber,result,title,description | Format-Table -AutoSize
}
