<#
    Remove unwanted Start->Run entries

    @guyrleech 2019
#>

<#
.SYNOPSIS

Remove entries from explorer's Most Recently Used (MRU) list exposed in Start->Run and File->Run New Task in task manager

.PARAMETER remove

Remove items which match this regular expression

.PARAMETER keep

Only keep items which match this regular expression, removing all others

.EXAMPLE

&' .\Trim run history.ps1' -keep 'powershell|regedit'

Remove all entries except those that contain 'powershell' or 'regedit'

.EXAMPLE

&' .\Trim run history.ps1' -remove 'mstsc' -confirm:$false

Remove all entries that contain 'mstsc' without asking for confirmation

#>

[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='High')]

Param
(
    [Parameter(Mandatory=$true,ParameterSetName='Remove')]
    [string]$remove ,
    [Parameter(Mandatory=$true,ParameterSetName='Keep')]
    [string]$keep ,
    ## Don't change these!
    [string]$key = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU' ,
    [string]$MRUList = 'MRUList'
)

[string]$newMRUList = Get-ItemProperty -Path $key -Name $MRUList -ErrorAction SilentlyContinue |  Select-Object -ExpandProperty $MRUList
[int]$removed = 0

Get-ItemProperty -Path $key -Exclude $MRUList | ForEach-Object `
{
    $_.PSObject.Properties | Where-Object { $_.Name -match '^[a-z]$' -and $_.MemberType -eq 'NoteProperty' } | ForEach-Object `
    {
        $thisEntry = $_
        if( ( ! [string]::IsNullOrEmpty( $remove ) -and $thisEntry.Value -match $remove ) -or ( ! [string]::IsNullOrEmpty( $keep ) -and $thisEntry.Value -notmatch $keep ) )
        {
            if( $PSCmdlet.ShouldProcess( $thisEntry.Value , 'Remove entry' ) )
            {
                Write-Verbose "Removing `"$($thisEntry.Value)`""
                Remove-ItemProperty -Path $key -Name $thisEntry.Name
                if( $? )
                {
                    $removed++
                    $newMRUList = $newMRUList -replace $thisEntry.Name , ''
                }
            }
        }
    }
}

Write-Verbose "Removed $removed entries"

if( $removed )
{
    Set-ItemProperty -Path $key -Name $MRUList -Value $newMRUList
}
