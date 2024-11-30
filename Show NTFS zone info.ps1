#requires -version 3.0
<#
    Find/remove files with Zone.Identifier NTFS alternate data stream and show the content

    This is how files are marked by certain browsers on Windows 10 when they have come from an "untrusted" source

    Pipe to Out-GridView or Export-Csv

    @guyrleech 2018

    Modification History:

    02/08/18  GRL  Added created and modified times and file owner
    06/08/18  GRL  Added remove and scrub options
    30/11/24  GRL  Non-fatal if not NTFS
#>

<#
.SYNOPSIS

Find files with Zone.Identifier NTFS alternate data stream and show the content

.PARAMETER path

The NTFS folder to search in

.PARAMETER recurse

Recurse the folder structure

.PARAMETER webOnly

Only show files which have come from the web

.PARAMETER remove

Remove the alternate data stream completely

.PARAMETER scrub

Remove the URL information from the alternate data stream, leaving just the zone information

.EXAMPLE

& '.\Show NTFS zone info.ps1' -path C:\Temp

Show all files in the c:\temp folder which have Zone.Identifier information attached

& '.\Show NTFS zone info.ps1' -path C:\users -Recurse -WebOnly | Out-GridView

Show all files in the c:\users folder and sub folders which have Zone.Identifier information attached which have come from the web

& '.\Show NTFS zone info.ps1' -path C:\users -Recurse -WebOnly -Scrub | Out-Null

Remove the URL information from all Zone.Identifier alternate data streams on files in the c:\users folder and sub folders

.NOTES

This is how files are marked by Windows when they have come from an "untrusted" source
Pipe the script to Out-GridView or Export-CSV if required

#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$true)]
    [string]$path ,
    [switch]$recurse ,
    [switch]$webOnly ,
    [switch]$remove ,
    [switch]$scrub ,
    [string]$zoneStream = 'Zone.Identifier'
)

[hashtable]$params = @{ 'Path' = $path ; 'Recurse' = $recurse }

## check path is on NTFS
[string]$drive = Get-Item -Path $path -ErrorAction Stop | Select-Object -ExpandProperty PSdrive -ErrorAction SilentlyContinue

[string]$fileSystem = 'UNKNOWN'
if( $null -ne $drive )
{
    $fileSystem = ( Get-CimInstance -ClassName win32_logicaldisk -Filter "DeviceId = '$($drive):'" ).FileSystem
}

if( $fileSystem -ne 'NTFS' )
{
    Write-Warning "File system type of `"$path`" is $fileSystem, not NTFS, so there may not be alternate data streams"
}

## read zone info from registry so we can display name rather than number
[hashtable]$zones = @{}

Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\*' -name DisplayName | ForEach-Object `
{
    $zones.Add( $_.PSChildName , $_.DisplayName )
}

Get-ChildItem @params -File | ForEach-Object `
{ 
    try
    {
        $zoneInfo = Get-Content -Path $_.FullName -Stream $zoneStream -ErrorAction SilentlyContinue 
        if( $zoneInfo )
        {
            $zone = $zoneInfo | Where-Object { $_ -match '^ZoneId=(.*)$' } | ForEach-Object { $matches[1] }
            $ReferrerURL = $zoneInfo | Where-Object { $_ -match '^ReferrerURL=(.*)$' } | ForEach-Object { $matches[1] }
            $HostURL = $zoneInfo | Where-Object { $_ -match '^HostURL=(.*)$' } | ForEach-Object { $matches[1] }
            if( ! $webOnly -or ( ! [string]::IsNullOrEmpty( $ReferrerURL ) -and $ReferrerURL -match '^https?:' ) -or ( ! [string]::IsNullOrEmpty( $HostURL ) -and $HostURL -match '^https?:' ) )
            {
                [pscustomobject]@{
                    'File' = $_.FullName ; 
                    'Zone' = $zones[ $zone ] ; 
                    'Referrer URL' = $ReferrerURL ; 
                    'Host URL' = $HostURL ; 
                    'Created' = $_.CreationTime ; 
                    'Modified' = $_.LastWriteTime ; 
                    'Owner' = ( Get-Acl -Path $_.FullName | Select -ExpandProperty Owner ) ;
                    'Size (KB)' = [int]( $_.Length / 1KB ) }
                
                if( $remove )
                {
                    Remove-Item -Path $_.FullName -Stream $zoneStream
                }
                elseif( $scrub )
                {
                    $lastModified = $_.LastWriteTime
                    $zoneInfo | Select-String -NotMatch 'URL=' | Set-Content -Path $_.FullName -Stream $zoneStream
                    $_.LastWriteTime = $lastModified
                }
            }
        }
    }
    catch {}
}
