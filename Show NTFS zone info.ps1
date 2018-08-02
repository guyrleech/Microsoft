#requires -version 3.0
<#
    Find files with Zone.Identifier NTFS alternate data stream and show the content

    This is how files are marked by Windows when they have come from an "untrusted" source

    Pipe to Out-GridView or Export-Csv

    @guyrleech 2018
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

.EXAMPLE

& '.\Show NTFS zone info.ps1' -path C:\Temp

Show all files in the c:\temp folder which have Zone.Identifier information attached

& '.\Show NTFS zone info.ps1' -path C:\users -Recurse -WebOnly | Out-GridView

Show all files in the c:\users folder and sub folders which have Zone.Identifier information attached which have come from the web

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
    [string]$zoneStream = 'Zone.Identifier'
)

[hashtable]$params = @{ 'Path' = $path ; 'Recurse' = $recurse }

## check path is on NTFS
[string]$drive = (Get-Item -Path $path -ErrorAction Stop).PSdrive

[string]$fileSystem = ( Get-CimInstance -ClassName win32_logicaldisk -Filter "DeviceId = '$($drive):'" ).FileSystem

if( $fileSystem -ne 'NTFS' )
{
    Write-Error "File system type of `"$path`" is $fileSystem, not NTFS, so there will be no alternate data streams"
    Exit 1
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
        $zoneInfo = Get-Content $_.FullName -Stream $zoneStream -ErrorAction SilentlyContinue 
        if( $zoneInfo )
        {
            $zone = $zoneInfo | Where-Object { $_ -match '^ZoneId=(.*)$' } | ForEach-Object { $matches[1] }
            $ReferrerURL = $zoneInfo | Where-Object { $_ -match '^ReferrerURL=(.*)$' } | ForEach-Object { $matches[1] }
            $HostURL = $zoneInfo | Where-Object { $_ -match '^HostURL=(.*)$' } | ForEach-Object { $matches[1] }
            if( ! $webOnly -or ( ! [string]::IsNullOrEmpty( $ReferrerURL ) -and $ReferrerURL -match '^https?:' ) -or ( ! [string]::IsNullOrEmpty( $HostURL ) -and $HostURL -match '^https?:' ) )
            {
                [pscustomobject]@{ 'File' = $_.FullName ; 'Zone' = $zones[ $zone ] ; 'Referrer URL' = $ReferrerURL ; 'Host URL' = $HostURL }
            }
        }
    }
    catch {}
}
