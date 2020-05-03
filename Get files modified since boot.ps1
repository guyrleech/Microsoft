#requires -version 3

<#
.SYNOPSIS

 Find files modified since boot, or another date/time, but excluding sparse files which may be App-V cache files and junction points

.DESCRIPTION

Useful for determining what files are using Citrix Provisioning Services write cache although a direct correlation between cache used
and sizes of files written cannot be calculated by this script since the write cache works at the file system block level so the file
size reported by the script is not necessariluy how much data has been updated and is thus in the write cache.

.PARAMETER folder

Folder to search. Defaults to the root of the system drive.

.PARAMETER boot

The time to search from. If not specified the script will retrieve the last boot time

.PARAMETER until

The time to search until. If not specified will search up to the present time.
Cannot be used with -duration

.PARAMETER duration

The length of time to look for changed files
Cannot be used with -until

.PARAMETER csv

The path to a non-existent csv file will be written to 

.PARAMETER createdOnly

Only show files which have been created in the time window, not modified

.PARAMETER nogridview

Do not output the results to an on screen sortable/filterable grid view. Output to the pipeline instead

.PARAMETER hashAlgorithm

The hash algorithm to use for taking a checksum of each file allow duplicate files to be flagged.
If not specified, no checksums will be calculated which will be faster

.EXAMPLE

& '.\Get files modified since boot.ps1'

Find all files modified or created since the last boot time and display in a grid view

.EXAMPLE

& '.\Get files modified since boot.ps1' -createdOnly -boot ((Get-Date).AddHours( -2 )) -csv c:\recentlycreatedfiles.csv -hashAlgorithm MD5

Find all files created in the last two hours, writing the results to the csv file specified and display in a grid view including MD5 checksums

.NOTES

To determine the amount of write cache used, use poolmon.exe - https://twitter.com/guyrleech/status/1255450263596474368?s=20

#>

[CmdletBinding()]

Param
(
    [string]$folder = (Join-Path -Path $env:SystemDrive -ChildPath '\' ) ,
    [datetime]$boot = ( Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object -ExpandProperty LastBootUpTime) ,
    [datetime]$until ,
    [string]$duration ,
    [string]$csv ,
    [switch]$createdOnly ,
    [switch]$nogridview ,
    [ValidateSet( 'ACTripleDES' , 'MD5' , 'RIPEMD160' , ' SHA1' , 'SHA256' , 'SHA384' , 'SHA512' )]
    [string]$hashAlgorithm
)

$properties =  [System.Collections.Generic.List[psobject]]@( 'DirectoryName','Name','Length','LastWriteTime','CreationTime'  )

if( $PSBoundParameters[ 'hashAlgorithm' ] )
{
    $properties.Add( (@{n="$hashAlgorithm Hash";'e'={Get-FileHash -Path $_.FullName -Algorithm $hashAlgorithm -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Hash}} ) )
}

if( $PSBoundParameters[ 'until'] )
{
    if( $PSBoundParameters[ 'duration' ] )
    {
        Throw 'Must not use both -until and -duration'
    }
    $end = $until
}
else
{
    $end = $null
}

if( $PSBoundParameters[ 'duration' ] )
{
    ## see what last character is as will tell us what units to work with
    [int]$multiplier = 0
    switch( $duration[-1] )
    {
        's' { $multiplier = 1 }
        'm' { $multiplier = 60 }
        'h' { $multiplier = 3600 }
        'd' { $multiplier = 86400 }
        'w' { $multiplier = 86400 * 7 }
        'y' { $multiplier = 86400 * 365 }
        default { Throw "Unknown multiplier `"$($duration[-1])`"" }
    }

    if( $duration.Length -le 1 )
    {
        $secondsAgo = $multiplier
    }
    else
    {
        $secondsAgo = ( ( $duration.Substring( 0 , $duration.Length - 1 ) -as [decimal] ) * $multiplier )
    }
    
    $end = $boot.AddSeconds( $secondsAgo )
}

[string]$datestamp = $(if( $createdOnly ) { 'CreationTime' } else { 'LastWriteTime' } )

## Do it this way so we can exclude junction point folders and sym linked files
$items = Get-Item $folder -Force
$results = New-Object -TypeName System.Collections.Generic.List[psobject]

While( $items )
{
    $newitems = @( $items | Get-ChildItem -Attributes !ReparsePoint+!SparseFile -ErrorAction SilentlyContinue -Force )
    $results += $newitems.Where( { $_.$datestamp -gt $boot -and ( ! $end -or $_.$datestamp -le $end ) -and ! $_.LinkType -and ! $_.PSIsContainer -and $_.Length -gt 0 } )
    $items = $newitems | Where-Object -Property Attributes -Like *Directory* 
}

if( $results.Count -gt 0 )
{
    [long]$diskUsage =  $results | Measure-Object -Property Length -Sum | Select-Object -ExpandProperty Sum
    [string]$message = "Found $($results.Count) items $(if( $createdOnly ) { 'created' } else { 'modified' })"
    $message += $(if( $end )
    {
        " between $(Get-Date -Date $boot -Format G) and $(Get-Date -Date $end -Format G)"
    }
    else
    {
        " since $(Get-Date -Date $boot -Format G)"
    })
    $message += " in `"$folder`" consuming $([math]::Round( $diskUsage / 1GB , 2 ))GB"
    
    [int]$duplicateCount = 0
    [hashtable]$duplicates = @{}
    if( $PSBoundParameters[ 'hashAlgorithm' ] )
    {
        [string]$hashproperty = "$hashAlgorithm Hash"

        ForEach( $file in $results )
        {
            if( $file.$hashproperty )
            {
                Try
                {
                    $duplicates.Add( $file.$hashproperty , 0 )
                }
                Catch
                {
                    $duplicates.Set_Item( $file.$hashproperty , $duplicates[ $file.$hashproperty ] + 1 )
                }
            }
        }
        ForEach( $file in $results )
        {
            if( $file.$hashproperty -and ( [int]$value = $duplicates[ $file.$hashproperty ] ) -ge 0 )
            {
                Add-Member -InputObject $file -MemberType NoteProperty -Name 'Duplicates' -Value $value
                if( $value -gt 0 )
                {
                    $duplicateCount++
                }
            }
        }

        $message += " with $duplicateCount potentially duplicate files"
    }

    Write-Verbose -Message $message

    if( ! [string]::IsNullOrEmpty( $csv ) )
    {
        $results | Select-Object -Property $properties | Sort-Object -property Length -Descending| Export-Csv -NoTypeInformation -NoClobber -Path $csv 
    }
    
    if( $nogridview )
    {
        $results | Select-Object -Property $properties
    }
    else
    {
        $results | Select-Object -Property $properties | Out-GridView -Title $message -PassThru
    }
}
