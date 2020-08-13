<#
.SYNOPSIS

Convert graphics files to other graphics formats

.DESCRIPTION

Some code from https://hazzy.techanarchy.net/posh/powershell/bmp-to-jpg-the-powershell-way/

.PARAMETER path

Directory containing the files to convert

.PARAMETER destFormat

The graphics format to convert the files to

.PARAMETER sourceFormat

The graphics format to convert the files from - will only act on files of this type found in the folder specified via -path

.PARAMETER destPath

Save destination file to this folder rather than the same folder as the source file

.PARAMETER recurse

Recurse the folder hierarchy specified via -path otherwise just operate on files in the specified folder

.EXAMPLE

& '.\Convert graphics files.ps1' -path "C:\Silly Screen Savers" -destFormat jpeg -sourceFormat bitmap

Find all bmp files in the folder "C:\Silly Screen Savers" and produce the correpsonding jpg file

#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$true,HelpMessage="Directory containing files to be converted")]
    [string]$path ,
    [validateset( “bitmap”,”emf”,”exif”,”gif”,”icon”,”jpeg”,”png”,”tiff”,”wmf” )]
    [string]$destFormat = 'jpeg' ,
    [string]$destPath ,
    [validateset( “bitmap”,”emf”,”exif”,”gif”,”icon”,”jpeg”,”png”,”tiff”,”wmf” )]
    [string]$sourceFormat = 'bitmap' ,
    [switch]$recurse
)

[hashtable]$formatToExtension = @{
    “bitmap” = '.bmp'
    ”emf” = '.emf'
    ”exif” = '.exif'
    ”gif” = '.gif'
    ”icon” = '.ico'
    ”jpeg” = '.jpg'
    ”png” = '.png'
    ”tiff” = '.tif'
    ”wmf”  = '.wmf'
}

if( $sourceFormat -eq $destFormat )
{
    Throw "Source and destination formats are the same"
}

[string]$oldExtension = $formatToExtension[ $sourceFormat ]
[string]$newExtension = $formatToExtension[ $destFormat ]

if( [string]::IsNullOrEmpty( $oldExtension ) )
{
    Throw "Unknown source format $sourceFormat"
}

if( [string]::IsNullOrEmpty( $newExtension ) )
{
    Throw "Unknown destination format $destFormat"
}

[void][Reflection.Assembly]::LoadWithPartialName( 'System.Windows.Forms' )

Get-ChildItem -Path $path -Recurse:$recurse -Force -Filter "*$oldExtension" | ForEach-Object `
{
    Write-Verbose "Converting `"$($_.FullName)`""

    [string]$newFile = $(if( $PSBoundParameters[ 'destPath' ] )
    {
        if( ! ( Test-Path -Path $destPath -PathType Container -ErrorAction SilentlyContinue) )
        {
            if( ! ( $newFolder = New-Item -Path $destPath -Force -ItemType Directory ) )
            {
                Throw "Failed to create destination path `"$destPath`""
            }
        }
        Join-Path -Path $destPath -ChildPath ( $_.BaseName + $newExtension )
    }
    else
    {
        Join-Path -Path $_.Directory.FullName -ChildPath ( $_.BaseName + $newExtension )
    })

    if( Test-Path -Path $newFile -ErrorAction SilentlyContinue ) 
    {
        Write-Warning "`"$newFile`" already exists"
    }
    else
    {
        $file = New-Object -TypeName System.Drawing.Bitmap( $_.FullName )
        $file.Save( $newFile , $destFormat )
        $file.Dispose()
    }
}
