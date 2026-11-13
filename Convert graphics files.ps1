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

.NOTES
    Modification History:

    2019/11/13  @guyrleech  Script born
    2026/11/13  @guyrleech  Added output .ico support and -overwrite
#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory,HelpMessage="Directory containing files to be converted")]
    [string]$path ,
    [validateset( "bitmap","emf","exif","gif","icon","jpeg","png","tiff","wmf",'ico' )]
    [string]$destFormat = 'jpeg' ,
    [string]$destPath ,
    [validateset( "bitmap","emf","exif","gif","icon","jpeg","png","tiff","wmf" )]
    [string]$sourceFormat = 'bitmap' ,
    [ValidateSet(16, 24, 32, 48, 64, 128, 256)]
    [int[]]$iconSizes ,
    [switch]$recurse ,
    [switch]$overwrite
)

[hashtable]$formatToExtension = @{
    "bitmap" = '.bmp'
    "emf"    = '.emf'
    "exif"   = '.exif'
    "gif"    = '.gif'
    "icon"   = '.ico'
    "jpeg"   = '.jpg'
    "png"    = '.png'
    "tiff"   = '.tif'
    "wmf"    = '.wmf'
    "ico"    = '.ico'
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

function ConvertTo-Ico
{
    param
    (
        [Parameter(Mandatory)][string]$Path,
        [string]$OutFile = [IO.Path]::ChangeExtension($Path, '.ico'),
        [int[]]$Sizes = @(16, 24, 32, 48, 64, 128, 256)
    )
    Add-Type -AssemblyName System.Drawing
    $src = [System.Drawing.Image]::FromFile((Resolve-Path $Path))
    try {
        $images = foreach ($s in $Sizes) {
            $bmp = New-Object System.Drawing.Bitmap $s, $s
            $g = [System.Drawing.Graphics]::FromImage($bmp)
            $g.InterpolationMode = 'HighQualityBicubic'
            $g.SmoothingMode     = 'HighQuality'
            $g.PixelOffsetMode   = 'HighQuality'
            $g.DrawImage($src, 0, 0, $s, $s)
            $g.Dispose()
            $ms = New-Object IO.MemoryStream
            $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
            $bmp.Dispose()
            [pscustomobject]@{ Size = $s; Bytes = $ms.ToArray() }
        }
    } finally { $src.Dispose() }

    $fs = [IO.File]::Create($OutFile)
    $w  = New-Object IO.BinaryWriter $fs
    try {
        # ICONDIR header: reserved, type 1 = icon, image count
        $w.Write([uint16]0); $w.Write([uint16]1); $w.Write([uint16]$images.Count)
        $offset = 6 + 16 * $images.Count
        foreach ($i in $images) {
            # ICONDIRENTRY: width/height (0 = 256), colours, reserved, planes, bpp, size, offset
            $dim = if ($i.Size -ge 256) { 0 } else { $i.Size }
            $w.Write([byte]$dim); $w.Write([byte]$dim)
            $w.Write([byte]0);    $w.Write([byte]0)
            $w.Write([uint16]1);  $w.Write([uint16]32)
            $w.Write([uint32]$i.Bytes.Length); $w.Write([uint32]$offset)
            $offset += $i.Bytes.Length
        }
        foreach ($i in $images) { $w.Write($i.Bytes) }
    } finally { $w.Dispose() }
    Get-Item $OutFile
}
[void][Reflection.Assembly]::LoadWithPartialName( 'System.Windows.Forms' )

[hashtable]$icoSizeParam = @{}
if( $PSBoundParameters[ 'iconSizes' ] )
{
    $icoSizeParam.Add( 'iconSizes' , $iconSizes )
}

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

    if( ( Test-Path -Path $newFile -ErrorAction SilentlyContinue ) -and -Not $overwrite )
    {
        Write-Warning "`"$newFile`" already exists"
    }
    elseif( $destFormat -ieq 'ico' )
    {
        ConvertTo-Ico -Path $_.FullName -OutFile $newFile @iconSizes
    }
    else
    {
        $file = New-Object -TypeName System.Drawing.Bitmap( $_.FullName )
        $file.Save( $newFile , $destFormat )
        $file.Dispose()
    }
}
