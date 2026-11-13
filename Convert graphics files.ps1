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
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$OutFile,
        [ValidateNotNullOrEmpty()]
        [ValidateRange(1, 256)]
        [int[]]$Sizes = @(16, 24, 32, 48, 64, 128, 256)
    )
    Add-Type -AssemblyName System.Drawing

    # Resolve against the current PowerShell location (.NET would use the process folder instead)
    $resolvedSourcePath = (Resolve-Path -LiteralPath $Path).ProviderPath
    if (-not $OutFile)
    {
        $OutFile = [System.IO.Path]::ChangeExtension($resolvedSourcePath, '.ico')
    }
    $resolvedOutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutFile)

    $sourceImage = [System.Drawing.Image]::FromFile($resolvedSourcePath)
    try
    {
        # @( ) keeps this an array even when only one size is requested,
        # otherwise .Count fails under Set-StrictMode
        $iconImages = @(foreach ($iconSize in $Sizes)
        {
            $resizedBitmap = New-Object -TypeName System.Drawing.Bitmap -ArgumentList $iconSize, $iconSize
            $graphics = [System.Drawing.Graphics]::FromImage($resizedBitmap)
            $graphics.InterpolationMode = 'HighQualityBicubic'
            $graphics.SmoothingMode     = 'HighQuality'
            $graphics.PixelOffsetMode   = 'HighQuality'
            $graphics.DrawImage($sourceImage, 0, 0, $iconSize, $iconSize)
            $graphics.Dispose()

            $pngStream = New-Object -TypeName System.IO.MemoryStream
            $resizedBitmap.Save($pngStream, [System.Drawing.Imaging.ImageFormat]::Png)
            $resizedBitmap.Dispose()

            [pscustomobject]@{
                Size     = $iconSize
                PngBytes = $pngStream.ToArray()
            }
        })
    }
    finally
    {
        $sourceImage.Dispose()
    }

    $outputStream = [System.IO.File]::Create($resolvedOutputPath)
    $binaryWriter = New-Object -TypeName System.IO.BinaryWriter -ArgumentList $outputStream
    try
    {
        # ICONDIR header: reserved, type (1 = icon), number of images
        $binaryWriter.Write([uint16]0)
        $binaryWriter.Write([uint16]1)
        $binaryWriter.Write([uint16]$iconImages.Count)

        $headerSize = 6
        $directoryEntrySize = 16
        $imageDataOffset = $headerSize + $directoryEntrySize * $iconImages.Count

        foreach ($iconImage in $iconImages)
        {
            # ICONDIRENTRY: width/height (0 means 256), colour count, reserved,
            # colour planes, bits per pixel, image data size, image data offset
            if ($iconImage.Size -ge 256)
            {
                $dimensionByte = 0
            }
            else
            {
                $dimensionByte = $iconImage.Size
            }
            $binaryWriter.Write([byte]$dimensionByte)
            $binaryWriter.Write([byte]$dimensionByte)
            $binaryWriter.Write([byte]0)
            $binaryWriter.Write([byte]0)
            $binaryWriter.Write([uint16]1)
            $binaryWriter.Write([uint16]32)
            $binaryWriter.Write([uint32]$iconImage.PngBytes.Length)
            $binaryWriter.Write([uint32]$imageDataOffset)
            $imageDataOffset += $iconImage.PngBytes.Length
        }

        foreach ($iconImage in $iconImages)
        {
            $binaryWriter.Write($iconImage.PngBytes)
        }
    }
    finally
    {
        $binaryWriter.Dispose()
    }

    Get-Item -LiteralPath $resolvedOutputPath
}

[void][Reflection.Assembly]::LoadWithPartialName( 'System.Windows.Forms' )

[hashtable]$icoSizeParam = @{}
if( $PSBoundParameters[ 'iconSizes' ] )
{
    $icoSizeParam.Add( 'Sizes' , $iconSizes )
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
        ConvertTo-Ico -Path $_.FullName -OutFile $newFile @icoSizeParam
    }
    else
    {
        $file = New-Object -TypeName System.Drawing.Bitmap( $_.FullName )
        $file.Save( $newFile , $destFormat )
        $file.Dispose()
    }
}
