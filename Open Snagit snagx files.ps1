#requires -version 3

<#
.SYNOPSIS
    Open a TechSmith Snagit .snagx file and open the graphics files contained in it

.DESCRIPTION
    Designed to be run via explorer Send To or Open With
    Unless you have changed the default FTA so that .ps1 files can be run directly, which is not recommended, put the following into a .cmd file and use that as the app to open the .snagx files via explorer

    @echo off

    REM Call powershell script that opens .snagx files by unzipping, finding the graphics files and opening them in the default viewer for that file type

    start "" C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noprofile -windowstyle minimized -file "C:\Scripts\Open Snagit snagx files.ps1" %1

.NOTES
    Debug output can be viewed in SysInternals dbgview utility

    Modification History

    2024/01/12  @guyrleech  Script born & delivered
#>

<#
Copyright © 2024 Guy Leech

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, 
including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

[string]$temporaryUnzipFolder = $null

try
{
    Add-Type -AssemblyName System.IO.Compression.FileSystem -Debug:$false -Verbose:$false

    ForEach( $arg in $args )
    {
        $properties = $null
        $properties = Get-ChildItem -Path $arg
        if( $null -ne $properties -and -not $properties.PSIsContainer -and $properties.Length -gt 42 )
        {
            $temporaryUnzipFolder = Join-Path -Path $env:temp -ChildPath "$($properties.Extension -replace '\.')_$($properties.BaseName)_$((New-GUID).Guid)"
            ## can't use Expand-Archive as errors on non-zip file in PS 5.1
            [System.IO.Compression.ZipFile]::ExtractToDirectory( $arg , $temporaryUnzipFolder )
            if( $? -and ( Test-Path -Path $temporaryUnzipFolder -PathType Container ) )
            {
                ## could look at index.json to get other files but just look for GUID files that aren't json
                [array]$graphicFiles = @( Get-ChildItem -Path $temporaryUnzipFolder -Filter '{*}.*' -File | Where-Object { $_.Extension -ine '.json' -and $_.Length -gt 42 -and $_.Name.Length -ge 42 } | Sort-Object -Property Length -Descending )
                [string]$thumbnailFile = Join-Path -Path $temporaryUnzipFolder -ChildPath 'thumbnail.png'
                ## for some images (cropped?) the thumbnail seems to hold the image
                $thumbNailProperties = Get-Item -Path $thumbnailFile -ErrorAction SilentlyContinue
                if( $thumbNailProperties -and $graphicFiles.Count -gt 0 -and $thumbNailProperties.Length -gt $graphicFiles[ 0 ].Length )
                {
                    [System.Diagnostics.Debug]::WriteLine( "Opening thumbnail graphic file `"$($thumbNailProperties.Fullname)`"" )
                    Start-Process -FilePath $thumbNailProperties.FullName -Verb Open
                }
                else
                {
                    ForEach( $file in $graphicFiles )
                    {
                        ## Debug output so tools like SysInternals dbgview can get it
                        [System.Diagnostics.Debug]::WriteLine( "Opening graphic file `"$($file.Fullname)`"" )
                        Start-Process -FilePath $file.FullName -Verb Open
                    }
                }
            }
        }
    }
}
catch
{
    [System.Diagnostics.Debug]::WriteLine( "Exception occurred: $_" )
    Throw
}
finally
{
    if( $null -ne $temporaryUnzipFolder -and ( Test-Path -Path $temporaryUnzipFolder -PathType Container ) )
    {
        ## process to open them may not have opened the file yet so give it some time to do so
        Start-Sleep -Seconds 15
        [System.Diagnostics.Debug]::WriteLine( "Removing folder $temporaryUnzipFolder" )
        Remove-Item -Path $temporaryUnzipFolder -Recurse -Force -Confirm:$false
    }
}
