<#
.SYNOPSIS
    Hash the given string with SHA256

.NOTES
    Modification History:

    2024/11/11  @guyrleech  Script born and raised
#>

[CmdletBinding()]

Param
(
    [string]$string = 'CORNED BEEF' ,
    [switch]$alternativeMethod 
)

if( $alternativeMethod )
{
    $sha256 = [System.Security.Cryptography.SHA256]::Create()

    [BitConverter]::ToString($sha256.ComputeHash([System.Text.Encoding]::UTF8.GetBytes('CORNED BEEF'))) -replace '-', ''
}
else
{
    $stringAsStream = [System.IO.MemoryStream]::new()

    $writer = [System.IO.StreamWriter]::new($stringAsStream)

    $writer.Write($string)

    $writer.Flush()

    $stringAsStream.Position = 0

    (Get-FileHash -InputStream $stringAsStream -Algorithm SHA256).Hash
}