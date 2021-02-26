
<#
.SYNOPSIS

    Output the long or short 8.3 NTFS names for the given path(s)

.DESCRIPTION

    Can get issues where %temp% is an NTFS short name, eg. when there are accented characters in a username, but other utilities/cmdlets return the long name so cannot check if the same with a simple string comparison

.PARAMETER path

    The path(s) to convert to long or short name depending on if -toShort is specified

.PARAMETER toShort

    Convert the path(s) specified to NTFS 8.3 short names. If not specified, expects -path to specify short names and will output their long names

.EXAMPLE

    & 'C:\Guys Scripts\Scripts\8.3 name converter.ps1' -path c:\progra~1,c:\progra~2 

    Convert the two given 8.3 NTFS short names to their long names

.EXAMPLE

    & 'C:\Guys Scripts\Scripts\8.3 name converter.ps1' -path 'C:\Program Files','C:\Program Files (x86)' -toShort

    Convert the two given long names to NTFS short names

.NOTES

    Uses Microsoft APIs GetLongPathName and GetShortPathName
#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$true,HelpMessage='8.3 shortname to expand')]
    [string[]]$path ,
    [switch]$toShort
)

## http://csharparticles.blogspot.com/2005/07/long-and-short-file-name-conversion-in.html

Add-Type -ErrorAction Stop -TypeDefinition @'
    using System;
    using System.Text;
    using System.Runtime.InteropServices;

    public static class kernel32
    {
        // https://docs.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-getlongpathnamea

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
            public static extern int GetLongPathName(
                [MarshalAs(UnmanagedType.LPTStr)]
                string path,
                [MarshalAs(UnmanagedType.LPTStr)]
                System.Text.StringBuilder longPath,
                int longPathLength
                );

        // https://docs.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-getshortpathnamew
        
        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
            public static extern int GetShortPathName(
                [MarshalAs(UnmanagedType.LPTStr)]
                string path,
                [MarshalAs(UnmanagedType.LPTStr)]
                System.Text.StringBuilder shortPath,
                int shortPathLength
                );
    }
'@

[Text.Stringbuilder]$newname = New-Object -TypeName System.Text.Stringbuilder -ArgumentList 260 ## MAX_PATH
[string]$otherName = $null
[int]$returned = -1

ForEach( $filename in $path )
{
    if( $toShort )
    {
        $returned = [kernel32]::GetShortPathName( $filename , $newname , $newname.Capacity )
    }
    else
    {
        $returned =  [kernel32]::GetLongPathName( $filename , $newname , $newname.Capacity )
    }

    if( $returned -gt 0 )
    {
        $newname.ToString()
        [void]$newname.Clear()
    }
    else
    {
        Write-Warning -Message "Failed to convert `"$filename`""
    }
}