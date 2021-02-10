<#
    Convert file/folder names to Title Case. Designed to run from Explorer right click SendTo menu

    Set shortcut target in shell:sendto explorer folder to this and set to run minimised so as not to see the PowerShell window

        C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -file "C:\Scripts\Send to Title Case.ps1"

    Uses MoveFile() API because Rename-Item and .NET error renaming folder when just case is different https://docs.microsoft.com/en-gb/windows/win32/api/winbase/nf-winbase-movefile

    @guyrleech 10/02/21
#>

[CmdletBinding()]

## do not call with -verbose, set $VerbosePreference to 'Continue' since we do not have named arguments as there is no Param block
[string]$dummyToAllowCmdletBindingWithNoParamBlock = $null

if( $args -and $args.Count )
{
    Add-Type -ErrorAction Stop -TypeDefinition @'
    
        using System;
        using System.Runtime.InteropServices;

        public static class Kernel32
        {
            [DllImport("kernel32", SetLastError = true)]
            public static extern bool MoveFile(string lpExistingFileName, string lpNewFileName);
        }
'@
    $type = $null
    $culture = Get-Culture
    [bool]$result = $false

    ForEach( $item in $args )
    {
        [string]$relativeItem = Split-Path -Path $item -Leaf
        ## if all caps then ToTitleCase doesn't lower any characters
        [string]$titleCase = $culture.TextInfo.ToTitleCase( $culture.TextInfo.ToLower( $relativeItem ) )
        if( $titleCase -cne $relativeItem )
        {
            [string]$folder = Split-Path -Path $item -Parent
            Write-Verbose -Message "Moving `"$relativeItem`" to `"$titleCase`""
            $result = [kernel32]::MoveFile( $item , (Join-Path -Path $folder -ChildPath $titleCase) )
            $lastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
            if( ! $result )
            {
                if( ! $type )
                {
                    $type = Add-Type -AssemblyName System.Windows.Forms -PassThru
                }
                
                [string]$errorText = "Failed to rename `"$item`" to `"$titleCase`"`n`n$lastError"
                Write-Error -Message $errorText 
                [void][Windows.MessageBox]::Show( $errorText , 'Rename Error' , 'OK' , 'Error' )
            }
        }
        else
        {
            Write-Verbose -Message "`"$titleCase`" same as `"$relativeItem`""
        }
    }
}
else
{
    Write-Error -Message "No file/folder paths passed"
}
