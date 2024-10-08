<#
.SYNOPSIS
    Take file name argument and put into clipboard as file

.DESCRIPTION
    So can copy path in VS Code, run this script then paste file in Explorer or a path in the clipboard from using "Copy as Path" in Explorer

.NOTES
    Modification History:

    2024/10/07 @guyrleech  Script born (original code from ChatGPT)
#>

[CmdletBinding()]

Param
(
)

[string]$message = $null

Add-Type -AssemblyName PresentationCore , PresentationFramework

if( [System.Windows.Clipboard]::ContainsText() )
{
    $files = New-Object -TypeName System.Collections.Specialized.StringCollection

    [string]$filePath = $null
    $filePath = Get-Clipboard

    if( -Not [string]::IsNullOrEmpty( $filePath ))
    {
        ## maybe quoted via copy as path if has spaces
        $filePath = $filePath.Trim( ' "''' )

        if( Test-Path -Path $filePath )
        {
            $null = $files.Add( $filePath )

            $dataObject = New-Object -TypeName System.Windows.DataObject
            $dataObject.SetFileDropList( $files )

            Write-Verbose -Message "Putting `"$filePath`" in file drop list"
            # Set the DataObject to the clipboard
            [System.Windows.Clipboard]::SetDataObject( $dataObject , $true )
        }
        else
        {
            $message = "`"$filePath`" - file not found"
        }
    }
    else
    {
        $message = "No text on clipboard"
    }
}
else
{
    $message = "Data on clipboard is not text"
}

if( -Not [string]::IsNullOrEmpty( $message ) )
{
    [System.Windows.MessageBox]::Show( $message , (( Split-Path -Path (& { $MyInvocation.ScriptName }) -Leaf ) -replace '\.\w+$' ), [System.Windows.MessageBoxButton]::OK , [System.Windows.MessageBoxImage]::Error )
    Throw $message
}
