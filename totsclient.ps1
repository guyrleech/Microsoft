<#
.SYNOPSIS
    Take an absolute local path and change to client mapped drive path

.DESCRIPTION
    Designed to be placed in $profile

.PARAMETER path
    The absolute local path to process. If not provided the clipboard contents will be used (eg copy from Visual Code)

.PARAMETER noClipboard
    Do not place result on clipboard instad putting on stdout

.PARAMETER prefix
    Prefix to use. Eg \\tsclient for RDP connections or \\client for Citrix

.PARAMETER suffix
    Suffix to use. Eg $ for Citrix client mapped drives (must escape with ` or put in single quotes '$'

.NOTES
    Modification History:

    2024/08/20  @guyrleech  Script born
#>

Function totsclient( [string]$path , [switch]$noClipboard , [string]$prefix = 'tsclient' , [string]$suffix )
{
    if( -Not $PSBoundParameters[ 'path' ] )
    {
        $path = Get-Clipboard -Format Text
    }
    if( [string]::IsNullOrEmpty( $path ) )
    {
        Write-Error "No path passed or in clipboard"
    }
    else
    {
        ## deal with absolute path which may or may not be quoted at start with "
	    [string]$tsclientPath = $path.Trim() -replace '^(?<quote>"?)(?<drive>\w):' , "`${quote}\\$prefix\`${drive}$suffix"
        if( $tsclientPath -ieq $path )
        {
            Write-Warning -Message "No change made to path $path"
        }
        elseif( $noClipboard )
        {
            $tsclientPath
        }
        else
        {
            ## quote it if contains spaces and not already quoted
            if( $tsclientPath -match '^[^"].*\s+' )
            {
                "`"$tsclientPath`"" | Set-Clipboard
            }
            else ## no need for quoting
            {
                $tsclientPath | Set-Clipboard
            }
        }
    }
}
