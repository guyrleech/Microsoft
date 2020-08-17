<#
.SYNOPSIS

Find shortcuts with target or icon path or arguments matching a given regular expression and change to a new string

.PARAMETER old

Regular expression to match in target or icon paths or arguments in shortcuts

.PARAMETER new

Text to replace regular expression in shortcut properties

.PARAMETER path

The path to the shortcuts

.PARAMETER name

Name or pattern of shortcuts to operate on

.PARAMETER force

Change the shortcut properties even if the changed folder does not exist

.PARAMETER recurse

Search sub folders of the path specified

.EXAMPLE

& '.\Fix shortcuts.ps1' -path "C:\Users\guy\AppData\Roaming\Microsoft\Windows\SendTo" -old '\\OneDrive\\' -new '\OneDrive - contoso.com\'

Find shortcuts in the folder specified and any that contain '\OneDrive\' in the target path, arguments or icon path, will be changed to '\OneDrive - contoso.com\' if that path exists
for the target or icon paths.

.EXAMPLE

& '.\Fix shortcuts.ps1' -path "C:\Scripts" -old '\\OneDrive\\' -new '\OneDrive - contoso.com\' -force -confirm:$false -name '*fred*.lnk'

Find shortcuts in the folder specified, and sub folders, named '*fred*.lnk' and any of these that contain '\OneDrive\' in the target path, arguments or icon path, will be changed to 
'\OneDrive - contoso.com\' even if that path does not exist for the target or icon paths.

#>

[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='High')]

Param
(
    [Parameter(Mandatory,HelpMessage='regex to search for')]
    [string]$old ,
    [Parameter(Mandatory,HelpMessage='Text to replace matched regex with')]
    [string]$new ,
    [string]$path = '.' ,
    [string]$name = '*.lnk' ,
    [switch]$force ,
    [switch]$recurse
)

$shellObject = New-Object -ComObject Wscript.Shell

[int]$found = 0
[int]$fixed = 0

Get-ChildItem -Path $path -Filter $name -Recurse:$recurse | ForEach-Object `
{
    $found++
    if( $shortcut = $shellObject.CreateShortcut( $_.FullName ) )
    {
        Write-Verbose -Message "$found : `"$($_.FullName)`""

        [string]$changedTargetPath    = $shortcut.TargetPath   -replace $old , $new
        [string]$changedArguments     = $shortcut.Arguments    -replace $old , $new
        [string]$changedIconLocation  = $shortcut.IconLocation -replace $old , $new
        [int]$changes = 0

        if( $changedTargetPath -ne $shortcut.TargetPath )
        {
            ## only change if -force or source doesn't exist and target does exist
            if( $force -or ( ! ( Test-Path -Path $shortcut.TargetPath -ErrorAction SilentlyContinue ) -and ( Test-Path -Path $changedTargetPath -ErrorAction SilentlyContinue ) ) )
            {
                Write-Verbose -Message "`"$($_.FullName)`" : target `"$($shortcut.TargetPath)`" changed to `"$changedTargetPath`""
                $shortcut.TargetPath = $changedTargetPath
                $changes++
            }
            else
            {
                Write-Verbose -Message "`"$($_.FullName)`" : target `"$($shortcut.TargetPath)`" NOT changed to `"$changedTargetPath`""
            }
        }
        if( $changedArguments -ne $shortcut.Arguments )
        {
            ## difficult to parse arguments to check if paths exist - they may not and it may be correct (eg argument is a file/path that is created)
            Write-Verbose -Message "`"$($_.FullName)`" : arguments `"$($shortcut.Arguments)`" changed to `"$changedArguments`""
            $shortcut.Arguments = $changedArguments
            $changes++
        }
        if( $changedIconLocation -ne $shortcut.IconLocation )
        {
            if( $force -or ( ! ( Test-Path -Path ($shortcut.IconLocation -split ',')[0] -ErrorAction SilentlyContinue ) -and ( Test-Path -Path ($changedIconLocation -split ',')[0] -ErrorAction SilentlyContinue ) ) )
            {
                Write-Verbose -Message "`"$($_.FullName)`" : icon location `"$($shortcut.IconLocation)`" changed to `"$changedIconLocation`""
                $shortcut.IconLocation = $changedIconLocation
                $changes++
            }
            else
            {
                Write-Verbose -Message "`"$($_.FullName)`" : icon location `"$($shortcut.IconLocation)`" NOT changed to `"$changedIconLocation`""
            }
        }
        if( $changes -gt 0 )
        {
            if( $PSCmdlet.ShouldProcess( $_.FullName , 'Fix' ) )
            {
                $fixed++
                $shortcut.Save()
            }
        }
    }
    else
    {
        Write-Warning -Message "Failed to read shortcut `"$($_.FullName)`""
    }
}

Write-Verbose -Message "Changed $fixed out of $found shortcuts examined"
