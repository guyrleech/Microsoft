#requires -version 3

<#
.SYNOPSIS
    Read Pending File Rename operations to see what the files are and if they are in use and optionally perform the deletes/replacements

.DESCRIPTION
    Use of this script is entirely at your own risk

.PARAMETER processEntries
    Try and perform the operations and remove the entries from the registry value if successful.

.EXAMPLE
    & '.\Process Pending File Rename Operations.ps1' | ogv

    Show details of the current queued pending file rename operations and present in a grid view
    
.EXAMPLE
    & '.\Process Pending File Rename Operations.ps1' -processEntries

    Show details of the current queued pending file rename operations and prompt to perform each one

.NOTES
    Modification History:

    @guyrleech  2021/05/11  Script born
    @guyrleech  2024/04/24  Warning if -ShowUsage used as not yet implemented. Only take file copy if not a directory. Fixed issue with internal copy of file with -whatif.
                            Removed "usage" output properties. Check if number of entries in value is odd and report/abort if so
#>

<#
Copyright © 2024 Guy Leech

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, 
including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='High')]

Param
(
    [switch]$showUsage ,
    [switch]$processEntries ,
    [string]$findModuleScript = 'Find loaded modules.ps1' 
)

if( $PSBoundParameters[ 'showUsage' ] )
{
    Write-Warning -Message "Awfully sorry but -showUsage is not yet implemented"
}

$errorVariabubble = $null

[string]$registryKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager'
[string]$registryValue = 'PendingFileRenameOperations'

[string[]]$pendingFileRenameOperations = @( Get-ItemProperty -Path $registryKey -Name $registryValue -ErrorAction SilentlyContinue -ErrorVariable errorVariabubble | Select-Object -ExpandProperty $registryValue -ErrorAction SilentlyContinue)

if( $null -ne $errorVariabubble -and $errorVariabubble.Count )
{
    Write-Warning -Message "No pendingFileRenameOperations value found: $errorVariabubble"
    exit 1
}

if( $null -eq $pendingFileRenameOperations -or ! $pendingFileRenameOperations.Count )
{
    Write-Warning -Message "Nothing found in pendingFileRenameOperations value"
    exit 1
}

Write-Verbose -Message "Got $($pendingFileRenameOperations.Count) total entries"

if( $pendingFileRenameOperations.Count -band 0x1 )
{
    [string]$message = "Count of values in pendingFileRenameOperations is $($pendingFileRenameOperations.Count) which is odd but should be even" 
    if( $processEntries -or $pendingFileRenameOperations.Count -lt 2 )
    {
        Throw $message
    }
    else
    {
        Write-Warning -Message $message
    }
}

## get all loaded modules from all processes including the opposite bitness of this PowerShell process

[string]$scriptPath = Split-Path -Path (& { $myInvocation.ScriptName }) -Parent
[string]$findModuleScriptPath = $null
[hashtable]$modulesToProcesses = @{}

if( $showUsage )
{
    $findModuleScriptPath = [System.IO.Path]::Combine( $scriptPath , $findModuleScript )
    if( ! ( Test-Path -Path $findModuleScriptPath -PathType Leaf ) )
    {
        Throw "Unable to find `"$findModuleScript`" which is needed when -showUsage is used"
    }
    & $findModuleScriptPath -moduleName . -regex | . { Process `
    {
        $module = $_
        $thisProcess = "$($module.Process) ($($module.PID))"
        if( $existing = $modulesToProcesses[ $module.'Module Path' ] )
        {
            $existing.Add( $thisProcess )
        }
        else
        {
            $modulesToProcesses.Add( $module.'Module Path' , [System.Collections.Generic.List[string]]@( $thisProcess ) )
        }
    }}
}

[int]$counter = 0

## build up new value so we can remove entries where the source file doesn't exist or we've been able to do the update
$newValue = New-Object -TypeName System.Collections.Generic.List[string]

## TODO do we need to do anything about 8.3 names?

## These are in pairs where first entry is the file to operate on and the second entry is either the file to overwrite or the entry is to be deleted if empty
For( [int]$index = 0 ; $index -lt $pendingFileRenameOperations.Count ; $index += 2 )
{
    $counter++
    [bool]$deletion = [string]::IsNullOrEmpty( $pendingFileRenameOperations[$index + 1] )

    Write-Verbose -Message "$counter / $($pendingFileRenameOperations.Count / 2) : $($pendingFileRenameOperations[$index]) -> $(if( $deletion ) { 'DELETE' } else { $pendingFileRenameOperations[$index + 1] })"
    [string]$sourceFile = $pendingFileRenameOperations[$index] -replace '^\\\?\?\\'
    $sourceProperties = $null
    if( [string]::IsNullOrEmpty( $sourceFile ) )
    {
        Write-Warning -Message "File name at index $index is unexpectedly blank"
    }
    else
    {
        $sourceProperties = Get-ItemProperty -Path $sourceFile -ErrorAction SilentlyContinue
    }
    ## getting version info doesn't work if has resources but it is not exe/dll extension so make a renamed copy with that extension
    [string]$copyOfSourceFile = $null
    if( $null -ne $sourceProperties -and $sourceFile -notmatch '\.(exe|dll|ocx|cpl|scr|mui)$' -and $sourceProperties.Attributes -notmatch '\bdirectory\b')
    {
        $copyOfSourceFile = [System.IO.Path]::GetTempFileName() + '.dll'
        try
        {
            Copy-Item -Path $sourceFile -Destination $copyOfSourceFile -WhatIf:$false -Confirm:$false ## internal copy so hide from the -whatif -confirm
        }
        catch
        {
            Write-Warning -Message  "Failed to copy $sourceFile to $copyOfSourceFile : $_"
            $copyOfSourceFile = $null
        }
    }

    $sourceVersionInfo = $null
    $sourceDetails = $null
    $sourceUsage = $null
    try
    {
        if( $sourceDetails = Get-ItemProperty -Path $(if( $copyOfSourceFile ) { $copyOfSourceFile } else { $sourceFile } ) -ErrorAction SilentlyContinue )
        {
            $sourceVersionInfo = $sourceDetails | Select -ExpandProperty VersionInfo -ErrorAction SilentlyContinue
            if( $showUsage )
            {
                $sourceUsage = $modulesToProcesses[ $sourceFile ] -join ','
            }
        }
    }
    catch
    {
    }

    ## get file details of the files
    ## if not a deletion then get details of the destination file
    $destinationDetails = $null
    $destinationVersionInfo = $null
    $destinationUsage = $null

    [string]$destinationFile = $pendingFileRenameOperations[ $index + 1 ] -replace '^!?\\\?\?\\'
    if( -Not $deletion -and ( $destinationDetails = Get-ItemProperty -Path $destinationFile -ErrorAction SilentlyContinue -ErrorVariable errorVariabubble ) )
    {
        $destinationVersionInfo = $destinationDetails | Select -ExpandProperty VersionInfo -ErrorAction SilentlyContinue
        ## TODO if doesn't exist then we must report this
    }
    if( $destinationDetails -and $showUsage )
    {
        $destinationUsage = $modulesToProcesses[ $destinationFile ] -join ','
    }
    ## TODO report if any destinations are in use by searching loaded modules

    $result = [PSCustomObject]@{
        'Source' = $sourceFile
        'SourceExtension' = [System.IO.Path]::GetExtension( $pendingFileRenameOperations[ $index ] )
 ##       'SourceUsage' = $sourceUsage
        'Destination' = $destinationFile
        'DestinationExtension' = $(if( ! [string]::IsNullOrEmpty( $pendingFileRenameOperations[ $index + 1 ] ) ) { [System.IO.Path]::GetExtension( $pendingFileRenameOperations[ $index + 1 ] ) })
 ##       'DestinationUsage' = $destinationUsage
        'Destination Size (KB)' = $(if( $destinationDetails ) { [math]::round( $destinationDetails.Length / 1KB , 1 ) })
        'Deletion' = $deletion
        'Deleted' = $false
        'Moved' = $false
        'SourceExists' = $null -ne $sourceProperties
        'SourceAttributes' = $sourceDetails | Select-Object -ExpandProperty Attributes
        'Source Size (KB)' = $(if( $sourceProperties -and $sourceProperties.Attributes -notmatch '\bdirectory\b' ) { [math]::round( $sourceProperties.Length / 1KB , 1 ) })
        'SourceFileVersion' = $sourceVersionInfo | Select-Object -ExpandProperty FileVersion
        'SourceFileVersionRaw' = $sourceVersionInfo | Select-Object -ExpandProperty FileVersionRaw
        'DestinationFileVersion' = $destinationVersionInfo | Select-Object -ExpandProperty FileVersion
        'DestinationFileVersionRaw' = $destinationVersionInfo | Select-Object -ExpandProperty FileVersionRaw
        'ProductName' = $sourceVersionInfo | Select-Object -ExpandProperty ProductName
        'CompanyName' = $sourceVersionInfo | Select-Object -ExpandProperty CompanyName
        'FileDescription' = $sourceVersionInfo | Select-Object -ExpandProperty FileDescription
    }

    if( $showUsage -and $destinationDetails )
    {
        ## TODO find the destination file in loaded modules of all processes
    }

    if( $processEntries )
    {
        ## if source file doesn't exist then remove the entry
        [bool]$addEntry = $null -ne $sourceProperties

        if( $addEntry ) ## only do it if source still exists
        {
            ## see if we can do the update and if so remove the entry
            if( $deletion )
            {
                if( $PSCmdlet.ShouldProcess( $sourceFile , 'Delete' ) )
                {
                    $errorVariabubble = $null
                    Write-Verbose -Message "Deleting $sourceFile"
                    Remove-Item -Path $sourceFile -Force -Recurse -ErrorVariable errorVariabubble
                    $addEntry =  ! ( $result.Deleted = ($? -or $null -eq $errorVariabubble -or $errorVariabubble.Count -eq 0) )
                }
            }
            else ## a move
            {
                if( $PSCmdlet.ShouldProcess( "`"$sourceFile`" to `"$destinationFile`"" , 'Move' ) )
                {
                    $errorVariabubble = $null
                    Write-Verbose -Message "Moving `"$sourceFile`" to `"$destinationFile`""
                    ## rename destination in case move fails
                    [string]$tempFile = $destinationFile + ".$pid.$((Get-Date).Ticks)"
                    if( Move-Item -Path $destinationFile -Destination $tempFile -Force -PassThru -ErrorVariable errorVariabubble )
                    {
                        if( ! ( Move-Item -Path $sourceFile -Destination $destinationFile -Force -ErrorVariable errorVariabubble -PassThru ) )
                        {
                            if( ! ( Move-Item -Destination $destinationFile -Path $tempFile -Force -PassThru -ErrorVariable errorVariabubble ))
                            {
                                Write-Error -Message "Failed to move `"$tempFile`" back to `"$destinationFile`""
                            }
                        }
                        else
                        {
                            Remove-Item -Path $tempFile -Force
                        }
                    }
                    else
                    {
                        Write-Error -Message "Failed to move `"$destinationFile`" to `"$tempFile`""
                    }
                    $addEntry =  ! ( $result.Moved = ($? -or $null -eq $errorVariabubble -or $errorVariabubble.Count -eq 0) )
                }
            }
        }
        else
        {
            Write-Verbose -Message "Removing entry for non-existent file $sourceFile"
        }

        if( $addEntry )
        {
            $newValue.Add( $pendingFileRenameOperations[ $index ] )
            $newValue.Add( $pendingFileRenameOperations[ $index + 1 ] )
        }
    }

    $result
    if( $copyOfSourceFile )
    {
        Remove-Item -Path $copyOfSourceFile -Force -ErrorAction SilentlyContinue -WhatIf:$false -Confirm:$False
        $copyOfSourceFile = $null
    }
}

if( $processEntries )
{
    if( $newValue -and $newValue.Count )
    {
        Set-ItemProperty -Path $registryKey -Name $registryValue -Value $newValue
    }
    else
    {
        ## remove value as nothing left to process
        Remove-ItemProperty -Path $registryKey -Name $registryValue
    }
}
