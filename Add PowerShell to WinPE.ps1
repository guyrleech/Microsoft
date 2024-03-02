#requires -RunAsAdministrator

<#
.SYNOPSIS
    Add PowerShell to mounted Windows PE image

.DESCRIPTION
    A version of the Windows ADK with PE Add-ons must be present

.PARAMETER mountPath
    The root path of where the WinPE wim file is mounted which will be discovered if not specified

.PARAMETER architecture
    The architecture of the WinPE image

.PARAMETER language
    The language of the packages to add

.PARAMETER winPEoptionalComponents
    The relative folder in the ADK installation where the optional component folders are located

.PARAMETER noSecureStartup
    Do not include the BitLocker components

.PARAMETER useDism
    Use dism.exe rather than the PowerShell cmdlets

.PARAMETER winPESourcePath
    The root folder of the WIndows ADK installation. Will be discovered by looking at the WIMMOUNT device driver path if not specified
    
.PARAMETER abort
    Abort immediately if an error encountered rather than continuing

.EXAMPLE
    & '.\Add PowerShell to WinPE.ps1'

    Add x64/amd64 PowerShell , WMI and secure startup packages to the x64 WIM image already mounted
    
.EXAMPLE
    & '.\Add PowerShell to WinPE.ps1' -architecture arm64 -language en-gb

    Add ARM64 PowerShell , WMI and secure startup packages in Proper, UK, English to the ARM WIM image already mounted

.NOTES
    Modification History:

    2023/12/16  @guyrleech  Script born
    2024/03/02  @guyrleech  Help added
#>

<#
Copyright © 2024 Guy Leech

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, 
including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

[CmdletBinding()]

Param
(
	[String]$mountPath ,
    [ValidateSet('arm64','amd64')]
    [String]$architecture = 'amd64' ,
    [String]$language = 'en-us' ,
    [String]$winPEoptionalComponents = 'WinPE_OCs' ,
    [switch]$noSecureStartup ,
    [switch]$useDism ,
    [string]$winPESourcePath ,## = 'C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Windows Preinstallation Environment'
    [switch]$abort
)

[int]$counter = 0
[string[]]$packages = @(
    "WinPE-WMI.cab"
    "$language\WinPE-WMI_$language.cab"
    "WinPE-NetFX.cab"
    "$language\WinPE-NetFX_$language.cab"
    "WinPE-Scripting.cab"
    "$language\WinPE-Scripting_$language.cab"
    "WinPE-PowerShell.cab"
    "$language\WinPE-PowerShell_$language.cab"
    "WinPE-StorageWMI.cab"
    "$language\WinPE-StorageWMI_$language.cab"
    "WinPE-DismCmdlets.cab"
    "$language\WinPE-DismCmdlets_en-us.cab"
)

if( [string]::IsNullOrEmpty( $mountPath ) )
{
    Import-Module -Name Dism -Verbose:$false
    $mounted = $null
    $mounted = Get-WindowsImage -Mounted
    if( $null -eq $mounted )
    {
        Throw "No mounted images found"
    }
    if( $mounted -is [array] -and $mounted.Count -ne 1 )
    {
        Throw "Cannot select mounted image as there are $($mounted.Count)"
    }
    if( $mounted.MountMode -ine 'ReadWrite' )
    {
        Throw "Image mounted $($mounted.MountMode) not read/write"
    }
    if( $mounted.MountStatus -ine 'OK ' )
    {
        Throw "Image mounted is status $($mounted.MountStatus) not OK"
    }
    $mountPath = $mounted.Path
}

if( -Not ( Test-Path -Path $mountPath -PathType Container) )
{
    Throw "Can't find mounted WinPE image folder $mountPath"
}

if( [string]::IsNullOrEmpty( $winPESourcePath ) )
{
    [string]$driverPath = Get-CimInstance -ClassName win32_systemdriver -Filter "Description = 'WIMMOUNT'" -ErrorAction SilentlyContinue|Select-Object -ExpandProperty PathName
    if( [string]::IsNullOrEmpty( $driverPath ) )
    {
        Throw "Unable to determine ADK installation folder"
    }
    ## path will be something like \??\C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\ARM64\DISM\wimmount.sys
    ## go 4 levels up
    [string]$rootPath = Split-Path -Path (Split-Path -Path (Split-Path -Path (Split-Path -Path ($driverPath -replace '^\\\?\?\\') -Parent) -Parent) -Parent) -Parent
    $winPESourcePath = Join-Path -Path $rootPath -ChildPath 'Windows Preinstallation Environment'
}

if( -Not ( Test-Path -Path $winPESourcePath -PathType Container) )
{
    Throw "Can't find base WinPE folder at $winPESourcePath"
}

if( -Not $noSecureStartup )
{
    $packages += @( 
        "WinPE-SecureStartup.cab" , 
        "$language\WinPE-SecureStartup_$language.cab" , 
        "WinPE-PlatformId.cab"
    )
}

ForEach( $package in $packages )
{
    $counter++
    [string]$packagePath = [System.IO.Path]::Combine( $winPESourcePath , $architecture , $winPEoptionalComponents , $package )

    if( Test-Path -Path $packagePath -PathType Leaf )
    {
        Write-Verbose -Message "$counter / $($packages.Count) : $package"

        if( $useDism )
        {
            $dismOutput = Dism.exe /Add-Package /Image:"$mountPath" /PackagePath:$packagePath
        }
        else
        {
            $dismOutput = $null
            $result = Add-WindowsPackage -PackagePath $packagePath -Path $mountPath -ErrorVariable dismOutput
        }
        $dismStatus = $?
        if( -Not $dismStatus )
        {
            [string]$message = "Operation on $package may have had problems - $dismOutput"
            if( $abort )
            {
                Throw $message
            }
            else
            {
                Write-Warning -Message $message
            }
        }
    }
    else
    {
        Throw "Can't find package $packagePath"
    }
}
