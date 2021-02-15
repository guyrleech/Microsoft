#requires -version 3

<#
.SYNOPSIS

    Wrapper for CIM File System (CIMFS) calls - requires minimum Win10 2004

.DESCRIPTION

    CIMFS APIS are documented at https://docs.microsoft.com/en-us/windows/win32/api/_cimfs/

.PARAMETER action

    The action to undertake

.PARAMETER path

    The containing folder which contains a single .cim file

.PARAMETER flags

    The flags to pass to the action

.PARAMETER guid

    The GUID of the volume id to take the action on. If one is not specified for the mount action, a random one will be generated

.EXAMPLE

    & .\cimfs.ps1 -action mount -path "\\grl-nas02\MSIX App Attach\Chrome v88 x64"

    Mount the only .cim file in the folder "\\grl-nas02\MSIX App Attach\Chrome v88 x64". If zero or more than one .cim files are found, the mount will fail

.EXAMPLE

    & .\cimfs.ps1 -action mount -path "\\grl-nas02\MSIX App Attach\Chrome v88 x64\Chrome.msix.cim"

    Mount the specified .cim file contained in the folder given.

.EXAMPLE

    & .\cimfs.ps1 -action list

    Show all currently mounted .cim files and how much disk space they are consuming
    
.EXAMPLE

    & .\cimfs.ps1 -action unmount -guid '{ab703362-8cd6-421b-8042-6f75dc0e9572}'

    Prompt to unmount the .cim file mounted with the given volume id (available from a -action list operation)
   
.EXAMPLE

    & .\cimfs.ps1 -action unmountall -confirm:$false

    Unmount all mounted .cim volumes without prompting for confirmation

.NOTES

    Once the .cim file is mounted, its contents can be examined thus where the Guid comes from a -action list operation:

        Get-ChildItem -LiteralPath '\\?\Volume{82c329d3-155f-4a26-abdf-717accc0652f}\'

    Mount with:

        cmd.exe /c mklink /j c:\temp\kim "\\?\Volume{82c329d3-155f-4a26-abdf-717accc0652f}\"

    Delete the mount point, not the contents with:

        [io.directory]::Delete( "c:\temp\kim" )

    CIM files can be created from MSIX packages using the msixmgr.exe tool - https://techcommunity.microsoft.com/t5/windows-virtual-desktop/simplify-msix-image-creation-with-the-msixmgr-tool/m-p/2118585


    Modification History:

        @guyrleech 15/02/21  First release, only implementing CimMountImage() and CimDismountImage()
#>

[CmdletBinding(DefaultParameterSetName='None',SupportsShouldProcess=$true,ConfirmImpact='High')]

Param
(
    [Parameter(Mandatory=$true,HelpMessage='CIMFS action')]
    [ValidateSet('Mount','Unmount','List','Unmountall')]
    [string]$action ,
    [string]$path ,
    [int]$flags = 0 ,
    [GUID]$guid
)

if( ! ( $os = Get-CimInstance -ClassName Win32_OperatingSystem ) )
{
    Throw 'Failed to get Win32_OperatingSystem'
}

if( $os.Caption -notmatch 'Windows 10 ' )
{
    Throw 'OS must be Windows 10 2004 minimum'
}

[version]$osVersion = $os.Version

if( $osVersion.Major -lt 10 -or $osVersion.Build -lt 19041 )
{
    Throw 'OS must be Windows 10 2004 minimum'
}

## https://docs.microsoft.com/en-us/windows/win32/api/_cimfs/

Add-Type -ErrorAction Stop -TypeDefinition @'
    using System;
    using System.Runtime.InteropServices;

    public static class cimfs
    {
        [DllImport( "cimfs.dll" , CallingConvention = CallingConvention.StdCall , ExactSpelling=true, PreserveSig=false , SetLastError = true , CharSet = CharSet.Auto )]
        public static extern long CimMountImage(
            [MarshalAs(UnmanagedType.LPWStr)]
            String imageContainingPath ,
            [MarshalAs(UnmanagedType.LPWStr)]
            String imageName ,
            int mountImageFlags , 
            ref Guid volumeId );

        [DllImport( "cimfs.dll" , CallingConvention = CallingConvention.StdCall , ExactSpelling=true, PreserveSig=false , SetLastError = true , CharSet = CharSet.Auto )]
        public static extern long CimDismountImage(
            Guid volumeId );
    }
'@

[int]$result = 1

try
{
    if( $action -eq 'mount' )
    {
        ## mount function needs containing folder name and cim file so if path passed is a .cim file, split into these two. If it's a folder, if there's a single .cim file , mount that else error 
        if( ! $PSBoundParameters[ 'path' ] )
        {
            Throw 'Must specify .cim file or containing folder to mount'
        }
        [string]$CIMfolder = $null
        [string]$CIMfile = $null
        if( Test-Path -Path $path -PathType Container )
        {
            [array]$cimfiles = @( Get-ChildItem -Path $path | Where-Object -Property Extension -eq '.cim' ) # -include *.cim doesn't work reliably
            if( ! $cimfiles -or ! $cimfiles.Count )
            {
                Throw "No .cim files found in folder `"$path`""
            }
            if( $cimfiles.Count -gt 1 )
            {
                Throw "Found $($cimfiles.Count) .cim files in `"$path`" - must only be one if only specifying a folder"
            }
            $CIMfolder = $path
            $CIMfile = $cimfiles.Name
        }
        elseif( Test-Path -Path $path -PathType Leaf )
        {
            $CIMfolder = Split-Path -Path $path -Parent
            $CIMfile = Split-Path -Path $path -Leaf
        }
        else
        {
            Throw "Unable to find `"$path`""
        }
        if( ! $PSBoundParameters[ 'guid' ] )
        {
            $guid = (New-Guid).Guid
        }

        $result = [cimfs]::CimMountImage( $CIMfolder , $CIMfile , $flags , [ref]$guid )
    
        $lastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()

        if( ! $result ) ## ERROR_SUCCESS
        {
            if( ! ( $volume = Get-CimInstance -ClassName win32_volume | Where-Object { $_.filesystem -eq 'cimfs' -and $_.Name -match $guid.Guid } ) )
            {
                Write-Warning -Message "Mount of `"$fullFilePath`" succeeded but unable to find cimfs filesystem with guid $($guid.Guid)"
            }
            else
            {
                Write-Verbose -Message "Mounted $CIMfolder\$CIMfile with guid $($guid.Guid)"
                [pscustomobject]@{
                    'DeviceID' = $volume.Name
                    'CIMfile' = Join-Path -Path $CIMfolder -ChildPath $CIMfile
                    'Status' = 'Mounted'
                    'FileSystem' = $volume.FileSystem
                }
            }
        }
        else
        {
            Throw "Failed to mount `"$fullFilePath`" - $lastError"
        }
    }
    elseif( $action -eq 'unmount' )
    {
        if( ! $PSBoundParameters[ 'guid' ] )
        {
            Throw 'Must specify guid of mounted CIM file when unmounting'
        }
        if( ! ( $volume = Get-CimInstance -ClassName win32_volume | Where-Object { $_.filesystem -eq 'cimfs' -and $_.Name -match $guid.Guid } ) )
        {
            Write-Warning -Message "Unable to find cimfs filesystem with guid $($guid.Guid)"
        }

        if( $PSCmdlet.ShouldProcess( $volume.Name , 'Unmount' ) )
        {
            $result = [cimfs]::CimDismountImage( $guid )
    
            $lastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()

            if( $result )
            {
                Throw "Unmount of cimfs volume with GUID $($guid.Guid) failed - $lastError"
            }
            else
            {
                [pscustomobject]@{
                    'DeviceID' = $volume.Name
                    'Status' = 'Unmounted'
                    'FileSystem' = $volume.FileSystem
                }
            }
        }
    }
    elseif( $action -eq 'unmountall' )
    {
        [int]$unmounted = 0
        [int]$mounted = 0
        Get-CimInstance -ClassName win32_volume | Where-Object filesystem -eq 'cimfs' | ForEach-Object `
        {
            $volume = $_
            $mounted++
            if( $volume.name -match '^\\\\\?\\Volume(\{.*\})' -and ( $thisGuid = $Matches[1] -as [GUID] ) )
            {
                if( $PSCmdlet.ShouldProcess( $volume.Name , 'Unmount' ) )
                {
                    Write-Verbose -Message "Unmounting $mounted : $($thisGuid.Guid)"
                
                    if( ! ( $result = [cimfs]::CimDismountImage( $thisGuid ) ) )
                    {
                        $unmounted++          
                        [pscustomobject]@{
                            'DeviceID' = $volume.Name
                            'Status' = 'Unmounted'
                            'FileSystem' = $volume.FileSystem
                        }
                    }
                }
            }
            else
            {
                Write-Warning -Message "Failed to get GUID from $($volume.Name)"
            }
        }
        Write-Verbose -Message "Unmounted $unmounted out of $mounted cimfs volumes"
        if( ! $mounted )
        {
            Write-Warning -Message 'No mounted cimfs volumes found'
        }
        elseif( $unmounted -lt $mounted )
        {
            Write-Warning -Message "Only umounted $unmounted cimfs volumes out of $mounted found"
        }
    }
    elseif( $action -eq 'list' )
    {
        Get-CimInstance -ClassName win32_volume | Where-Object filesystem -eq 'cimfs' | Select-Object -Property DeviceID,@{n='Used (MB)';e={[math]::Round( (Get-ChildItem -LiteralPath $_.Name -File -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum / 1MB , 1 )}}
        ## TODO report where it is mounted
    }
    else
    {
        Throw "Unknown action `"$action`""
    }
}
catch
{
    Throw $_
}
