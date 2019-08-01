#requires -version 3
<#
    Show FSLogix mounted volume details & cross reference to FSLogix session information in the registry

    @guyrleech 2019

    Modification History:

    01/08/19   GRL   Initial public release
#>

<#
.SYNOPSIS

Show FSLogix currently mounted volume details & cross reference to FSLogix session information in the registry

.DESCRIPTION

Gets Windows disks, volumes and partitions information and correlates with HKEY_LOCAL_MACHINE\SOFTWARE\FSLogix\Profiles\Sessions to show disk sizes, capacities and free space

.PARAMETER label

Only include partitions whose label matches this regular expression. They are typically labelled "Profile-%username%"

.PARAMETER noUsedSpace

Do not iterate over the mounted disks contents to calculate how much data they contain.

.EXAMPLE

& '.\Show FSlogix volumes.ps1' -verbose -label billybob  -noUsedSpace | Out-GridView

Show the disk and associated information for the user billybob, except for the space consumed by the disk's contents, in an on screen sortable/filterable grid view

.EXAMPLE

& '.\Show FSlogix volumes.ps1' | Sort-Object -Property 'Free Capacity %' | Format-Table -Autosize

Show the disk and associated information for the all currently logged on users and sort on the free capacity so those listed first have the least free space

#>

[CmdletBinding()]

Param
(
    [string]$label ,
    [switch]$noUsedSpace
)

## TODO pull in login times from LSASS and disk mounting events from event log to show times and durations of each

Function Calculate-FolderSize( [string]$folderName )
{
    $items = @( $folderName )
    [array]$files = While( $items )
    {
        $newitems = $items | Get-ChildItem -Force -ErrorAction SilentlyContinue | Where-Object { ! ( $_.Attributes -band [System.IO.FileAttributes]::ReparsePoint ) }
        $newitems
        $items = $newitems | Where-Object { $_.Attributes -band [System.IO.FileAttributes]::Directory }
    }
    if( $files -and $files.Count )
    {
        [long]($files | Measure-Object -Property Length -Sum | Select -ExpandProperty Sum)
    }
    else
    {
        [long]0
    }
}

[array]$partitions = @( Get-Partition | Where-Object { $_.DiskId -match '&ven_msft&prod_virtual_disk' -and ! $_.DriveLetter -and $_.Type -eq 'Basic' } )

if( ! $partitions -or ! $partitions.Count )
{
    Throw "No partitions found mounted off virtual disks"
}

Write-Verbose "Found $($partitions.Count) virtual disk partitions"

[array]$fixedVolumes = @( Get-Volume | Where-Object { $_.DriveType -eq 'Fixed' } )

if( ! $fixedVolumes -or ! $fixedVolumes.Count )
{
    Write-Warning "Unable to find any fixed volumes"
}
else
{
    Write-Verbose "Found $($fixedVolumes.Count) fixed volumes"
}

[array]$virtualDisks = @( Get-Disk | Where-Object { $_.BusType -eq 'File Backed Virtual' } )

if( ! $virtualDisks -or ! $virtualDisks.Count )
{
    Write-Warning "Unable to find any file backed virtual disks"
}
else
{
    Write-Verbose "Found $($virtualDisks.Count) file backed virtual disks"
}

[int]$counter = 0

[array]$results = @( ForEach( $partition in $partitions )
{
    $counter++
    Write-Verbose "$counter / $($partitions.Count) : Partition GUID $($partition.Guid)"

    $volume = $fixedVolumes | Where-Object { $_.UniqueId -match $partition.Guid }
    if( ! $volume )
    {
        Write-Warning "Unable to find fixed volume with GUID $($partition.Guid)"
    }
    if( ! $PSBoundParameters[ 'label' ] -or ($volume -and $volume.FileSystemLabel -match $label ))
    {
        [string]$uniqueId = ($partition.UniqueId -split '[{}]')[-1]
        $disk = $virtualDisks | Where-Object { $_.UniqueId -eq $uniqueId }
        if( ! $disk )
        {
            Write-Warning "Unable to find disk with unique id $uniqueId"
        }
        $result = [pscustomobject][ordered]@{
            'Label' = $volume | Select-Object -ExpandProperty FileSystemLabel
            'Operational Status' = $volume | Select-Object -ExpandProperty OperationalStatus
            'Health Status' = $volume | Select-Object -ExpandProperty HealthStatus
            'Provisioning Type' = $disk | Select-Object -ExpandProperty ProvisioningType
            'Disk Capacity (GB)' = [math]::Round( ( $disk | Select-Object -ExpandProperty Size ) / 1GB , 2 )
            'Volume Size (GB)' = [math]::Round( ( $volume | Select-Object -ExpandProperty Size ) / 1GB , 2 )
            'Volume Capacity (GB)' = [math]::Round(  ( $volume | Select-Object -ExpandProperty SizeRemaining ) / 1GB , 2 )
            'Free Capacity %' = [math]::Round( ( $volume | Select-Object -ExpandProperty SizeRemaining ) / ( $volume | Select-Object -ExpandProperty Size ) * 100 , 2 )
        }
        [array]$paths = Get-ChildItem -LiteralPath ($partition|Select-Object -ExpandProperty AccessPaths) | . { Process `
        {
            [string]$folder = $_.FullName
            [string]$childFolder = $_.Name
        
            if( ! $noUsedSpace )
            {
                Add-Member -InputObject $result -MemberType NoteProperty -Name "`"$childFolder`" Folder Size (GB)" -Value ([math]::Round( (Calculate-FolderSize -folderName $folder) / 1GB , 2 ))
            }

            Add-Member -InputObject $result -MemberType NoteProperty -Name "`"$childFolder`" Folder Permissions" -Value ((Get-Acl -LiteralPath $folder | Select -ExpandProperty AccessToString) -replace "[`n`r]" , ' , ')
        }}
        $fslogixRegValue = Get-ItemProperty -Path "HKLM:\SOFTWARE\FSLogix\Profiles\Sessions\*" -ErrorAction SilentlyContinue | Where-Object { $_.Volume -eq $volume.Path }
        if( ! $fslogixRegValue )
        {
            Write-Warning "Couldn't find FSlogix registry key for volume $($volume.Path)"
        }
        else
        {
            Add-Member -InputObject $result -NotePropertyMembers @{
                'Username' = ([System.Security.Principal.SecurityIdentifier]($fslogixRegValue.PSChildName)).Translate([System.Security.Principal.NTAccount]).Value
                'Profile Path' = $fslogixRegValue.ProfilePath
                'Local Profile Path' = $fslogixRegValue.LocalProfilePath
                'Session Id' = $fslogixRegValue.WindowsSessionID
                'Last Profile Load Time (ms)' = $fslogixRegValue.LastProfileLoadTimeMS
            }
        }
        Add-Member -InputObject $result -NotePropertyMembers @{
            'Paths' = ($partition | Select-Object -ExpandProperty AccessPaths) -join ' , '
            'Location' = $disk | Select-Object -ExpandProperty Location
            'Physical Sector Size' = $disk | Select-Object -ExpandProperty PhysicalSectorSize
        }
        $result
    }
    else
    {
        Write-Verbose "Excluding $($volume.FileSystemLabel)"
    }
})

$results