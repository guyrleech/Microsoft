#requires -version 3

<#
.SYNOPSIS
    Get parent devices for devices matching query

.PARAMETER deviceRegex
    The regular expression to use to match the single device to retrieve information for

.PARAMETER recurse
    Work up the parent hierarchy to the top otherwise just gives device and immediate parent

.PARAMETER level
    Used internally to order the hierarchy and indent the device name text, no need to change unless you don't like negative numbers

.PARAMETER replace
    Regular expression used to change device property names, don't use unless required

.PARAMETER with
    String to replace instances of matches for the -replace regex in the device properties.
    Use the empty string to remove the matching text

.PARAMETER properties
    The device properties to output. Contains a built in list

.EXAMPLE
    . '.\Device Hierarchy.ps1' -deviceRegex "realtek usb gbe" -recurse

    Find the device which matches "realtek usb gbe" and get all of its parent devices

.NOTES
    Modification History:

    2023/12/28  @guyrleech  Script born
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
    [Parameter(Mandatory=$true,HelpMessage='Regex to match the device to enumerate parents')]
    [string]$deviceRegex ,
    [switch]$recurse ,
    [array]$properties = @( @{n='IndentedName';e={ "{0}{1}" -f ( ' ' * ($lowestLevel + $_.HierarchyLevel)), $_.FriendlyName }},'HierarchyLevel','DeviceId','Present','Device_DriverDate','Device_DriverProvider','Device_DriverVersion','Device_DriverInfSection','Device_DriverInfPath','Device_LocationInfo' ) ,
    [string]$replace = 'DEVPKEY_' ,
    [string]$with = '' ,
    [int]$level = 1
)

#region Functions
Function Add-DeviceProperties
{
    Param
    (
        [Parameter(Mandatory=$true)]
        $inputObject ,
        [Parameter(Mandatory=$true)]
        $deviceId ,
        $replace ,
        $with
    )
    [hashtable]$properties = @{}
    Get-PnpDeviceProperty -InstanceId $deviceId -ErrorAction SilentlyContinue | ForEach-Object `
    {
        $property = $_
        [string]$propertyName = $property.KeyName -replace $replace , $with
        if( -Not $inputObject.PSObject.Properties[ $propertyName ] )
        {
            $properties.Add( $propertyName , $property.Data )
        }
    }
    if( $properties.Count -gt 0 )
    {
        Add-Member -InputObject $inputObject -NotePropertyMembers $properties
    }
}

Function Get-ParentDevice
{
    Param
    (
        [string]$deviceId ,
        $level = 0 ,
        [switch]$recurse 
    )

    Write-Verbose -Message "Get-ParentDevice -level $level -deviceId $deviceId"
    $parent = $null
    if( -Not [string]::IsNullOrEmpty( $deviceId ) )
    {
        $parentDeviceId = $null
        $parentDeviceId = Get-PnpDeviceProperty -InstanceId $deviceId -ErrorAction SilentlyContinue -KeyName 'DEVPKEY_Device_Parent' | Select-Object -ExpandProperty Data -ErrorAction SilentlyContinue
        if( $null -ne $parentDeviceId )
        {
            if( $recurse )
            {
                Get-ParentDevice -deviceId $parentDeviceId -recurse -level ($level - 1)
            }
            $parentDevice = $script:allDevices | Where-Object DeviceId -ieq $parentDeviceId
            if( $null -eq $parentDevice )
            {
                Write-Warning -Message "Unable to find parent device id $parentDeviceId for device $deviceId"
            }
            elseif( $parentDevice.PSObject.Properties[ 'FriendlyName' ] -and -Not [string]::IsNullOrEmpty( $parentDevice.FriendlyName ) )
            {
                Write-Verbose -Message "`tGet-ParentDevice -level $level -deviceId $deviceId"
                $parent = Add-Member -PassThru -InputObject $parentDevice -NotePropertyMembers @{
                    HierarchyLevel = $level
                    ChildDeviceId = $deviceId
                }
                Add-DeviceProperties -InputObject $parent -deviceId $parentDeviceId -replace $replace -with $with
                $parent ##return
            }
        }
        else
        {
            Write-Warning -Message "Unable to find parent for device $deviceId"
        }
    }
}
#endregion Functions

Import-Module -Name PnpDevice -Verbose:$false -Debug:$false

[array]$script:allDevices = @( Get-PnpDevice )

Write-Verbose -Message "Got $($script:allDevices.Count) devices in total"

$matchingDevices = $script:allDevices | Where-Object FriendlyName -match $deviceRegex

if( $null -eq $matchingDevices )
{
    Throw "No devices found matching `"$deviceRegex`""
}
elseif( $matchingDevices -is [array] -and $matchingDevices.Count -ne 1 )
{
    Throw "$($matchingDevices.Count) devices found matching `"$deviceRegex`"`n`t$(($matchingDevices | Select-Object -ExpandProperty FriendlyName) -join "`n`t")"
}

$baseDevice = Add-Member -InputObject $matchingDevices -passThru -NotePropertyMembers @{
        HierarchyLevel = $level
        ChildDeviceId = $null ## TODO could see if has child as in this is the parent device for something(s)
}
Add-DeviceProperties -InputObject $baseDevice -deviceId $matchingDevices.DeviceId -replace $replace -with $with

$level--

$tree = @(
    $baseDevice
    Get-ParentDevice -device $matchingDevices.DeviceId -recurse:$recurse -level $level )

## get lowest level so we can figure out how many spaces to indent each child item's name
$lowestLevel = [math]::Abs( ($tree | Measure-Object -Property HierarchyLevel -Minimum | Select-Object -ExpandProperty Minimum ) )

## output highest level parents first
$tree | Sort-Object -Property HierarchyLevel | Select-Object -Property $properties
