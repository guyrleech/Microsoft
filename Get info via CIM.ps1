#require -version 3

<#
.SYNOPSIS

Gather info from one or more computers via CIM and write to CSV file

.DESCRIPTION

Originally written to help gather information during health check exercises

.PARAMETER computers

A comma separated list of computers to query. If none is specified, the local computer is used

.PARAMETER computersFile

A text file containing one computer per line. Blank lines and those starting with # will be ignored as will any characters like space or # after the computer name

.PARAMETER outputFolder

The folder to write the results files to. Will be created if it doesn't exist

.PARAMETER excludeClasses

A regular expression where any CIM class that matches will not be queried

.PARAMETER includeClasses

A regular expression where only CIM class that matches will be queried when -common is not specified

.PARAMETER namespace

The CIM namespace to query. If not specified the default is used

.PARAMETER overwrite

Overwrite existing results files if they are not empty

.PARAMETER common

Query the script's built-in list of 45+ most common/useful CIM classes

.PARAMETER authentication

The authentication method to use to establish the CIM session

.PARAMETER timeout

Override the default connections/query timeout (in seconds)

.PARAMETER classes

A comma separated list of CIM classes to query

.PARAMETER expandArrays

The delimiter to use to expand result values which are arrays. Specify the empty string '' or $null to stop array expansion

.EXAMPLE

& '.\Get info via CIM.ps1' -computersFile c:\temp\puters.txt -Verbose -outputFolder C:\Temp\CIM -common

Query the built in list of CIM classes on the list of computers contained in the file c:\temp\puters.txt and output the csv results file to c:\temp\CIM

.EXAMPLE

& '.\Get info via CIM.ps1' -computers machine1,machine2 -outputFolder C:\Temp\CIM -classes win32_service,win32_process

Query the win32_service and win32_process CIM classes on computers machine1 and machine2 and output the csv results file to c:\temp\CIM

.EXAMPLE

& '.\Get info via CIM.ps1' -outputFolder C:\Temp\CIM -classes win32_service,win32_process

Query the win32_service and win32_process CIM classes on the local computer and output the csv results file to c:\temp\CIM

.NOTES

To show all available CIM classes, run Get-CIMClass -ClassName *

Modification History:

11/03/20  @guyrleech  Initial public release

#>

[CmdletBinding()]

Param
(
    [System.Collections.Generic.List[string]]$computers ,
    [string]$computersFile ,
    [string]$outputFolder = $env:TEMP ,
    [string]$excludeClasses = '^__|Win32_PerfRawData|Win32_PerfFormatted' ,
    [string]$includeClasses ,
    [string]$namespace,
    [switch]$overwrite ,
    [Parameter(ParameterSetName='common')]
    [switch]$common ,
    [System.Management.Automation.PSCredential]$credential ,
    [ValidateSet('Basic','CredSsp','Default','Digest','Kerberos','Negotiate','NtlmDomain')]
    [string]$authentication ,
    [int]$timeout , ## seconds
    [Parameter(ParameterSetName='classes')]
    [string[]]$classes ,
    [AllowNull()]
    [AllowEmptyString()]
    [string]$expandArrays = ','
)

if( $common )
{
    $classes = @( 
        'Win32_ComputerSystemProduct' , 
        'Win32_ComputerSystem' , 
        'win32_operatingsystem' , 
        'Win32_NetworkAdapter' ,
        'Win32_NetworkProtocol' ,
        'Win32_NetworkClient' ,
        'Win32_NetworkAdapterConfiguration' ,
        'Win32_NetworkLoginProfile' ,
        'Win32_NetworkAdapterSetting' , 
        'win32_volume' , 
        'win32_userprofile' , 
        'win32_process' , 
        'win32_service' , 
        'win32_session' , 
        'win32_logicaldisk' , 
        'Win32_LogicalDiskToPartition' ,
        'win32_diskdrive', 
        'win32_diskpartition' , 
        'win32_bios' , 
        'win32_baseboard' , 
        'win32_currenttime' , 
        'win32_desktop' , 
        'win32_environment' , 
        'win32_group' , 
        'win32_groupuser' , 
        'Win32_IP4PersistedRouteTable' , 
        'win32_ip4routetable' , 
        'Win32_LoadOrderGroup' ,
        'Win32_LoadOrderGroupServiceDependencies' ,
        'Win32_LoadOrderGroupServiceMembers' , 
        'Win32_PageFileUsage' ,
        'Win32_SystemDevices' ,
        'Win32_SystemAccount' ,
        'Win32_SystemDriver' ,
        'Win32_SystemTimeZone' ,
        'Win32_LocalTime' ,
        'Win32_QuickFixEngineering' ,
        'Win32_SystemSetting' ,
        'Win32_DriverForDevice' ,
        'Win32_InstalledWin32Program' , ## This appears to be passive, unlike Win32_Product although returns fewer results
        'Win32_InstalledStoreProgram' ,
        'Win32_PhysicalMemory' ,
        'Win32_PrinterDriver' ,
        'Win32_Printer' ,
        'Win32_Processor' ,
        'Win32_ProtocolBinding' ,
        'win32_memorydevice' )
}

## stop duplicate computers in the text file
[hashtable]$comptersList = @{}

if( $PSBoundParameters[ 'computersfile' ] )
{
    if( ! $computers )
    {
        $computers = New-Object -TypeName System.Collections.Generic.List[string]
    }

    ForEach( $line in (Get-Content -Path $computersFile -ErrorAction Stop | Where-Object { $_ -notmatch '^#' }))
    {
        if( $line -match '([a-z0-9_\-\.]+)' -and ! [string]::IsNullOrEmpty( $Matches[1] ) )
        {
            try
            {
                $comptersList.Add( $Matches[1] , $true )
                $computers.Add( $Matches[1] )
            }
            catch
            {
                Write-Warning -Message "Duplicate computer `"$($Matches[1])`""
            }
        }
    }
}

[hashtable]$cimArguments = @{ 
    'Verbose' = $false
    'ErrorAction'= 'SilentlyContinue' }

if( $PSBoundParameters[ 'namespace' ] )
{
    $cimArguments.Add( 'namespace' , $namespace )
}

if( $PSBoundParameters[ 'timeout' ] )
{
    $cimArguments.Add( 'OperationTimeoutSec' , $timeout )
}

if( ! $PSBoundParameters[ 'classes' ] -and ! $common )
{
    [hashtable]$classArguments = @{}
    if( $PSBoundParameters[ 'namespace' ] )
    {
        $classArguments.Add( 'namespace' , $namespace )
    }
    $classes = @( Get-CimClass @classArguments ' , ' Where-Object { ( ! $PSBoundParameters[ 'includeClasses' ] -or $_.CimClassName -match $includeClasses ) -and $_.CimClassName -NotMatch $excludeClasses } ' , ' Select-Object -ExpandProperty CimClassName ' , ' Sort-Object )
}

[int]$counter = 0

if( ! ( Test-Path -Path $outputFolder -ErrorAction SilentlyContinue -PathType Container ) )
{
    [void]( New-Item -Path $outputFolder -Force -ItemType Directory )
}

[hashtable]$CIMSessionParameters = @{}
if( $computers -and $computers.Count )
{
    $CIMSessionParameters.Add( 'ComputerName' , $computers )
}

if( $PSBoundParameters[ 'credential' ] )
{
    $CIMSessionParameters.Add( 'Credential' , $credential )
}

if( $PSBoundParameters[ 'timeout' ] )
{
    $CIMSessionParameters.Add( 'OperationTimeoutSec' , $timeout )
}

if( $PSBoundParameters[ 'authentication' ] )
{
    $cimArguments.Add( 'authentication' , $authentication )
}

$CIMsession = New-CimSession @CIMSessionParameters

if( ! $CIMsession )
{
    Throw "Failed to create CIM session to $($computers -join ',')"
}

$cimArguments.Add( 'CIMSession' , $CIMsession )

Write-Verbose "Calling $($classes.Count) CIM methods for $($computers.Count) machines and writing results to `"$outputFolder`""

ForEach( $class in $classes )
{
    $counter++
    if( ! $PSBoundParameters[ 'IncludeClasses' ] -or $class -match $includeClasses )
    {
        Write-Verbose "$counter / $($classes.Count) : $class"
        
        [string]$outputFile = (Join-Path -Path $outputFolder -ChildPath "$class.csv")
        if( $overWrite  `
            -or ! ( Test-Path -Path $outputFile -PathType Leaf -ErrorAction SilentlyContinue ) `
            -or ! ( Get-ItemProperty -Path $outputFile | Select -ExpandProperty Length ))
        {
            [array]$results = @( Get-CimInstance @cimArguments -ClassName $class | Select-Object -Property PSComputerName,* -ExcludeProperty PSShowComputerName,CIM?* -ErrorAction SilentlyContinue )
            if( $results -and $results.Count )
            {
                if( ! [string]::IsNullOrEmpty( $expandArrays ) )
                {
                    ForEach( $result in $results )
                    {
                        ForEach( $property in $result.PSObject.Properties )
                        {
                            if( $property.Value -and $property.MemberType -eq 'NoteProperty' -and $property.Value -is [array] -and $property.Value.Count )
                            {
                                $result.$($property.Name) = $property.Value -join $expandArrays
                            }
                        }
                    }
                }
                $results | Export-CSV -Path $outputFile -NoTypeInformation
            }
            else
            {
                Write-Warning "No data returned from $class"
            }
        }
        else
        {
            Write-Verbose "File `"$outputFile`" already exists so not redoing"
        }
    }
}

Remove-CimSession -CimSession $CIMsession
$CIMsession = $null
