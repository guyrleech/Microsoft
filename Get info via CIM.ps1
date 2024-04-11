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

.PARAMETER delimiter

Delimiter used in class names to spearate an optional namespace from the class name

.PARAMETER classes

A comma separated list of CIM classes to query

.PARAMETER expandArrays

The delimiter to use to expand result values which are arrays. Specify the empty string '' or $null to stop array expansion

.PARAMETER noProcessDetail

Do not include extra process detail file which has version information

.PARAMETER noInstalledPrograms

Do not process the registry to get installed programs

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
31/03/20  @guyrleech  Added CIM_LogicalDevice and CIM_System to common
28/03/24  @guyrleech  Added PnP classes to common. Added support for different namespaces in common & added SecurityCenter(2) classes to common
11/04/24  @guyrleech  Added non-CIM queries to get installed programs and process exe & module versions
#>

[CmdletBinding(DefaultParameterSetName = 'common')]

Param
(
    [Parameter(ParameterSetName='common')]
    [switch]$common ,
    [System.Collections.Generic.List[string]]$computers ,
    [string]$computersFile ,
    [string]$outputFolder = $env:TEMP ,
    [string]$excludeClasses = '^__|Win32_PerfRawData|Win32_PerfFormatted' ,
    [string]$includeClasses ,
    [string]$namespace,
    [switch]$overwrite ,
    [switch]$noProcessDetail ,
    [switch]$noInstalledPrograms ,
    [System.Management.Automation.PSCredential]$credential ,
    [ValidateSet('Basic','CredSsp','Default','Digest','Kerberos','Negotiate','NtlmDomain')]
    [string]$authentication ,
    [int]$timeout , ## seconds
    [Parameter(ParameterSetName='classes')]
    [string[]]$classes ,
    [string]$delimiter = ':' ,
    [AllowNull()]
    [AllowEmptyString()]
    [string]$expandArrays = ','
)

## code from https://github.com/guyrleech/Microsoft/blob/master/Get%20installed%20software.ps1
Function Process-RegistryKey
{
    [CmdletBinding()]
    Param
    (
        [string]$hive ,
        $reg ,
        [string[]]$UninstallKeys ,
        [switch]$includeEmptyDisplayNames ,
        [AllowNull()]
        [string]$username ,
        [AllowNull()]
        [string]$productname ,
        [AllowNull()]
        [string]$vendor
    )

    ForEach( $UninstallKey in $UninstallKeys )
    {
        $regkey = $reg.OpenSubKey($UninstallKey) 
    
        if( $regkey )
        {
            [string]$architecture = if( $UninstallKey -match '\\wow6432node\\' ){ '32 bit' } else { 'Native' } 

            $subkeys = $regkey.GetSubKeyNames() 
    
            foreach($key in $subkeys)
            {
                $thisKey = Join-Path -Path $UninstallKey -ChildPath $key 

                $thisSubKey = $reg.OpenSubKey($thisKey) 

                if( $includeEmptyDisplayNames -or -Not [string]::IsNullOrEmpty( $thisSubKey.GetValue('DisplayName') ) )
                {
                    [string]$installDate = $thisSubKey.GetValue('InstallDate')
                    $installedOn = New-Object -TypeName 'DateTime'
                    if( [string]::IsNullOrEmpty( $installDate ) -or ! [datetime]::TryParseExact( $installDate , 'yyyyMMdd' , [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$installedOn ) )
                    {
                        $installedOn = $null
                    }
                    $size = New-Object -TypeName 'Int'
                    if( ! [int]::TryParse( $thisSubKey.GetValue('EstimatedSize') , [ref]$size ) )
                    {
                        $size = $null
                    }
                    else
                    {
                        $size = [math]::Round( $size / 1KB , 1 ) ## already in KB
                    }

                    if( $thisSubKey.GetValue('DisplayName') -match $productname -and $thisSubKey.GetValue('Publisher') -match $vendor )
                    {
                        [pscustomobject][ordered]@{
                            'ComputerName' = $computername
                            'Hive' = $Hive
                            'User' = $username
                            'Key' = $key
                            'Architecture' = $architecture
                            'DisplayName' = $($thisSubKey.GetValue('DisplayName'))
                            'DisplayVersion' = $($thisSubKey.GetValue('DisplayVersion'))
                            'InstallLocation' = $($thisSubKey.GetValue('InstallLocation'))
                            'InstallSource' = $($thisSubKey.GetValue('InstallSource'))
                            'Publisher' = $($thisSubKey.GetValue('Publisher'))
                            'InstallDate' = $installedOn
                            'Size (MB)' = $size
                            'System Component' = $($thisSubKey.GetValue('SystemComponent') -eq 1)
                            'Comments' = $($thisSubKey.GetValue('Comments'))
                            'Contact' = $($thisSubKey.GetValue('Contact'))
                            'HelpLink' = $($thisSubKey.GetValue('HelpLink'))
                            'HelpTelephone' = $($thisSubKey.GetValue('HelpTelephone'))
                            'Uninstall' = $($thisSubKey.GetValue('UninstallString'))
                        }
                    }
                }
                else
                {
                    Write-Warning "Ignoring `"$hive\$thisKey`" on $computername as has no DisplayName entry"
                }

                $thisSubKey.Close()
            } 
            $regKey.Close()
        }
        elseif( $hive -eq 'HKLM' )
        {
            Write-Warning "Failed to open `"$hive\$UninstallKey`" on $computername"
        }
    }
}

[string[]]$UninstallKeys = @( 'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall' , 'SOFTWARE\\wow6432node\\Microsoft\\Windows\\CurrentVersion\\Uninstall' )

if( $common )
{
    $classes = @(
        'root/SecurityCenter:AntiVirusProduct'
        'root/SecurityCenter:AntiSpywareProduct'
        'root/SecurityCenter:FirewallProduct'
        'root/SecurityCenter2:AntiVirusProduct'
        'root/SecurityCenter2:AntiSpywareProduct'
        'root/SecurityCenter2:FirewallProduct'
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
        'win32_memorydevice' ,
        'CIM_LogicalDevice' ,
        'Win32_PnPEntity' ,
        'Win32_PnPSignedDriver' ,
        'Win32_PnPDevice' ,
        'CIM_System' )
}

## stop duplicate computers in the text file
[hashtable]$computersAdded = @{}

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
                $computersAdded.Add( $Matches[1] , $true )
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

if( -Not $PSBoundParameters[ 'classes' ] -and -Not $common )
{
    [hashtable]$classArguments = @{}
    if( $PSBoundParameters[ 'namespace' ] )
    {
        $classArguments.Add( 'namespace' , $namespace )
    }
    $classes = @( Get-CimClass @classArguments | Where-Object { ( -Not $PSBoundParameters[ 'includeClasses' ] -or $_.CimClassName -match $includeClasses ) -and $_.CimClassName -NotMatch $excludeClasses } | Select-Object -ExpandProperty CimClassName | Sort-Object )
}

[int]$counter = 0

if( -Not ( Test-Path -Path $outputFolder -ErrorAction SilentlyContinue -PathType Container ) )
{
    if( -Not ( New-Item -Path $outputFolder -Force -ItemType Directory ) )
    {
        Throw "Failed to create output folder $outputFolder"
    }
}

[hashtable]$CIMSessionParameters = @{ 'Verbose' = $false }
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

$CIMsession = $null
$CIMsession = New-CimSession @CIMSessionParameters

if( $null -eq $CIMsession )
{
    Throw "Failed to create CIM session to $($computers -join ',')"
}

$cimArguments.Add( 'CIMSession' , $CIMsession )

Write-Verbose "Querying $($classes.Count) CIM classes for $(($computers|Measure-Object).Count) machines and writing results to `"$outputFolder`""

ForEach( $class in $classes )
{
    $counter++
    if( -Not $PSBoundParameters[ 'IncludeClasses' ] -or $class -match $includeClasses )
    {
        Write-Verbose "$counter / $($classes.Count) : $class"
        
        [string]$outputFile = (Join-Path -Path $outputFolder -ChildPath "$($class -replace '[/\\:]' , '_').csv" )
        if( $overWrite  `
            -or -Not ( Test-Path -Path $outputFile -PathType Leaf -ErrorAction SilentlyContinue ) `
            -or -Not ( Get-ItemProperty -Path $outputFile | Select -ExpandProperty Length ))
        {
            [hashtable]$differentNamespace = @{}
            [string[]]$nameSpaceAndClass = $class -split $delimiter , 2 ## only split on 1st delimiter
            if( $nameSpaceAndClass -is [array] -and $nameSpaceAndClass.Count -gt 1 )
            {
                if( $nameSpaceAndClass[ 0 ] -ine $namespace )
                {
                    $differentNamespace.Add( 'Namespace' , $nameSpaceAndClass[ 0 ] )
                }
                $class = $nameSpaceAndClass[ 1 ]
            }
            [array]$results = @( Get-CimInstance @cimArguments -ClassName $class @differentNamespace | Select-Object -Property PSComputerName,* -ExcludeProperty PSShowComputerName,CIM?* -ErrorAction SilentlyContinue )
            if( $results -and $results.Count )
            {
                if( -Not [string]::IsNullOrEmpty( $expandArrays ) )
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
            Write-Warning "File `"$outputFile`" already exists so not overwriting"
        }
    }
}

if( -Not $noProcessDetail )
{
    ## Get processes natively so we get exe and module version but save to json as will not be two dimensional
    $currentPrincipal = New-Object -TypeName Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $outputFile = Join-Path -Path $outputFolder -ChildPath 'processes.json'
    if( $overWrite  `
                -or -Not ( Test-Path -Path $outputFile -PathType Leaf -ErrorAction SilentlyContinue ) `
                -or ( Get-ItemProperty -Path $outputFile | Select -ExpandProperty Length ) -eq 0 )
    {
        [hashtable]$processParameters = @{ }
        if( $computers -and $computers.Count )
        {
            $processParameters.Add( 'ComputerName' , $computers )
        }
        elseif( $currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator) )
        {
            $processParameters.Add( 'IncludeUserName' , $true )
        }
        ## need to go a few levels deep to capture version information for loaded modules
        Get-Process @processParameters | ConvertTo-Json -Depth 4 | Out-File -FilePath $outputFile -Force
    }
    else
    {
        Write-Warning "File `"$outputFile`" already exists so not overwriting"
    }
}

Remove-CimSession -CimSession $CIMsession
$CIMsession = $null

$outputFile = Join-Path -Path $outputFolder -ChildPath 'applications.csv'
if( -Not $noInstalledPrograms )
{
    if( $overWrite  `
                -or -Not ( Test-Path -Path $outputFile -PathType Leaf -ErrorAction SilentlyContinue ) `
                -or ( Get-ItemProperty -Path $outputFile | Select -ExpandProperty Length ) -eq 0 )
    {
        if( $null -eq $computers -or $computers.Count -eq 0 )
        {
            $computers = @( $env:COMPUTERNAME )
        }
        [array]$installed = @( foreach($computerName in $computers)
        {
            Write-Verbose -Message "Processing installed apps on $computerName"
            $reg = $null
            $reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey( 'LocalMachine' , $computername )
    
            if( $? -and $reg )
            {
                Process-RegistryKey -Hive 'HKLM' -reg $reg -UninstallKeys $UninstallKeys -includeEmptyDisplayNames
                $reg.Close()
                $reg = $null
            }
            else
            {
                Write-Warning "Failed to open HKLM on $computername"
            }

            $reg = [microsoft.win32.registrykey]::OpenRemoteBaseKey( 'Users' , $computername )
    
            if( $? -and $reg )
            {
                ## get each user SID key and process that for per-user installed apps
                ForEach( $subkey in $reg.GetSubKeyNames() )
                {
                    try
                    {
                        if( $userReg = $reg.OpenSubKey( $subKey ) )
                        {
                            [string]$username = $null
                            try
                            {
                                $username = ([System.Security.Principal.SecurityIdentifier]($subKey)).Translate([System.Security.Principal.NTAccount]).Value
                            }
                            catch
                            {
                                $username = $null
                            }
                            Process-RegistryKey -Hive (Join-Path -Path 'HKU' -ChildPath $subkey) -reg $userReg -UninstallKeys $UninstallKeys -includeEmptyDisplayNames
                            $userReg.Close()
                            $userReg = $null
                        }
                    }
                    catch
                    {
                    }
                }
                $reg.Close()
                $reg = $null
            }
            else
            {
                Write-Warning "Failed to open HKU on $computername"
            }
        } ) | Sort -Property ComputerName, DisplayName

        if( $installed -and $installed.Count )
        {
            Write-Verbose -Message "Writing details of $($installed.Count) apps to $outputFile"
            $installed | Export-Csv -Path $outputFile -NoTypeInformation -Force
        }
    }
    else
    {
        Write-Warning "File `"$outputFile`" already exists so not overwriting"
    }
}