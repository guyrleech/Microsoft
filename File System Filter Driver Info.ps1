#requires -RunAsAdministrator

<#
.SYNOPSIS
    Show details of file system filter drivers by parsing fltmc.exe output and cross referencing to win32_systemdriver, Win32_LoadOrderGroupServiceMembers and file system

.NOTES
    Modification History:

    2024/04/12  @guyrleech  Script created
    2024/04/15  @guyrleech  Added -nonMicrosoft argument and load order group lookup
#>

[CmdletBinding()]

Param
(
    [switch]$nonMicrosoft ,
    [string]$microsoftOwnSubject = '^CN=Microsoft Windows, O=Microsoft Corporation, L=Redmond, S=Washington, C=US$'
)

[array]$filterDrivers = @( fltMC.exe | Where-Object { $_ -match '^(\w+)\s+(\d+)\s+([\S]+)\s+(\d+)' } | Select-Object -Property @{name='Name';expression={$Matches[1]}},
  @{name='Instances';expression={$Matches[2] -as [int]}}, @{name='Altitude';expression={$Matches[3] -as [decimal]}}, @{name='Frame';expression={$Matches[4] -as [int]}} )

## npsvctrig is kernel driver so can't filter on ServiceType = 'File System Driver' 
[array]$systemDrivers = @( Get-CimInstance -ClassName win32_systemdriver )

[hashtable]$loadOrderGroups = @{}
Get-CimInstance -ClassName Win32_LoadOrderGroupServiceMembers | ForEach-Object { $loadOrderGroups.Add( $_.PartComponent.Name , $_.GroupComponent.Name ) }

Write-Verbose -Message "Got $($filterDrivers.Count) filter drivers and $($systemDrivers.Count) system drivers"

ForEach( $filterDriver in $filterDrivers )
{
    $systemDriver = $systemDrivers | Where-Object -Property Name -ieq $filterDriver.Name
    $driverFile = $systemDriver | Select-Object -ExpandProperty PathName
    $signing = $null
    if( $null -ne $driverFile )
    {
        $driverFileProperties = Get-ItemProperty -Path $driverFile -ErrorAction SilentlyContinue
        $signing = Get-AuthenticodeSignature -FilePath $driverFile
    }
    if( $null -eq $signing -or -Not $nonMicrosoft -or $signing.SignerCertificate.Subject -notmatch $microsoftOwnSubject )
    {
        [pscustomobject]@{
            Name = $filterDriver.Name
            Description = $systemDriver | Select-Object -ExpandProperty Description
            Altitude = $filterDriver.Altitude
            Group = $loadOrderGroups[ $filterDriver.Name ]
            DriverFile = $driverFile
            StartMode = $systemDriver | Select-Object -ExpandProperty StartMode
            State = $systemDriver | Select-Object -ExpandProperty State
            AcceptPause = $systemDriver | Select-Object -ExpandProperty AcceptPause
            AcceptStop = $systemDriver | Select-Object -ExpandProperty AcceptStop
            Status = $systemDriver | Select-Object -ExpandProperty Status
            Company = $driverFileProperties | Select-Object -ExpandProperty VersionInfo -ErrorAction SilentlyContinue | Select-Object -ExpandProperty CompanyName -ErrorAction SilentlyContinue
            FileDescription = $driverFileProperties | Select-Object -ExpandProperty VersionInfo -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FileDescription -ErrorAction SilentlyContinue
            ProductName = $driverFileProperties | Select-Object -ExpandProperty VersionInfo -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ProductName -ErrorAction SilentlyContinue
            FileVersion = $driverFileProperties | Select-Object -ExpandProperty VersionInfo -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FileVersionRaw -ErrorAction SilentlyContinue
            ProductVersion = $driverFileProperties | Select-Object -ExpandProperty VersionInfo -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ProductVersionRaw -ErrorAction SilentlyContinue
            SigningStatus  = $signing | Select-Object -ExpandProperty Status
            SigningSubject = $signing | Select-Object -ExpandProperty SignerCertificate | Select-Object -ExpandProperty Subject
            SigningIssuer  = $signing | Select-Object -ExpandProperty SignerCertificate | Select-Object -ExpandProperty Issuer
            SigningNotBefore = $signing | Select-Object -ExpandProperty SignerCertificate | Select-Object -ExpandProperty NotBefore
            SigningNotAfter  = $signing | Select-Object -ExpandProperty SignerCertificate | Select-Object -ExpandProperty NotAfter
        }
    }
    else
    {
        Write-Verbose -Message "Excluding $($filterDriver.Name)"
    }
}

$null = $null
