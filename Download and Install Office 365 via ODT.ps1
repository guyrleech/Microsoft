#requires -version 3
<#
    Install Office 365 via ODT. Script to be used by Flexxible IT Apps2Digital formulas

    https://docs.microsoft.com/en-us/deployoffice/overview-of-the-office-2016-deployment-tool

    @guyrleech 2019
#>

[CmdletBinding()]

Param
(
    # no parameters as parameter mechanism is via #pattern#
)

$ProgressPreference = 'SilentlyContinue'

## TODO Replace "64" with a #Bitness# parameter
## TODO Option to install Visio
## TODO Option to not install OneDrive

# SourcePath="**officeDownloadFolder**" 
[string]$configXML = @'
<Configuration>
  <Add OfficeClientEdition="64" Channel="Monthly">
    <Product ID="O365ProPlusRetail">
      <Language ID="en-us" />
    </Product>
  </Add>
    <Updates Enabled="FALSE" Channel="Monthly" />
    <Display Level="None" AcceptEULA="TRUE" />
    <Property Name="AUTOACTIVATE" Value="0" />
    <Logging Level="Standard" Path="**workingFolder**" />
</Configuration>
'@

[bool]$createdFolder = $false
[string]$odtDownloadRegex = '/officedeploymenttool[a-z0-9_-]*\.exe$'

## Check running as admin
$myWindowsPrincipal = New-Object System.Security.Principal.WindowsPrincipal([System.Security.Principal.WindowsIdentity]::GetCurrent())
$adminRole = [System.Security.Principal.WindowsBuiltInRole]::Administrator
if ( ! ( $myWindowsPrincipal.IsInRole( $adminRole ) ) )
{
    Throw 'Script is not being run elevated'
}

try
{
    [string]$workingFolder = (Join-Path -Path $env:temp -ChildPath (([guid]::NewGuid()).Guid))

    # URL the ODT
    [string]$downloadURL = 'https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117'

    # TODO How do we deal with a proxy?

    if( ( Test-Path -Path $workingFolder -ErrorAction SilentlyContinue ) )
    {
        Throw "Random working folder name `"$workingFolder`" already exists"
    }

    $newFolder = New-Item -Path $workingFolder -ItemType Directory
    if( ! $newFolder )
    {
        Throw "Failed to create working folder `"$workingFolder`""
    }
    else
    {
        $createdFolder = $true
    }
    
    $html = Invoke-WebRequest -Uri $downloadURL -UseBasicParsing
    
    if( ! $html -or ! $html.Links )
    {
        Throw "Failed to download from $downloadURL"
    }

    [string[]]$downloadLinks = @( $html.Links | Where-Object { $_.href -Match $odtDownloadRegex } | Select-Object -ExpandProperty href | Sort-Object -Unique )

    if( ! $downloadLinks.Count )
    {
        Throw "No links on page $downloadURL matching regex $odtDownloadRegex"
    }

    if( $downloadLinks.Count -gt 1 )
    {
        Throw "Ambiguous download links on $downloadURL : $($downloadLinks -join ',')"
    }

    [string]$outputFile = Join-Path -Path $workingFolder -ChildPath ($downloadLinks[0] -split '/')[-1]

    Write-Verbose "$(Get-Date -Format G): Starting download from $($downloadLinks[0]) to `"$outputFile`""

    (New-Object System.Net.WebClient).Downloadfile( $downloadLinks[0] , $outputFile )

    if( ! ( Test-Path -Path $outputFile -PathType Leaf -ErrorAction SilentlyContinue ) )
    {
        Throw "Failed to download from $($downloadLinks[0]) to `"$outputFile`""
    }

    # Now check signature to ensure we are running the correct executable
    Unblock-File -Path $outputFile -ErrorAction SilentlyContinue

    $signing = Get-AuthenticodeSignature -FilePath $outputFile -ErrorAction SilentlyContinue
    if( ! $signing )
    {
        Throw "Could not get signing information from `"$outputFile`""
    }
    if( $signing.Status -ne 'Valid' )
    {
        Throw "Certificate status for `"$outputFile`" is $($signing.Status), not `"Valid`""
    }
    if( $signing.SignerCertificate.Subject -notmatch '^CN=Microsoft Corporation,' )
    {
        Throw "`"$outputFile`" is not signed by Microsoft Corporation, found $($signing.SignerCertificate.Subject)"
    }

    # Run silently to extract setup.exe
    $odtInstallProcess = Start-Process -FilePath $outputFile -ArgumentList "/extract:`"$($workingFolder)`" /quiet" -PassThru -Wait -WindowStyle Hidden
    if( ! $odtInstallProcess )
    {
        Throw "Failed to run `"$outputFile`""
    }
    if( ! $odtInstallProcess -or $odtInstallProcess.ExitCode )
    {
        Throw "Bad exit code $($odtInstallProcess.ExitCode) from `"$outputFile`""
    }
    [string]$setupExe = Join-Path -Path $workingFolder -ChildPath 'setup.exe'
    if( ! ( Test-Path -Path $setupExe -PathType Leaf -ErrorAction SilentlyContinue ) )
    {
        Throw "Unable to locate extracted setup.exe after running `"$outputFile`""
    }
    Unblock-File -Path $setupExe -ErrorAction SilentlyContinue

    # Construct XML config file for Office Deployment Kit setup.exe
    [string]$xmlFilePath = Join-Path -Path $workingFolder -ChildPath 'office365.xml'
    
    ## TODO verify bitness is 32 or 64 bit

    # Get the log file in our folder. 
    $configXML -replace '\*\*workingFolder\*\*' , $workingFolder | Out-File -FilePath $xmlFilePath -Force -Encoding ascii
    # Run silently to download media
    
    Write-Verbose "$(Get-Date -Format G): Running `"$setupExe`" to download office media"
    $odtProcess = Start-Process -FilePath $setupExe -ArgumentList "/download $(Split-Path -Path $xmlFilePath -Leaf)" -PassThru -Wait -WorkingDirectory $workingFolder -WindowStyle Hidden

    if( ! $odtProcess )
    {
        Throw "Failed to run `"$setupExe`" to download Office"
    }

    if( $odtProcess.ExitCode )
    {
        Throw "Bad exit code $($odtProcess.ExitCode) from `"$setupExe`" to download Office"
    }

    Write-Verbose -Message "$(Get-Date -Format G): download of Office media to `"$workingFolder`" finished, starting installation"

    ## Install
    $installProcess = Start-Process -FilePath $setupExe -ArgumentList "/configure $(Split-Path -Path $xmlFilePath -Leaf)" -PassThru -Wait -WorkingDirectory $workingFolder -WindowStyle Hidden

    if( ! $installProcess )
    {
        Throw "Failed to run `"$setupExe`" to install Office"
    }

    if( $odtProcess.ExitCode )
    {
        Throw "Bad exit code $($odtProcess.ExitCode) from `"$setupExe`" to install Office"
    }
    
    Write-Verbose -Message "$(Get-Date -Format G): installation of Office media from `"$workingFolder`" finished"
}
catch
{
    Throw $_
}
finally
{
    # Delete temp folder and files
    if( ( Get-Variable -Name createdFolder -ErrorAction SilentlyContinue ) -and $createdFolder -and ( Test-Path -Path $workingFolder -ErrorAction SilentlyContinue ) )
    {
       Remove-Item -Path $workingFolder -Force -Recurse -ErrorAction Continue
    }
}
