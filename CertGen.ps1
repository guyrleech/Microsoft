#requires -RunAsAdministrator
#requires -version 3

<#
.SYNOPSIS
    Create server certificate for this machine using local CA and optionally install and add to IIS bindings

.PARAMETER computername
    The computer name to use in the certificate subject and Subject Alternative Name fields

.PARAMETER certificateType
    The type of certificate to generate

.PARAMETER certificateLocation
    The certificate store to save the new certificate to

.PARAMETER friendlyName
    The friendly name for the certificate. If not specified the computer name will be used

.PARAMETER prefix
    The prefix for the subject field. Set to $null or empty string to remove the default

.PARAMETER noInstall
    Do not install the new certificate, just leave the file on disk

.PARAMETER webSite
    The web site(s) whose bindings will have the new certificate installed

.PARAMETER protocol
    The protocol for the binding

.PARAMETER keyLength
    The length of the key in the certificate

.PARAMETER exportPrivateKey
    Make the private key of the new certificate exportable

.PARAMETER noDeleteCertificateFile
    Do not delete the certificate file after successful installation

.EXAMPLE
    & . '.\CertGen.ps1' 

    Create and install a certificate with subject and SAN of the computer & domain that the script is run on
    
.EXAMPLE
    & . '.\CertGen.ps1' -website 'Default Web Site'

    Create and install a certificate with subject and SAN of the computer & domain that the script is run on & set for the https binding on the default web site

.EXAMPLE
    & . '.\CertGen.ps1' -friendlyName "Hello World"

    Create a certificate with subject and SAN of the computer that the script is run on & friendly name "Hello World".
    Do not install the generated certificate, the file containing it will be output to the pipeline
  
.EXAMPLE
    & . '.\CertGen.ps1' -website 'Default Web Site','Support' -computerName '*.guyrleech.local' -friendlyName 'Wild Thing' -confirm:$false

    Create and install a wildcard certificate for the guyrleech.local domain & friendly name of "Wild Thing" & set it for the https binding on the support & default web sites
   
.NOTES
    Method based on https://leeejeffries.com/request-an-ssl-certificate-from-a-windows-ca-without-web-en

    Modification History:

    2024/05/09  @guyrleech  Script born out of frustration that IIS mgmt doesn't do SANs !
    2024/09/10  @guyrleech  Removed DNS=127.0.0.1 as recommended by @ronin3510.
                            Change to regex used on certutil as quoting different on Win 11 compared with Server 2016
#>


<#
Copyright © 2024 Guy Leech

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, 
including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

[CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='High')]

Param
(
    [string]$computername ,
    [string]$certificateType = 'CertificateTemplate:webserver' ,
    [string]$certificateLocation = 'Cert:\LocalMachine\My' ,
    [string]$friendlyName ,
    [string]$prefix = 'CN=' ,
    [switch]$noInstall ,
    [switch]$noDeleteCertificateFile ,
    [string[]]$webSite ,
    [string]$protocol = 'https' ,
    [int]$keyLength = 4096 ,
    [switch]$exportPrivateKey
)

if( [string]::IsNullOrEmpty( $computername ) )
{
    $computername = "$($env:COMPUTERNAME).$([System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain().Name)"
}

if( [string]::IsNullOrEmpty( $friendlyName ) )
{
    $friendlyName = $computername
}

[string]$template = @"
[Version] 
Signature="`$Windows NT`$"
[NewRequest] 
Subject = "$prefix$computername" ; For a wildcard use "CN=*.DOMAIN.COM" 
FriendlyName = "$FriendlyName"
; For an empty subject use the following line instead or remove the Subject line entirely 
; Subject = 
Exportable = $($exportPrivateKey.ToString()); Private key is/is not exportable 
KeyLength = $keyLength ; Common key sizes: 512, 1024, 2048, 4096, 8192, 16384 
KeySpec = 1 ; AT_KEYEXCHANGE 
KeyUsage = 0xA0 ; Digital Signature, Key Encipherment 
MachineKeySet = True ; The key belongs to the local computer account 
ProviderName = "Microsoft RSA SChannel Cryptographic Provider" 
ProviderType = 12 
SMIME = FALSE 
RequestType = CMC
[Strings] 
szOID_SUBJECT_ALT_NAME2 = "2.5.29.17" 
szOID_ENHANCED_KEY_USAGE = "2.5.29.37" 
szOID_PKIX_KP_SERVER_AUTH = "1.3.6.1.5.5.7.3.1" 
szOID_PKIX_KP_CLIENT_AUTH = "1.3.6.1.5.5.7.3.2" 
[Extensions] 
%szOID_SUBJECT_ALT_NAME2% = "{text}dns=$computername&IPAddress=127.0.0.1&DNS=localhost" 
%szOID_ENHANCED_KEY_USAGE% = "{text}%szOID_PKIX_KP_SERVER_AUTH%,%szOID_PKIX_KP_CLIENT_AUTH%" 
"@

## TODO Needs to be IPAddress = 127.0.0.1 or you can use DNS=localhost

Write-Verbose -Message "Creating certificate subject `"$computername`" & friendly name `"$friendlyName`""

[string]$randomishBaseName = (New-Guid).Guid.ToString() + "-$pid"
$certificateTemplateFile   = Join-Path -Path $env:temp -ChildPath "$randomishBaseName.inf"
$certificateRequestFile    = Join-Path -Path $env:temp -ChildPath "$randomishBaseName.req"
$certificateFile           = Join-Path -Path $env:temp -ChildPath "$randomishBaseName.cer"

$template | Out-File -FilePath $certificateTemplateFile

if( -Not $WhatIfPreference -and ( -Not $? -or -Not (Test-Path -Path $certificateTemplateFile -PathType Leaf )) )
{
    Throw "Failed to write template to $certificateTemplateFile"
}

Write-Verbose -Message "Template file is $certificateTemplateFile"

$request = certreq.exe -new -q $certificateTemplateFile $certificateRequestFile

if( -Not $? )
{
    Throw "Failed to create request from $certificateTemplateFile to $certificateRequestFile : $request"
}

if( -Not ( Test-Path -Path $certificateRequestFile -PathType Leaf ) )
{
    Throw "Failed to create request file $certificateRequestFile from $certificateTemplateFile : $request"
}

$result = certreq.exe -attrib $certificateType -q -submit $certificateRequestFile $certificateFile 

if( -Not $? )
{
    ## This error can mean that there are more than 1 certificate authorities so we use certutil to show us the default which we grab via regex to pass to certreq again
    ## TODO probably should make it an optional parameter in case certutil gives us the wrong one
    if( Select-String -inputobject $result -SimpleMatch 'No Certification Authorities available' )
    {
        Write-Verbose -Message "Trying to get root CA via certutil as got error: $result"
        $configMatch = certutil.exe | Select-String -Pattern '^\s*Config:\s*[`"]?(.+)["'']?$'
        if( $null -ne $configMatch )
        {
            $rootCA = $null
            $rootCA = $configMatch.Matches.groups[1].value -replace "[`"']"
            if( $null -ne $rootCA )
            {
                Write-Verbose -Message "Submitting certificate request to root CA $rootCA"
                $submission = certreq.exe -attrib $certificateType -q -config $rootCA -submit $certificateRequestFile $certificateFile 
                if( -Not $? )
                {
                    Throw "Failed to create request from $certificateTemplateFile to $certificateRequestFile via Root CA $rootCA : $submission"
                }
            }
            else
            {
                Throw "Failed to find Root CA in certutil output"
            }
        }
        else
        {
            Write-Warning -Message "Failed to parse CA (via Config: line) from certutil.exe output"
        }
    }
    else
    {
        Throw "Failed to create request from $certificateTemplateFile to $certificateRequestFile : $result"
    }
}

if( -Not ( Test-Path -Path $certificateFile -PathType Leaf ) )
{
    Throw "certutil failed to create certificate file $certificateFile"
}

if( -Not $noInstall )
{
    $newCertificate = $null
    if( $PSCmdlet.ShouldProcess( $certificateLocation , "Install certificate file" ) )
    {
        $newCertificate = Import-Certificate -FilePath $certificateFile -CertStoreLocation $certificateLocation
        if( $null -eq $newCertificate )
        {
            Throw "Failed to import certificate from file $certificateFile to store $certificateLocation"
        }
        if( -Not $noDeleteCertificateFile )
        {
            Remove-Item -Path $certificateFile -Force
        }
        $newCertificate ## output

        if( $null -ne $webSite -and $webSite.Count -gt 0 )
        {
            Import-Module -Name IISAdministration -Verbose:$false -Debug:$false
            ForEach( $site in $webSite )
            {
                $binding = $null
                $binding = Get-WebBinding -Name $site -Protocol $protocol
                if( $null -ne $binding )
                {
                    if( $PSCmdlet.ShouldProcess( "Thumbprint $($newCertificate.Thumbprint)" , "Apply to $protocol binding on web site `"$site`"" ))
                    {
                        $binding.RebindSslCertificate( $newCertificate.Thumbprint , (Split-Path -Path $certificateLocation -Leaf) )
                        if( -Not $? )
                        {
                            Write-Warning -Message "Problem binding new certificate to binding for web site `"$site`""
                        }
                    }
                }
                else
                {
                    Write-Warning -Message "Failed to get binding for protocol $protocol for web site `"$site`""
                }
            }
        }
    }
}
else
{
    [pscustomobject]@{
        CertificateFile = $certificateFile
    }
}

Remove-Item -Path $certificateTemplateFile -ErrorAction SilentlyContinue
Remove-Item -Path $certificateRequestFile  -ErrorAction SilentlyContinue
