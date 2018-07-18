#requires -version 3.0
<#
    Find all instances of SQL and set firewall exceptions for them

    Use this script at your own risk. The author is absolved of all blame for any undesired effects.

    @guyrleech, (c) 2018
#>

[CmdletBinding()]

Param
(
    [string]$sqlInstance = 'MSSQL%' ,
    [string[]]$protocols = @( 'tcp' , 'udp')
)

Get-CimInstance -ClassName Win32_Service -Filter "name like '$sqlInstance'" | ForEach-Object `
{ 
    [string]$instancePath = $_.PathName -replace '"([^"]*)".*$' , '$1' ## may have -sINSTANCENAME on the end so we strip that out

    if( ! [string]::IsNullOrEmpty( $instancePath ) -and ( Test-Path -Path $instancePath ) )
    {
        ForEach( $protocol in $protocols )
        {
            Write-Output "Creating $protocol firewall exception for $instancePath for instance $($_.DisplayName)"
    
            netsh.exe advfirewall firewall add rule name="$($_.DisplayName) $protocol" protocol=$protocol dir=in action=allow program=$instancePath enable=yes profile=any localip=any localport=any remoteip=any remoteport=any
        }
    }
    else
    {
        Write-Warning "Bad instance path `"$instancePath`" for instance $($_.DisplayName)"
    }
}