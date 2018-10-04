<#
    Run sysprep, get logs and remove apps which it complains about

    2018-10-04 10:03:08, Error                 SYSPRP Package Microsoft.Windows.Photos_2018.18051.21218.0_x64__8wekyb3d8bbwe was installed for a user, but not provisioned for all users. This package will not function properly in the sysprep image.

    @guyrleech 2018

    Modification History:
#>

<#
.SYNOPSIS

Run sysprep and if it fails with package errors in the log file, remove that package and try again

.PARAMETER arguments

The arguments to pass to sysprep

.PARAMETER logfile

The sysprep logfile

.PARAMETER failedAppRegex

A regular expression used to match the log file line containing the application error

.EXAMPLE

& '.\Fix Sysprep Appx errors.ps1' -Verbose -Confirm:$false

Run sysrep with the default parameters (generalize, enter OOBE, shutdown), showing verbose messages but not prompting to confirm before removing any problematic AppX packages

#>

[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='High')]

Param
(   
    [string]$arguments = '/quiet /generalize /oobe /shutdown' ,
    [string]$logFile = (Join-Path ([environment]::getfolderpath('system')) 'Sysprep\Panther\setupact.log') ,
    [string]$failedAppRegex = ',\s+Error\s+SYSPRP Package (.*) was installed for a user, but not provisioned for all users\. This package will not function properly in the sysprep image\.' ,
    [string]$transcriptFile
)

if( ! [string]::IsNullOrEmpty( $transcriptFile ) )
{
    Start-Transcript -Path $transcriptFile
}

try
{
    Import-Module Appx,Dism

    [string]$lastError = $null
    [int]$counter = 1

    While( $true )
    {
        Write-Verbose "$counter : $(Get-Date -Format G)"
        $sysprep = Start-Process -FilePath ( Join-Path  ([environment]::getfolderpath('system')) 'Sysprep\sysprep.exe' ) -ArgumentList $arguments -PassThru -Wait -ErrorAction Stop
        if( ! ( Test-Path -Path $logFile -PathType Leaf -ErrorAction Stop ) )
        {
            Throw "Failed to find sysprep log file $logFile"
        }
        $sysprepError = Get-Content $logFile | Where-Object { $_ -match $failedAppRegex } | Select -Last 1
        if( [string]::IsNullOrEmpty( $sysprepError ) )
        {
            Throw "Could not find log line in $logfile matching regex `"$regex`""
        }
        if( ! [string]::IsNullOrEmpty( $matches[1] ) )
        {
            if( $lastError -eq $Matches[1] )
            {
                Throw "Unable to continue as cannot get rid of error for package $($matches[1])"
            }
            if( $PSCmdlet.ShouldProcess( $matches[1] , 'Remove AppX package' ) )
            {
                Write-Verbose "Removing package $($matches[1]) ..."
                Remove-AppxPackage -Package $Matches[1]
                $lastError = $Matches[1]
            }
            $counter++
        }
        else
        {
            Throw "Could not find package in $logfile line `"$($matches[0])`""
        }
    }
}
catch
{
    Throw $_
}
Finally
{
    if( ! [string]::IsNullOrEmpty( $transcriptFile ) )
    {
        Stop-Transcript
    }
}
