#requires -version 3

<#
.SYNOPSIS

Monitor a specific registry key and alert and/or delete specified values or the key when it changes

.PARAMETER rootkey

The root registry key to monitor

.PARAMETER subkey

The sub key of the root registry key to monitor

.PARAMETER key

The whole registry key to monitor

.PARAMETER deleteValues

Delete values which match the regex specified

.PARAMETER deleteKey

Delete the key specified. Does not need to exist before the script is run

.PARAMETER timeout

Number of seconds to wait for registry changes to occur. Default is infinite

.PARAMETER alert

Produce a message box when the registry change occurs

.PARAMETER retries

Number of retries to perform when waiting for the parent registry key to exist

.PARAMETER retryInterval

Wait in milliseconds between retries when waiting for the parent registry key to exist

.PARAMETER logfile

Log file to write to

.PARAMETER username

User key to watch in HKEY_USERS

.EXAMPLE

& '.\reg watcher.ps1' -key 'HKCU:\SOFTWARE\Guy Leech\Event Actioner' -alert

When any changes occur in the given registry key, produce a message box with the time of the change

.EXAMPLE

& '.\reg watcher.ps1' -key 'HKCU:\SOFTWARE\Guy Leech\Event Actioner' -deleteKey

When the given key is created, delete it after prompting for confirmation

.EXAMPLE

& '.\reg watcher.ps1' -key 'HKU:\Bob\SOFTWARE\Guy Leech\Event Actioner' -deleteValues '^History' -confirm:$false

When values are changed/created/deleted in the given name where their name starts with "History", delete the values without prompting for confirmation

.NOTES

Modification History:

    06/03/2021  @guyrleech  First public release
#>

[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='High')]

Param
(
    [Parameter(Mandatory=$true,ParameterSetName='Separate',HelpMessage="The root registry key to monitor (e.g. HKEY_LOCAL_MACHINE)")]
    [ValidateSet('HKLM','HKU','HKCR','HKCU','HKEY_LOCAL_MACHINE','HKEY_CLASSES_ROOT','HKEY_CURRENT_USER','HKEY_USERS')]
    [string]$rootkey = 'HKEY_LOCAL_MACHINE' ,
    [Parameter(Mandatory=$true,ParameterSetName='Separate',HelpMessage="The registry key to monitor (e.g.software\...)")]
    [string]$subkey ,
    [Parameter(Mandatory=$true,ParameterSetName='Combined',HelpMessage="The registry key to monitor (e.g.HKEY_LOCAL_MACHINE\software\...)")]
    [string]$key ,
    [string]$deleteValues ,
    [switch]$deleteKey ,
    [switch]$noInitialDelete ,
    [int]$timeout = -1 ,
    [switch]$alert ,
    [string]$eventName = 'Guys_reg_key_change' ,
    [int]$retries = 60 ,
    [int]$retryInterval = 500 ,
    [string]$logfile ,
    [string]$username ## for HKCU to override %username%
)

Function
Delete-Values
{
    Param
    (
        [Parameter(Mandatory=$true)]
        [string]$key ,
        [Parameter(Mandatory=$true)]
        [string]$values
    )
    
    Get-Item -Path $key -ErrorAction SilentlyContinue| Select-Object -ExpandProperty property | Where-Object { $_ -match $values } | % `
    {
        Write-Verbose "Deleting `"$_`" from `"$key`""
        Remove-ItemProperty -Path $key -Name $_
    }
}

if( ! [string]::IsNullOrEmpty( $logfile ) )
{
    Start-Transcript -Path $logfile
}

Try
{
    if( ! [string]::IsNullOrEmpty( $deleteValues ) -and $deleteKey )
    {
        Write-Warning "-deletekey will triumph over -deletevalues"
    }

    ## if passed a single key then we split into root and subkey since WMI call needs it that way
    if( ! [string]::IsNullOrEmpty( $key ))
    {
        [string[]]$parts = $key -split '\\'
        if( ! $parts -or $parts.Count -lt 2 )
        {
            Throw "Invalid key `"$key`""
        }
        try
        {
            $rootKey = $parts[0] -replace ':' , '' ## in case PoSH format key
        }
        catch
        {
            ## error most likely because $rootKey is invalid thanks to ValidateSet parameter
            Throw "Unknown base key `"$($parts[0])`""
        }
        $subkey = $key -replace ( '^' + $parts[0]) , ''
    }

    [hashtable]$keyConversion = @{
        'HKEY_LOCAL_MACHINE' = 'HKLM'
        'HKEY_USERS' = 'HKU'
        'HKEY_CLASSES_ROOT' = 'HKCR'
    }

    [string]$poshRootKey = switch -regex ( $rootkey)
    {
        'HKEY_LOCAL_MACHINE' 
        {
            'HKLM'
        }
        'HKEY_USERS'
        {
            'HKU'
        }
        'HKEY_CLASSES_ROOT'
        {
            'HKCR'
        }
        'HKEY_CURRENT_USER|HKCU'
        {
            ## WMI call won't take HKCU, since operating in system space effectively, so convert to HKU\Sid if we haven't been passed a username to use 
            [string]$name = $username
            if( [string]::IsNullOrEmpty( $name ) )
            {
                $name = $env:USERNAME
            }
            [string]$sid = $null
            try
            {
                $sid = (New-Object System.Security.Principal.NTAccount($name)).Translate([System.Security.Principal.SecurityIdentifier]).value
            }
            catch
            {
                $sid = $null
            }
            if( [string]::IsNullOrEmpty( $sid ))
            {
                Throw "Unable to get sid for user $name"
            }
            $rootkey = 'HKEY_USERS'
            $subkey = $sid + $subkey
            'HKU'
        }
        'HKU'
        {
            $rootkey = 'HKEY_USERS'
            'HKU'
        }
        'HKLM'
        {
            $rootkey = 'HKEY_LOCAL_MACHINE'
            'HKLM'
        }
        'HKCR'
        {
            $rootkey = 'HKEY_CLASSES_ROOT'
            'HKCR'
        }
    }

    if( [string]::IsNullOrEmpty( $poshRootKey ) )
    {
        Throw "Unknown root key $rootkey"
    }
    elseif( $poshRootKey ) ##-eq 'HKU' -and ( ! [string]::IsNullOrEmpty( $deleteValues) ) -or $deleteKey )
    {
        try
        {
            $provider = Get-PSDrive -Name HKU -ErrorAction SilentlyContinue
        }
        catch
        {
            $provider = $null
        }

        if( ! $provider )
        {
            if( ! ( New-PSDrive -Name HKU -PSProvider Registry -Root Registry::HKEY_USERS ) )
            {
                Throw "Failed to create PS drive for HKU"
            }
        }
    }

    [string]$poshKey = $null
    [string]$keyToDelete = $null
    $formsLoaded = $null

    if( $deleteKey )
    {
        ## We need to monitor the parent for changes and then delete this key if it is then present
        $keyToDelete = Split-Path $subkey -Leaf ## need to add this to PoshKey after it is constructed
        $subkey = Split-Path $subkey -Parent
    }

    ## must escape backslashes for WMI except one at the start
    if( $subkey[0] -eq '\' )
    {
        $poshKey = "$($poshRootKey):$subkey"
        $subkey = $subkey.Substring(1,$subkey.Length - 1) -replace '\\', '\\' ## get rid of leading \ as will confuse WMI
    }
    else ## no leading backslash
    {
        $poshKey = "$($poshRootKey):\$subkey"
        $subkey =  $subkey -replace '\\', '\\' 
    }

    if( $deleteKey )
    {
        $keyToDelete = Join-Path -Path $poshKey -ChildPath $keyToDelete
    }

    [int]$retry = 0

    Write-Verbose "Starting monitoring $poshKey @ $(Get-Date) as $($env:username)"

    [string]$title = Split-Path -Path ( & { $MyInvocation.ScriptName } ) -Leaf

    while( $true )
    {
        if( ! $noInitialDelete )
        {
            if( ! [string]::IsNullOrEmpty( $keyToDelete ) )
            {
                if( Test-Path -Path $keyToDelete )
                {
                    if( $pscmdlet.ShouldProcess( $keyToDelete , "Initial delete of key" ) )
                    {
                        Remove-Item -Path $keyToDelete -Recurse -Force
                    }
                }
            }
            elseif( ! [string]::IsNullOrEmpty( $deleteValues ) -and $pscmdlet.ShouldProcess( $poshKey , "Initial delete of values matching `"$deleteValues`"" ) )
            {
                Delete-Values -key $poshKey -values $deleteValues 
            }
        }

        Unregister-Event -SourceIdentifier $eventName -ErrorAction SilentlyContinue ## Just in case we exited uncleanly
        [string]$changeEvent = 'RegistryKeyChangeEvent'
        [string]$pathType = 'KeyPath'
        if( $deleteKey )
        {
            $changeEvent ='RegistryTreeChangeEvent'
            $pathType = 'RootPath'
        }

        [bool]$registered = $false
        Try
        {
            Register-CimIndicationEvent -Namespace 'root\default' -Query "SELECT * FROM $changeEvent WHERE Hive='$rootkey' AND $pathType='$subkey'" -SourceIdentifier $eventName
            $registered = $?
        }
        Catch
        {
        }

        if( $registered )
        {
            $retry = 0
            ## For MessageBox()
            if( $alert -and ! $formsLoaded )
            {
                $formsLoaded = [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
            }
            [bool]$keyPresent = $true
            While( $keyPresent )
            {
                ## We waited for any event so ensure it is ours - it may be one we've generated by deleting values
                if( ($eventRaised = Wait-Event -Timeout $timeout) -And $eventRaised.Sender.ToString() -eq 'Microsoft.Management.Infrastructure.CimCmdlets.CimIndicationWatcher' -And $eventRaised.SourceIdentifier -match ( '^' + $eventName + '$' ) )
                {
                    [string]$keyChanged = '{0}:\{1}' -f $keyConversion[ $eventRaised.SourceEventArgs.NewEvent.Hive ] , ( $eventRaised.SourceEventArgs.NewEvent | Select-Object -ExpandProperty *Path)
                    [datetime]$dateChanged = [datetime]::FromFileTime( $eventRaised.SourceEventArgs.NewEvent.TIME_CREATED )
                    [string]$message = "`"$keyChanged`" changed at $(Get-Date -Date $dateChanged -Format G)"
                    Write-Verbose -Message $message
                    ## see if key still exists since if a delete action we need to re-register our event handler
                    if( ! ( Test-Path -Path $keyChanged ) )
                    {
                        Write-Warning "`"$keyChanged`" not present"
                        $keyPresent = $false
                    }
                    elseif( ! [string]::IsNullOrEmpty( $keytoDelete ) )
                    {
                        if( ( Test-Path $keyToDelete ) -and $pscmdlet.ShouldProcess( $keyToDelete , "Delete key" ) )
                        {
                            Remove-Item -Path $keyToDelete -Recurse -Force
                        }
                    }
                    elseif( ! [string]::IsNullOrEmpty( $deleteValues ) -and $pscmdlet.ShouldProcess( $poshKey , "Delete values matching `"$deleteValues`"" ) )
                    {
                        Delete-Values -key $poshKey -values $deleteValues 
                    }
                    if( $alert )
                    {
                        $null = [System.Windows.Forms.MessageBox]::Show( $message , $title , 0 , [System.Windows.Forms.MessageBoxIcon]::Exclamation)
                    }
                }
                elseif( ! $eventRaised )
                {
                    Write-Verbose "Timeout after $timeout seconds @ $(Get-Date)"
                }
                else
                {
                    Write-Warning "Unexpected event received @$(Get-Date):`n$eventRaised"
                }
                $eventRaised | Remove-Event -ErrorAction SilentlyContinue
            }
        }
        else
        {
            if( $retry++ -lt $retries )
            {
                ## If event was a key deletion then we need to wait and see if it gets recreated
                Write-Warning "Failed to set notification on key `"$rootkey\$subkey`" will retry $retry / $retries in $retryInterval milliseconds"
                Start-Sleep -Milliseconds $retryInterval
            }
            else
            {
                Write-Error "Failed to set notification on key $rootkey\$subkey"
                break
            }
        }

        Unregister-Event $eventName -ErrorAction SilentlyContinue
    }

    Write-Verbose "Finished monitoring $poshKey @ $(Get-Date)"
}
Catch
{
    Throw $_
}
Finally
{
    if( ! [string]::IsNullOrEmpty( $logfile ) )
    {
        Stop-Transcript
    }
}
