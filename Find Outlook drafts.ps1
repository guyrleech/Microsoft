<#
    Find recent draft emails and flag up - e.g. you start an email, reboot and forget to finish it

    Guy Leech, 2018

    Modification history:

    28/05/18  GL  Added install and uninstall ability
#>

<#
.SYNOPSIS

See if there are any draft emails in Outlook which have been created in the last x days, default of 7, prompt and open them if selected.

.DESCRIPTION

If the looping option is selected via the -wait option, it will only prompt once per Outlook session so will only prompt a second time
if a new Outlook process starts. This is designed for the scenario where an Outlook draft email is open on screen but Outlook is restarted and that
draft is no longer open so may be forgotten about.
The script can be installed to run at logon via the -install option and removed via the -uninstall option

.PARAMETER withinDays

Find draft emails created within this many days of the current date/time

.PARAMETER waitForOutlook

If Outlook is not running then do not check for recent drafts unless the -wait option is also specified

.PARAMETER wait

Will wait for Outlook to start, if it is not already running, when the script is first launched and then loop infintely checking
to see if a new Outlook instance is launched and if so it will check for recent drafts

.PARAMTER nag

How often to nag the user, by displaying a popup if there are any draft emails found. 
If not specified then nagging will not occur unless a new Outlook process is detected.

.PARAMETER checkInterval

How often, in minutes, that Outlook processes are checked

.PARAMETER install

Name of the autorun entry that will be created in the registry to run this script at logon with the additional options specified. 
Will install for the user running it unless -allusers is specified which requires administrative rights.

.PARAMETER uninstall

Name of the existing autorun entry that will be removed from the registry. 
Will remove just for the user running it unless -allusers is specified which requires administrative rights.

.EXAMPLE

& '.\Find Outlook drafts.ps1'

Check once for any Outlook drafts created in the last 7 days and prompt to open any found.
If Outlook is not running it will be started if the user selects to open the draft emails

.EXAMPLE

& '.\Find Outlook drafts.ps1' -waitForOutlook -withinDays 14

If Outlook is running, check once for any Outlook drafts created in the last 14 days and prompt to open any found.
If Outlook is not running, the script will exit without checking.

.EXAMPLE

& '.\Find Outlook drafts.ps1' -waitForOutlook -withinDays 365 -wait -checkInterval 5

If Outlook is not running, wait for it to start and then check for any Outlook drafts created in the last 365 days and prompt to open any found.
Then every 5 minutes it will check if a new Outlook process has been launched and if so will check for drafts again.
If there are no new Outlook processes then no draft email checks will be performed.

.EXAMPLE

& '.\Find Outlook drafts.ps1' -waitForOutlook -withinDays 365 -wait -checkInterval 5 -nag 10

If Outlook is not running, wait for it to start and then check for any Outlook drafts created in the last 365 days and prompt to open any found.
Then every 5 minutes it will check if a new Outlook process has been launched and if so will check for drafts again.
Every 10 minutes it will prompt the user to open draft emails again.

.EXAMPLE

& '.\Find Outlook drafts.ps1' -waitForOutlook -withinDays 365 -wait -checkInterval 5 -install "Outlook Drafts Checker"

Will run the script at every logon with the following options:
If Outlook is not running, wait for it to start and then check for any Outlook drafts created in the last 365 days and prompt to open any found.
Then every 5 minutes it will check if a new Outlook process has been launched and if so will check for drafts again.
If there are no new Outlook processes then no draft email checks will be performed.

.EXAMPLE

& '.\Find Outlook drafts.ps1' -uninstall "Outlook Drafts Checker"

Remove the logon script entry called "Outlook Drafts Checker"

.NOTES

Run via a logon script with -waitForOutlook and -wait to bring recent draft emails to the user's attention when Outlook is launched or relaunched.
Autorun entries are created in the registry.

#>

[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='High')]

Param
(
    [ValidateRange(1, [int]::MaxValue)]
    [int]$withinDays = 7 ,
    [switch]$waitForOutlook ,
    [switch]$wait ,
    [ValidateRange(1, [int]::MaxValue)]
    [int]$nag ,
    [ValidateRange(1, [int]::MaxValue)]
    [int]$checkInterval = 1 ,
    [string]$install ,
    [string]$uninstall ,
    [switch]$allusers
)

[int]$ERROR_INVALID_PARAMETER = 87

if( $nag -and $nag -lt $checkInterval )
{
    $checkInterval = $nag
}

if( ! [string]::IsNullOrEmpty( $install ) -or ! [string]::IsNullOrEmpty( $uninstall ) )
{
    if( [string]::IsNullOrEmpty( $install ) -and [string]::IsNullOrEmpty( $uninstall ) )
    {
        Write-Error "Must only specify one of -install or -uninstall"
        Exit $ERROR_INVALID_PARAMETER
    }
    [string]$regKey = '\Software\Microsoft\Windows\CurrentVersion\Run'
    if( $allusers )
    {
        $regKey = 'HKLM:' + $regKey
    }
    else
    {
        $regKey = 'HKCU:' + $regKey
    }
    
    [string]$name = $install
    [bool]$installing = $true
    if( [string]::IsNullOrEmpty( $name ) )
    {
        $name = $uninstall
        $installing = $false
    }
    ## See if it exists already
    [string]$existing = $null
    try
    {
        $existing = (Get-ItemProperty -Path $regKey -Name $name -ErrorAction SilentlyContinue).$name
    }
    catch{}

    ## Strip out -install and -allusers parameters, convert single quotes to double quotes
    [string]$run = 'powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File ' + $script:MyInvocation.Line -Replace "'" , "`""  -replace '&', '' -replace '-allusers','' -replace '-install\s+[''"][^''"]*[''"]','' -replace '-install\s+[^\s]*',''
            
    [bool]$carryOn = $true
    if( ! [string]::IsNullOrEmpty( $existing ) )
    {
        if( $installing )
        {
            if( $run -eq $existing )
            {
                Write-Warning "Value `"$name`" already exists as `"$existing`" in $regKey so no change required"
                $carryOn = $false
            }
            elseif( ! $pscmdlet.ShouldProcess( "Value `"$name`" already exists as `"$existing`"" , 'Overwrite' ) )
            {
                $carryOn = $false
            }
        }
        else
        {
            [string]$thisScript = & { $MyInvocation.ScriptName }
            if( $existing.IndexOf( $thisScript ) -lt 0 )
            {
                Write-Error "Script `"$thisScript`" not found in `"$existing`" in $regKey\$uninstall so not removing"
                $carryOn = $false
            }
        }
    }
    elseif( ! $installing )
    {
        Write-Error "Cannot uninstall `"$name`" from $regKey as does not exist"
        $carryOn = $false
    }

    if( $carryOn )
    {
        if( $installing )
        {
            Set-ItemProperty -Path $regKey -Name $name -Value $run -Force
        }
        elseif( $pscmdlet.ShouldProcess( "`"$name`" from $regKey" , 'Remove' ) )
        {
            Remove-ItemProperty -Path $regKey -Name $name -Force
        }
    }
    Exit 0
}

$thisSession = Get-Process -Id $pid | Select -ExpandProperty sessionid
[datetime]$nextNagTime = (Get-Date).AddMinutes( $nag )
[datetime]$lastOutlookProcessCheck = (Get-Date).AddYears( -50 )
[bool]$firstTime = $true

[void] (Add-type -assembly 'Microsoft.Office.Interop.Outlook')
[void] (Add-type -assembly 'Microsoft.VisualBasic')

do
{
    [array]$outlooksNow = @( Get-Process -name 'outlook' -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $thisSession } )
    [bool]$doChecks = $false
    [bool]$outlookWasRunning = ( $outlooksNow -and $outlooksNow.Count )

    if( $waitForOutlook )
    {       
        do
        {
            if( $outlooksNow -and $outlooksNow.Count )
            {
                break
            }
            elseif( $wait )
            {
                Write-Verbose "$(Get-Date -Format G) : no outlook processes found so sleeping for $checkInterval minutes"
                Start-Sleep -Seconds ( $checkInterval * 60 )
            }
            $outlooksNow = @( Get-Process -name 'outlook' -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $thisSession } )
        } while( $wait )

        if(  ! $outlooksNow -or ! $outlooksNow.Count )
        {
            Write-Warning "No outlook processes detected in session $thisSession so aborting"
            exit 1
        }
        elseif( $firstTime )
        {
            Write-Verbose "Outlook already running at start of script $($outlooksNow | Select id,starttime | Out-String)"
            $doChecks = $true ## used to flag up that we should do drafts checks
            $firstTime = $false
        } 
        else
        {
            Write-Verbose "Outlooks found are: $( $outlooksNow | Select Name,Id,StartTime | Out-String )"
        }
    }
    elseif( $firstTime ) ## not been asked to wait so set flag so checking block is invoked
    {
        $doChecks = $true
        $firstTime = $false
    }

    ## see if new outlook processes so we don't warn if they are the same ones as last time
    if( $wait -and $outlooksNow -and $outlooksNow.Count )
    {
        $outlookWasRunning = $true
        if( ! $doChecks )
        {
            $outlooksNow | ForEach-Object `
            {
                $outlookNow = $_
                if( $outlookNow.StartTime -gt $lastOutlookProcessCheck )
                {
                    Write-Verbose "$(Get-Date -Format G) : found outlook id $($outlookNow.Id) launched $(Get-Date -Date $outlookNow.StartTime -Format G) after $(Get-Date -Date $lastOutlookProcessCheck -Format G)"
                    $doChecks = $true
                }
            }
        }
        $lastOutlookProcessCheck = Get-Date
    }

    ## see if we are due to nag the user again
    if( $nag -and (Get-Date) -ge $nextNagTime )
    {
        $doChecks = $true
        $nextNagTime = (Get-Date).AddMinutes( $nag )
        Write-Verbose "$(Get-Date -Format G) : nag time reached, next nag scheduled for $(Get-Date $nextNagTime -Format G)"
    }

    if( $doChecks )
    {
        $outlook = New-Object -ComObject outlook.application -Verbose:$false
        $namespace = $outlook.GetNameSpace("MAPI")

        if( ! $outlook -or ! $namespace )
        {
            Write-Error "Failed to create Outlook objects"
            return 1
        }

        try
        {
            $draftsFolder = $namespace.GetDefaultFolder( [Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderDrafts )

            if( $draftsFolder )
            {
                [datetime]$createdSince = (Get-Date).AddDays(-$withinDays) 
                [string]$strFilter = ( "[LastModificationTime] >= '{0}'" -f (Get-Date -Date $createdSince -Format d ) )
                [array]$recentDrafts = @( $draftsFolder.Items.Restrict($strFilter) )
                if( $recentDrafts -and $recentDrafts.Count )
                {
                    [string]$drafts = "Open these drafts ?`n`n"
                    $recentDrafts | ForEach-Object { $drafts += "$(Get-Date -Date $_.CreationTime -Format G) `"$($_.Subject)`"`n" } 
                    [string]$answer = [Microsoft.VisualBasic.Interaction]::MsgBox( $drafts , 'YesNo,SystemModal,Exclamation' , "Found $($recentDrafts.Count) draft emails created since $(Get-Date -Date $createdSince -Format G)" )
                    if( $answer -eq 'Yes' )
                    {
                        $recentDrafts | ForEach-Object `
                        {
                            $_.Display()
                        }
                    }
                    $lastNagTime = Get-Date
                }
                $draftsFolder = $null
            }
            else
            {
                Write-Warning "Failed to get drafts folder"
            }
        }
        catch
        {
            Write-Warning "Outlook operations failed: $($_.Exception.Message)"
        }
        
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject( $namespace )
        $namespace = $null
        Remove-Variable -Name namespace
        if( ! $outlookWasRunning -and $answer -ne 'Yes' )
        {
            ## double check by ensuring parents are svchost which is when it is called via COM
            [bool]$comLaunched = $false
            Get-WmiObject -class win32_process -filter "name = 'outlook.exe' and sessionid = '$thisSession'" | ForEach-Object `
            {
                $parent = Get-Process -Id $_.ParentProcessId -ErrorAction SilentlyContinue
                if( $parent )
                {
                    if( $parent.Name -eq 'svchost' )
                    {
                        $comLaunched = $true
                    }
                }
            }
            if( $comLaunched )
            {
                Write-Verbose "Quitting Outlook COM object as this script started it"
                $outlook.Quit()
            }
        }
        [void][System.Runtime.Interopservices.Marshal]::FinalReleaseComObject( $outlook ) ## nobody else should have it so ensure we do actually release it
        $outlook = $null
        Remove-Variable -Name Outlook
    }
    if( $wait )
    {
        Write-Verbose "$(Get-Date -Format G) : waiting for $checkInterval minutes before looking for new Outlook instances"
        Start-Sleep -Seconds ( $checkInterval * 60 )
    }
} While( $wait )
