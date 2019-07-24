#requires -version 3
<#
    Show process durations via security event logs when process creation/termination auditing is enabled

    @guyrleech 2019

    Modification History:

    10/05/2019  GRL   Added subject logon id to grid view output
                      Added enable/disable of process creatiuon and termination auditing

    12/05/2019  GRL   Remove process termination events from array for speed increase
                      Enable cmd line auditing when -enable specified

    13/05/2019  GRL   Added creation and modification times of executables & executable summary option
                      Added multiple computer capability

    23/05/2019  GRL   Added option for processing saved event log files
                      Added option for having no exe file details

    13/06/2019  GRL   Moved end event cache to hash table for considerable speed improvement

    14/06/2019  GRL   Fixed bug giving negative durations. Filtering on process stop collection building

    24/07/2019  GRL   Added time frames via logon sessions from LSASS
                      Added -listSessions to just show LSASS sessions retrieved
                      Changed -logon to -logonTimes and -boot to -bootTimes
#>

<#
.SYNOPSIS

Retrieve process start events from the security event log, try and find corresponding process exit and optionally also show start time relative to that user's logon and/or computer boot

.PARAMETER usernames

Only include processes for users which match this regular expression

.PARAMETER processName

Only include processes which match this regular expression

.PARAMETER eventLog

The path to an event log file containing saved events

.PARAMETER start

Only retrieve processes started after this date/time

.PARAMETER end

Only retrieve processes started before this date/time

.PARAMETER last

Show processes started in the preceding period where 's' is seconds, 'm' is minutes, 'h' is hours, 'd' is days, 'w' is weeks and 'y' is years so 12h will retrieve all in the last 12 hours.

.PARAMETER listSessions

Just list the interactive logon sessions found on each computer

.PARAMETER logonTimes

Include the logon time and the time since logon for the process creation for the logon session this process belongs to

.PARAMETER bootTimes

Include the boot time and the time since boot for the process creation 

.PARAMETER enable

Enable process creation and termination auditing

.PARAMETER disable

Disable process creation and termination auditing

.PARAMETER outputFile

Write the results to the specified csv file

.PARAMETER noGridview

Output the results to the pipeline rather than a grid view

.PARAMETER excludeSystem

Do not include processes run by the system account

.PARAMETER noFileInfo

Do not include exe file information

.PARAMETER summary

Show a summary by executable including number of executions and file details

.PARAMETER computers

A comma separated list of computers to run query the security event logs of. Firewall must allow Remote Eventlog.

.EXAMPLE

& '.\Get Process Durations.ps1' -last 2d -logon -username billybob -boot

Find all process creations and corresponding terminations for the user billybob in the last 2 days, calculate the start time relative to logon for that user's session and relative to the the boot time and display in a grid view

.EXAMPLE

& '.\Get Process Durations.ps1' -enable

Enable process creation and termination auditing

.EXAMPLE

& '.\Get Process Durations.ps1' -listSessions

Show all interactive logon sessions from LSASS since last boot

.EXAMPLE

& '.\Get Process Durations.ps1' -summary -start "08:00" -computers xa1,xa2

Find all process creations on computers xa1 and xa2 since 0800 today and produce a grid view summarising per unique executable

.NOTES

Must have process creation and process termination auditing enabled although this script can enable/disable if required

If process command line auditing is enabled then the command line will be included. See https://docs.microsoft.com/en-us/windows-server/identity/ad-ds/manage/component-updates/command-line-process-auditing

Enable/Disable of auditing will not work in non-English locales

When run for multiple computers, the file information is only taken from the first instance of that executable encountered so not compared across computers
#>

[CmdletBinding()]

Param
(
    [string]$username ,
    [string]$processName ,
    [string]$start ,
    [string]$end ,
    [string]$last ,
    [string]$eventLog ,
    [string]$logonOf ,
    [int]$beforeSeconds = 30 ,
    [int]$afterSeconds = 120 ,
    [string]$logonAround ,
    [int]$skipLogons = 0 ,
    [switch]$listSessions ,
    [string[]]$computers = @( $env:COMPUTERNAME ) ,
    [switch]$logonTimes ,
    [switch]$bootTimes ,
    [switch]$enable ,
    [switch]$disable ,
    [switch]$summary ,
    [string]$outputFile ,
    [switch]$nogridview ,
    [switch]$noFileInfo ,
    [switch]$excludeSystem 
)

[string[]]$startPropertiesMap = @(
    'SubjectUserSid' , 
    'SubjectUserName' ,
    'SubjectDomainName' ,
    'SubjectLogonId' ,
    'NewProcessId' ,
    'NewProcessName' ,
    'TokenElevationType' ,
    'ProcessId' ,
    'CommandLine' ,
    'TargetUserSid' ,
    'TargetUserName' ,
    'TargetDomainName' ,
    'TargetLogonId' ,
    'ParentProcessName'
)

Set-Variable -Name 'endSubjectUserSid' -Value 0 -Option ReadOnly -ErrorAction SilentlyContinue
Set-Variable -Name 'endSubjectUserName' -Value 1 -Option ReadOnly -ErrorAction SilentlyContinue
Set-Variable -Name 'endSubjectDomainName' -Value 2 -Option ReadOnly -ErrorAction SilentlyContinue
Set-Variable -Name 'endSubjectLogonId' -Value 3 -Option ReadOnly -ErrorAction SilentlyContinue
Set-Variable -Name 'endStatus' -Value 4 -Option ReadOnly -ErrorAction SilentlyContinue
Set-Variable -Name 'endProcessId' -Value 5 -Option ReadOnly -ErrorAction SilentlyContinue
Set-Variable -Name 'endProcessName' -Value 6 -Option ReadOnly -ErrorAction SilentlyContinue

[hashtable]$auditingGuids = @{
    'Process Creation'    = '{0CCE922C-69AE-11D9-BED3-505054503030}'
    'Process Termination' = '{0CCE922C-69AE-11D9-BED3-505054503030}' }

## https://www.codeproject.com/Articles/18179/Using-the-Local-Security-Authority-to-Enumerate-Us
$LSADefinitions = @'
    [DllImport("secur32.dll", SetLastError = false)]
    public static extern uint LsaFreeReturnBuffer(IntPtr buffer);

    [DllImport("Secur32.dll", SetLastError = false)]
    public static extern uint LsaEnumerateLogonSessions
            (out UInt64 LogonSessionCount, out IntPtr LogonSessionList);

    [DllImport("Secur32.dll", SetLastError = false)]
    public static extern uint LsaGetLogonSessionData(IntPtr luid, 
        out IntPtr ppLogonSessionData);

    [StructLayout(LayoutKind.Sequential)]
    public struct LSA_UNICODE_STRING
    {
        public UInt16 Length;
        public UInt16 MaximumLength;
        public IntPtr buffer;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct LUID
    {
        public UInt32 LowPart;
        public UInt32 HighPart;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct SECURITY_LOGON_SESSION_DATA
    {
        public UInt32 Size;
        public LUID LoginID;
        public LSA_UNICODE_STRING Username;
        public LSA_UNICODE_STRING LoginDomain;
        public LSA_UNICODE_STRING AuthenticationPackage;
        public UInt32 LogonType;
        public UInt32 Session;
        public IntPtr PSiD;
        public UInt64 LoginTime;
        public LSA_UNICODE_STRING LogonServer;
        public LSA_UNICODE_STRING DnsDomainName;
        public LSA_UNICODE_STRING Upn;
    }

    public enum SECURITY_LOGON_TYPE : uint
    {
        Interactive = 2,        //The security principal is logging on 
                                //interactively.
        Network,                //The security principal is logging using a 
                                //network.
        Batch,                  //The logon is for a batch process.
        Service,                //The logon is for a service account.
        Proxy,                  //Not supported.
        Unlock,                 //The logon is an attempt to unlock a workstation.
        NetworkCleartext,       //The logon is a network logon with cleartext 
                                //credentials.
        NewCredentials,         //Allows the caller to clone its current token and
                                //specify new credentials for outbound connections.
        RemoteInteractive,      //A terminal server session that is both remote 
                                //and interactive.
        CachedInteractive,      //Attempt to use the cached credentials without 
                                //going out across the network.
        CachedRemoteInteractive,// Same as RemoteInteractive, except used 
                                // internally for auditing purposes.
        CachedUnlock            // The logon is an attempt to unlock a workstation.
    }
'@

<#
Function Load-GUI( $inputXml )
{
    $form = $NULL
    [xml]$XAML =  $inputXML -replace 'mc:Ignorable="d"' , '' -replace 'x:N' ,'N'  -replace '^<Win.*' , '<Window'
    $reader = New-Object Xml.XmlNodeReader $xaml

    try
    {
        $form = [Windows.Markup.XamlReader]::Load( $reader )
    }
    catch
    {
        Write-Error -Message "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .NET is installed.`n$_"
        return $null
    }
 
    if( $form )
    {
        $xaml.SelectNodes('//*[@Name]') | ForEach-Object `
        {
            Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -Scope Global
        }
    }

    return $form
}
#>

Function Get-AuditSetting
{
    [CmdletBinding()]
    Param
    (
        [string]$GUID
    )
    [string[]]$fields = ( auditpol.exe /get /subcategory:"$GUID" /r | Select-Object -Skip 1 ) -split ',' ## Don't use ConvertFrom-CSV as makes it harder to get the column we want
    if( $fields -and $fields.Count -ge 6 )
    {
        ## Machine Name,Policy Target,Subcategory,Subcategory GUID,Inclusion Setting,Exclusion Setting
        ## DESKTOP2,System,Process Termination,{0CCE922C-69AE-11D9-BED3-505054503030},No Auditing,
        $fields[5] ## get a blank field at the start
    }
    else
    {
        Write-Warning "Unable to determine audit setting"
    }
}

[datetime]$startTime = Get-Date

if( $enable -and $disable )
{
    Throw 'Cannot enable and disable in same call'
}

if( $summary -and $noFileInfo )
{
    Throw 'Cannot specify -noFileInfo with -summary'
}

if( $enable -or $disable )
{
    [hashtable]$requiredAuditEvents = @{
        'Process Creation'    = '0cce922b-69ae-11d9-bed3-505054503030'
        'Process Termination' = '0cce922c-69ae-11d9-bed3-505054503030'
    }

    [string]$state = $(if( $enable ) { 'Enable' } else { 'Disable' })

    [int]$errors = 0

    ForEach( $requiredAuditEvent in $requiredAuditEvents.GetEnumerator() )
    {
        $process = Start-Process -FilePath auditpol.exe -ArgumentList "/set /subcategory:{$($requiredAuditEvent.Value)} /success:$state" -Wait -WindowStyle Hidden -PassThru
        if( ! $process -or $process.ExitCode )
        {
            Write-Error "Error running auditpol.exe to set $($requiredAuditEvent.Name) auditing to $state - error $($process|Select-Object -ExpandProperty ExitCode)"
            $errors++
        }
    }

    if( $enable )
    {
        [void](New-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\Audit' -Name 'ProcessCreationIncludeCmdLine_Enabled' -Value 1 -PropertyType 'DWord' -Force)
    }

    Exit $errors
}

if( $PSBoundParameters[ 'last' ] -and ( $PSBoundParameters[ 'start' ] -or $PSBoundParameters[ 'end' ] ) )
{
    Throw "Cannot use -last when -start or -end are also specified"
}

[hashtable]$startEventFilter = @{
    'Id' = 4688
}

[int]$secondsAgo = 0

if( ! [string]::IsNullOrEmpty( $last ) )
{
    ## see what last character is as will tell us what units to work with
    [int]$multiplier = 0
    switch( $last[-1] )
    {
        "s" { $multiplier = 1 }
        "m" { $multiplier = 60 }
        "h" { $multiplier = 3600 }
        "d" { $multiplier = 86400 }
        "w" { $multiplier = 86400 * 7 }
        "y" { $multiplier = 86400 * 365 }
        default { Throw "Unknown multiplier `"$($last[-1])`"" }
    }
    $endDate = Get-Date
    if( $last.Length -le 1 )
    {
        $secondsAgo = $multiplier
    }
    else
    {
        $secondsAgo = ( ( $last.Substring( 0 , $last.Length - 1 ) -as [decimal] ) * $multiplier )
    }

    $startDate = $endDate.AddSeconds( -$secondsAgo )
    $startEventFilter.Add( 'StartTime' , $startDate )
    ## if using event log file so -last is relative to the latest event in the file so -1h means 0700 if latest event is 0800 but we'll calculate this when we process that computer (which is probably local anyway) and change STartTime
}

if( $PSBoundParameters[ 'logonOf' ] )
{
    if( $computers -and ( $computers.Count -gt 1 -or ( $computers[0] -ne '.' -and $computers[0] -ne $env:COMPUTERNAME ) ) )
    {
        Throw "Cannot use -logonOf with -computers"
    }

    if( ! ( ([System.Management.Automation.PSTypeName]'Win32.Secure32').Type ) )
    {
        Add-Type -MemberDefinition $LSADefinitions -Name 'Secure32' -Namespace 'Win32' -UsingNamespace System.Text -Debug:$false
    }

    $count = [UInt64]0
    $luidPtr = [IntPtr]::Zero

    [uint64]$ntStatus = [Win32.Secure32]::LsaEnumerateLogonSessions( [ref]$count , [ref]$luidPtr )

    if( $ntStatus )
    {
        Write-Error "LsaEnumerateLogonSessions failed with error $ntStatus"
    }
    elseif( ! $count )
    {
        Write-Error "No sessions returned by LsaEnumerateLogonSessions"
    }
    elseif( $luidPtr -eq [IntPtr]::Zero )
    {
        Write-Error "No buffer returned by LsaEnumerateLogonSessions"
    }
    else
    {   
        Write-Debug "$count sessions retrieved from LSASS"
        [IntPtr]$iter = $luidPtr
        $earliestSession = $null
        [array]$lsaSessions = @( For ([uint64]$i = 0; $i -lt $count; $i++)
        {
            $sessionData = [IntPtr]::Zero
            $ntStatus = [Win32.Secure32]::LsaGetLogonSessionData( $iter , [ref]$sessionData )

            if( ! $ntStatus -and $sessionData -ne [IntPtr]::Zero )
            {
                $data = [System.Runtime.InteropServices.Marshal]::PtrToStructure( $sessionData , [type][Win32.Secure32+SECURITY_LOGON_SESSION_DATA] )

                if ($data.PSiD -ne [IntPtr]::Zero)
                {
                    $sid = New-Object -TypeName System.Security.Principal.SecurityIdentifier -ArgumentList $Data.PSiD

                    #extract some useful information from the session data struct
                    [datetime]$loginTime = [datetime]::FromFileTime( $data.LoginTime )
                    $thisUser = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($data.Username.buffer) #get the account name
                    $thisDomain = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($data.LoginDomain.buffer) #get the domain name
                    try
                    { 
                        $secType = [Win32.Secure32+SECURITY_LOGON_TYPE]$data.LogonType
                    }
                    catch
                    {
                        $secType = 'Unknown'
                    }

                    if( ! $earliestSession -or $loginTime -lt $earliestSession )
                    {
                        $earliestSession = $loginTime
                    }
                    if( $secType -match 'Interactive' )
                    {
                        $authPackage = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($data.AuthenticationPackage.buffer) #get the authentication package
                        $session = $data.Session # get the session number
                        $logonServer = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($data.LogonServer.buffer) #get the logon server
                        $DnsDomainName = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($data.DnsDomainName.buffer) #get the DNS Domain Name
                        $upn = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($data.upn.buffer) #get the User Principal Name

                        [pscustomobject]@{
                            'Sid' = $sid
                            'Username' = $thisUser
                            'Domain' = $thisDomain
                            'Session' = $session
                            'LoginId' = [uint64]( $loginID = [Int64]("0x{0:x8}{1:x8}" -f $data.LoginID.HighPart , $data.LoginID.LowPart) )
                            'LogonServer' = $logonServer
                            'DnsDomainName' = $DnsDomainName
                            'UPN' = $upn
                            'AuthPackage' = $authPackage
                            'SecurityType' = $secType
                            'Type' = $data.LogonType
                            'LoginTime' = [datetime]$loginTime
                        }
                    }
                }
                [void][Win32.Secure32]::LsaFreeReturnBuffer( $sessionData )
                $sessionData = [IntPtr]::Zero
            }
            $iter = $iter.ToInt64() + [System.Runtime.InteropServices.Marshal]::SizeOf([type][Win32.Secure32+LUID])  # move to next pointer
        }) | Sort-Object -Descending -Property 'LoginTime'

        [void]([Win32.Secure32]::LsaFreeReturnBuffer( $luidPtr ))
        $luidPtr = [IntPtr]::Zero

        Write-Verbose "Found $(if( $lsaSessions ) { $lsaSessions.Count } else { 0 }) LSA sessions, earliest session $(if( $earliestSession ) { Get-Date $earliestSession -Format G } else { 'never' })"

        ## Now find the requested session
        $closest = $null
        if( $PSBoundParameters[ 'logonAround' ] )
        {
            [datetime]$targetLogonTime = Get-Date -Date $logonAround -ErrorAction Stop
        }
        [int]$logonCount = 0

        ForEach( $lsaSession in $lsaSessions )
        {
            if( $lsaSession.Username -eq $logonOf )
            {
                $logonCount++
                if( $PSBoundParameters[ 'logonAround' ] )
                {        
                    if( $closest )
                    {
                        $thisCloseness = [math]::Abs( ( New-TimeSpan -Start $lsaSession.LoginTime -End $targetLogonTime ).TotalSeconds )
                        $otherCloseness = [math]::Abs( ( New-TimeSpan -Start $closest.LoginTime -End $targetLogonTime ).TotalSeconds )
                        if( $thisCloseness -lt $otherCloseness )
                        {
                            $closest = $lsaSession
                        }
                    }
                    elseif( $logonCount -gt $skipLogons )
                    {
                        $closest = $lsaSession
                    }
                }
                elseif( $logonCount -gt $skipLogons ) ## get the last logon for the user
                {
                    $closest = $lsaSession
                    break
                }
            }
        }

        if( $closest )
        {
            Write-Verbose "Using logon for $($closest.Domain)\$($closest.Username) at $(Get-Date -Date $closest.LoginTime -Format G)"
            $startEventFilter.Add( 'StartTime' , $closest.LoginTime.AddSeconds( -$beforeSeconds ) )
            $startEventFilter.Add( 'EndTime' , $closest.LoginTime.AddSeconds( $afterSeconds ) )
        }
        else
        {
            [datetime]$bootTime = Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object -ExpandProperty LastBootupTime
            Throw "Unable to find a logon for $logonOf in $(if( $lsaSessions ) { $lsaSessions.Count } else { 0 }) LSA sessions, earliest session $(if( $earliestSession ) { Get-Date $earliestSession -Format G } else { 'never' }), boot at $(Get-Date -Date $bootTime -Format G)"
        }
    }
}

if( $PSBoundParameters[ 'start' ] )
{
    $startEventFilter.Add( 'StartTime' , (Get-Date -Date $start ))
}

if( $PSBoundParameters[ 'end' ] )
{
    $startEventFilter.Add( 'EndTime' , (Get-Date -Date $end ))
}

[bool]$differentUserName = $false

[int]$counter = 0
[int]$index = 0
[hashtable]$fileProperties = @{}
[hashtable]$allSessions = @{}

if( $PSBoundParameters[ 'eventLog' ] )
{
    $startEventFilter.Add( 'Path' , $eventLog )
}
else
{
    $startEventFilter.Add( 'LogName', 'Security' )
}

[array]$processes = @( ForEach( $computer in $computers )
{
    if( $computer -eq '.' )
    {
        $computer = $env:COMPUTERNAME
    }

    $counter++
    Write-Verbose "Checking $counter / $($computers.Count ) : $computer"
    
    [string]$machineAccount = $computer + '$'

    [hashtable]$remoteParam = @{}
    if( $computer -ne '.' -and $computer -ne $env:COMPUTERNAME )
    {
        $remoteParam.Add( 'ComputerName' , $computer )
    }
    
    [hashtable]$systemAccounts = @{}

    Get-CimInstance @remoteParam -ClassName win32_SystemAccount | ForEach-Object `
    {
        $systemAccounts.Add( $_.SID , $_.Name )
    }

    ## If using event log file and -last then we need to get the date of the newest event as -last will be relative to that
    if( $PSBoundParameters[ 'last' ] -and $PSBoundParameters[ 'eventLog' ] )
    {
        ## Remove start time from hash table
        $startEventFilter.Remove( 'StartTime' )
        $latestEventHere = Get-WinEvent @remoteParam -FilterHashtable $startEventFilter -ErrorAction SilentlyContinue -MaxEvents 1
        if( $latestEventHere )
        {
            $startEventFilter.Add( 'StartTime' , $latestEventHere.TimeCreated.AddSeconds( - $secondsAgo ) )
        }
    }

    ## Get Oldest event before we filter on date so can report oldest
    $earliestEvent = $null
    $earliestEventHere = Get-WinEvent @remoteParam -FilterHashtable $startEventFilter -Oldest -ErrorAction SilentlyContinue -MaxEvents 1
    if( ! $earliestEvent -or $earliestEventHere -lt $earliestEvent )
    {
        $earliestEvent = $earliestEventHere
    }
    
    [hashtable]$logons = @{}
    ## get logons so we can cross reference to the id of the logon
    Get-WmiObject @remoteParam win32_logonsession -Filter "LogonType='10' or LogonType='12' or LogonType='2' or LogonType='11'" | ForEach-Object `
    {
        $session = $_
        [array]$users = @( Get-WmiObject @remoteParam win32_loggedonuser -filter "Dependent = '\\\\.\\root\\cimv2:Win32_LogonSession.LogonId=`"$($session.LogonId)`"'" | ForEach-Object `
        {
            if( $_.Antecedent -match 'Domain="(.*)",Name="(.*)"$' ` )
            {
                [pscustomobject]@{ 'LogonTime' = (([WMI] '').ConvertToDateTime( $session.StartTime )) ; 'Domain' = $Matches[1] ; 'UserName' = $Matches[2] ; 'Computer' = $computer }
            }
            else
            {
                Write-Warning "Unexpected antecedent format `"$($_.Antecedent)`""
            }
        })
        if( $users -and $users.Count )
        {
            $logons.Add( $session.LogonId , $users )
        }
    }

    if( $listSessions -and $logons.Count )
    {
        $allSessions += $logons
    }
    else
    {
        [hashtable]$stopEventFilter = $startEventFilter.Clone()
        $stopEventFilter[ 'Id' ] = 4689
        $eventError = $null
        $error.Clear()
        [int]$multiplePids = 0

        [hashtable]$endEvents = @{}
        Get-WinEvent @remoteParam -FilterHashtable $stopEventFilter -Oldest -ErrorAction SilentlyContinue | ForEach-Object `
        {
            $event = $_
            if( ( ! $username -or $event.Properties[ $endSubjectUserName ].Value -match $username ) -and ( ! $processName -or $event.Properties[$endProcessName].value -match $processName ) )
            {
                try
                {
                    $endEvents.Add( $event.Properties[ $endProcessId ].Value -as [int] , $event )
                }
                catch
                {
                    ## already got it so will need to have an array so we can look for the right start time and or user
                    $existing = $endEvents[ [int]$event.Properties[ $endProcessId ].Value ]
                    if( $existing -is [System.Collections.ArrayList] )
                    {
                        [void]($endEvents[ $event.Properties[ $endProcessId ].Value -as [int] ]).Add( $event ) ## appends
                        $multiplePids++
                    }
                    elseif( $existing )
                    {
                        $endEvents[ [int]$event.Properties[ $endProcessId ].Value ] = [System.Collections.ArrayList]@( $existing , $event ) ## oldest first
                    }
                    else
                    {
                        Throw $_
                    }
                }
            }
        }

        Write-Verbose "Got $($endEvents.Count) + $multiplePids end events from $computer"

        [hashtable]$runningProcesses = @{}
        if( $remoteParam.Count )
        {
            $processes = @( Invoke-Command @remoteParam -ScriptBlock { Get-Process -IncludeUserName } )
        }
        else
        {
            $processes = @( Get-Process -IncludeUserName )
        }
    
        $bootTime = $null

        ## Don't cache if saved event log not for this machine 
        if( ( $earliestEventHere -and ( $earliestEventHere.MachineName -eq $computer -or ( $earliestEventHere.MachineName -split '\.' )[0] -eq $computer ) ) )
        {
            ## get the boot time so we can add a column relative to it and also check a process start is after latest boot otherwise LSASS won't have login details for it
            $bootTime = Get-CimInstance @remoteParam -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue | Select-Object -ExpandProperty LastBootupTime
    
            ForEach( $process in $processes )
            {
                if( ! $bootTime -or $process.StartTime -ge $bootTime )
                {
                    $runningProcesses.Add( [int]$process.Id , $process )
                }
            }
        }

        ## Find all process starts then we'll look for the corresponding stops
        Get-WinEvent @remoteParam -FilterHashtable $startEventFilter -Oldest -ErrorAction SilentlyContinue -ErrorVariable 'eventError'  | ForEach-Object `
        {
            $event = $_
            if( ( ! $username -or $event.Properties[ 1 ].Value -match $username ) -and ( ! $excludeSystem -or ( $event.Properties[ 1 ].Value -ne $machineAccount -and $event.Properties[ 1 ].Value -ne '-' )) -and ( ! $processName -or $event.Properties[5].value -match $processName ) )
            {
                [hashtable]$started = @{ 'Start' = $event.TimeCreated ; 'Computer' = $computer }
                For( $index = 0 ; $index -lt [math]::Min( $startPropertiesMap.Count , $event.Properties.Count ) ; $index++ )
                {
                    $started.Add( $startPropertiesMap[ $index ] , $event.Properties[ $index ].value )
                }
                if( $started[ 'SubjectUserName' ] -eq '-' )
                {
                    $started.Set_Item( 'SubjectUserName' , $systemAccounts[ ($event.Properties[ 0 ].Value | Select-Object -ExpandProperty Value) ] )
                    $started.Set_Item( 'SubjectDomainName' , $env:COMPUTERNAME )
                }

                ## now find corresponding termination event
                $terminate = $endEvents[ [int]$started.NewProcessId ]
                if( $terminate -is [System.Collections.ArrayList] -and $terminate.Count )
                {
                    ## need to find the right event as have multiple pids but oldest first so pick the first one after the time we need
                    $thisTerminate = $null
                    $index = 0
                    do
                    {
                        try
                        {
                            if( $terminate[$index].TimeCreated -ge $event.TimeCreated )
                            {
                                $thisTerminate = $terminate[$index]
                            }
                        }
                        catch
                        {
                            Write-Error $_
                        }
                    } while( ! $thisTerminate -and (++$index) -lt $terminate.Count )

                    if( $thisTerminate )
                    {
                        $terminate.RemoveAt( $index )
                        $terminate = $thisTerminate
                    }
                    else
                    {
                        $terminate = $null
                    }
                }
                elseif( $terminate -and $terminate.TimeCreated -lt $event.TimeCreated ) ## This is not the event you are looking for (because it is prior to the launch)
                {
                    $terminate = $null
                }
            
                if( $terminate )
                {
                    $started += @{
                        'Exit Code' = $terminate.Properties[ $endStatus ].value
                        'End' = $terminate.TimeCreated
                        'Duration' = (New-TimeSpan -Start $event.TimeCreated -End $terminate.TimeCreated | Select-Object -ExpandProperty TotalMilliSeconds) / 1000 }
                }

                if( ! $noFileInfo )
                {
                    $exeProperties = $fileProperties[ $started.NewProcessName ]
                    if( ! $exeProperties )
                    {
                        if( $remoteParam.Count )
                        {
                            $result = Invoke-Command @remoteParam -ScriptBlock { Get-ItemProperty -Path $($using:started).NewProcessName -ErrorAction SilentlyContinue }
                            if( $result )
                            {
                                ## This is a deserialised object which doesn't seem to persist new properties added so we will make a local copy
                                $exeProperties = New-Object -TypeName 'PSCustomObject'
                                $result.PSObject.Properties | Where-Object MemberType -Match 'property$' | ForEach-Object `
                                {
                                    if( $_.Name -eq 'VersionInfo' ) ## has been flattened into a string so need to unflatten
                                    {
                                        [int]$added = 0
                                        $versionInfo = New-Object -TypeName 'PSCustomObject'
                                        $_.Value -split "`n" | ForEach-Object `
                                        {
                                            [string[]]$split = $_ -split ':',2 ## will be : in file names so only split on first
                                            if( $split -and $split.Count -eq 2 )
                                            {
                                                Add-Member -InputObject $versionInfo -MemberType NoteProperty -Name $split[0] -Value ($split[1]).Trim()
                                                $added++
                                            }
                                        }
                                        if( $added )
                                        {
                                            Add-Member -InputObject $exeProperties -MemberType NoteProperty -Name VersionInfo -Value $versionInfo 
                                        }
                                    }
                                    else
                                    {
                                        Add-Member -InputObject $exeProperties -MemberType NoteProperty -Name $_.Name -Value $_.Value
                                    }
                                }
                            }
                        }
                        else
                        {
                            $exeProperties = Get-ItemProperty -Path $started.NewProcessName -ErrorAction SilentlyContinue
                        }
                        if( $exeProperties ) 
                        {  
                            try
                            {
                                if( $remoteParam.Count )
                                {
                                    $signatureStatus = Invoke-Command @remoteParam -ScriptBlock { Get-AuthenticodeSignature -FilePath $($using:exeProperties).FullName -ErrorAction SilentlyContinue } | Select-Object -ExpandProperty 'Status' | Select-Object -ExpandProperty 'Value'
                                }
                                else
                                {
                                    $signatureStatus = Get-AuthenticodeSignature -FilePath $exeProperties.FullName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty 'Status'
                                }
                            }
                            catch
                            {
                                $signatureStatus = '?'
                            }
                            if( $remoteParam.Count )
                            {
                                $owner = Invoke-Command @remoteParam -ScriptBlock {  Get-Acl -Path $($using:started).NewProcessName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Owner }
                            }
                            else
                            {
                                $owner = Get-Acl -Path $started.NewProcessName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Owner
                            }
                            $extraProperties = Add-Member -PassThru -InputObject $exeProperties -Force -NotePropertyMembers @{
                                'Signed' = if( $signatureStatus.ToString() -eq 'Valid' ) { 'Yes' } else { 'No' }
                                'Occurrences' = ([int]1)
                                'Owner' = $owner
                            }
                            $fileProperties.Add( $started.NewProcessName , $extraProperties )
                        }
                    }
                    else
                    {
                        $exeProperties.Occurrences += 1
                    }
                    if( $exeProperties )
                    {     
                        $started += @{
                            'Exe Signed' = $exeProperties.Signed
                            'Exe Created' = $exeProperties.CreationTime
                            'Exe Modified' = $exeProperties.LastWriteTime 
                            'Exe File Owner' = $exeProperties.Owner }
                    }
                }

                if( ! $terminate -and ! $PSBoundParameters[ 'eventLog' ] ) ## probably still running
                {
                    $existing = $runningProcesses[ [int]$started.NewProcessId ]
                    if( $existing )
                    {
                        ## check user running now is same as launched
                        if( $existing.UserName -and $existing.UserName -ne "$($started.SubjectDomainName)\$($started.SubjectUserName)" )
                        {
                            $differentUserName = $true
                            $started.Add( 'User Name Now' , $existing.UserName )
                        }
                    }
                    else
                    {
                        Write-Warning "Cannot find process terminated event for pid $($started.NewProcessId) and not currently running"
                    }
                }

                ## check on same computer and logon after last boot othwerwise LSASS won't have it
                $theLogon = $null
                if( $logonTimes -and  ( $earliestEventHere -and ( $earliestEventHere.MachineName -eq $computer -or ( $earliestEventHere.MachineName -split '\.' )[0] -eq $computer ) ) -and ( ! $bootTime -or $started.Start -ge $bootTime ) )
                {
                    ## get the logon time so we can add a column relative to it
                    $logonTime = $null
                    $thisLogon = $logons[ $started.SubjectLogonId.ToString() ]

                    if( $thisLogon )
                    {
                        if( $thisLogon -is [array] )
                        {
                            ## Need to find this user
                            ForEach( $alogon in $thisLogon )
                            {
                                if( $started.SubjectDomainName -eq $alogon.Domain -and $started.SubjectUserName -eq $alogon.Username )
                                {
                                    if( $theLogon )
                                    {
                                        Write-Warning "Multiple logons for same user $($started.SubjectDomainName)\$($started.SubjectUserName)"
                                    }
                                    $theLogon = $alogon
                                }
                            }
                        }
                        elseif( $started.SubjectDomainName -eq $thisLogon.Domain -and $started.SubjectUserName -eq $thisLogon.Username )
                        {
                            $theLogon -eq $thisLogon
                        }
                        if( ! $theLogon )
                        {
                            Write-Warning "Couldn't find logon for user $($started.SubjectDomainName)\$($started.SubjectUserName) for process $($started.NewProcessId) started @ $(Get-Date -Date $started.Start -Format G)"
                        }
                        $started.Add( 'Logon Time' , $(if( $theLogon ) { $theLogon.LogonTime } ) )
                        $started.Add( 'After Logon (s)' , $(if( $theLogon ) { New-TimeSpan -Start $theLogon.LogonTime -End $started.Start | Select-Object -ExpandProperty TotalSeconds } ) )
                    }
                }
                if( $bootTimes -and $bootTime )
                {
                    $started.Add( 'After Boot (s)' , $(if( $theLogon ) { New-TimeSpan -Start $bootTime -End $started.Start | Select-Object -ExpandProperty TotalSeconds } ) )
                }
                [pscustomobject]$started
            }
        }

        if( $eventError )
        {
            $oldestEvent = Get-WinEvent @remoteParam -LogName $startEventFilter[ 'Logname' ] -Oldest -ErrorAction SilentlyContinue | Select-Object -First 1
            Write-Warning "Failed to get any events with ids $($startEventFilter[ 'Id' ] -join ',') from $($startEventFilter[ 'Logname' ]) event log on $computer between $(Get-Date -Date $startEventFilter[ 'StartTime' ] -Format G) and $(Get-Date -Date $startEventFilter[ 'EndTime' ] -Format G), oldest event is $(Get-Date -Date $oldestEvent.TimeCreated -Format G)"

            ## Look to see if there's an error otherwise check if required auditing is enabled
            if( $eventError -and $eventError.Count )
            {
                if( $eventError[0].Exception.Message -match 'No events were found that match the specified selection criteria' )
                {
                    ForEach( $auditGuid in $auditingGuids.GetEnumerator() )
                    {
                        [string]$result = Get-AuditSetting -GUID $auditGuid.Value
                        if( $result -notmatch 'Success' )
                        {
                            Write-Warning "Auditing for $($auditGuid.Name) is `"$result`" so will not generate required events"
                        }
                    }
                }
                Throw $eventError[0]
            }
        }
    }
})

[string]$status = $null
if( $PSBoundParameters[ 'username' ] )
{
    $status += " matching user `"$username`""
}
if( $startEventFilter[ 'StartTime' ] )
{
    $status += " from $(Get-Date -Date $startEventFilter[ 'StartTime' ] -Format G)"
}
if( $startEventFilter[ 'EndTime' ] )
{
    $status += " to $(Get-Date -Date $startEventFilter[ 'EndTime' ] -Format G)"
}
if( $earliestEvent )
{
    $status += " earliest event $(Get-Date $earliestEvent.TimeCreated -Format G)"
}

Write-Verbose "Got $($processes.Count) processes $status in $((New-TimeSpan -Start $startTime -End (Get-Date)).TotalSeconds) seconds"

if( $listSessions )
{
    if( $allSessions.Count )
    {
        $allSessions.GetEnumerator() | Select-Object -ExpandProperty Value | Format-Table -AutoSize
    }
    else
    {
        Write-Warning "No logon sessions found"
    }
}

if( ! $processes.Count )
{
    if( ! $listSessions )
    {
        Write-Warning "No process start events found $status"
    }
}
elseif( $summary )
{
    $output = @( $fileProperties.GetEnumerator() | ForEach-Object `
    {
        $exeFile = $_.Value
        [pscustomobject][ordered]@{
            'Executable' = $exeFile.FullName
            'Executions' = $exeFile.Occurrences
            'Created' = $exeFile.CreationTime
            'Last Modified' = $exeFile.LastWriteTime
            'File Owner' = $exeFile.Owner
            'Size (KB)' = [int]( $exeFile.Length / 1KB )
            'Product Name' = $exeFile.VersionInfo | Select-Object -ExpandProperty ProductName -ErrorAction SilentlyContinue
            'Company Name' = $exeFile.VersionInfo | Select-Object -ExpandProperty CompanyName -ErrorAction SilentlyContinue
            'File Description'  = $exeFile.VersionInfo | Select-Object -ExpandProperty FileDescription -ErrorAction SilentlyContinue
            'File Version' = $exeFile.VersionInfo | Select-Object -ExpandProperty FileVersion -ErrorAction SilentlyContinue
            'Product Version' = $exeFile.VersionInfo | Select-Object -ExpandProperty ProductVersion -ErrorAction SilentlyContinue
            'Signed' = $exeFile.Signed
        }
    })

    if( $nogridview )
    {
        $output
    }
    else
    {
        [string]$title = "$($output.Count) unique executable process starts $status"
        [array]$selected = @( $output | Out-GridView -PassThru -Title $title )
        if( $selected -and $selected.Count )
        {
            $selected | Set-Clipboard
        }
    }
}
else
{
    $headings = [System.Collections.ArrayList]@()
    if( $computers.Count -gt 1 )
    {
        [void]$headings.Add( 'Computer' )
    }
    [void]$headings.Add( @{n='User Name';e={'{0}\{1}' -f $_.SubjectDomainName , $_.SubjectUserName}} )
    if( $differentUserName )
    {
        [void]($headings.Add( 'User Name Now' ))
    }
    $headings += @( @{n='Process';e={$_.NewProcessName}} , @{n='PID';e={$_.NewProcessId}} , `
        'CommandLine' , @{n='Parent Process';e={$_.ParentProcessName}} , 'SubjectLogonId' , @{n='Parent PID';e={$_.ProcessId}} , @{n='Start';e={('{0}.{1}' -f (Get-Date -Date $_.start -Format G) , $_.start.Millisecond)}} , `
        @{n='End';e={('{0}.{1}' -f (Get-Date -Date $_.end -Format G) , $_.end.Millisecond)}} , 'Duration' ,  @{n='Exit Code';e={if( $_.'Exit Code' -ne $null ) { '0x{0:x}' -f $_.'Exit Code'}}} )
    if( ! $noFileInfo )
    {
        $headings += @( 'Exe Created' , 'Exe Modified' , 'Exe File Owner' )
    }
    if( $logonTimes )
    {
        $headings += @( @{n='Logon Time';e={('{0}.{1}' -f (Get-Date -Date $_.'Logon Time' -Format G) , $_.'Logon Time'.Millisecond)}} , 'After Logon (s)' )
    }
    if( $bootTimes -and $bootTime )
    {
        $headings += @( @{n='Boot Time';e={Get-Date -Date $bootTime -Format G}} , 'After Boot (s)' )
    }
    if( $PSBoundParameters[ 'outputfile' ] )
    {
        $processes | Export-Csv -Path $outputFile -NoTypeInformation -NoClobber
    }
    if( $nogridview )
    {
        $processes
    }
    else
    {
        <#
        [void][Reflection.Assembly]::LoadWithPartialName('Presentationframework')

        $mainForm = Load-GUI $mainwindowXAML

        if( ! $mainForm )
        {
            return
        }

        [array]$processes = @( Get-CimInstance -ClassName win32_process -ComputerName $machine | Select-Object Name,ProcessId,@{n='Owner';e={Invoke-WmiMethod -InputObject $_ -Name GetOwner | Select -ExpandProperty user}},ParentProcessId,SessionId,
                                @{n='WorkingSetKB';e={[math]::Round( $_.WorkingSetSize / 1KB)}},@{n='PeakWorkingSetKB';e={$_.PeakWorkingSetSize}},
                                @{n='PageFileUsageKB';e={$_.PageFileUsage}},@{n='PeakPageFileUsageKB';e={$_.PeakPageFileUsage}},
                                @{n='IOReadsMB';e={[math]::Round($_.ReadTransferCount/1MB)}},@{n='IOWritesMB';e={[math]::Round($_.WriteTransferCount/1MB)}},
                                @{n='BasePriority';e={$processPriorities[ [int]$_.Priority ] }},@{n='ProcessorTime';e={[math]::round( ($_.KernelModeTime + $_.UserModeTime) / 10e6 )}},HandleCount,ThreadCount,
                                @{n='StartTime';e={Get-Date $_.CreationDate -Format G}},CommandLine )

        $mainForm.Title =  "$($processes.Count) processes on $machine"
        $WPFProcessList.ItemsSource = $processes
        $WPFProcessList.IsReadOnly = $true
        $WPFProcessList.CanUserSortColumns = $true
        #>

        [string]$title = "$($processes.Count) process starts $status"
        [array]$selected = @( $processes | Select-Object -Property $headings| Out-GridView -PassThru -Title $title )
        if( $selected -and $selected.Count )
        {
            $selected | Set-Clipboard
        }
    }
}
