#requires -version 3
<#
    Check already running processes and then watch for process created events and if the process is in a specified list and have been launched elevated then terminate them and audit to event log

    Modification History:

    02/08/19    @guyrleech   Initial Release
#>

<#
.SYNOPSIS

Check already running processes and then watch for process created events and if the process is in a specified list and have been launched elevated then terminate them and audit to event log

.PARAMETER processes

A comma separated list of processes which must not be run elevated so if they are they will be terminated

.PARAMETER allowedUsers

A comma separated list of domain\user names that are allowed to run the prohibited processes elevated

.PARAMETER disallowedUsers

A comma separated list of domain\user names that are not allowed to run the prohibited processes elevated so if a user is not in this list they will be allowed to run the prohibited processes elevated

.PARAMETER noStartupCheck

Do not check for already running elvated prohibited processes when the script starts

.PARAMETER  noAuditing

Do not audit process terminations to the event log

.PARAMETER eventLogName

Name of the event log to audit to. If it does not exist it will be created.

.PARAMETER eventSource

Name of the event source to audit against. If it does not exist it will be created.

.PARAMETER eventId

The id of the event to write to the event log

.PARAMETER eventLogSize

The maximum size of the event log if it is created by the script

.EXAMPLE

& '.\Kill elevated processes.ps1'

Check if any of the default processes are running elevated, or are subsequently launched elevated, and if so kill them, including any child processes, and write an event to the event log

.EXAMPLE

& '.\Kill elevated processes.ps1' -eventLogName 'Guy Leech' -eventSource 'Bad Stuff' -allowedUsers contoso\guyl

Check if any of the default processes are running elevated, or are subsequently launched elevated, unless they are running as contoso\guyl, and if so kill them, including any child processes, 
and write an event to the  'Guy Leech' event log with event source "Bad Stuff"

.NOTES

Script must be run elevated otherwise it won't be able to terminate elevated processes

#>

[CmdletBinding()]

Param
(
    [string[]]$processes  = @( 'chrome.exe' , 'iexplore.exe' , 'firefox.exe' , 'microsoftedge.exe' , 'microsoftedgecp.exe' ) ,
    [AllowNull()]
    [string[]]$allowedUsers ,
    [AllowNull()]
    [string[]]$disallowedUsers ,
    [switch]$noStartupCheck ,
    [switch]$noAuditing ,
    [string]$eventLogName = 'Application' ,
    [string]$eventSource = 'SecurityCenter' ,
    [int]$eventId = '666' ,
    [long]$eventLogSize = 8MB
)

Function Kill-ChildProcesses
{
    [CmdletBinding()]
    Param
    (
        [int]$parentPid
    )

    if( $parentPid -gt 0 )
    {
        Get-CimInstance -ClassName win32_process -Filter "ParentProcessId = '$parentPid'" -ErrorAction SilentlyContinue -Verbose:$false | . { Process `
        {
            Write-Verbose "Killing child process $($_.Name) (pid $($_.ProcessId)) of pid $parentPid"
            Kill-ChildProcesses -parentPid $_.ProcessId
        }}

        Stop-Process -Force -Id $parentPid -ErrorAction SilentlyContinue
    }
}

Function Check-ElevatedProcess
{
    [CmdletBinding()]
    Param
    (
        $thisProcess ,
        [string[]]$allowedUsers ,
        [string[]]$disallowedUsers ,
        [AllowNull()]
        [string]$eventLogName ,
        [AllowNull()]
        [string]$eventSource ,
        [int]$eventId
    )

    ## TODO should we add groups capability?
    if( ! [string]::IsNullOrEmpty( $thisProcess.UserName ) -and $allowedUsers -and $allowedUsers.Count )
    {
        if( $allowedUsers -contains $thisProcess.UserName )
        {
            Write-Verbose "User $($thisProcess.UserName) is in the allowed list so allowing"
            return
        }
    }
    if( ! [string]::IsNullOrEmpty( $thisProcess.UserName ) -and $disallowedUsers -and $disallowedUsers.Count )
    {
        if( ! ( $disallowedUsers -contains $thisProcess.UserName ) )
        {
            Write-Verbose "User $($thisProcess.UserName) is not in the disallowed list so allowing"
            return
        }
    }
    if( $thisProcess -and $thisProcess.Handle )
    {
        ## Now we need to get the process token
        [IntPtr]$hToken = [IntPtr]::Zero
        $returned = [ProcessTokens]::OpenProcessToken( $thisProcess.Handle , [ProcessTokens]::TOKEN_QUERY, [ref]$hToken);$LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
        ## Now we need to see if elevated 
        if( $returned -and $hToken -ne [IntPtr]::Zero )
        {
            [IntPtr]$ptr = [IntPtr]::Zero
            [int]$returnedLength = 0
	        $tokenInfo = New-Object -TypeName ProcessTokens+TOKEN_ELEVATION
            [int]$size = [System.Runtime.InteropServices.Marshal]::SizeOf($tokenInfo)
            [IntPtr]$ptrTokenInfo = [System.Runtime.InteropServices.Marshal]::AllocHGlobal([System.Runtime.InteropServices.Marshal]::SizeOf( $tokenInfo ))
            $marshalResult = [System.Runtime.InteropServices.Marshal]::StructureToPtr( $tokenInfo , $ptrTokenInfo , $false )
            ## https://docs.microsoft.com/en-us/windows/win32/api/securitybaseapi/nf-securitybaseapi-gettokeninformation
            ## https://technet.microsoft.com/en-us/windowsserver/aa379626(v=vs.100)
            ## https://technet.microsoft.com/en-us/windowsserver/bb530717(v=vs.100)
            $returned = [ProcessTokens]::GetTokenInformation(
                $hToken ,
                [ProcessTokens+TOKEN_INFORMATION_CLASS]::TokenElevation ,
                $ptrTokenInfo ,
                $size ,
                [ref]$returnedLength );$LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
            if( $returned -and $returnedLength -eq $size )
            {
                $tokenElevationInfo = [System.Runtime.InteropServices.Marshal]::PtrToStructure( $ptrTokenInfo , [Type]$tokenInfo.GetType() )
                if( $tokenElevationInfo )
                {
                    if( $tokenElevationInfo.TokenIsElevated )
                    {
                        ## See if it has any child processes and terminate those too
                        Kill-ChildProcesses -parentPid $thisProcess.Id
                        Write-Verbose "$(Get-Date -Format G): Process $($thisProcess.Name) (pid $($thisProcess.id)) running as $($thisProcess.UserName) is elevated so terminated"
                        if( ! [string]::IsNullOrEmpty( $eventLogName ) )
                        {
                            [string]$message = "Terminated elevated process $($thisProcess.Name) (pid $($thisProcess.id)) started at $(Get-Date -Date $thisProcess.StartTime -Format G) running as $($thisProcess.UserName) in session $($thisProcess.SessionId)"
                            Write-EventLog -LogName $eventLogName -Source $eventSource -EventId $eventId -EntryType Error -Message $message
                        }
                    }
                }
                else
                {
                    Write-Warning "$(Get-Date -Format G): Failed to marshal token elevation result for $($thisProcess.Name) (pid $($thisProcess.id)) running as $($thisProcess.UserName)"
                }
            }
            else
            {
                Write-Warning "$(Get-Date -Format G): Failed to get token information for $($thisProcess.Name) (pid $($thisProcess.id)) running as $($thisProcess.UserName): $LastError"
            }
            [Runtime.InteropServices.Marshal]::FreeHGlobal( $ptrTokenInfo )
            $ptrTokenInfo = [IntPtr]::Zero
            [void][ProcessTokens]::CloseHandle( $hToken )
            $hToken = [IntPtr]::Zero
        }
        else
        {
            Write-Warning "$(Get-Date -Format G): Failed to open process token for $($thisProcess.Name) (pid $($thisProcess.id)) running as $($thisProcess.UserName): $LastError"
        }
    }
    else
    {
        Write-Warning "$(Get-Date -Format G): No process handle for $($event.TargetInstance.Name) (pid $($event.TargetInstance.ProcessId)) so cannot determine elevation"
    }
}

$myWindowsPrincipal = New-Object System.Security.Principal.WindowsPrincipal([System.Security.Principal.WindowsIdentity]::GetCurrent())

if( ! $myWindowsPrincipal.IsInRole(([System.Security.Principal.WindowsBuiltInRole]::Administrator)))
{
    Throw "Script must run as an administrator otherwise it won't be able to terminate elevated processes"
}

$definition = @'
using System;
using System.Runtime.InteropServices;

public class ProcessTokens
{
    public enum TOKEN_INFORMATION_CLASS 
    { 
      TokenUser                             = 1,
      TokenGroups,
      TokenPrivileges,
      TokenOwner,
      TokenPrimaryGroup,
      TokenDefaultDacl,
      TokenSource,
      TokenType,
      TokenImpersonationLevel,
      TokenStatistics,
      TokenRestrictedSids,
      TokenSessionId,
      TokenGroupsAndPrivileges,
      TokenSessionReference,
      TokenSandBoxInert,
      TokenAuditPolicy,
      TokenOrigin,
      TokenElevationType,
      TokenLinkedToken,
      TokenElevation,
      TokenHasRestrictions,
      TokenAccessInformation,
      TokenVirtualizationAllowed,
      TokenVirtualizationEnabled,
      TokenIntegrityLevel,
      TokenUIAccess,
      TokenMandatoryPolicy,
      TokenLogonSid,
      TokenIsAppContainer,
      TokenCapabilities,
      TokenAppContainerSid,
      TokenAppContainerNumber,
      TokenUserClaimAttributes,
      TokenDeviceClaimAttributes,
      TokenRestrictedUserClaimAttributes,
      TokenRestrictedDeviceClaimAttributes,
      TokenDeviceGroups,
      TokenRestrictedDeviceGroups,
      TokenSecurityAttributes,
      TokenIsRestricted,
      MaxTokenInfoClass
    };
    
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    public struct TOKEN_ELEVATION 
    {
        public int TokenIsElevated;
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool CloseHandle( [In] IntPtr hHandle );
            
    [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool OpenProcessToken(IntPtr h, int acc, ref IntPtr phtok);

    [DllImport("advapi32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetTokenInformation(
        IntPtr                  TokenHandle,
        TOKEN_INFORMATION_CLASS TokenInformationClass,
        IntPtr                  TokenInformation,
        int                     TokenInformationLength,
        ref int                 ReturnLength
    );

    public const int TOKEN_QUERY = 0x00000008;
}

'@

[void](Add-Type $definition -PassThru -ErrorAction Stop)

if( ! $noAuditing )
{
    try
    {
        $eventLog = Get-EventLog -LogName $eventLogName -ErrorAction SilentlyContinue ## will also error if present but no events
    }
    catch
    {
        $eventLog = $null
    }

    if( ! $eventLog )
    {
        New-EventLog -LogName $eventLogName -Source $eventSource
        Limit-EventLog -OverflowAction OverWriteAsNeeded -MaximumSize $eventLogSize -LogName $eventLogName
    }
}
else
{
    $eventLogName = $null
}

if( ! $noStartupCheck )
{
    Write-Verbose "Checking already running processes ..."
    ## Need to use CIM as need full process name, don't assume can just add .exe
    Get-CimInstance -ClassName Win32_process -Filter "SessionId != 0" -ErrorAction SilentlyContinue | Where-Object { $_.Name -in $processes } | . { Process `
    {
        $thisProcess = Get-Process -Id $_.ProcessId -ErrorAction SilentlyContinue -IncludeUserName
        if( $thisProcess )
        {
            Check-ElevatedProcess -thisProcess $thisProcess -eventLogName $eventLogName -eventSource $eventSource -eventId $eventId -allowedUsers $allowedUsers -disallowedUsers $disallowedUsers
        }
    }}
}

## https://fleexlab.blogspot.com/2017/04/watching-for-new-processes-with.html
$query = New-Object System.Management.WqlEventQuery -ArgumentList "__InstanceCreationEvent", (New-Object TimeSpan 0,0,1), "TargetInstance isa 'Win32_Process'"
$processWatcher = New-Object System.Management.ManagementEventWatcher -ArgumentList $query
$processWatcher.Options.Timeout = [System.Management.ManagementOptions]::InfiniteTimeout

Write-Verbose "Watching for new processes ..."

While( $true )
{
    $event = $processWatcher.WaitForNextEvent()
    if( $event -and $event.TargetInstance -and $event.TargetInstance.ProcessId -ne $pid )
    {
        if( $event.TargetInstance.Name -in $processes )
        {
            $thisProcess = Get-Process -Id $event.TargetInstance.ProcessId -IncludeUserName -ErrorAction SilentlyContinue
            Check-ElevatedProcess -thisProcess $thisProcess -eventLogName $eventLogName -eventSource $eventSource -eventId $eventId -allowedUsers $allowedUsers -disallowedUsers $disallowedUsers
        }
        else
        {
            Write-Verbose "$(Get-Date -Format G): Ignoring $($event.TargetInstance.Name) (pid $($event.TargetInstance.ProcessId)) as not in prohibited list"
        }
    }
    else
    {
        Write-Warning "No target instance for event at $(Get-Date -Format G)"
    }
}