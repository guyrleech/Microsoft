#requires -version 3
<#
    Use WTSAPi32.dll to get sessions

    Some code adapted from https://www.pinvoke.net/default.aspx/wtsapi32.wtsenumeratesessions

    @guyrleech 2018

    Modification History:

    21/01/19  GRL Changed Get-WTSSessionInformation to take array of computer names
#>

Add-Type -ErrorAction Stop -TypeDefinition @'
    using System;
    using System.Runtime.InteropServices;
    public enum WTS_CONNECTSTATE_CLASS
    {
        WTSActive,
        WTSConnected,
        WTSConnectQuery,
        WTSShadow,
        WTSDisconnected,
        WTSIdle,
        WTSListen,
        WTSReset,
        WTSDown,
        WTSInit
    }
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    public struct WTSINFOEX_LEVEL1_W {
        public Int32                  SessionId;
        public WTS_CONNECTSTATE_CLASS SessionState;
        public Int32                   SessionFlags; // 0 = locked, 1 = unlocked , ffffffff = unknown
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 33)]
        public string WinStationName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 21)]
        public string UserName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 18)]
        public string DomainName;
        public UInt64           LogonTime;
        public UInt64           ConnectTime;
        public UInt64           DisconnectTime;
        public UInt64           LastInputTime;
        public UInt64           CurrentTime;
        public Int32            IncomingBytes;
        public Int32            OutgoingBytes;
        public Int32            IncomingFrames;
        public Int32            OutgoingFrames;
        public Int32            IncomingCompressedBytes;
        public Int32            OutgoingCompressedBytes;
    }
    [StructLayout(LayoutKind.Sequential)]
    public struct WTS_SESSION_INFO
    {
        public Int32 SessionID;

        [MarshalAs(UnmanagedType.LPStr)]
        public String pWinStationName;

        public WTS_CONNECTSTATE_CLASS State;
    }
    [StructLayout(LayoutKind.Explicit)]
    public struct WTSINFOEX_LEVEL_W
    { //Union
        [FieldOffset(0)]
        public WTSINFOEX_LEVEL1_W WTSInfoExLevel1;
    } 
    [StructLayout(LayoutKind.Sequential)]
    public struct WTSINFOEX
    {
        public Int32 Level ;
        public WTSINFOEX_LEVEL_W Data;
    }
    public enum WTS_INFO_CLASS
    {
        WTSInitialProgram,
        WTSApplicationName,
        WTSWorkingDirectory,
        WTSOEMId,
        WTSSessionId,
        WTSUserName,
        WTSWinStationName,
        WTSDomainName,
        WTSConnectState,
        WTSClientBuildNumber,
        WTSClientName,
        WTSClientDirectory,
        WTSClientProductId,
        WTSClientHardwareId,
        WTSClientAddress,
        WTSClientDisplay,
        WTSClientProtocolType,
        WTSIdleTime,
        WTSLogonTime,
        WTSIncomingBytes,
        WTSOutgoingBytes,
        WTSIncomingFrames,
        WTSOutgoingFrames,
        WTSClientInfo,
        WTSSessionInfo,
        WTSSessionInfoEx,
        WTSConfigInfo,
        WTSValidationInfo,   // Info Class value used to fetch Validation Information through the WTSQuerySessionInformation
        WTSSessionAddressV4,
        WTSIsRemoteSession
    }
    public static class wtsapi
    {
        [DllImport("wtsapi32.dll", SetLastError=true)]
        public static extern int WTSQuerySessionInformationW(
                 System.IntPtr hServer,
                 int SessionId,
                 int WTSInfoClass ,
                 ref System.IntPtr ppSessionInfo,
                 ref int pBytesReturned );

        [DllImport("wtsapi32.dll", SetLastError=true)]
        public static extern int WTSEnumerateSessions(
                 System.IntPtr hServer,
                 int Reserved,
                 int Version,
                 ref System.IntPtr ppSessionInfo,
                 ref int pCount);

        [DllImport("wtsapi32.dll", SetLastError=true)]
        public static extern IntPtr WTSOpenServer(string pServerName);
        
        [DllImport("wtsapi32.dll", SetLastError=true)]
        public static extern void WTSCloseServer(IntPtr hServer);
        
        [DllImport("wtsapi32.dll", SetLastError=true)]
        public static extern void WTSFreeMemory(IntPtr pMemory);
    }
'@ 

Function Get-WTSSessionInformation
{
    [cmdletbinding()]

    Param
    (
        [string[]]$computers = @( $null )
    )

    [long]$count = 0
    [IntPtr]$ppSessionInfo = 0
    [IntPtr]$ppQueryInfo = 0
    [long]$ppBytesReturned = 0
    $wtsSessionInfo = New-Object -TypeName 'WTS_SESSION_INFO'
    $wtsInfoEx = New-Object -TypeName 'WTSINFOEX'
    [int]$datasize = [system.runtime.interopservices.marshal]::SizeOf( [Type]$wtsSessionInfo.GetType() )

    ForEach( $computer in $computers )
    {
        [string]$machineName = $(if( $computer ) { $computer } else { $env:COMPUTERNAME })
        [IntPtr]$serverHandle = [wtsapi]::WTSOpenServer( $computer )

        ## If the function fails, it returns a handle that is not valid. You can test the validity of the handle by using it in another function call.

        [long]$retval = [wtsapi]::WTSEnumerateSessions( $serverHandle , 0 , 1 , [ref]$ppSessionInfo , [ref]$count );$LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()

        if ($retval -ne 0)
        {
             for ([int]$index = 0; $index -lt $count; $index++)
             {
                 $element = [system.runtime.interopservices.marshal]::PtrToStructure( [long]$ppSessionInfo + ($datasize * $index), [type]$wtsSessionInfo.GetType())
                 if( $element -and $element.SessionID -ne 0 ) ## session 0 is non-interactive (session zero isolation)
                 {
                     $retval = [wtsapi]::WTSQuerySessionInformationW( $serverHandle , $element.SessionID , [WTS_INFO_CLASS]::WTSSessionInfoEx , [ref]$ppQueryInfo , [ref]$ppBytesReturned );$LastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
                     if( $retval -and $ppQueryInfo )
                     {
                        $value = [system.runtime.interopservices.marshal]::PtrToStructure( $ppQueryInfo , [Type]$wtsInfoEx.GetType())
                        if( $value -and $value.Data -and $value.Data.WTSInfoExLevel1.SessionState -ne [WTS_CONNECTSTATE_CLASS]::WTSListen -and $value.Data.WTSInfoExLevel1.SessionState -ne [WTS_CONNECTSTATE_CLASS]::WTSConnected )
                        {
                            $wtsinfo = $value.Data.WTSInfoExLevel1
                            $idleTime = New-TimeSpan -End ([datetime]::FromFileTimeUtc($wtsinfo.CurrentTime)) -Start ([datetime]::FromFileTimeUtc($wtsinfo.LastInputTime))
                            Add-Member -InputObject $wtsinfo -Force -NotePropertyMembers @{
                                'IdleTimeInSeconds' =  $idleTime | Select -ExpandProperty TotalSeconds
                                'IdleTimeInMinutes' =  $idleTime | Select -ExpandProperty TotalMinutes
                                'Computer' = $machineName
                                'LogonTime' = [datetime]::FromFileTime( $wtsinfo.LogonTime )
                                'DisconnectTime' = [datetime]::FromFileTime( $wtsinfo.DisconnectTime )
                                'LastInputTime' = [datetime]::FromFileTime( $wtsinfo.LastInputTime )
                                'ConnectTime' = [datetime]::FromFileTime( $wtsinfo.ConnectTime )
                                'CurrentTime' = [datetime]::FromFileTime( $wtsinfo.CurrentTime )
                            }
                            $wtsinfo
                        }
                        [wtsapi]::WTSFreeMemory( $ppQueryInfo )
                        $ppQueryInfo = [IntPtr]::Zero
                     }
                     else
                     {
                        Write-Error "$($machineName): $LastError"
                     }
                 }
             }
        }
        else
        {
            Write-Error "$($machineName): $LastError"
        }

        if( $ppSessionInfo -ne [IntPtr]::Zero )
        {
            [wtsapi]::WTSFreeMemory( $ppSessionInfo )
            $ppSessionInfo = [IntPtr]::Zero
        }
        [wtsapi]::WTSCloseServer( $serverHandle )
        $serverHandle = [IntPtr]::Zero
    }
}
