#requires -version 3.0
<#
    Give some machine info & stats for server core login when PowerShell is set as the shell in "Shell" value in HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon

    Copy file to $PSHOME (C:\Windows\System32\WindowsPowerShell\v1.0)

    Set $samples to 0 if you don't want to delay to measure CPU usage

    Guy Leech, 2018
#>

[int]$samples = 5

$parent = Get-Process -Id (Get-CimInstance -ClassName Win32_Process -Filter "processid = $pid"|select -ExpandProperty ParentProcessId) -EA SilentlyContinue

[console]::Title = "Process id $pid started at $(Get-Date -Format G)"

## Only display for logon shell or if we can't find the parent
if( ! $parent -or $parent.Name -eq 'userinit' )
{
	Get-CimInstance -ClassName Win32_ComputerSystem | Select Name,Domain,@{n='Installed Memory (GB)';e={[math]::round($_.TotalPhysicalMemory/1GB,2)}},NumberOfLogicalProcessors,NumberOfProcessors|Format-Table -Autosize

	Get-CimInstance -ClassName Win32_OperatingSystem|select @{n='Free Memory (GB)';e={[math]::round( $_.FreePhysicalMemory / 1MB,2)}},@{n='Free Pagefile (GB)';e={[math]::round($_.FreeSpaceInPagingFiles/1MB,2)}},LastBootUpTime,InstallDate,LocalDateTime|Format-Table -Autosize

	Get-CimInstance -ClassName Win32_LogicalDisk -Filter "drivetype = 3"|select deviceid,@{n='Size (GB)';e={[math]::round($_.size/1GB)}},@{n='Free (GB)';e={[math]::round($_.freespace/1GB)}}|Format-Table -Autosize

	Get-NetAdapter

	Get-NetIPAddress | Where-Object { $_.InterfaceAlias -notmatch '^Loopback' -and $_.InterfaceAlias -notmatch '^isatap'  } | select InterfaceAlias,IPAddress,PrefixLength| sort InterfaceAlias | Format-Table -Autosize

    if( $samples )
    {
        Write-Output "Sampling CPU for $samples seconds ..."

        [decimal]$cpuUsage = [math]::Round( ( Get-Counter -Counter '\Processor(*)\% Processor Time' -SampleInterval 1 -MaxSamples $samples | select -ExpandProperty CounterSamples| Where-Object { $_.InstanceName -eq '_total' } | select -ExpandProperty CookedValue  | Measure-Object -Average ).Average , 1 )

        Write-Output "Average CPU over $samples seconds is $cpuUsage %"
    }
}

## For safety, let us not sit in system folders
if( (Get-Location).Path -match "^$windir" )
{
	Set-Location $env:userprofile
}
