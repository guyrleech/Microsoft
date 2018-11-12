<#
    Check Lync/Skype for Business signed in and alert if not

    Require Lync SDK - see notes section below

    NO WARRANTY SUPPLIED. USE THIS SCRIPT AT YOUR OWN RISK.

    @guyrleech, November 2016

    Modification History:

    12/11/18  GRL  Added checks for excessive "Do not disturb" status
#>

<#
.SYNOPSIS

Use the Microsoft Lync SDK to monitor Skype for Business status and alert with popup and optional sound file if it is not signed in.

.DESCRIPTION

Reduce the risk of missing messages, etc because Skype for Business has somehow become signed out.

.PARAMETER pollPeriod

How often to check, in seconds, that Skype for Business is signed in. The default is 120 seconds.

.PARAMETER excessiveDoNotDisturbPeriod

Number of minutes over which a dialogue will be shown asking if the user wants to continue in Do Not Disturb, change their status or snooze notifications

.PARAMETER doNotDisturbText

The Lync status text which is monitored for excessive time in this state

.PARAMETER lyncPath

Specifies the full path to the "Microsoft.Lync.Model.dll" file required from the Lync SDK. Only specify it if the SDK is not installed in the standard location.

.PARAMETER soundFile

Path to an audio file (e.g. .wav or .mp3) that will be played if Skype for Business is detected as not being signed in

.NOTES

Needs Lync 2013 SDK from https://www.microsoft.com/en-us/download/details.aspx?id=36824 but must extract with 7zip and run the msi otherwise it will complain Lync/Skype is not installed.

There is no Lync 2016 SDK (yet)

If you get the following error, your Lync client needs updating:

    InvalidCastException: Unable to cast COM object of type 'System.__ComObject' to interface type 'Microsoft.Office.Uc.UCOfficeIntegration'. This operation failed because the QueryInterface call on the COM component for the interface with IID '{6A222195-F65E-467F-8F77-EB180BD85288}' failed due to the following error: No such interface supported (Exception from HRESULT: 0x80004002 (E_NOINTERFACE)).

#>

[CmdletBinding()]

Param
(
    [int]$pollperiod = 120 , ## seconds
    [int]$excessiveDoNotDisturbPeriod , ## minutes
    [string]$doNotDisturbText = 'Do not disturb' ,
    [string]$lyncPath  ,
    [string]$soundFile 
)

## https://docs.microsoft.com/en-us/previous-versions/office/developer/lync-2010/hh380072(v%3Doffice.14)
[hashtable]$availabilityId = @{
    'Available' = 3500
    'Busy' = 6500
    'Do not disturb' = 9500
    'Be right back' = 12500
    'Away' = 15500
    'Offline' = 18500
}

#region XAML&Modules

[string]$mainWindowXAML = @'
<Window x:Name="wndMain" x:Class="SkyperChecker.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:SkyperChecker"
        mc:Ignorable="d"
        Title="Skype for Business Checker" Height="361" Width="689">
    <Grid>
        <TextBox x:Name="txtNotificationArea" HorizontalAlignment="Left" Height="72" Margin="25,34,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="620"/>
        <Button x:Name="btnStayBusy" Content="Stay Busy" HorizontalAlignment="Left" Height="35" Margin="25,142,0,0" VerticalAlignment="Top" Width="154" IsDefault="True"/>
        <Button x:Name="btnSnooze" Content="Snooze" HorizontalAlignment="Left" Height="34" Margin="25,206,0,0" VerticalAlignment="Top" Width="154" RenderTransformOrigin="0.487,2.118"/>
        <TextBox x:Name="txtSnoozePeriod" HorizontalAlignment="Left" Height="34" Margin="208,206,0,0" TextWrapping="Wrap" Text="10" VerticalAlignment="Top" Width="63"/>
        <Label Content="Minutes" HorizontalAlignment="Left" Height="34" Margin="289,206,0,0" VerticalAlignment="Top" Width="102"/>
        <Button x:Name="btnChangeStatus" Content="Change Status" HorizontalAlignment="Left" Height="33" Margin="25,266,0,0" VerticalAlignment="Top" Width="154"/>
        <ComboBox x:Name="comboStatus" HorizontalAlignment="Left" Height="33" Margin="208,266,0,0" VerticalAlignment="Top" Width="105">
            <ComboBoxItem Content="Available" IsSelected="True"/>
            <ComboBoxItem Content="Busy"/>
            <ComboBoxItem Content="Be Right Back"/>
            <ComboBoxItem Content="Appear Away"/>
        </ComboBox>

    </Grid>
</Window>
'@

Function Load-GUI( $inputXml )
{
    $form = $NULL
    $inputXML = $inputXML -replace 'mc:Ignorable="d"' , '' -replace 'x:N' ,'N'  -replace '^<Win.*' , '<Window'
 
    [xml]$XAML = $inputXML
 
    $reader = New-Object Xml.XmlNodeReader $xaml

    try
    {
        $Form = [Windows.Markup.XamlReader]::Load( $reader )
    }
    catch
    {
        Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .NET is installed.`n$_"
        return $null
    }
 
    $xaml.SelectNodes('//*[@Name]') | ForEach-Object `
    {
        Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -Scope Global
    }

    return $form
}

#endregion XAML&Modules

Function Initialise-SkypeObject
{
    do
    {
        try
        {
            $client = [Microsoft.Lync.Model.LyncClient]::GetClient()
        }
        catch
        {
            ## See if process is running
            if( ! $self )
            {
                $self = Get-WmiObject -Class win32_process -filter "processid=$pid"
            }
            $proc = Get-Process -Name 'lync' -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $self.SessionId }
            if( $proc )
            {
                $msg = "Failed to get Lync object`n$($error[0].Exception.Message)"
            }
            else
            {
                $msg = "Lync.exe not running - launch it manually please"
            }
                    
            $response = [Microsoft.VisualBasic.Interaction]::MsgBox( ( $msg + ". Abort monitoring?" ), 'YesNo,SystemModal,Critical' , $MyInvocation.MyCommand.Name )

            if( $response -eq 'yes' )
            {
                Exit
            }
        }
    }
    while( ! $client )

    return $client
}

$null = [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

if( [string]::IsNullOrEmpty( $lyncPath ) )
{
    $lyncPath = "${env:ProgramFiles(x86)}\Microsoft Office 2013\LyncSDK\Assemblies\Desktop\Microsoft.Lync.Model.dll"
    if( ! ( Test-Path $lyncPath ) )
    {
        $lyncPath = "${env:ProgramFiles}\Microsoft Office 2013\LyncSDK\Assemblies\Desktop\Microsoft.Lync.Model.dll"
    }
}

if( ! ( Test-Path -Path $lyncPath -PathType Leaf ) )
{
    [Microsoft.VisualBasic.Interaction]::MsgBox( "Unable to locate Lync SDK module in `"$lyncPath`" - aborting" , 'OkOnly,SystemModal,Critical' , $MyInvocation.MyCommand.Name )
    return
}

Import-Module $lyncPath

if( ! $? )
{
    [Microsoft.VisualBasic.Interaction]::MsgBox( "Failed to import Lync SDK module from `"$lyncPath`" - aborting" , 'OkOnly,SystemModal,Critical' , $MyInvocation.MyCommand.Name )
    return
}

$parent = Get-Process -Id (Get-WmiObject win32_process -filter "processid=$pid").ParentProcessId -ErrorAction SilentlyContinue
[dateTime]$startTime = (Get-Date).AddSeconds( -$pollperiod )

if( $parent -And ( $parent.StartTime -ge $startTime ) )
{
    ## Probably just logged on so give it time to get logged in
    Write-Verbose "Sleeping as parent process started at $($parent.StartTime) so probably just logged on"
    Start-Sleep -Seconds $pollperiod
}

$client = Initialise-SkypeObject

$firstSawDoNotDisturb = $null
$global:snoozeUntil = $null

[void][Reflection.Assembly]::LoadWithPartialName('Presentationframework')

While( $true )
{
    ## see if we recently resumed from sleep/hibernate as this will probably cause the previous wait to complete immediately so network may not be back up and Skype signed in yet
    $startTime = (Get-Date).AddSeconds( -$pollperiod )
    [array]$events = @( Get-WinEvent -FilterHashtable @{Logname='System';ID=131,107;StartTime=$startTime;ProviderName='Microsoft-Windows-Kernel-Power'} -ErrorAction SilentlyContinue ) ## sleeps
    $events += @( Get-WinEvent -FilterHashtable @{Logname='System';ID=1;StartTime=$startTime;ProviderName='Microsoft-Windows-Power-Troubleshooter'} -ErrorAction SilentlyContinue ) ## hibernations
    if( $events -and $events.Count -gt 0 )
    {
        Write-Verbose "Sleeping as got $($events.Count) power events after $startTime :`n $events"
        Start-Sleep -Seconds $pollperiod
    }
    
    Write-Verbose "$(Get-Date): client state is $($client.State)" ## See if we just resumed from sleep/hibernation
    if( $client.State -eq 'Invalid' )
    {
        ## This will happen if Skype exits
       $client = Initialise-SkypeObject
    }

    if( ! $client -or $client.State -ne 'SignedIn' )
    {
        if( ! [string]::IsNullOrEmpty( $soundFile) )
        {
            if( ! $sound )
            {
                Add-Type -AssemblyName PresentationCore 
                $sound = New-Object System.Windows.Media.MediaPlayer
                $sound.Open( $soundFile )
            }
            $sound.Play()
        }
        $response = [Microsoft.VisualBasic.Interaction]::MsgBox( "Lync client $($client.State) @ $(Get-Date -Format G). Abort monitoring?" , 'YesNo,SystemModal,Exclamation' , $MyInvocation.MyCommand.Name )
        if( $sound )
        {
            $sound.Stop()
        }

        if( $response -eq 'yes' )
        {
            break
        }
    }
    else
    {
        if( $excessiveDoNotDisturbPeriod -and $client.Self.Contact.GetContactInformation([Microsoft.Lync.Model.ContactInformationType]::Availability) -eq $availabilityId[ $doNotDisturbText ] )
        {
            Write-Verbose "$(Get-Date -Format G) : do not disturb detected, first noticed at $firstSawDoNotDisturb"
            if( $firstSawDoNotDisturb )
            {
                if( [datetime]::Now -ge $firstSawDoNotDisturb.AddMinutes( $excessiveDoNotDisturbPeriod ) -and ( ! $snoozeUntil -or [datetime]::Now -ge $snoozeUntil ) )
                {
                    $self = $client.Self
                    $status = $self.Contact.GetContactInformation([Microsoft.Lync.Model.ContactInformationType]::Activity)
                    $mainForm = Load-GUI $mainwindowXAML
                    
                    if( $mainForm )
                    {
                        $WPFbtnStayBusy.add_Click({
                            $_.Handled = $true
                            $snoozeUntil = $null
                            $mainForm.DialogResult = $true
                            $mainForm.Close()
                        })
                    
                        $WPFbtnChangeStatus.add_Click({
                            $_.Handled = $true
                            ## https://gallery.technet.microsoft.com/lync/Configuring-Lync-presence-5a8fa90a
                            $ContactInfo = New-Object 'System.Collections.Generic.Dictionary[Microsoft.Lync.Model.PublishableContactInformationType, object]' 
                            $ContactInfo.Add( [Microsoft.Lync.Model.PublishableContactInformationType]::Availability , $availabilityId[ $WPFcomboStatus.SelectedItem.Content ] )
                            $publish = $self.BeginPublishContactInformation( $ContactInfo , $null , $null ) 
                            $self.EndPublishContactInformation( $Publish )
                            $snoozeUntil = $null

                            $mainForm.DialogResult = $true
                            $mainForm.Close()
                        })
                    
                        $WPFbtnSnooze.add_Click({
                            $_.Handled = $true
                            $global:snoozeUntil = ([datetime]::Now).AddMinutes( $WPFtxtSnoozePeriod.Text )
                            $mainForm.DialogResult = $true
                            $mainForm.Close()
                        })

                        $WPFtxtNotificationArea.Text = "Lync client been in $status since at least $(Get-Date -Date $firstSawDoNotDisturb -Format G)"

                        $mainForm.Topmost = $true
                        $result = $mainForm.ShowDialog()
                    }
                    else
                    {
                        $response = [Microsoft.VisualBasic.Interaction]::MsgBox( "Unable to display dialogue for $status options" , 'OKOnly,SystemModal,Error' , $MyInvocation.MyCommand.Name )
                    }
                }
            }
            else
            {
                $firstSawDoNotDisturb = [datetime]::Now
            }
        }
        else
        {
            $firstSawDoNotDisturb = $null
        }
        Start-Sleep -Seconds $pollperiod
    }
}
