#requires -version 3

<#
.SYNOPSIS
    Show monitor resolutions and allow to pick one for mstsc launch where it uses most of the screen or other dimensions from parameters

.NOTES
    @guyrleech 2022/08/21

    Modification History:

    2022/08/24 @guyrleech  Added msrdc support
    2022/08/25 @guyrleech  Added GUI
    2022/09/23 @guyrleech  Fixed bug where : in .rdp file
    2022/10/13 @guyrleech  Added Other RDP options capability & persist to registry. Added username text box
    2022/10/14 @guyrleech  Added editable Comment column to display view
    2022/11/09 @guyrleech  Detect if window maximised and undo that before sizing/positioning.
                           Added -useOtherOptions otherwise tries to login locally with no password if AZ options in rdp file
                           Added logic to find existing windows as msrdc re-uses same process
                           Added Launch button on other tabs
    2022/11/14 @guyrleech  Fix for when UserFriendlyName is empty/null. Added BOE manufacturer   
    2022/11/15 @guyrleech  Change method of getting width and height as was relative to DPI scaling which made size too small
    2022/11/16 @guyrleech  Added VMware tab
    2022/12/14 @guyrleech  Changed -percentage to support x:y
    2022/12/19 @guyrleech  Added code to work for profile install of msrdc
    2023/02/21 @guyrleech  Persist computers list to HKCU
    2023/02/24 @guyrleech  Add VMware VMs to main computers list if connected to
    2023/03/10 @guyrleech  Added looking for msrdc.exe in program files x86
    2024/09/18 @guyrleech  Added Hyper-V support
    2024/09/19 @guyrleech  Added Hyper-V console button
    2024/09/23 @guyrleech  Added context menu for Hyper-V VMs
    2024/10/04 @guyrleech  Added more Hyper-V context menu items
                           Changed temp rdpfile naming & location
                           Uses primary monitor if no monitor selected
    2024/11/08 @guyrleech  Reverse name in .rdp file
                           Fix window title
    2024/11/12 @guyrleech  -remove added to remove characters from address for icon display differentiation improvement functionality
                           Added Hyper-V Clear Filter button
    2024/11/13 @guyrleech  Fixed not finding existing mstsc window
    2024/12/10 @guyrleech  Added snapshot management dialog
    2024/12/16 @guyrleech  Fixed VMware VM list not showing names
    2025/02/24 @guyrleech  Fixed snapshot issues
    2025/03/14 @guyrleech  Re-enable support for msrdc
    2025/03/24 @guyrleech  msrdc (Windows (365) App, was Remote Desktop (store) app) autodetection and greyed out if not available
    2025/03/25 @guyrleech  No Hyper-V host specified causes it to use localhost
    2025/03/26 @guyrleech  Added prompts to shutdown running VM before taking snapshot and to start after taking snapshot
                           Added nested virtualistion enablement
    2025/03/27 @guyrleech  Added message box for error if msrdc copy errors
    2025/04/09 @guyrleech  Added Buy Me A Coffee button
    2025/09/23 @guyrleech  Removed Hyper-V Connect button as Apply Filter does the same
    2025/05/29 @guyrleech  Added Window Title text box
    2025/06/19 @guyrleech  Added Azure tab
    2025/07/25 @guyrleech  Added long press functionality in Hyper-V VM list to clear selections and launch remote session if ctrl or alt pressed too
    2025/08/29 @guyrleech  Added DNS resolution to VM detail pane
    2025/09/08 @guyrleech  Fixed bug in enabling nested virtualisation
    2025/09/26 @guyrleech  Added keyboard handling radio buttons
    2025/10/05 @guyrleech  Double click fixed
    2026/05/14 @guyrleech  Added  -asyncActions
    2026/06/09 @guyrleech  Added Azure VMs
    2026/06/18 @guyrleech  Added AVD checkbox on Azure tab and AVD columns
    2026/06/30 @guyrleech  Added Azure context menu option to open selected VM(s) in Azure Portal
    2026/07/03 @guyrleech  Added Azure AVD Run context menu item to run PowerShell on selected VMs and show output
    2026/07/06 @guyrleech  Added support for Azure hibernate context menu action
    2026/07/07 @guyrleech  Changed AVD tab a lot
    2026/07/08 @guyrleech  Added AVD session process list view with remote kill context menu action
    2026/07/09 @guyrleech  Added Azure Config context menu item to change disk type on deallocated VMs
    2026/07/09 @guyrleech  Added OS disk type changing
    2026/07/16 @guyrleech  Added Azure Edit Tags context menu item to view/edit/add/delete VM tags
    2026/07/16 @guyrleech  Added Assigned User column to AVD list view
    2026/07/17 @guyrleech  Added Application Groups context menu under Host Pool showing apps and assignments in grid view. Added disk type at AZ top level
    
    ## TODO persist the "comment" column in memory so that it is available when undocked and redocked
    ## TODO make hypervisor operations async with a watcher thread
    ## TODO add history tab which is disabled by default (and thus audit)
    ## TODO add VMware console to that tab, make mstsc.exe configurable so could use with other exes
    ## TODO can we embed mstsx ax control so we can resize windows natively without mstsc.exe etc?
    ## TODO implement persistent tags so can make comments on machines in grid view. Persist to file so could be on a share

#>

<#
Copyright © 2026 Guy Leech

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, 
including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

[CmdletBinding()]

Param
(
    [string]$address ,
    [string]$displayModel ,
    [string]$displayManufacturer ,
    [string]$displayManufacturerCode ,
    [string]$username ,
    [string]$remoteDesktopName ,
    [string]$hypervHost ,
    [switch]$primary ,
    [switch]$asyncActions ,
    [string]$percentage ,
    [switch]$usemsrdc ,
    [switch]$noFriendlyName ,
    [switch]$keepRdpFile ,
    [switch]$noResize , ## use mstsc with no width/height parameters
    [string]$widthHeight , ## colon delimited
    [string]$xy , ## colon delimited
    [switch]$avd ,
    [decimal]$longPressSeconds = 1 ,
    [string]$sessionHostMgmtAPIVersion = '2026-04-01-preview' , ## hopefully will get a released version soon!
    [string]$drivesToRedirect = '*' ,
    [string]$extraMsrdcParameters = '/SkipAvdSignatureChecks' ,
    [string]$msrdcCopyPath ,
    [string]$msrdcCopyFolder ,
    [string]$msrdcCopyName = 'Copy of msrdc' ,
    [switch]$usegridviewpicker ,
    [switch]$fullScreen ,
    [string]$youAreHereSnapshot = 'YouAreHere'  ,
    [switch]$showDisplays ,
    [switch]$showManufacturerCodes ,
    [switch]$useOtherOptions ,
    [switch]$noMove ,
    [switch]$reverse ,
    [string]$remove , ## ^GL([AH]V)?[SW]\d+
    [int]$windowWaitTimeSeconds = 20 ,
    [ValidateRange(30,3600)]
    [int]$azureRunJobTimeoutSeconds = 120 ,
    [ValidateRange(30,3600)]
    [int]$azureAuthenticateTimeoutSeconds = 120 ,
    [int]$pollForWindowEveryMilliseconds = 333 ,
    [string]$tempFolder = $(Join-Path -Path $env:temp -ChildPath 'Guy Leech mstsc Sizer') ,
    [string]$exe = 'mstsc.exe' ,
    [string]$AVDAPIversion = '2022-09-09' ,
    [string]$configKey = 'HKCU:\SOFTWARE\Guy Leech\mstsc wrapper'
)

Function StuffToDo()
{
    ## add signing option with auto generate certificate or existing
    Get-ChildItem -Path Cert:\CurrentUser\My -CodeSigningCert|Where NotAfter -gt (Get-Date)
}
#region data

[array]$script:vms = $null
$script:vmwareConnection = $null
[array]$script:theseSnapshots = @()
$script:remoteSession = $null
$script:credentials = $null
$script:leftButtonClickedTime = $null
$script:targetItemData = $null
$script:azureColumnFilters = @{}
$script:azureColumnHeaders = @{}
$script:azureSelectedSubscription = $null
$script:azureLastGetAzVmCall = $null

# keep user added comments so can set when displays change
##$script:itemscopy = New-Object -TypeName System.Collections.Generic.List[object]

## https://docs.microsoft.com/en-us/windows-server/remote/remote-desktop-services/clients/rdp-files
## full address:s:$address ## removed since also using /v: so causes doubling of address in mstsc title bar
[string]$rdpTemplate = @'
desktopwidth:i:$width
desktopheight:i:$height
full address:s:$address
window title:s:$address
use multimon:i:0
screen mode id:i:$screenmode
dynamic resolution:i:1
smart sizing:i:0
drivestoredirect:s:$drivesToRedirect
'@

[string]$mainwindowXAML = @'
<Window x:Class="mstsc_msrdc_wrapper.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:mstsc_msrdc_wrapper"
        mc:Ignorable="d"
        Title="Guy's mstsc Wrapper Script" Height="520" Width="850" MinWidth="850" MinHeight="520">
    <Grid HorizontalAlignment="Stretch" VerticalAlignment="Stretch">
        <TabControl HorizontalAlignment="Stretch" VerticalAlignment="Stretch" x:Name="tabControl" >
        
            <TabItem Header="Main">
                <Grid  HorizontalAlignment="Stretch" VerticalAlignment="Stretch">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="23*"/>
                        <ColumnDefinition Width="50*"/>
                        <ColumnDefinition Width="103*"/>
                        <ColumnDefinition Width="68*"/>
                        <ColumnDefinition Width="518*"/>
                    </Grid.ColumnDefinitions>
                    <DataGrid x:Name="datagridDisplays" Grid.ColumnSpan="5" HorizontalAlignment="Stretch" VerticalAlignment="Top" Height="110" Margin="15,10,10,0" SelectionMode="Single" />
                    <Label Content="Computer" HorizontalAlignment="Left" Height="38" Margin="14,132,0,0" VerticalAlignment="Top" Width="71" Grid.ColumnSpan="3"/>
                    <CheckBox x:Name="chkboxmsrdc" Grid.Column="4" Content="Use msrdc instead of mstsc" HorizontalAlignment="Left" Height="21" Margin="145,189,0,0" VerticalAlignment="Top" Width="292" IsEnabled="true"/>
                    <ComboBox x:Name="comboboxComputer" Grid.Column="2" HorizontalAlignment="Left" Height="27" Margin="14,137,0,0" VerticalAlignment="Top" Width="254" IsEditable="True" IsDropDownOpen="False" Grid.ColumnSpan="3">
                        <ComboBox.ContextMenu>
                            <ContextMenu>
                                <MenuItem Header="Delete" x:Name="deleteComputersContextMenu"/>
                            </ContextMenu>
                        </ComboBox.ContextMenu>
                    </ComboBox>
                    <Label Content="Window&#xA;Title" HorizontalAlignment="Left" Height="46" Margin="14,320,0,0" VerticalAlignment="Top" Width="71" Grid.ColumnSpan="3"/>
                    <TextBox x:Name="txtboxWindowTitle" Grid.Column="2" HorizontalAlignment="Left" Height="26" Margin="14,330,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="254" Text="" Grid.ColumnSpan="3"/>
                    <CheckBox x:Name="chkboxPrimary" Grid.Column="4" Content="Use primary monitor" HorizontalAlignment="Left" Height="21" Margin="145,215,0,0" VerticalAlignment="Top" Width="292"/>
                    <TextBox x:Name="txtboxDrivesToRedirect" Grid.Column="2" HorizontalAlignment="Left" Height="26" Margin="14,230,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="254" Text="*" Grid.ColumnSpan="3"/>
                    <Label Content="Drive&#xA;Redirection" HorizontalAlignment="Left" Height="46" Margin="14,220,0,0" VerticalAlignment="Top" Width="71" Grid.ColumnSpan="3"/>
                    <CheckBox x:Name="chkboxNoMove" Grid.Column="4" Content="Do not move window" HorizontalAlignment="Left" Height="21" Margin="145,245,0,0" VerticalAlignment="Top" Width="292"/>
                    <RadioButton x:Name="radioFullScreen" Grid.Column="4" Content="Fullscreen" HorizontalAlignment="Left" Height="24" Margin="145,296,0,0" VerticalAlignment="Top" Width="206" GroupName="WindowSize"/>
                    <RadioButton x:Name="radioPercentage" Grid.Column="4" Content="Screen Percentage (X:Y)" HorizontalAlignment="Left" Height="24" Margin="145,272,0,0" VerticalAlignment="Top" Width="206" GroupName="WindowSize"/>
                    <RadioButton x:Name="radioWidthHeight" Grid.Column="4" Content="Width &amp; Height" HorizontalAlignment="Left" Height="24" Margin="145,324,0,0" VerticalAlignment="Top" Width="206" GroupName="WindowSize" />
                    <TextBox x:Name="txtboxWindowPosition" Grid.Column="2" HorizontalAlignment="Left" Height="26" Margin="14,283,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="254" Text="0,0" Grid.ColumnSpan="3"/>
                    <Label Content="Window&#xA;Position" HorizontalAlignment="Left" Height="46" Margin="14,273,0,0" VerticalAlignment="Top" Width="71" Grid.ColumnSpan="3"/>
                    <TextBox x:Name="txtboxScreenPercentage" Grid.Column="4" HorizontalAlignment="Left" Height="23" Margin="314,273,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="158"/>
                    <TextBox x:Name="txtboxWidthHeight" Grid.Column="4" HorizontalAlignment="Left" Height="23" Margin="314,319,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="158">
                        <TextBox.InputBindings>
                            <MouseBinding Gesture="LeftDoubleClick" />
                        </TextBox.InputBindings>
                    </TextBox>
                    <RadioButton x:Name="radioFillScreen" Grid.Column="4" Content="Fill Screen" HorizontalAlignment="Left" Height="24" Margin="145,348,0,0" VerticalAlignment="Top" Width="206" GroupName="WindowSize"/>
                    <CheckBox x:Name="chkboxRdpSigning" Grid.Column="4" Content="RDP File Signing" HorizontalAlignment="Left" Height="21" Margin="145,378,0,0" VerticalAlignment="Top" Width="160" ToolTip="Sign the RDP file before launching so mstsc/msrdc does not show an untrusted publisher warning"/>
                    <ComboBox x:Name="comboboxSigningCert" Grid.Column="4" HorizontalAlignment="Left" Height="25" Margin="314,375,0,0" VerticalAlignment="Top" Width="250" IsEnabled="False" ToolTip="Code signing certificate to use for RDP file signing"/>

                    <Grid Grid.ColumnSpan="5" VerticalAlignment="Bottom" Margin="5">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*" />
                            <ColumnDefinition Width="*" />
                            <ColumnDefinition Width="*" />
                            <ColumnDefinition Width="*" />
                        </Grid.ColumnDefinitions>
                        <Button x:Name="btnLaunch" Content="_Launch" Grid.Column="0" Margin="5" Height="32" VerticalAlignment="Center"/>
                        <Button x:Name="btnRefresh" Content="_Refresh" Grid.Column="1" Margin="5" Height="32" VerticalAlignment="Center"/>
                        <Button x:Name="btnCreateShortcut" Content="_Create Shortcut" Grid.Column="2" Height="32"  Margin="5" VerticalAlignment="Center"/>
                        <Button x:Name="buttonBuyMeACoffee" Cursor="Pen" ToolTip="Buy me a coffee!" Grid.Column="3" Height="32" Margin="5" VerticalAlignment="Center">
                            <Viewbox Stretch="Uniform">
                                <Image x:Name="CoffeeImage2" Stretch="Fill"/>
                            </Viewbox>
                        </Button>
                    </Grid>
                    <Label Content="User" HorizontalAlignment="Left" Height="38" Margin="14,181,0,0" VerticalAlignment="Top" Width="71" Grid.ColumnSpan="3"/>
                    <TextBox x:Name="textboxUsername" Grid.Column="2" HorizontalAlignment="Left" Height="27" Margin="14,186,0,0" VerticalAlignment="Top" Width="254" Grid.ColumnSpan="3"/>
                </Grid>
            </TabItem>

            <TabItem Header="Mstsc Options">
                <Grid Margin="10" HorizontalAlignment="Stretch" VerticalAlignment="Stretch">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <!-- Checkboxes spanning both columns -->
                    <CheckBox x:Name="chkboxMultimon" Content="_Multi Monitor" Grid.Row="0" Grid.ColumnSpan="2" Height="28" Margin="5"/>
                    <CheckBox x:Name="chkboxSpan" Content="_Span" Grid.Row="1" Grid.ColumnSpan="2" Height="28" Margin="5"/>
                    <CheckBox x:Name="chkboxAdmin" Content="_Admin" Grid.Row="2" Grid.ColumnSpan="2" Height="28" Margin="5"/>
                    <CheckBox x:Name="chkboxPublic" Content="_Public" Grid.Row="3" Grid.ColumnSpan="2" Height="28" Margin="5"/>
                    <CheckBox x:Name="chkboxRemoteGuard" Content="Remote _Guard" Grid.Row="4" Grid.ColumnSpan="2" Height="28" Margin="5"/>
                    <CheckBox x:Name="chkboxRestrictedAdmin" Content="_Restricted Admin" Grid.Row="5" Grid.ColumnSpan="2" Height="28" Margin="5"/>
                    
                    <!-- Key Handling label in left column -->
                    <Label Content="Key Handling:" Grid.Row="6" Grid.Column="0" HorizontalAlignment="Left" VerticalAlignment="Center" FontWeight="Bold" Margin="5"/>
                    
                    <!-- Radio buttons indented to the right in left column -->
                    <RadioButton x:Name="radioKeysLocal" Content="Keys _Local" Grid.Row="7" Grid.Column="0" Height="24" Margin="25,5,5,5" GroupName="KeyHandling"/>
                    <RadioButton x:Name="radioKeysRemote" Content="Keys _Remote" Grid.Row="8" Grid.Column="0" Height="24" Margin="25,5,5,5" GroupName="KeyHandling" IsChecked="True"/>
                    <RadioButton x:Name="radioKeysFullScreen" Content="Keys Remote in _Fullscreen" Grid.Row="9" Grid.Column="0" Height="24" Margin="25,5,5,5" GroupName="KeyHandling"/>

                    <!-- Launch button -->
                    <Button x:Name="btnLaunchMstscOptions" Content="_Launch" Grid.Row="12" Grid.ColumnSpan="2" Height="32" Width="100" HorizontalAlignment="Center" Margin="10" IsDefault="True"/>
                </Grid>
            </TabItem>

            <TabItem Header="Other Options">
                <Grid x:Name="OtherRDPOptions" Margin="55,0,528,0"  HorizontalAlignment="Stretch" VerticalAlignment="Stretch">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="146*"/>
                        <ColumnDefinition Width="33*"/>
                    </Grid.ColumnDefinitions>
                    <CheckBox x:Name="chkboxDoNotSave" Content="Do Not Save" Width="196" Grid.Column="1" Margin="46,133,-209,154"/>
                    <CheckBox x:Name="chkboxDoNotApply" Content="Do Not Apply" Width="196" Grid.Column="1" Margin="46,106,-209,181"/>
                    <Label Content="Other RDP File Options:" HorizontalAlignment="Center" Height="49" Margin="0,19,0,0" VerticalAlignment="Top" Width="144"/>
                    <TextBox x:Name="txtBoxOtherOptions" HorizontalAlignment="Left" Margin="10,57,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Height="230" Width="168" ForceCursor="True" IsManipulationEnabled="True" AcceptsReturn="True"  VerticalScrollBarVisibility="Visible" Grid.ColumnSpan="2"/>
                    <Button x:Name="btnLaunchOtherOptions" Content="_Launch" HorizontalAlignment="Left" Height="25" VerticalAlignment="Bottom" Width="96" Margin="10,0,0,-19" IsDefault="True"/>
                </Grid>
            </TabItem>

            <TabItem Header="VMware" x:Name="tabVMware">
                 <Grid x:Name="gridVMware" Margin="10" HorizontalAlignment="Stretch" VerticalAlignment="Stretch">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto" />
                        <RowDefinition Height="Auto" />
                        <RowDefinition Height="*" />
                        <RowDefinition Height="Auto" />
                    </Grid.RowDefinitions>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="2*" />
                        <ColumnDefinition Width="1*" />
                        <ColumnDefinition Width="1*" />
                        <ColumnDefinition Width="1*" />
                    </Grid.ColumnDefinitions>
                    
                    <ListView x:Name="listViewVMwareVMs" Grid.Row="2" Grid.ColumnSpan="4" Margin="5" VerticalAlignment="Stretch" SelectionMode="Multiple">
                        <ListView.View>
                            <GridView>
                                <GridViewColumn Header="Name" DisplayMemberBinding="{Binding Name}"/>
                            </GridView>
                        </ListView.View>
                    </ListView>
                    
                    <TextBox x:Name="textBoxVMwarevCenter" Grid.Row="1" Grid.Column="0" Margin="5" TextWrapping="Wrap" VerticalAlignment="Top" />
                    <Label x:Name="labelVMwareVMs" Content="0 VMs" Grid.Row="1" Grid.Column="0" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="5"/>
                    <Label Content="Filter" Grid.Row="0" Grid.Column="1" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="5"/>
                    <CheckBox x:Name="checkBoxVMwareRegEx" Content="RegE_x" Grid.Row="1" Grid.Column="1" Margin="5,40,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="textBoxVMwareFilter"  Grid.Row="1" Grid.Column="1" Margin="5" TextWrapping="Wrap" VerticalAlignment="Top" />
                    <Label Content="vCenter" Grid.Row="0" Grid.Column="0" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="5" />
                    <RadioButton x:Name="radioButtonVMwareConnectByIP" Content="Connect by _IP" Grid.Row="1" Grid.Column="3" Margin="5,40,0,0"  GroupName="GroupBy"/>
                    <RadioButton x:Name="radioButtonVMwareConnectByName"   Content="Connect by _Name" Grid.Row="1" Grid.Column="3" Margin="5,70,0,0"  GroupName="GroupBy" IsChecked="True"/>
                    <Label Content="RDP Port" Grid.Row="0" Grid.Column="3" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="5" />
                    <TextBox x:Name="textBoxVMwareRDPPort" Grid.Row="1" Grid.Column="3" Margin="5" TextWrapping="Wrap" VerticalAlignment="Top"  />

                    <StackPanel Grid.Row="4" Grid.ColumnSpan="4" Orientation="Horizontal" HorizontalAlignment="Center" Margin="5">
                        <Button x:Name="buttonVMwareConnect" Content="_Connect" Width="100" Margin="5" />
                        <Button x:Name="buttonVMwareDisconnect" Content="_Disconnect" Width="100" Margin="5" />
                        <Button x:Name="btnLaunchVMwareOptions" Content="_Launch" Width="100" Margin="5" />
                        <Button x:Name="buttonVMwareApplyFilter" Content="Apply _Filter" Width="100" Margin="5" />
                    </StackPanel>
                </Grid>
            </TabItem>

            <TabItem Header="Hyper-V" x:Name="tabHyperV">
                <Grid x:Name="HyperVOptions" Margin="10" HorizontalAlignment="Stretch" VerticalAlignment="Stretch">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto" />
                        <RowDefinition Height="Auto" />
                        <RowDefinition Height="*" />
                        <RowDefinition Height="Auto" />
                    </Grid.RowDefinitions>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="2*" />
                        <ColumnDefinition Width="1*" />
                        <ColumnDefinition Width="1*" />
                        <ColumnDefinition Width="1*" />
                    </Grid.ColumnDefinitions>

                    <Label Content="Host" Grid.Row="0" Grid.Column="0" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="5" />
                    <TextBox x:Name="textBoxHyperVHost" Grid.Row="1" Grid.Column="0" Margin="5" TextWrapping="Wrap" VerticalAlignment="Top" />

                    <Label x:Name="labelHyperVVMs" Content="0 VMs" Grid.Row="1" Grid.Column="0" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="5" />
                    <Label Content="Filter" Grid.Row="0" Grid.Column="1" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="5" />
                    <TextBox x:Name="textBoxHyperVFilter" Grid.Row="1" Grid.Column="1" Margin="5" TextWrapping="Wrap" VerticalAlignment="Top" />
                    <CheckBox x:Name="checkBoxHyperVRegEx" Content="RegE_x" Grid.Row="1" Grid.Column="1" Margin="5,40,0,0" VerticalAlignment="Top" />
                    <CheckBox x:Name="checkBoxHyperVAllVMs" Content="_All VMs" Grid.Row="1" Grid.Column="1" Margin="5,70,0,0" VerticalAlignment="Top" />

                    <RadioButton x:Name="radioButtonHyperVConnectByIP" Content="Connect by _IP" Grid.Row="1" Grid.Column="3" Margin="5,40,0,0"  GroupName="GroupBy"/>
                    <RadioButton x:Name="radioButtonHyperVConnectByName"   Content="Connect by _Name" Grid.Row="1" Grid.Column="3" Margin="5,70,0,0"  GroupName="GroupBy" IsChecked="True"/>

                    <Label Content="RDP Port" Grid.Row="0" Grid.Column="3" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="5" />
                    <TextBox x:Name="textBoxHyperVRDPPort" Grid.Row="1" Grid.Column="3" Margin="5" TextWrapping="Wrap" VerticalAlignment="Top" />

                    <ListView x:Name="listViewHyperVVMs" Grid.Row="2" Grid.ColumnSpan="4" Margin="5" VerticalAlignment="Stretch" SelectionMode="Multiple">
                        <ListView.View>
                            <GridView>
                                <GridViewColumn Header="Name" DisplayMemberBinding="{Binding Name}" />
                                <GridViewColumn Header="Power State" DisplayMemberBinding="{Binding PowerState}" />
                            </GridView>
                        </ListView.View>
                        <ListView.ContextMenu>
                            <ContextMenu>
                                <MenuItem Header="Power" x:Name="PowerContextMenu">
                                    <MenuItem Header="Power On" x:Name="HyperVPowerOnContextMenu" />
                                    <MenuItem Header="Power Off" x:Name="HyperVPowerOffContextMenu" />
                                    <MenuItem Header="Shutdown" x:Name="HyperVShutdownContextMenu" />
                                    <MenuItem Header="Restart" x:Name="HyperVRestartContextMenu" />
                                    <MenuItem Header="Resume" x:Name="HyperVResumeContextMenu" />
                                    <MenuItem Header="Save" x:Name="HyperVSaveContextMenu" />
                                    <MenuItem Header="Suspend" x:Name="HyperVSuspendContextMenu" />
                                </MenuItem>
                                <MenuItem Header="Config" x:Name="ConfigContextMenu">
                                    <MenuItem Header="Run" x:Name="HyperVRunContextMenu" />
                                    <MenuItem Header="Detail" x:Name="HyperVDetailContextMenu" />
                                    <MenuItem Header="Rename" x:Name="HyperVRenameMenu" />
                                    <MenuItem Header="Reconfigure" x:Name="HyperVReconfigureMenu" />
                                    <MenuItem Header="Enable Nested Virtualisation" x:Name="HyperVEnableNestedVirtualisationContextMenu" />
                                    <MenuItem Header="Enable Resource Metering" x:Name="HyperVEnableResourceMeteringContextMenu" />
                                    <MenuItem Header="Performance Data" x:Name="HyperVMeasureContextMenu" />
                                    <MenuItem Header="Disable Resource Metering" x:Name="HyperVDisableResourceMeteringContextMenu" />
                                </MenuItem>
                                <MenuItem Header="Delete" x:Name="DeletionContextMenu">
                                    <MenuItem Header="Delete VM" x:Name="HyperVDeleteContextMenu" />
                                    <MenuItem Header="Delete VM + Disks" x:Name="HyperVDeleteAllContextMenu" />
                                </MenuItem>
                                <MenuItem Header="CD" x:Name="CDContextMenu">
                                    <MenuItem Header="Mount" x:Name="HyperVMountCDContextMenu" />
                                    <MenuItem Header="Eject" x:Name="HyperVEjectCDContextMenu" />
                                </MenuItem>
                                <MenuItem Header="Snapshots" x:Name="SnapshotsContextMenu">
                                    <MenuItem Header="Manage" x:Name="HyperVManageSnapshotContextMenu" />
                                    <MenuItem Header="Take Snapshot" x:Name="HyperVTakeSnapshotContextMenu" />
                                    <MenuItem Header="Revert to Latest Snapshot" x:Name="HyperVRevertLatestSnapshotContextMenu" />
                                    <MenuItem Header="Delete Latest Snapshot" x:Name="HyperVDeleteLatestSnapshotContextMenu" />
                                </MenuItem>
                                <MenuItem Header="New" x:Name="NewContextMenu">
                                    <MenuItem Header="Brand New" x:Name="HyperVNewVMContextMenu" />
                                    <MenuItem Header="Templated" x:Name="HyperVNewVMFromTemplateContextMenu" />
                                </MenuItem>
                                <MenuItem Header="Name to Clipboard" x:Name="HyperVNameToClipboard" />
                                <MenuItem Header="NICS" x:Name="NICSContextMenu">
                                    <MenuItem Header="Disconnect NIC" x:Name="HyperVDisconnectNICContextMenu" />
                                    <MenuItem Header="Connect To" x:Name="ConnectNICContextMenu">
                                        <MenuItem Header="Internal" x:Name="HyperVConnectNICInternalContextMenu" />
                                        <MenuItem Header="External" x:Name="HyperVConnectNICExternalContextMenu" />
                                        <MenuItem Header="Private" x:Name="HyperVConnectNICPrivateContextMenu" />
                                    </MenuItem>
                                </MenuItem>
                            </ContextMenu>
                        </ListView.ContextMenu>
                    </ListView>

                    <StackPanel Grid.Row="4" Grid.ColumnSpan="4" Orientation="Horizontal" HorizontalAlignment="Center" Margin="5">
                        <Button x:Name="btnLaunchHyperVOptions" Content="_Launch" Width="100" Margin="5" />
                        <Button x:Name="btnLaunchHyperVConsole" Content="C_onsole" Width="100" Margin="5" />
                      <!-- Apply Filter does the same  <Button x:Name="btnConnectHyperV" Content="_Connect" Width="100" Margin="5" />   -->
                        <Button x:Name="buttonHyperVApplyFilter" Content="Apply _Filter" Width="100" Margin="5" />
                        <Button x:Name="buttonHyperVClearFilter" Content="Clea_r Filter" Width="100" Margin="5" />
                        <Button x:Name="buttonHyperBuyMeACoffee" Cursor="Pen" ToolTip="Buy me a coffee!">
                            <Viewbox Stretch="Uniform">
                                <Image x:Name="CoffeeImage" Stretch="Fill"/>
                            </Viewbox>
                        </Button>
                    </StackPanel>
                </Grid>
            </TabItem>
            
            <TabItem Header="Azure" x:Name="tabAzure">
                <Grid x:Name="AzureOptions" Margin="10" HorizontalAlignment="Stretch" VerticalAlignment="Stretch">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto" />
                        <RowDefinition Height="Auto" />
                        <RowDefinition Height="*" />
                        <RowDefinition Height="Auto" />
                    </Grid.RowDefinitions>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="1*" />
                        <ColumnDefinition Width="1*" />
                        <ColumnDefinition Width="1*" />
                        <ColumnDefinition Width="1*" />
                    </Grid.ColumnDefinitions>

                    <Label x:Name="labelAzureLastGetAzVmCall" Content="Last Fetch: n/a" Grid.Row="0" Grid.Column="0" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="5" />
                    <Label x:Name="labelAzureVMs" Content="0 VMs" Grid.Row="1" Grid.Column="0" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="5" />

                    <Label Content="Tenant" Grid.Row="0" Grid.Column="1" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="5" />
                    <Grid Grid.Row="1" Grid.Column="1" Margin="5" VerticalAlignment="Top">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto" />
                            <RowDefinition Height="Auto" />
                            <RowDefinition Height="Auto" />
                        </Grid.RowDefinitions>
                        <TextBox x:Name="textBoxAzureTenant" Grid.Row="0" TextWrapping="Wrap" VerticalAlignment="Top" IsReadOnly="True" Background="#F0F0F0" />
                        <CheckBox x:Name="checkBoxAzureAllVMs" Content="_All VMs" Grid.Row="1" Margin="0,6,0,0" VerticalAlignment="Top" />
                        <CheckBox x:Name="checkBoxAzureAVD" Content="_AVD" Grid.Row="2" Margin="0,4,0,0" VerticalAlignment="Top" />
                    </Grid>

                    <RadioButton x:Name="radioButtonAzureConnectByIP" Content="Connect by _IP" Grid.Row="1" Grid.Column="3" Margin="5,40,0,0"  GroupName="GroupBy" IsEnabled="False"/>
                    <RadioButton x:Name="radioButtonAzureConnectByName"   Content="Connect by _Name" Grid.Row="1" Grid.Column="3" Margin="5,70,0,0"  GroupName="GroupBy" IsChecked="True" IsEnabled="False"/>

                    <Label Content="RDP Port" Grid.Row="0" Grid.Column="3" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="5" />
                    <TextBox x:Name="textBoxAzureRDPPort" Grid.Row="1" Grid.Column="3" Margin="5" TextWrapping="Wrap" VerticalAlignment="Top" IsEnabled="False" />

                    <ListView x:Name="listViewAzureVMs" Grid.Row="2" Grid.ColumnSpan="4" Margin="5" VerticalAlignment="Stretch" SelectionMode="Multiple">
                        <ListView.View>
                            <GridView>
                                <GridViewColumn Header="Name" DisplayMemberBinding="{Binding Name}" />
                                <GridViewColumn Header="Location" DisplayMemberBinding="{Binding Location}" />
                                <GridViewColumn Header="Resource Group" DisplayMemberBinding="{Binding ResourceGroup}" />
                                <GridViewColumn Header="Subscription" DisplayMemberBinding="{Binding Subscription}" />
                                <GridViewColumn Header="Created" DisplayMemberBinding="{Binding Created}" />
                                <GridViewColumn Header="Power State" DisplayMemberBinding="{Binding PowerState}" />
                                <GridViewColumn Header="Size" DisplayMemberBinding="{Binding Size}" />
                                <GridViewColumn Header="OS Disk Type" DisplayMemberBinding="{Binding OSDiskType}" />
                            </GridView>
                        </ListView.View>
                        <ListView.ContextMenu>
                            <ContextMenu>
                                <MenuItem Header="Power" x:Name="AzurePowerContextMenu">
                                    <MenuItem Header="Power On" x:Name="AzurePowerOnContextMenu" />
                                    <MenuItem Header="Power Off &amp; Deallocate" x:Name="AzurePowerOffContextMenu" />
                                    <MenuItem Header="Shutdown &amp; Deallocate" x:Name="AzureShutdownContextMenu" />
                                    <MenuItem Header="Restart" x:Name="AzureRestartContextMenu" />
                                    <MenuItem Header="Hibernate" x:Name="AzureHibernateContextMenu" />
                                </MenuItem>
                                <MenuItem Header="Config" x:Name="AzureConfigContextMenu">
                                    <MenuItem Header="Run" x:Name="AzureRunContextMenu" />
                                    <MenuItem Header="Detail" x:Name="AzureDetailContextMenu" />
                                    <MenuItem Header="Open in Portal" x:Name="AzureOpenInPortalContextMenu" />
                                    <MenuItem Header="Extensions + Applications" x:Name="AzureExtensionsApplicationsContextMenu" />
                                    <MenuItem Header="Rename" x:Name="AzureRenameMenu" />
                                    <MenuItem Header="Reconfigure" x:Name="AzureReconfigureMenu" />
                                    <MenuItem Header="Change Disk Type" x:Name="AzureChangeDiskTypeContextMenu" />
                                    <MenuItem Header="Edit Tags" x:Name="AzureEditTagsContextMenu" />
                                    <MenuItem Header="Activity Logs" x:Name="AzureVMActivityLogsContextMenu" />
                                    <MenuItem Header="AVD Logs" x:Name="AzureAVDLogsContextMenu" IsEnabled="False" />
                                </MenuItem>
                                <MenuItem Header="Delete" x:Name="AzureDeletionContextMenu">
                                    <MenuItem Header="Delete VM" x:Name="AzureDeleteContextMenu" />
                                    <MenuItem Header="Delete Session Host" x:Name="AzureDeleteSessionHostContextMenu" />
                                    <MenuItem Header="Delete Session Host &amp; VM" x:Name="AzureDeleteSessionHostAndVMContextMenu" />
                                     <!-- <MenuItem Header="Delete VM + Disks" x:Name="AzureDeleteAllContextMenu" /> -->
                                </MenuItem>
                                <MenuItem Header="Sessions" x:Name="AzureSessionContextMenu">
                                    <MenuItem Header="Detail" x:Name="AzureDetailSessionContextMenu" />
                                    <MenuItem Header="Message" x:Name="AzureMessageSessionContextMenu" />
                                    <MenuItem Header="Disconnect" x:Name="AzureDisconnectSessionContextMenu" />
                                    <MenuItem Header="Logoff" x:Name="AzureLogoffSessionContextMenu" />
                                    <MenuItem Header="Force Logoff" x:Name="AzureForceLogoffSessionContextMenu" />
                                    <MenuItem Header="Drain Mode On" x:Name="AzureDrainModeOnSessionContextMenu" />
                                    <MenuItem Header="Drain Mode Off" x:Name="AzureDrainModeOffSessionContextMenu" />
                                </MenuItem>
                                <MenuItem Header="Host Pool" x:Name="AzureHostPoolContextMenu" IsEnabled="False">
                                    <MenuItem Header="Detail" x:Name="AzureHostPoolDetailContextMenu" />
                                    <MenuItem Header="Increase Size" x:Name="AzureChangeHostPoolSizeContextMenu" />
                                    <MenuItem Header="Activity Logs" x:Name="AzureHostPoolActivityLogsContextMenu" />
                                    <MenuItem Header="Application Groups" x:Name="AzureAppGroupsContextMenu" />
                                </MenuItem>
                                <!-- 
                                <MenuItem Header="Snapshots" x:Name="AzureSnapshotsContextMenu">
                                    <MenuItem Header="Manage" x:Name="AzureManageSnapshotContextMenu" />
                                    <MenuItem Header="Take Snapshot" x:Name="AzureTakeSnapshotContextMenu" />
                                    <MenuItem Header="Revert to Latest Snapshot" x:Name="AzureRevertLatestSnapshotContextMenu" />
                                    <MenuItem Header="Delete Latest Snapshot" x:Name="AzureDeleteLatestSnapshotContextMenu" />
                                </MenuItem>
                                <MenuItem Header="New" x:Name="AzureNewContextMenu">
                                    <MenuItem Header="Brand New" x:Name="AzureNewVMContextMenu" />
                                    <MenuItem Header="Templated" x:Name="AzureNewVMFromTemplateContextMenu" />
                                </MenuItem>
                                -->
                                <MenuItem Header="Name to Clipboard" x:Name="AzureNameToClipboard" />
                                <MenuItem Header="NICS" x:Name="AzureNICSContextMenu">
                                    <MenuItem Header="Disconnect NIC" x:Name="AzureDisconnectNICContextMenu" />
                                    <MenuItem Header="Connect To" x:Name="AzureConnectNICContextMenu">
                                        <MenuItem Header="Internal" x:Name="AzureConnectNICInternalContextMenu" />
                                        <MenuItem Header="External" x:Name="AzureConnectNICExternalContextMenu" />
                                        <MenuItem Header="Private" x:Name="AzureConnectNICPrivateContextMenu" />
                                    </MenuItem>
                                </MenuItem>
                            </ContextMenu>
                        </ListView.ContextMenu>
                    </ListView>

                    <StackPanel Grid.Row="4" Grid.ColumnSpan="4" Orientation="Horizontal" HorizontalAlignment="Center" Margin="5">
                        <Button x:Name="btnLaunchAzureOptions" Content="_Launch" Width="100" Margin="5" IsEnabled="False"/>
                        <Button x:Name="buttonAzureConnect" Content="_Authenticate" Width="100" Margin="5" IsEnabled="True" />
                                                <Button x:Name="buttonAzureSubscription" Content="_Subscription" Width="100" Margin="5" />
                      <!-- Apply Filter does the same  <Button x:Name="btnConnectAzure" Content="_Connect" Width="100" Margin="5" />   -->
                        <Button x:Name="buttonAzureApplyFilter" Content="_Refresh" Width="100" Margin="5" />
                        <Button x:Name="buttonAzureClearFilter" Content="_Clear Filter" Width="100" Margin="5" />
                        <Button x:Name="buttonAzureBuyMeACoffee" Cursor="Pen" ToolTip="Buy me a coffee!">
                            <Viewbox Stretch="Uniform">
                                <Image x:Name="CoffeeImageAzure" Stretch="Fill"/>
                            </Viewbox>
                        </Button>
                    </StackPanel>
                </Grid>
            </TabItem>

            <TabItem Header="Active Directory" x:Name="tabActiveDirectory" IsEnabled="true">
            
                 <Grid x:Name="gridActiveDirectory" Margin="10" HorizontalAlignment="Stretch" VerticalAlignment="Stretch">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto" />
                        <RowDefinition Height="Auto" />
                        <RowDefinition Height="Auto" />
                        <RowDefinition Height="*" />
                        <RowDefinition Height="Auto" />
                    </Grid.RowDefinitions>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="2*" />
                        <ColumnDefinition Width="1*" />
                        <ColumnDefinition Width="1*" />
                        <ColumnDefinition Width="1*" />
                    </Grid.ColumnDefinitions>
                    
                    <ListView x:Name="listViewAD" Grid.Row="3" Grid.ColumnSpan="4" Margin="5" VerticalAlignment="Stretch" SelectionMode="Multiple">
                        <ListView.View>
                            <GridView>
                                <GridViewColumn Header="Name" DisplayMemberBinding="{Binding Name}"/>
                                <GridViewColumn Header="Container" DisplayMemberBinding="{Binding Container}"/>
                            </GridView>
                        </ListView.View>
                    </ListView>
                    
                    <!-- put first so gets input cursor by default -->
                    <TextBox x:Name="textBoxADFilter"  Grid.Row="1" Grid.Column="1" Margin="5" TextWrapping="Wrap" VerticalAlignment="Top"  Width="Auto" HorizontalAlignment="Stretch"/>
                    <TextBox x:Name="textBoxDomainController" Grid.Row="1" Grid.Column="0" Margin="5" TextWrapping="Wrap" VerticalAlignment="Top" />
                    <Label x:Name="labelADComputers" Content="0 Machines" Grid.Row="2" Grid.Column="0" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="5"/>
                    <Label Content="Filter" Grid.Row="0" Grid.Column="1" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="5"/>
                    <!-- <CheckBox x:Name="checkBoxADRegEx" Content="RegE_x" Grid.Row="1" Grid.Column="1" Margin="5,40,0,0" VerticalAlignment="Top"/> -->
                    <CheckBox x:Name="checkBoxADRecurse" Content="_Recurse" Grid.Row="1" Grid.Column="3" Margin="5,40,0,0" VerticalAlignment="Top"/>
                    <Label Content="Domain Controller" Grid.Row="0" Grid.Column="0" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="5" />

                    <StackPanel Grid.Row="2" Grid.Column="1" Grid.ColumnSpan="2" Orientation="Horizontal" HorizontalAlignment="Stretch" >
                        <RadioButton x:Name="radioButtonADTypeGroup"    Content="_Group"     Margin="5,0" GroupName="SearchBy"/>
                        <RadioButton x:Name="radioButtonADTypeOU"       Content="_OU"        Margin="5,0" GroupName="SearchBy"/>
                        <RadioButton x:Name="radioButtonADTypeComputer" Content="_Computer"  Margin="5,0" GroupName="SearchBy" IsChecked="True"/>
                    </StackPanel>

                    <Label Content="RDP Port" Grid.Row="0" Grid.Column="3" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="5" />
                    <TextBox x:Name="textBoxADRDPPort" Grid.Row="1" Grid.Column="3" Margin="5" TextWrapping="Wrap" VerticalAlignment="Top"  />

                    <StackPanel Grid.Row="5" Grid.ColumnSpan="4" Orientation="Horizontal" HorizontalAlignment="Center" Margin="5">
                        <Button x:Name="buttonADSearch" Content="_Search" Width="100" Margin="5" />
                        <Button x:Name="btnLaunchADOptions" Content="_Launch" Width="100" Margin="5" />
                    </StackPanel>
                </Grid>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
'@
#>

[string]$textInputXAML = @'
<Window x:Class="WPF_Scratchpad.Window1"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WPF_Scratchpad"
        mc:Ignorable="d"
        Title="Window1" Height="450" Width="800" MinHeight="260" MinWidth="460">
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Label x:Name="lblInputTextLabel" Content="Label" Grid.Row="0" Margin="0,0,0,8" HorizontalAlignment="Stretch" VerticalAlignment="Top"/>

        <TextBox x:Name="textboxInputText" Grid.Row="1" Margin="0" TextWrapping="Wrap" HorizontalAlignment="Stretch" VerticalAlignment="Stretch"/>

        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button x:Name="btnInputTextOK" Content="OK" Height="32" MinWidth="90" Margin="0,0,8,0" IsDefault="True"/>
            <Button x:Name="btnInputTextCancel" Content="Cancel" Height="32" MinWidth="90" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
'@

[string]$comboSelectXAML = @'
<Window x:Class="WPF_Scratchpad.Window1"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WPF_Scratchpad"
        mc:Ignorable="d"
        Title="Window1" Height="220" Width="460" MinHeight="180" MinWidth="360" ResizeMode="NoResize">
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Label x:Name="lblComboHeader" Content="" Grid.Row="0" Margin="0,0,0,4" HorizontalAlignment="Stretch" FontWeight="Bold"/>
        <Label x:Name="lblComboCurrentValue" Content="" Grid.Row="1" Margin="0,0,0,8" HorizontalAlignment="Stretch"/>
        <Label x:Name="lblComboLabel" Content="New disk type:" Grid.Row="2" Margin="0,0,0,2" HorizontalAlignment="Stretch"/>
        <ComboBox x:Name="comboBoxSelect" Grid.Row="3" Margin="0,0,0,0" HorizontalAlignment="Stretch" IsReadOnly="True"/>

        <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button x:Name="btnComboOK" Content="OK" Height="32" MinWidth="90" Margin="0,0,8,0" IsDefault="True"/>
            <Button x:Name="btnComboCancel" Content="Cancel" Height="32" MinWidth="90" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
'@

[string]$azureTagsEditorXAML = @'
<Window x:Class="WPF_Scratchpad.TagsWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        mc:Ignorable="d"
        Title="Edit Azure Tags" Height="540" Width="640" MinHeight="300" MinWidth="440" WindowStartupLocation="CenterOwner">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header label -->
        <Label x:Name="lblTagsHeader" Content="" Grid.Row="0" FontWeight="Bold" Margin="0,0,0,4" FontSize="13"/>

        <!-- Column header row -->
        <Grid Grid.Row="1" Margin="0,0,0,2">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="210"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="32"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="Tag Name" FontWeight="Bold" Grid.Column="0" Margin="4,0,0,0"/>
            <TextBlock Text="Value" FontWeight="Bold" Grid.Column="1" Margin="2,0,0,0"/>
        </Grid>

        <!-- Scrollable tag rows -->
        <ScrollViewer Grid.Row="2" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Margin="0,0,0,4">
            <ItemsControl x:Name="tagsItemsControl"/>
        </ScrollViewer>

        <!-- Separator -->
        <Separator Grid.Row="3" Margin="0,2,0,6"/>

        <!-- Add new tag row -->
        <Grid Grid.Row="4" Margin="0,0,0,8">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="210"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBox x:Name="txtNewTagName" Grid.Column="0" Margin="0,0,4,0" ToolTip="New tag name" Height="26" VerticalContentAlignment="Center"/>
            <TextBox x:Name="txtNewTagValue" Grid.Column="1" Margin="0,0,4,0" ToolTip="New tag value" Height="26" VerticalContentAlignment="Center"/>
            <Button x:Name="btnAddTag" Content="+ _Add Tag" Grid.Column="2" Padding="8,2" Height="26"/>
        </Grid>

        <!-- OK / Cancel buttons -->
        <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="btnTagsOK" Content="OK" Width="90" Margin="0,0,8,0" Height="32" IsDefault="True"/>
            <Button x:Name="btnTagsCancel" Content="Cancel" Width="90" Height="32" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
'@

[string]$hostPoolSizeXAML = @'
<Window x:Class="WPF_Scratchpad.HostPoolSize"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        mc:Ignorable="d"
        Title="Change Host Pool Size" Height="240" Width="440" MinHeight="220" MinWidth="340"
        ResizeMode="NoResize" WindowStartupLocation="CenterOwner">
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Label x:Name="lblHostPoolSizeHeader" Content="" Grid.Row="0" FontWeight="Bold" FontSize="12" Margin="0,0,0,2"/>
        <Label x:Name="lblHostPoolCurrentSize" Content="" Grid.Row="1" Margin="0,0,0,10"/>
        <StackPanel Grid.Row="2" Orientation="Horizontal">
            <Label Content="New instance count:" Width="160" VerticalContentAlignment="Center" Padding="0"/>
            <TextBox x:Name="txtHostPoolNewSize" Width="90" Height="26" VerticalContentAlignment="Center"/>
        </StackPanel>
        <Label x:Name="lblHostPoolSizeError" Content="" Grid.Row="3" Foreground="Red" Margin="0,4,0,0" Height="20" Padding="0"/>
        <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,8,0,0">
            <Button x:Name="btnHostPoolSizeOK" Content="OK" Width="90" Margin="0,0,8,0" Height="32" IsDefault="True" IsEnabled="False"/>
            <Button x:Name="btnHostPoolSizeCancel" Content="Cancel" Width="90" Height="32" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
'@

[string]$hostPoolDetailXAML = @'
<Window x:Class="WPF_Scratchpad.HostPoolDetail"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        mc:Ignorable="d"
        Title="Host Pool Detail" Height="520" Width="720" MinHeight="300" MinWidth="500"
        ResizeMode="CanResize" WindowStartupLocation="CenterOwner">
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Label x:Name="lblHostPoolDetailHeader" Content="" Grid.Row="0" FontWeight="Bold" FontSize="12" Margin="0,0,0,8"/>
        <DataGrid x:Name="dgHostPoolDetail" Grid.Row="1"
                  AutoGenerateColumns="False" IsReadOnly="True"
                  CanUserSortColumns="True" CanUserResizeColumns="True"
                  ScrollViewer.VerticalScrollBarVisibility="Auto"
                  ScrollViewer.HorizontalScrollBarVisibility="Auto"
                  AlternatingRowBackground="#F5F5F5" GridLinesVisibility="Horizontal"
                  HeadersVisibility="Column">
            <DataGrid.Columns>
                <DataGridTextColumn Header="Property" Binding="{Binding Property}" Width="230"/>
                <DataGridTextColumn Header="Value" Binding="{Binding Value}" Width="*"/>
            </DataGrid.Columns>
        </DataGrid>
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,8,0,0">
            <Button x:Name="btnHostPoolDetailClose" Content="Close" Width="90" Height="32" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
'@

[string]$hostPoolActivityLogsXAML = @'
<Window x:Class="WPF_Scratchpad.HostPoolActivityLogs"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        mc:Ignorable="d"
        Title="Host Pool Activity Logs" Height="600" Width="960" MinHeight="400" MinWidth="600"
        ResizeMode="CanResize" WindowStartupLocation="CenterOwner">
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Label x:Name="lblHostPoolActLogsHeader" Content="" Grid.Row="0" FontWeight="Bold" FontSize="12" Margin="0,0,0,4"/>
        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,6">
            <Label Content="Go back:" VerticalContentAlignment="Center" Padding="0,0,6,0"/>
            <TextBox x:Name="txtHostPoolActLogsValue" Width="70" Height="26" VerticalContentAlignment="Center" Text="24"/>
            <RadioButton x:Name="radHostPoolActLogsHours" Content="Hours" IsChecked="True" Margin="8,0,0,0" VerticalContentAlignment="Center"/>
            <RadioButton x:Name="radHostPoolActLogsDays" Content="Days" Margin="8,0,0,0" VerticalContentAlignment="Center"/>
            <Button x:Name="btnHostPoolActLogsRetrieve" Content="Retrieve" Width="80" Height="26" Margin="16,0,0,0"/>
        </StackPanel>
        <Label x:Name="lblHostPoolActLogsStatus" Content="" Grid.Row="2" Foreground="Gray" Margin="0,0,0,4" Padding="0"/>
        <DataGrid x:Name="dgHostPoolActLogs" Grid.Row="3"
                  AutoGenerateColumns="False" IsReadOnly="True"
                  CanUserSortColumns="True" CanUserResizeColumns="True" CanUserReorderColumns="True"
                  ScrollViewer.VerticalScrollBarVisibility="Auto"
                  ScrollViewer.HorizontalScrollBarVisibility="Auto"
                  AlternatingRowBackground="#F5F5F5" GridLinesVisibility="Horizontal"
                  SelectionMode="Extended">
            <DataGrid.Columns>
                <DataGridTextColumn Header="Time"          Binding="{Binding Time}"          Width="140"/>
                <DataGridTextColumn Header="Caller"        Binding="{Binding Caller}"        Width="200"/>
                <DataGridTextColumn Header="Operation"     Binding="{Binding Operation}"     Width="240"/>
                <DataGridTextColumn Header="Status"        Binding="{Binding Status}"        Width="90"/>
                <DataGridTextColumn Header="Level"         Binding="{Binding Level}"         Width="80"/>
                <DataGridTextColumn Header="Resource Type" Binding="{Binding ResourceType}"  Width="180"/>
                <DataGridTextColumn Header="Description"   Binding="{Binding Description}"   Width="*"/>
            </DataGrid.Columns>
        </DataGrid>
        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,8,0,0">
            <Button x:Name="btnHostPoolActLogsClose" Content="Close" Width="90" Height="32" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
'@

[string]$vmActivityLogsXAML = @'
<Window x:Class="WPF_Scratchpad.VMActivityLogs"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        mc:Ignorable="d"
        Title="VM Activity Logs" Height="600" Width="960" MinHeight="400" MinWidth="600"
        ResizeMode="CanResize" WindowStartupLocation="CenterOwner">
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Label x:Name="lblVMActLogsHeader" Content="" Grid.Row="0" FontWeight="Bold" FontSize="12" Margin="0,0,0,4"/>
        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,6">
            <Label Content="Go back:" VerticalContentAlignment="Center" Padding="0,0,6,0"/>
            <TextBox x:Name="txtVMActLogsValue" Width="70" Height="26" VerticalContentAlignment="Center" Text="24"/>
            <RadioButton x:Name="radVMActLogsHours" Content="Hours" IsChecked="True" Margin="8,0,0,0" VerticalContentAlignment="Center"/>
            <RadioButton x:Name="radVMActLogsDays" Content="Days" Margin="8,0,0,0" VerticalContentAlignment="Center"/>
            <Button x:Name="btnVMActLogsRetrieve" Content="Retrieve" Width="80" Height="26" Margin="16,0,0,0"/>
        </StackPanel>
        <Label x:Name="lblVMActLogsStatus" Content="" Grid.Row="2" Foreground="Gray" Margin="0,0,0,4" Padding="0"/>
        <DataGrid x:Name="dgVMActLogs" Grid.Row="3"
                  AutoGenerateColumns="False" IsReadOnly="True"
                  CanUserSortColumns="True" CanUserResizeColumns="True" CanUserReorderColumns="True"
                  ScrollViewer.VerticalScrollBarVisibility="Auto"
                  ScrollViewer.HorizontalScrollBarVisibility="Auto"
                  AlternatingRowBackground="#F5F5F5" GridLinesVisibility="Horizontal"
                  SelectionMode="Extended">
            <DataGrid.Columns>
                <DataGridTextColumn Header="Time"          Binding="{Binding Time}"          Width="140"/>
                <DataGridTextColumn Header="Caller"        Binding="{Binding Caller}"        Width="200"/>
                <DataGridTextColumn Header="Operation"     Binding="{Binding Operation}"     Width="240"/>
                <DataGridTextColumn Header="Status"        Binding="{Binding Status}"        Width="90"/>
                <DataGridTextColumn Header="Level"         Binding="{Binding Level}"         Width="80"/>
                <DataGridTextColumn Header="Resource Type" Binding="{Binding ResourceType}"  Width="180"/>
                <DataGridTextColumn Header="Description"   Binding="{Binding Description}"   Width="*"/>
            </DataGrid.Columns>
        </DataGrid>
        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,8,0,0">
            <Button x:Name="btnVMActLogsClose" Content="Close" Width="90" Height="32" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
'@

[string]$avdLogsXAML = @'
<Window x:Class="WPF_Scratchpad.AVDLogs"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        mc:Ignorable="d"
        Title="AVD Logs" Height="600" Width="960" MinHeight="400" MinWidth="600"
        ResizeMode="CanResize" WindowStartupLocation="CenterOwner">
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Label x:Name="lblAVDLogsHeader" Content="" Grid.Row="0" FontWeight="Bold" FontSize="12" Margin="0,0,0,4"/>
        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,6">
            <Label Content="Go back:" VerticalContentAlignment="Center" Padding="0,0,6,0"/>
            <TextBox x:Name="txtAVDLogsValue" Width="60" Height="26" VerticalContentAlignment="Center" Text="24"/>
            <RadioButton x:Name="radAVDLogsHours" Content="Hours" IsChecked="True" Margin="8,0,0,0" VerticalContentAlignment="Center"/>
            <RadioButton x:Name="radAVDLogsDays" Content="Days" Margin="8,0,0,0" VerticalContentAlignment="Center"/>
            <Button x:Name="btnAVDLogsRetrieve" Content="Retrieve" Width="80" Height="26" Margin="16,0,0,0"/>
        </StackPanel>
        <Label x:Name="lblAVDLogsStatus" Content="" Grid.Row="2" Foreground="Gray" Margin="0,0,0,4" Padding="0"/>
        <DataGrid x:Name="dgAVDLogs" Grid.Row="3"
                  AutoGenerateColumns="True" IsReadOnly="True"
                  CanUserSortColumns="True" CanUserResizeColumns="True" CanUserReorderColumns="True"
                  ScrollViewer.VerticalScrollBarVisibility="Auto"
                  ScrollViewer.HorizontalScrollBarVisibility="Auto"
                  AlternatingRowBackground="#F5F5F5" GridLinesVisibility="Horizontal"
                  SelectionMode="Extended"/>
        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,8,0,0">
            <Button x:Name="btnAVDLogsClose" Content="Close" Width="90" Height="32" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
'@

[string]$appGroupsXAML = @'
<Window x:Class="WPF_Scratchpad.AppGroups"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        mc:Ignorable="d"
        Title="Application Groups" Height="640" Width="1100" MinHeight="400" MinWidth="700"
        ResizeMode="CanResize" WindowStartupLocation="CenterOwner">
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Label x:Name="lblAppGroupsHeader" Content="" Grid.Row="0" FontWeight="Bold" FontSize="12" Margin="0,0,0,4"/>
        <Label x:Name="lblAppGroupsStatus" Content="" Grid.Row="1" Foreground="Gray" Margin="0,0,0,4" Padding="0"/>
        <TabControl Grid.Row="2">
            <TabItem Header="Applications">
                <DataGrid x:Name="dgAppGroupsApplications"
                          AutoGenerateColumns="False" IsReadOnly="True"
                          CanUserSortColumns="True" CanUserResizeColumns="True" CanUserReorderColumns="True"
                          ScrollViewer.VerticalScrollBarVisibility="Auto"
                          ScrollViewer.HorizontalScrollBarVisibility="Auto"
                          AlternatingRowBackground="#F5F5F5" GridLinesVisibility="Horizontal"
                          SelectionMode="Extended">
                    <DataGrid.Columns>
                        <DataGridTextColumn Header="App Group"       Binding="{Binding AppGroup}"       Width="180"/>
                        <DataGridTextColumn Header="Group Type"      Binding="{Binding GroupType}"      Width="100"/>
                        <DataGridTextColumn Header="Workspace"       Binding="{Binding Workspace}"      Width="140"/>
                        <DataGridTextColumn Header="App Name"        Binding="{Binding AppName}"        Width="160"/>
                        <DataGridTextColumn Header="Display Name"    Binding="{Binding DisplayName}"    Width="160"/>
                        <DataGridTextColumn Header="Description"     Binding="{Binding Description}"    Width="160"/>
                        <DataGridTextColumn Header="App Type"        Binding="{Binding AppType}"        Width="90"/>
                        <DataGridTextColumn Header="Path"            Binding="{Binding FilePath}"       Width="*"/>
                    </DataGrid.Columns>
                </DataGrid>
            </TabItem>
            <TabItem Header="Assignments">
                <DataGrid x:Name="dgAppGroupsAssignments"
                          AutoGenerateColumns="False" IsReadOnly="True"
                          CanUserSortColumns="True" CanUserResizeColumns="True" CanUserReorderColumns="True"
                          ScrollViewer.VerticalScrollBarVisibility="Auto"
                          ScrollViewer.HorizontalScrollBarVisibility="Auto"
                          AlternatingRowBackground="#F5F5F5" GridLinesVisibility="Horizontal"
                          SelectionMode="Extended">
                    <DataGrid.Columns>
                        <DataGridTextColumn Header="App Group"      Binding="{Binding AppGroup}"      Width="180"/>
                        <DataGridTextColumn Header="Group Type"     Binding="{Binding GroupType}"     Width="100"/>
                        <DataGridTextColumn Header="Workspace"      Binding="{Binding Workspace}"     Width="140"/>
                        <DataGridTextColumn Header="Principal"      Binding="{Binding Principal}"     Width="220"/>
                        <DataGridTextColumn Header="Principal Type" Binding="{Binding PrincipalType}" Width="100"/>
                        <DataGridTextColumn Header="Role"           Binding="{Binding Role}"           Width="*"/>
                    </DataGrid.Columns>
                </DataGrid>
            </TabItem>
        </TabControl>
        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,8,0,0">
            <Button x:Name="btnAppGroupsClose" Content="Close" Width="90" Height="32" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
'@

[string]$snapshotsXAML = @'
<Window x:Class="mstsc_GUI.Snapshots"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:VMWare_GUI"
        mc:Ignorable="d"
        Title="Snapshots" Height="450" Width="800">
    <Grid>
        <TreeView x:Name="treeSnapshots" HorizontalAlignment="Left" Height="246" Margin="85,74,0,0" VerticalAlignment="Top" Width="471"/>
        <Grid Margin="593,88,85,128" >
            <Button x:Name="btnTakeSnapshot" Content="_Take Snapshot" HorizontalAlignment="Left" VerticalAlignment="Top" Width="92"/>
            <Button x:Name="btnDeleteSnapshot" Content="De_lete" HorizontalAlignment="Left" Margin="0,172,0,0" VerticalAlignment="Top" Width="92"/>
            <Button x:Name="btnRevertSnapshot" Content="_Revert" HorizontalAlignment="Left" Margin="0,83,0,0" VerticalAlignment="Top" Width="92"/>
            <Button x:Name="btnConsolidateSnapshot" Content="_Consolidate" HorizontalAlignment="Left" Margin="0,129,0,0" VerticalAlignment="Top" Width="92"/>
            <Button x:Name="btnDetailsSnapshot" Content="_Details" HorizontalAlignment="Left" Margin="0,42,0,0" VerticalAlignment="Top" Width="92"/>
        </Grid>
        <Button x:Name="btnSnapshotsOk" Content="OK" HorizontalAlignment="Left" Margin="95,365,0,0" VerticalAlignment="Top" Width="75" IsDefault="True"/>
        <Button x:Name="btnSnapshotsCancel" Content="Cancel" HorizontalAlignment="Left" Margin="198,365,0,0" VerticalAlignment="Top" Width="75" IsCancel="True"/>
        <Label x:Name="lblLastRevert" Content="Last Revert" HorizontalAlignment="Left" Height="29" Margin="85,23,0,0" VerticalAlignment="Top" Width="622"/>
        <Button x:Name="btnDeleteSnapShotTree" Content="Delete _Tree" HorizontalAlignment="Left" Margin="593,300,0,0" VerticalAlignment="Top" Width="92"/>
    </Grid>
</Window>
'@

[string]$buyMeACoffee = @'
/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQEBQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGBYUGBIUFRT/2wBDAQMEBAUEBQkFBQkUDQsNFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBT/wAARCAAcAIADASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD6F+IPxB1rxJ4nvpJL6eGCKZ44beOQqqKGIHA7+prmH1/UI1LPqNyqgZJadgAPfmn67/yG9Q/6+JP/AEI1m38FveaZcW08BnSVGR0BxvUggjOR1HFfz9Vr1a1eUpzd2+/mf1Dh8NQoYaMadNWS7eRnad8U5taikn02/N5apLJCJz4g0+3DMjFWwk12j4yDyVAPUZHNbHgfxbf/ABB1O603Rb+4utUtZXins4r5JnjKjJJaKRkIwRhgxBz1rj/hJ+zzpOt/DPw1qFx8DfF2sSXVjFMNQt9XtlS4VhlXAOoqRlSONowe1Q+E/BUXw4+L3i5bHwrq/ge6jtLIwW1/erNIsLiXJDR3EwBLq2Ru6bPWvdr0IQhLmw9WCj9puNnrbSze+60PgMtzjEY3GfV/3bveys9P8z0KTWdUidka+u1dTgqZnBB/Om/27qX/AEELr/v+3+Ndv8MNKgu9D8RXk1npl5PGbdIW1ZwsSEudx3E+g6DrVzxL8K/tniOI6c0FnZXt1HbwBMtG2Yg8kiHPKDmvPjgMTUoRr05N36dd2v0Pq55ngqOJlhq0FHl62Vnon69fwZ55/bmpf9BC6/7/ADf40f25qX/QQuv+/wA3+NdCvw3vZtVsLCK5haS6szfF2BCxRckFjg9QAfxqHwZ4IHi2DUriTUodNtdORZZ5ZULYQkjIA6njp71zLC4xzVOzu79eyu+p3PGZcqbq3VlZvTu7Lp1Zif25qX/QQuv+/wA3+NH9u6l/0ELr/v8At/jXZ698Ip9HhIi1W2vrsTwxmCJGGEmZhC+7pltv3e3rWX4z8AyeF9XsNPtrr+0ZbtF2bYjGd5YrtwexI4Pcc1dTB42inKadlbr3+ZnRzDLa8lGm0279LbavddvvMD+3NS/6CF1/3+b/ABo/t3Uv+ghdf9/m/wAa6XxX8N5fDOmQXceow6izztayQwRsGSVVy4H94D1os/hy194esbuLUYf7UvJI1h00jDMjuyK2fqjE8cDk1LwmMU3Tad0r7/8AB/Aax2XOnGrdcrdlp1+78djmv7c1LAzqF1/3+b/Gj+3dS/6CF1/3/b/GvTYvhEnhtoNSmvbfWrE2lw8irEQqsIXK4Jzn5l68dKxF8HJrlr4K0yzjjgvL6Geee42ZJTeTubHJwqNgfhXVLLsZBe+7SfS/ml6dTihm+X1Je5FOK3drW0k3o1fS34nG/wBu6kP+Yhdf9/2/xrQ0Lxzrvh3Uor201O5EiEEq8rMjj0YE8iuq1f4OtpOnazetqqPFZDMI8gjzsKrMDz8hG9Rg9TXnGMcVx1qeLwMouo3F7rX/AIJ6GHq4DM6clSSlHZ6d/VF7Xf8AkN6h/wBfEn/oRqGwgguLjZc3ItItrHzChbkDIGB6nj8a7j44+HrPw948ukskaOO4H2hkJyAzE5x6D2rz+ssTTeHxM6ctWmzowVVYvB06kNOaK9VoanwV+L/jbSvD+p+GYfG/hjQ7Tw3ftptrbajpBeYwGOOaNixukz8swXIUD5a5HQPFus/EDUtf8U63f2mo3t/fy26XFhbmCF4LdjbxlVLvwfLLfeP3qvz6VZXUpkms7eWQ9XeJWJ/EirEMMdvGscSLHGvARBgD8K2rZljK8Zwq1nKErWi7Wjbe1km7vuzxcv4eoYDFyxUbX1tvpf5/odJ4e8WQaPpN3p11pUGqW1xKkxWWRkwyggY2/U11Vj8cby1tIo5NHsppLZyLQ5ZY7eIqFaNVHqAfmzn5j1rzSilSzHE0Eo05Wt5I9SvlGCxUnOtC7bvu/Tv237n0H4G8XnVfDms+IL/S9Pt7eG2ks0jtoiDJGkZYI7EnCAYUYGSSK8in8d3VxZ+IIpLeJX1jyVd4vkWJIyCEVfTAA+grCi1O8t7OWziu547SU5kgWRhG59SucH8aq114nNqteFOKdmk7vu3v+B5+CyKhhqtWckmpNWWuiVml9/5HoOj/ABdns7y/u9Q0u21GaeS3miXe0aRSQghDjnI5zjPUVW8R/E2TU7+2uNPtXs/KulvXNxM07yyr93JOPlUZAUetcRRXK8yxTp+zc9PRX777ncsmwMantVT123dtrbXtsegaj8XptR02S3/saztrhWla2nhLDyPNGJTtOdzHLc9smqGk/EiXSLbTGj0qzl1PT18qG+mDMVj3E7ducZ5Iz1wfxrjqKHmWKlLnc9fRBHJsDGHs1T033f8Antq9NtT274ffEafxJqM+jJpMK6UljMyaYjlzM7OpYbm56FuBwBXT+NvEkHgmK31O00y2kitVS0juIgE8yNmmDxRsB8hTapyPx6182QXEttKskUjRSA8MhII/GnPdTSRLE80jxqSQjMSAe5xXsUs+qwoOEleXR6adtLHgV+F6NTFKpCXLDrHXXvrfr3/4B6TL8YbX+zltYdDdRbzNcW/m3ruHlOD5kwx+8IYZHSvNZ5nuZ5JpG3SSMXZj3JOSaZXRfDzQ7bxH4w03T7wMbaaTDhDgkdcZrxamIr4+cKdSXktEvyPoaWEwuVU6lWlG2l3q29NerP/Z
'@

#endregion data
<#
    desktopscalefactor:i:200
    compression:i:0
#>

#region pre-main

## https://sirconfigmgr.de/display-inventory/
[hashtable]  $ManufacturerHash = @{ 
    'AAC' =	'AcerView'
    'ACR' = 'Acer'
    'AOC' = 'AOC'
    'AIC' = 'AG Neovo'
    'APP' = 'Apple Computer'
    'AST' = 'AST Research'
    'AUO' = 'Asus'
    'BNQ' = 'BenQ'
    'CMO' = 'Acer'
    'CPL' = 'Compal'
    'CPQ' = 'Compaq'
    'CPT' = 'Chunghwa Picture Tubes, Ltd.'
    'CTX' = 'CTX'
    'DEC' = 'DEC'
    'DEL' = 'Dell'
    'DPC' = 'Delta'
    'DWE' = 'Daewoo'
    'EIZ' = 'EIZO'
    'ELS' = 'ELSA'
    'ENC' = 'EIZO'
    'EPI' = 'Envision'
    'FCM' = 'Funai'
    'FUJ' = 'Fujitsu'
    'FUS' = 'Fujitsu-Siemens'
    'GSM' = 'LG Electronics'
    'GWY' = 'Gateway 2000'
    'HEI' = 'Hyundai'
    'HIT' = 'Hyundai'
    'HSL' = 'Hansol'
    'HTC' = 'Hitachi/Nissei'
    'HWP' = 'HP'
    'IBM' = 'IBM'
    'ICL' = 'Fujitsu ICL'
    'IVM' = 'Iiyama'
    'KDS' = 'Korea Data Systems'
    'LEN' = 'Lenovo'
    'LGD' = 'Asus'
    'LPL' = 'Fujitsu'
    'MAX' = 'Belinea' 
    'MEI' = 'Panasonic'
    'MEL' = 'Mitsubishi Electronics'
    'MS_' = 'Panasonic'
    'NAN' = 'Nanao'
    'NEC' = 'NEC'
    'NOK' = 'Nokia Data'
    'NVD' = 'Fujitsu'
    'OPT' = 'Optoma'
    'PHL' = 'Philips'
    'REL' = 'Relisys'
    'SAN' = 'Samsung'
    'SAM' = 'Samsung'
    'SBI' = 'Smarttech'
    'SGI' = 'SGI'
    'SNY' = 'Sony'
    'SRC' = 'Shamrock'
    'SUN' = 'Sun Microsystems'
    'SEC' = 'Hewlett-Packard'
    'TAT' = 'Tatung'
    'TOS' = 'Toshiba'
    'TSB' = 'Toshiba'
    'VSC' = 'ViewSonic'
    'ZCM' = 'Zenith'
    'UNK' = 'Unknown'
    '_YV' = 'Fujitsu'
    ## not in original
    'TMX' = 'Huawei'
    'HSD' = 'Hannspree'
    'BOE' = 'BOE Technology'
 }

 Function Convert-Base64ToImageSource {
    Param
    (
        [string]$base64
    )

    $bytes = [Convert]::FromBase64String($base64)
    $stream = New-Object System.IO.MemoryStream( , $bytes)
    $image = New-Object System.Windows.Media.Imaging.BitmapImage
    $image.BeginInit()
    $image.StreamSource = $stream
    $image.CacheOption = 'OnLoad'
    $image.EndInit()
    $image.Freeze()
    return $image
}

 Function Get-Msrdc
 {
    [string]$msrdc = $null

    if( -Not [string]::IsNullOrEmpty( $msrdcCopyPath ) )
    {
        $exe = $msrdcCopyPath
    }
    elseif( -Not ( Get-Command -Name ($exe = 'msrdc.exe') -CommandType Application -ErrorAction SilentlyContinue ) )
    {
        if( $apppathskey = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msrdc.exe' -ErrorAction SilentlyContinue ) 
        {
            if( $apppathskey.psobject.Properties[ '(default)' ] )
            {
                $exe = $apppathskey.'(default)'
            }
            elseif( $apppathskey.psobject.Properties[ 'path' ] )
            {
                $exe = Join-Path -Path $apppathskey.path -ChildPath 'msrdc.exe'
            }
            else
            {
                Throw "App Paths key found for msrdc.exe but it contains no usable paths"
            }
        }
        elseif( $installPath = Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.PSObject.Properties[ 'DisplayName' ] -and $_.DisplayName -eq 'Remote Desktop' -and $_.Publisher -eq 'Microsoft Corporation' } | Select-Object -ExpandProperty InstallLocation )
        {
            $exe = Join-Path -Path $installPath -ChildPath 'msrdc.exe'
        }
        elseif( $appx = Get-AppxPackage -Name '*.Windows365' | Sort-Object -Property Version -Descending | Select-Object -First 1 )
        {
            ## cannot execute it directly from here so we take a copy (hopefully Ivanti Application Control's Trusted Ownership won't bite us :-) )
            [string]$copyToFolder = $msrdcCopyFolder
            if( [string]::IsNullOrEmpty( $copyToFolder ) )
            {
                $copyToFolder = [System.IO.Path]::Combine( $env:LOCALAPPDATA , 'Programs' , $msrdcCopyName )
            }
            if( -Not (Test-Path -Path $copyToFolder) )
            {
                New-Item -Path $copyToFolder -ItemType Directory -Force ## if it errors, so be it
            }
            $appxMsrdcVersion = Get-ItemProperty -Path ( Join-Path $appx.InstallLocation -ChildPath 'msrdc\msrdc.exe' ) | Select-Object -ExpandProperty VersionInfo | Select-Object -ExpandProperty FileVersionRaw
            if( -Not ( $copyProperties = Get-ItemProperty -Path (Join-Path -Path $copyToFolder -ChildPath 'msrdc.exe') -ErrorAction SilentlyContinue) -or $appxMsrdcVersion -gt $copyProperties.VersionInfo.FileVersionRaw ) 
            {
                Write-Verbose "Copying from $($appx.InstallLocation ) to $copyToFolder"
                $copyErrors = $null
                Copy-Item -Path (Join-Path $appx.InstallLocation -ChildPath 'msrdc\*') -Destination $copyToFolder -Recurse -Force -ErrorVariable copyErrors
                if( -Not $? )
                {
                    [void][Windows.MessageBox]::Show( "Errors copying msrdc`n$($copyErrors -join "`n")", 'Error Copying msrdc' , 'Ok' ,'Error' )
                    ## TODO we could use a different folder but would have to change logic to find it.
                }
            }
            else
            {
                Write-Verbose -Message "msrdc copy folder `"$copyToFolder`" already exists"
            }
            $exe = Join-Path -Path $copyToFolder -ChildPath 'msrdc.exe' 
        }
        else
        {
            $exe = [System.IO.Path]::Combine( ([Environment]::GetFolderPath( [Environment+SpecialFolder]::ProgramFiles )) , 'Remote Desktop' , 'msrdc.exe' )
            if( -Not ( Test-Path -Path $exe -PathType Leaf ) )
            {
                $exe = [System.IO.Path]::Combine( ([Environment]::GetFolderPath( [Environment+SpecialFolder]::ProgramFilesX86 )) , 'Remote Desktop' , 'msrdc.exe' )
                ## TODO what if per user install?
            }
        }
    }

    if( -Not [string]::IsNullOrEmpty( $exe ) -and ( Test-Path -Path $exe -PathType Leaf -ErrorAction SilentlyContinue ) )
    {
        $exe ## return
    }
}

Function Set-WindowToFront
{
    Param
    (
        [Parameter(Mandatory)]
        [IntPtr]$windowHandle
    )
    
    ## first restore window
    if( [bool]$setForegroundWindow = [user32]::ShowWindowAsync( $windowHandle , 9 )) ## 9 = SW_RESTORE
    {
        ## https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowpos
        ## now set window top most to bring it to front
        if( $setForegroundWindow = [user32]::SetWindowPos( $windowHandle, [IntPtr]-1 , 0 ,0 , 0 , 0 , 0x4043 ) ) ## -1 = HWND_TOPMOST , -2 = HWND_NOTOPMOST , 0x4043 = SWP_ASYNCWINDOWPOS | SWP_SHOWWINDOW | SWP_NOSIZE | SWP_NOMOVE
        {
            ## now set window to not top most but will stay on top for now (otherwise would always be on top like task manager can be)
            $setForegroundWindow  = [user32]::SetWindowPos( $windowHandle, [IntPtr]-2 , 0 ,0 , 0 , 0 , 0x4043 ) 
        }
    }

    $setForegroundWindow ## return
}

Function Process-Snapshot
{
    Param
    (
        $GUIobject ,
        $Operation ,
        $VM
    )
    
    $_.Handled = $true
    
    [bool]$closeDialogue = $true

    if( $null -eq $VM )
    {
        [void][Windows.MessageBox]::Show( 'No VM passed to Process-Snapshot' , $Operation , 'Ok' ,'Error' )
        return
    }

    if( $operation -eq 'ConsolidateSnapshot' )
    {
        ## $VM = Get-VM -Id $VMId
        if( ! ( $task = $VM.ExtensionData.ConsolidateVMDisks_Task() ) `
            -or ! ($taskStatus = Get-Task -Id $task) `
                -or $taskStatus.State -eq 'Error' )
        {
            [void][Windows.MessageBox]::Show( 'Task Failed' , $Operation , 'Ok' ,'Error' )
        }
        else
        {
            [void][Windows.MessageBox]::Show( 'Task Started' , $Operation , 'Ok' ,'Information' )
        }
        $closeDialogue = $false
    }
    elseif( $Operation -eq 'TakeSnapShot' )
    {
        if( $takeSnapshotForm = New-Form -inputXaml $takeSnapshotXAML )
        {
            $takeSnapshotForm.Title += " of $($vm.name)"

            $WPFbtnTakeSnapshotOk.Add_Click({ 
                $_.Handled = $true
                $takeSnapshotForm.DialogResult = $true 
                $takeSnapshotForm.Close()  })

            if( $VM = Get-VM -Id $VMId )
            {
                if( $VM.PowerState -eq 'PoweredOff' ) 
                {
                    $wpfchkSnapshotQuiesce.IsEnabled = $false
                    $WPFchckSnapshotMemory.IsEnabled = $false
                }
                elseif( $VM.Guest.State -eq 'NotRunning' )
                {
                    $wpfchckSnapshotShutdownStart.IsEnabled = $false ## won't be able to shut it down cleanly so don't offer it
                }

                if( $takeSnapshotForm.ShowDialog() )
                {
                    if( $VM = Get-VM -Id $VMId )
                    {
                        ## Get Data from form and take snapshot
                        [hashtable]$parameters = @{ 'VM' = $vm ; 'RunAsync' = $true }
                        if( ! [string]::IsNullOrEmpty( $WPFtxtSnapshotName.Text ) )
                        {
                            $parameters.Add( 'Name' , $WPFtxtSnapshotName.Text )
                        }
                        else
                        {
                            return
                        }
                        if( ! [string]::IsNullOrEmpty( $WPFtxtSnapshotDescription.Text ) )
                        {
                            $parameters.Add( 'Description' , $WPFtxtSnapshotDescription.Text )
                        }
                        $parameters.Add( 'Quiesce' , $wpfchkSnapshotQuiesce.IsChecked )
                        $parameters.Add( 'Memory' , $WPFchckSnapshotMemory.IsChecked )
         
                        [string]$answer = 'yes'

                        if( $wpfchckSnapshotShutdownStart.IsChecked )
                        {
                            if( $VM.PowerState -ne 'PoweredOff' -and ( $answer = [Windows.MessageBox]::Show( "VM $($VM.Name)" , "Confirm Shutdown & Startup" , 'YesNo' ,'Question' ) ) -and $answer -ieq 'yes' )
                            {
                                $shutdownError = $null
                                if( -Not ( $guest = Shutdown-VMGuest -VM $VM -ErrorVariable shutdownError -Confirm:$false ) )
                                {
                                    $answer = 'abort'
                                    [void][Windows.MessageBox]::Show( $shutdownError , "Shutdown Error for $($VM.Name)" , 'Ok' ,'Error' )
                                }
                                else
                                {
                                    [datetime]$endWaitTime = [datetime]::Now.AddSeconds( $powerActionWaitTimeSeconds )
                                    Write-Verbose -Message "$(Get-Date -Format G): waiting for VM to shutdown until $(Get-Date -Date $endWaitTime -Format G)"
                                    do
                                    {
                                        Start-Sleep -Seconds 5
                                        $VM = Get-VM -Id $VMId
                                    }
                                    while( $VM -and $VM.PowerState -ne 'PoweredOff' -and [datetime]::Now -le $endWaitTime )
                                    
                                    Write-Verbose -Message "$(Get-Date -Format G): finished waiting for VM $($VM.Name) to shutdown - power state is $($VM|Select-Object -ExpandProperty PowerState)"

                                    if( -Not $VM -or $VM.PowerState -ne 'PoweredOff' )
                                    {
                                        [void][Windows.MessageBox]::Show( "Timed out waiting $powerActionWaitTimeSeconds seconds for shutdown to complete" , "Shutdown Error for $($VM.Name)" , 'Ok' ,'Error' )
                                        $answer = 'abort'
                                    }
                                    $parameters[ 'RunAsync' ] = $false
                                }
                            }
                        }

                        if( $answer -ieq 'yes' )
                        {
                            New-Snapshot @parameters
                            if( $? -and $wpfchckSnapshotShutdownStart.IsChecked )
                            {
                                Start-VM -VM $VM -RunAsync
                            }
                        }
                    }
                    else
                    {
                        [void][Windows.MessageBox]::Show( "Failed to get VM" , "Snapshot Error for $($VM.Name)" , 'Ok' ,'Error' )
                    }
                }
            }
            else
            {
                [void][Windows.MessageBox]::Show( "Failed to get VM" , "Snapshot Error for $($VM.Name)" , 'Ok' ,'Error' )
            }
        }
    }
    elseif( $Operation -eq 'DetailsSnapshot' )
    {
        $closeDialogue = $false
        if( $VM ) ## =  Get-VM -Id $VMId )
        {
            [string]$tag = $null
            if( $GUIobject.SelectedItem -and $GUIobject.PSObject.Properties[ 'SelectedItem' ] )
            {
                $tag = $GUIobject.SelectedItem.Tag
            }
            elseif( $GUIobject.Items.Count -eq 1 )
            {
                ## if only one snapshot then report on that one
                $tag = $GUIobject.Items[0].Tag
            }
            else
            {
                [void][Windows.MessageBox]::Show( 'No snapshot selected' , $Operation , 'Ok' ,'Error' )
                return
            }
            if( $tag -eq $youAreHereSnapshot )
            {
                return
            }
            if( $snapshot = $script:theseSnapshots | Where-Object Id -eq $tag)
            {
                [uint64]$size = 0
                ForEach( $disk in $snapshot.HardDrives  )
                {
                    $file = $null
                    ## needs backslashes escaping - file could be remote
                    $file = Get-CimInstance -ClassName cim_datafile -Filter "Name = '$($disk.Path -replace '\\' , '\\')'" -CimSession $snapshot.CimSession
                    if( $null -ne $file )
                    {
                        $size += $file.FileSize
                    }
                }
                [string]$details = "Name = $($snapshot.Name)`n`rNotes = $($snapshot.Notes)`n`rCreated = $($snapshot.CreationTime.ToString('G'))`n`rSize = $([math]::Round( $size / 1GB , 2))GB`n`rType = $($snapshot.SnapshotType)`n`rPower State = $($snapshot.State)`n`rAutomatic = $(if( $snapshot.IsAutomaticCheckpoint ) { 'Yes' } else {'No' })"
                [void][Windows.MessageBox]::Show( $details , 'Snapshot Details' , 'Ok' ,'Information' )
            }
            else
            {
                Write-Warning "Unable to get snapshot $($GUIobject.SelectedItem.Tag) for `"$($VM.Name)`""
            }
        }
        else
        {
            Write-Warning "Unable to get vm for vm id $vmid"
        }
    }
    elseif( ! $GUIobject -or ( $GUIobject.SelectedItem -and $GUIobject.SelectedItem.Tag -and $GUIobject.SelectedItem.Tag -ne $youAreHereSnapshot ) )
    {
        ##$VM = Get-VM -Id $VMId
        if( $VM )
        {
            if( $Operation -eq 'LatestSnapshotRevert' )
            {
                if( ! ( $snapshot = Get-VMSnapshot -VM $vm -verbose:$false | Sort-Object -Property CreationTime -Descending|Select-Object -First 1 ))
                {
                    [Windows.MessageBox]::Show( "No snapshots found for $($vm.Name)" , 'Snapshot Revert Error' , 'OK' ,'Error' )
                    return
                }
            }
            else
            {
                ##$snapshot = Get-VMSnapshot -Name $GUIobject.SelectedItem.Tag -VM $vm -verbose:$false
                $snapshot = $script:theseSnapshots | Where-Object Id -eq $GUIobject.SelectedItem.Tag
            }
            if( $snapshot )
            {
                [string]$answer = 'no'
                [string]$questionText = $null

                if( $Operation -eq 'DeleteSnapShotTree' )
                {
                    $questionText = 'From snapshot'
                }
                else
                {
                    $questionText = 'Snapshot'
                }

                $questionText += " `"$($snapshot.Name)`" on $($vm.Name), taken $($snapshot.CreationTime.ToString('G')) ?"
                $answer = [Windows.MessageBox]::Show( $questionText , "Confirm $($operation -creplace '([a-zA-Z])([A-Z])' , '$1 $2')" , 'YesNo' ,'Question' )
        
                if( $answer -eq 'yes' )
                {
                    if( $Operation -eq 'DeleteSnapShot' )
                    {
                        Remove-VMSnapshot -VMSnapshot $snapshot -AsJob -Confirm:$false
                    }
                    elseif( $Operation -eq 'DeleteSnapShotTree' )
                    {
                        Remove-VMSnapshot -VMSnapshot $snapshot -Confirm:$false -IncludeAllChildSnapshots -AsJob
                    }
                    elseif( $Operation -eq 'RevertSnapShot' -or $Operation -eq 'LatestSnapshotRevert' )
                    {
                        $answer = $null
                        if( $snapshot.State -ieq 'Off' )
                        {
                            $answer = [Windows.MessageBox]::Show( "Power on after snapshot restored on $($vm.Name)?" , 'Confirm Power Operation' , 'YesNo' ,'Question' )
                        }
                        [hashtable]$revertParameters = @{ 'VMSnapshot' = $snapshot ;'Confirm' = $false ; AsJob = ($answer -ne 'Yes') }
                        Restore-VMSnapshot @revertParameters
                        if( $answer -eq 'Yes' )
                        {
                            Start-VM -VM $vm -Confirm:$false
                        }
                    }
                    else
                    {
                        Write-Warning "Unexpected snapshot operation $operation"
                    }
                }
                else
                {
                    $closeDialogue = $false
                }
            }
            else
            {
                Write-Warning "Unable to get snapshot $($GUIobject.SelectedItem.Tag) for `"$($VM.Name)`""
            }
        }
        else
        {
            Write-Warning "Unable to get vm for vm id $vmid"
        }
    }
    else
    {
        $closeDialogue = $false
    }

    if( $closeDialogue -and (Get-Variable -Name snapshotsForm -ErrorAction SilentlyContinue) -and $snapshotsForm )
    {
        ## Close dialog since needs refreshing
        $snapshotsForm.DialogResult = $true 
        $snapshotsForm.Close()
        $snapshotsForm = $null
    }
}

Function Find-TreeItem
{
    Param
    (
        [array]$controls ,
        [string]$tag
    )

    $result = $null

    ForEach( $control in $controls )
    {
        if( $null -eq $result )
        {
            if( $control.Tag -eq $tag )
            {
                $result = $control
            }
            elseif( $control.PSobject.Properties[ 'Items' ] )
            {
                ForEach( $item in $control.Items )
                {
                    if( $item.Tag -eq $tag )
                    {
                        $result = $item
                    }
                    elseif( $item.PSobject.Properties[ 'Items' ] -and $item.Items.Count )
                    {
                        $result = Find-TreeItem -control $item.Items -tag $tag
                    }
                }
            }
        }
    }

    $result
}

## https://blog.ctglobalservices.com/powershell/kaj/powershell-wpf-treeview-example/

Function Add-TreeItem
{
    Param
    (
          $Name,
          $Parent,
          $Tag 
    )

    $ChildItem = New-Object System.Windows.Controls.TreeViewItem
    $ChildItem.Header = $Name
    $ChildItem.Name = $Name -replace '[/\s,;:\.\-\+)(]' , '_' -replace '%252f' , '_' -replace '&' , '_' # default snapshot names have / for date which are escaped
    $ChildItem.Tag = $Tag
    $ChildItem.IsExpanded = $true
    ##[Void]$ChildItem.Items.Add("*")
    [Void]$Parent.Items.Add($ChildItem)
}

Function Show-SnapShotWindow
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        $vmName
    )

    $thisVM = Hyper-V\Get-VM -Name $vmName @hypervParameters
    $script:theseSnapshots = @( Get-VMCheckpoint -VMName $vmName @hypervParameters )
    if( $null -eq $script:theseSnapshots -or $script:theseSnapshots.Count -eq 0)
    {
        [void][Windows.MessageBox]::Show( "No snapshots found for $vmName" , 'Snapshot Management' , 'Ok' ,'Warning' )
        return
    }

    $snapshotsForm = New-WPFWindow -inputXAML $snapshotsXAML

    [bool]$result = $false
    if( $snapshotsForm )
    {
        $snapshotsForm.Title += " for $vmName"

        ForEach( $snapshot in $script:theseSnapshots )
        {
            ## if has a parent we need to find that and add to it
            if( $snapshot.ParentSnapshotId )
            {
                ## find where to add our node
                $parent = $script:theseSnapshots | Where-Object Id -eq $snapshot.ParentSnapshotId
                if( $parent )
                {
                    $parentNode = Find-TreeItem -control $WPFtreeSnapshots -tag $snapshot.ParentSnapshotId
                    if( $parentNode )
                    {
                        Add-TreeItem -Name $snapshot.Name -Parent $parentNode -Tag $snapshot.Id 
                    }
                    else
                    {
                        Write-Warning "Unable to locate tree view item for parent snapshot `"$($snapshot.parent)`""
                    }
                }
                else
                {
                    Write-Warning "Unable to locate parent snapshot `"$($snapshot.Parent)`" for snapshot `"$($snapshot.Name)`" for $vmName"
                }
            }
            else ## no parent then needs to be top level node but check not already created because we enountered a child previously
            {
                Add-TreeItem -Name $snapshot.Name -Parent $WPFtreeSnapshots -Tag $snapshot.Id
            }
        }
        
        [string]$currentSnapShotId = $null
        if( $currentSnapShot = $script:theseSnapshots | Where-Object Id -eq $thisVM.ParentCheckpointId)
        {
            $currentSnapShotId = $currentSnapShot.Id
        }
        if( $currentSnapShotId )
        {
            if( ($currentSnapShotItem = Find-TreeItem -control $WPFtreeSnapshots -tag $currentSnapShotId ))
            {
                Add-TreeItem -Name '__You are here__' -Parent $currentSnapShotItem -Tag $youAreHereSnapshot
            }
            else
            {
                Write-Warning "Unable to locate tree view item for current snapshot"
            }
        }
        else
        {
            Write-Warning "No current snapshot set for $vmName"
        }

        $WPFbtnSnapshotsOk.Add_Click({
            $_.Handled = $true 
            $snapshotsForm.DialogResult = $true 
            $snapshotsForm.Close()  })

        ## get last revert operation
## TODO event logs?
<#
        $lastRevert = Get-VIEvent -Entity $vm.Name -ErrorAction SilentlyContinue | Where-Object { $_.PSObject.Properties[ 'EventTypeId' ] -and $_.EventTypeId -eq 'com.vmware.vc.vm.VmStateRevertedToSnapshot' -and $_.FullFormattedMessage -match 'has been reverted to the state of snapshot (.*), with ID \d' } | Select-Object -First 1
        [string]$text = $null
        if( $lastRevert )
        {
            $text = "Last revert was to snapshot `"$($Matches[1])`" on $(Get-Date -Date $lastRevert.CreatedTime -Format G)"
        }
        else
        {
            $text = "No snapshot revert event found"
        }
        $wpflblLastRevert.Content = $text
        ## see if consolidation is required so that we enable/disable the consolidation button
        $wpfbtnConsolidateSnapshot.IsEnabled = $vm.Extensiondata.Runtime.ConsolidationNeeded
                     
#>   
        $WPFbtnTakeSnapshot.Add_Click( { Process-Snapshot -GUIobject $WPFtreeSnapshots -Operation 'TakeSnapShot' -VM $thisVM } )
        $WPFbtnDeleteSnapshot.Add_Click( { Process-Snapshot -GUIobject $WPFtreeSnapshots -Operation 'DeleteSnapShot' -VM $thisVM} )
        $WPFbtnDeleteSnapShotTree.Add_Click( { Process-Snapshot -GUIobject $WPFtreeSnapshots -Operation 'DeleteSnapShotTree' -VM $thisVM} )
        $WPFbtnRevertSnapshot.Add_Click( { Process-Snapshot -GUIobject $WPFtreeSnapshots -Operation 'RevertSnapShot' -VM $thisVM} )
        $WPFbtnDetailsSnapshot.Add_Click( { Process-Snapshot -GUIobject $WPFtreeSnapshots -Operation 'DetailsSnapShot' -VM $thisVM} )
        $WPFbtnConsolidateSnapshot.Add_Click( { Process-Snapshot -GUIobject $WPFtreeSnapshots -Operation 'ConsolidateSnapShot' -VMId $thisVM } )

        $result = $snapshotsForm.ShowDialog()
     }
}

Function Get-GridViewColumnHeader
{
    Param
    (
        [Parameter(Mandatory=$true)]
        $eventArgs
    )

    $element = $eventArgs.OriginalSource -as [System.Windows.DependencyObject]
    while( $element -and $element -isnot [System.Windows.Controls.GridViewColumnHeader] )
    {
        $element = [System.Windows.Media.VisualTreeHelper]::GetParent( $element )
    }

    $element -as [System.Windows.Controls.GridViewColumnHeader]
}

Function Get-GridViewColumnBindingName
{
    Param
    (
        [Parameter(Mandatory=$true)]
        $column
    )

    if( $column.DisplayMemberBinding -and $column.DisplayMemberBinding.Path )
    {
        [string]$column.DisplayMemberBinding.Path.Path
    }
    elseif( $column.Header )
    {
        [string]$column.Header
    }
}

Function Sort-Columns
{
    Param
    (
        [Parameter(Mandatory=$true)]
        $control ,

        [Parameter(Mandatory=$true)]
        $eventArgs
    )

    $columnHeader = Get-GridViewColumnHeader -eventArgs $eventArgs
    if( $null -eq $columnHeader -or $null -eq $columnHeader.Column )
    {
        return
    }

    $itemsSource = $control.ItemsSource
    if( $null -eq $itemsSource )
    {
        $itemsSource = $control.Items
    }

    if( $view = [Windows.Data.CollectionViewSource]::GetDefaultView( $itemsSource ) )
    {
        Try
        {
            [string]$column = $columnHeader.Column.DisplayMemberBinding.Path.Path ## has to be name of the binding, not the header unless no binding
            if( [string]::IsNullOrEmpty( $column ) )
            {
                $column = $columnHeader.Column.Header
            }
            if( [string]::IsNullOrEmpty( $column ) )
            {
                return
            }

            [string]$direction = 'Ascending'
            if( $view.PSObject.Properties[ 'SortDescriptions' ] -and $view.SortDescriptions -and $view.SortDescriptions.Count -gt 0 )
            {
                $currentSort = $view.SortDescriptions[0]
                if( $currentSort.PropertyName -eq $column -and $currentSort.Direction -eq [System.ComponentModel.ListSortDirection]::Ascending )
                {
                    $direction = 'Descending'
                }
                $view.SortDescriptions.Clear()
            }

            $view.SortDescriptions.Add( ( New-Object ComponentModel.SortDescription( $column , $direction ) ) )
            $view.Refresh()
        }
        Catch
        {
        }
    }
}

Function Initialize-AzureColumnHeaders
{
    $azureGridView = $WPFlistViewAzureVMs.View -as [System.Windows.Controls.GridView]
    if( $null -eq $azureGridView )
    {
        return
    }

    ForEach( $column in $azureGridView.Columns )
    {
        [string]$bindingName = Get-GridViewColumnBindingName -column $column
        if( -Not [string]::IsNullOrEmpty( $bindingName ) -and -Not $script:azureColumnHeaders.ContainsKey( $bindingName ) )
        {
            $script:azureColumnHeaders[ $bindingName ] = [string]$column.Header
        }
    }
}

Function Update-AzureColumnHeaders
{
    $azureGridView = $WPFlistViewAzureVMs.View -as [System.Windows.Controls.GridView]
    if( $null -eq $azureGridView )
    {
        return
    }

    Initialize-AzureColumnHeaders

    ForEach( $column in $azureGridView.Columns )
    {
        [string]$bindingName = Get-GridViewColumnBindingName -column $column
        if( -Not [string]::IsNullOrEmpty( $bindingName ) )
        {
            [string]$baseHeader = $script:azureColumnHeaders[ $bindingName ]
            if( [string]::IsNullOrEmpty( $baseHeader ) )
            {
                $baseHeader = [string]$column.Header
                $script:azureColumnHeaders[ $bindingName ] = $baseHeader
            }

            if( $script:azureColumnFilters.ContainsKey( $bindingName ) -and -Not [string]::IsNullOrEmpty( $script:azureColumnFilters[ $bindingName ] ) )
            {
                $column.Header = "$baseHeader *"
            }
            else
            {
                $column.Header = $baseHeader
            }
        }
    }
}

Function Test-AzureColumnFilters
{
    Param
    (
        [Parameter(Mandatory=$true)]
        $item
    )

    ForEach( $filter in $script:azureColumnFilters.GetEnumerator() )
    {
        if( [string]::IsNullOrEmpty( $filter.Value ) )
        {
            continue
        }

        [string]$itemValue = ''
        if( $item.PSObject.Properties[ $filter.Key ] )
        {
            $itemValue = [string]$item.( $filter.Key )
        }

        [string]$pattern = $filter.Value
        if( $pattern -notmatch '[\*\?\[]' )
        {
            $pattern = "*$pattern*"
        }

        if( $itemValue -inotlike $pattern )
        {
            return $false
        }
    }

    return $true
}

Function Update-AzureVMLabel
{
    if( $null -ne $WPFlabelAzureLastGetAzVmCall )
    {
        if( $null -ne $script:azureLastGetAzVmCall )
        {
            $WPFlabelAzureLastGetAzVmCall.Content = "Last Fetch: $( $script:azureLastGetAzVmCall.ToString( 'G' ) )"
        }
        else
        {
            $WPFlabelAzureLastGetAzVmCall.Content = 'Last Fetch: n/a'
        }
    }

    [int]$totalCount = $WPFlistViewAzureVMs.Items.Count
    [int]$visibleCount = $totalCount

    if( $script:azureColumnFilters.Count -gt 0 )
    {
        $visibleCount = @( $WPFlistViewAzureVMs.Items | Where-Object { Test-AzureColumnFilters -item $_ } ).Count
    }

    if( $visibleCount -eq $totalCount )
    {
        $WPFlabelAzureVMs.Content = "$totalCount VMs"
    }
    else
    {
        $WPFlabelAzureVMs.Content = "$visibleCount of $totalCount VMs"
    }
}

Function Apply-AzureColumnFilters
{
    $itemsSource = $WPFlistViewAzureVMs.ItemsSource
    if( $null -eq $itemsSource )
    {
        $itemsSource = $WPFlistViewAzureVMs.Items
    }

    if( $view = [Windows.Data.CollectionViewSource]::GetDefaultView( $itemsSource ) )
    {
        if( $script:azureColumnFilters.Count -gt 0 )
        {
            $view.Filter = [Predicate[object]]{ param( $item ) Test-AzureColumnFilters -item $item }
        }
        else
        {
            $view.Filter = $null
        }
        $view.Refresh()
    }

    Update-AzureColumnHeaders
    Update-AzureVMLabel
}

Function Set-AzureColumnFilter
{
    Param
    (
        [Parameter(Mandatory=$true)]
        $columnHeader
    )

    if( $null -eq $columnHeader.Column )
    {
        return
    }

    [string]$bindingName = Get-GridViewColumnBindingName -column $columnHeader.Column
    if( [string]::IsNullOrEmpty( $bindingName ) )
    {
        return
    }

    Initialize-AzureColumnHeaders
    [string]$headerText = $script:azureColumnHeaders[ $bindingName ]
    if( [string]::IsNullOrEmpty( $headerText ) )
    {
        $headerText = [string]$columnHeader.Column.Header
    }

    $textInputWindow = New-Object -TypeName System.Windows.Window
    $textInputWindow.Title = "Filter $headerText"
    $textInputWindow.SizeToContent = [System.Windows.SizeToContent]::WidthAndHeight
    $textInputWindow.ResizeMode = [System.Windows.ResizeMode]::NoResize
    $textInputWindow.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterOwner

    if( $null -ne $mainWindow )
    {
        $textInputWindow.Owner = $mainWindow
    }

    $textInputBox = New-Object -TypeName System.Windows.Controls.TextBox
    $textInputBox.Width = 340
    $textInputBox.Margin = New-Object -TypeName System.Windows.Thickness -ArgumentList 10
    $textInputBox.Text = if( $script:azureColumnFilters.ContainsKey( $bindingName ) ) { $script:azureColumnFilters[ $bindingName ] } else { '' }

    $textInputWindow.Content = $textInputBox

    $textInputWindow.Add_ContentRendered({
        $textInputBox.Focus()
        $textInputBox.SelectAll()
    })

    $textInputWindow.Add_PreviewKeyDown({
        param( $windowObject , $keyEventArgs )

        if( $keyEventArgs.Key -eq [System.Windows.Input.Key]::Enter )
        {
            $windowObject.DialogResult = $true
            $keyEventArgs.Handled = $true
            $windowObject.Close()
        }
        elseif( $keyEventArgs.Key -eq [System.Windows.Input.Key]::Escape )
        {
            $windowObject.DialogResult = $false
            $keyEventArgs.Handled = $true
            $windowObject.Close()
        }
    })

    if( $textInputWindow.ShowDialog() )
    {
        [string]$filterValue = $textInputBox.Text.Trim()
        if( [string]::IsNullOrEmpty( $filterValue ) )
        {
            if( $script:azureColumnFilters.ContainsKey( $bindingName ) )
            {
                $script:azureColumnFilters.Remove( $bindingName )
            }
        }
        else
        {
            $script:azureColumnFilters[ $bindingName ] = $filterValue
        }
        Apply-AzureColumnFilters
    }
}

Function New-RemoteSession
{
    [CmdletBinding()]

    Param
    (
        [switch]$rethrow
    )
    
    [string]$commandLine = $null

    if( -Not [string]::IsNullOrEmpty( $address ) )
    {
        $commandLine = -join ($commandLine , " /v:$address" )
    }

    [int]$width = -1
    [int]$height = -1

    try
    {
        if( $fullScreen )
        {
            $commandLine = -join ($commandLine ,  ' /f' )
        }
        elseif( -Not $noResize )
        {
            [int]$reservedWidth  = $chosenDisplay.ScreenBounds.Width  - $chosenDisplay.ScreenWorkingArea.Width
            [int]$reservedHeight = $chosenDisplay.ScreenBounds.Height - $chosenDisplay.ScreenWorkingArea.Height
            if( -Not [string]::IsNullOrEmpty( $widthHeight ) )
            {
                [string[]]$dimensions = @( $widthHeight -split '[:x,\s]' )
                if( $dimensions.Count -ne 2 )
                {
                    Throw "Invalid parameter `"$widthHeight`" specified - must be width:height"
                }
                if( ( $width = $dimensions[0] -as [int] ) -le 0 )
                {
                    Throw "Invalid width in `"$widthHeight`""
                }
                if( ( $height = $dimensions[1] -as [int] ) -le 0 )
                {
                    Throw "Invalid height in `"$widthHeight`""
                }

            }
            elseif( -Not [string]::IsNullOrEmpty( $percentage ) )
            {
                [int]$xpercentage = 100
                [int]$ypercentage = 100
                [int]$percentageAsInt = -1

                if( $percentage -match '^(\d+):(\d+)' )
                {
                    $xpercentage = $Matches[1]
                    $ypercentage = $Matches[2]
                }
                elseif( [int]::TryParse( $percentage , [ref]$percentageAsInt ) )
                {
                    $xpercentage = $ypercentage = $percentageAsInt
                }
                else
                {
                    Write-Warning -Message "Percentage `"$percentage`" is not valid - ignoring"
                }
                $width  = $chosenDisplay.dmPelsWidth  * $xpercentage / 100
                $height = $chosenDisplay.dmPelsHeight * $ypercentage / 100
            }
            else ## fill the chosen screen
            {
                $width  = $chosenDisplay.dmPelsWidth  - $reservedWidth
                $height = $chosenDisplay.dmPelsHeight - $reservedHeight
            }
            <# ## doesn't work when font scaling not 100%
            elseif( $percentage -gt 0 )
            {
                $width  = $chosenDisplay.ScreenWorkingArea.Width  * $percentage / 100
                $height = $chosenDisplay.ScreenWorkingArea.Height * $percentage / 100
            }
            else ## fill the chosen screen
            {
                $width  = $chosenDisplay.ScreenWorkingArea.Width  - $widthReduction
                $height = $chosenDisplay.ScreenWorkingArea.Height - $heightReduction
            }
            #>
        }

        if( $width -gt 0 )
        {
            $commandLine = -join ($commandLine , " /w:$width" )
        }

        if( $height -gt 0 )
        {
            $commandLine = -join ($commandLine , " /h:$height" )
        }

        Write-Verbose -Message "Running $exe $commandLine"

        $process = $null
        
        ## recreate in case temp been tidied since script started
        if( -Not ( Test-Path -Path $tempFolder -PathType Container ) -and -Not ( New-Item -Path $tempFolder -ItemType Directory -Force ) )
        {
            Throw "Failed to create temp folder $tempFolder"
        }

        ## mstsc title bar uses filename before first dot and address so if we use address in filename it duplicates it
        ## as differentiator in host names is at the right hand side then reverse the name used in the rdp file so the last characters are first
        [string]$filename = $address.ToUpper()
        if( $reverse )
        {
            [array]$array = $address.ToCharArray()
            [array]::Reverse( $array )
            $filename = ($array -join '').ToUpper()
        }
        elseif( -Not [string]::IsNullOrEmpty( $remove ) )
        {
            $filename = $address -replace $remove ## designed to remove the leading characters of the name which will all be the same/similar eg GLHVS22 and makes icons difficult to distinguish
        }
        [string]$tempRdpFile = Join-Path -Path $tempFolder -ChildPath "$filename.rdp"
        [string]$windowTitle = "^$filename - $address - Remote Desktop Connection$"

        ## using our own folder so drop tmp from name since makes right click on mstsc taskbar icon look crap
        <#
        if( -Not ( $tempFile = New-TemporaryFile  ) )
        {
            Throw "Unable to create temporary file for rdp settings"
        }
        ## change file extension (probably .tmp but don't assume)
        $tempRdpFile = $tempFile.FullName -replace '\.[^\.]+$' , '.rdp'
        if( -Not ( Move-Item -Path $tempFile -Destination $tempRdpFile -PassThru ) )
        {
            Throw "Failed to move $tempFile to $tempRdpFile"
        }
        #>

        ## mstsc file will have things in it doesn't understand which it silently ignores
        
        [int]$screenmode = 1
        if( $fullScreen )
        {
            $screenmode = 2
        }
    
        [string]$rdpFileContents = $ExecutionContext.InvokeCommand.ExpandString( $rdpTemplate )

        [int]$keyboardHook = 0
        if( $WPFradioKeysRemote.IsChecked )
        {
            $keyboardHook = 1
        }
        elseif( $WPFradioKeysFullScreen.IsChecked )
        {
            $keyboardHook = 2
        }
        $rdpFileContents += "`nkeyboardhook:i:$keyboardHook`n"

    ##### TDODO implement multi monitor, etc mstsc options
    
        if( $usemsrdc )
        {
           $exe = Get-Msrdc
            
            if( [string]::IsNullOrEmpty( $remoteDesktopName ))
            {
                ## if there is no remote desktop name specified then the temp rdp file name is included in the Window title which is fugly
                $remoteDesktopName = "$address - Remote Desktop"
                $rdpFileContents += "`nremotedesktopname:s:$remoteDesktopName`n"
                $windowTitle = $remoteDesktopName
            }

            $commandLine = "`"$tempRdpFile`""
            if( -Not [string]::IsNullOrEmpty( $username ) )
            {
                $commandLine = "$commandLine /u:$username"
            }
            if( -Not [string]::IsNullOrEmpty( $extraMsrdcParameters ) )
            {
                $commandLine = "$commandLine $extraMsrdcParameters"
            }
            if( -Not $nofriendlyName )
            {
                [string]$newName = $wpftxtboxWindowTitle.Text
                if( [string]::IsNullOrEmpty( $newName ) )
                {
                    $newName = $filename
                }
                $commandLine = "$commandLine /friendlyname:`"$newName`""
                $windowTitle = $newName
            }
        }
        else ## mstsc
        {
        <#
            ## window title comes from the base name of the .rdp file so if we don't rename the temp file, that will be the name in the title bar which is fugly
            [string]$tempRdpFileWithName = Join-Path -Path (Split-Path -Path $tempRdpFile -Parent) -ChildPath "$($address -replace ':' , '.').$(Split-Path -Path $tempRdpFile -Leaf)"
            if( Move-Item -Path $tempRdpFile -Destination $tempRdpFileWithName -PassThru )
            {
                $tempRdpFile = $tempRdpFileWithName
            }
        #>
            $commandline = "`"$tempRdpFile`"" ## $commandLine" ## everything is in the rdp file
        }
        
        ## see if we already have a window with this title so we can offer to switch to that or create a new one
        $existingWindows = $null
        $existingWindows = [Api.Apidef]::GetWindows( -1 ) | Where-Object WinTitle -imatch $windowTitle
        $otherProcess = $null

        if( $null -ne $existingWindows )
        {
            Write-Verbose -Message "Already have window `"$windowTitle`" in process $($existingWindows.PID)"
            
            $otherprocess = Get-Process -Id $existingWindows.PID | Where-Object Name -in @( 'mstsc' , 'msrdc' , 'vmconnect' )
            try
            {
                $answer = [Windows.MessageBox]::Show(  "Activate Existing Window ?`nLaunched $($otherprocess.StartTime.ToString('G'))" , "Already Connected to $address" , 'YesNoCancel' ,'Question' )
            }
            catch
            {
                Write-Warning "Exception for existing process $($otherprocess.Name) (PID $($otherprocess.Id)) for $address : $_"
                $answer = 'no'
            }
            if( $answer -ieq 'yes' )
            {
                if( $otherprocess )
                {
                    if( -Not ( Set-WindowToFront -windowHandle $otherprocess.MainWindowHandle ))
                    {
                        [void][Windows.MessageBox]::Show( 'Failed to Activate Window' , "$($otherprocess.Name) (PID $($otherprocess.Id))" , 'Ok' ,'Error' )
                    }
                }
                else
                {
                    [void][Windows.MessageBox]::Show( 'Failed to Get Process' , "PID $($otherprocess.Id)" , 'Ok' ,'Error' )
                }
            
                return
            }
            elseif( $answer -ieq 'Cancel' )
            {
                return
            }
            else
            {
                $otherProcess = $null
            }
        }

        if( -Not [string]::IsNullOrEmpty( $rdpFileContents ) )
        {
            Write-Verbose -Message "Writing $($rdpFileContents.Length) bytes to $tempRdpFile"
            ## TODO do we need to make sure no duplicates?
            [string]$fullRdpContent = $rdpFileContents + "`n" + $(if( -Not $wpfchkboxDoNotApply.IsChecked ) { $WPFtxtBoxOtherOptions.Text })
            $fullRdpContent | Set-Content -Path $tempRdpFile -Force
            if( -Not $? )
            {
                Throw "Failed to write rdp file contents to $tempRdpFile"
            }

            ## Sign the RDP file if the user has opted in and a certificate is selected
            if( $WPFchkboxRdpSigning.IsChecked -and $null -ne $WPFcomboboxSigningCert.SelectedItem )
            {
                [string]$rdpsignExe = Join-Path -Path $env:SystemRoot -ChildPath 'System32\rdpsign.exe'
                if( Test-Path -Path $rdpsignExe -PathType Leaf )
                {
                    [string]$signingThumbprint = $WPFcomboboxSigningCert.SelectedItem.Tag
                    if( [string]::IsNullOrEmpty( $signingThumbprint ) )
                    {
                        [void][Windows.MessageBox]::Show( 'Could not retrieve certificate thumbprint from selection' , 'RDP Signing Error' , 'Ok' , 'Error' )
                    }
                    else
                    {
                        Write-Verbose -Message "Signing $tempRdpFile with certificate thumbprint $signingThumbprint"
                        $rdpsignProcess = Start-Process -FilePath $rdpsignExe -ArgumentList "/sha256 $signingThumbprint /q `"$tempRdpFile`"" -Wait -PassThru -WindowStyle Hidden
                        ## rdpsign.exe zeroes the file on failure, so always check and restore if needed
                        if( $null -eq $rdpsignProcess -or $rdpsignProcess.ExitCode -ne 0 -or (Get-Item -Path $tempRdpFile -ErrorAction SilentlyContinue).Length -eq 0 )
                        {
                            Write-Verbose -Message "rdpsign.exe failed (exit code $($rdpsignProcess.ExitCode)) - restoring RDP file content"
                            $fullRdpContent | Set-Content -Path $tempRdpFile -Force
                            [void][Windows.MessageBox]::Show( "rdpsign.exe failed (exit code $($rdpsignProcess.ExitCode))`nThe RDP file will not be signed" , 'RDP Signing Error' , 'Ok' , 'Warning' )
                        }
                    }
                }
                else
                {
                    [void][Windows.MessageBox]::Show( "rdpsign.exe not found at $rdpsignExe" , 'RDP Signing Error' , 'Ok' , 'Error' )
                }
            }
        }
        
        $process = $null

        ## use Start-Process so we can get a pid and thus window handle to move to chosen display
        $process = Start-Process -FilePath $exe -ArgumentList $commandLine -PassThru -WindowStyle Normal

        if( -Not $process )
        {
            Throw "Failed to launch $exe $commandLine"
        }

        try
        {
            [void]$process.WaitForInputIdle() ## but this may only be authentication
        }
        catch
        {
            Write-Warning -Message "$_"
        }

#      if( $process.HasExited -or -not $process.MainWindowHandle ) ## msrdc reuses existing process and new process exits so have to look for this window title in another process
#      {
        ## There may be more than one, either open or closed, so we need to find the new one which will be the one not in the existingWindows collection
        [int]$windowPid = -1
        [datetime]$endTime = [datetime]::MaxValue
        [string]$baseExe = (Split-Path -Path $exe -Leaf) -replace '\.[^\.]+$'
        if( $windowWaitTimeSeconds -gt 0 )
        {
            $endTime = [datetime]::Now.AddSeconds( $windowWaitTimeSeconds )
        }
        $existingProcesses = @()
        if( $usemsrdc )
        {
            $existingProcesses = @( Get-Process -Name msrdc -ErrorAction SilentlyContinue )
        }
        ## can take a little time for the window to appear and get the title so we poll :-(
## TODO change search string when connecting to IP - window title is "192 - 192.168.1.32 - Remote Desktop Connection"
        do
        {
            $allWindowsNow = @( [Api.Apidef]::GetWindows( -1 ) | Where-Object WinTitle -match $windowTitle )
            ForEach( $window in $allWindowsNow )
            {
                ## Need to find our new window, not any existing one
                if( ( $existingProcess = Get-Process -Id $window.PID -ErrorAction SilentlyContinue ) -and $existingProcess.Name -eq $baseExe  )
                {
                    ## if msrdc then it may have used an existing process :(
                    if( $existingProcess.StartTime -ge $process.StartTime -or $existingProcesses.Count -gt 0 )
                    {
                        $windowPid = $window.PID
                        break
                    }
                }
            }
            if( $windowPid -lt 0 )
            {
                if( $usemsrdc )
                {
                    ## TODO can't simply check if process has exited as msrdc can re-use existing plus if prompting for credentials it will be a process CredentialUIBroker.exe which is not a child of msrdc
                    if( $null -eq $existingProcesses -or $existingProcesses.Count -eq 0 )
                    {
                        Write-Warning -Message "Process $($process.Id) has exited and no previous instances"
                        break
                    }

                }
                else
                {
                    if( $process.HasExited )
                    {
                        Write-Warning -Message "Process $($process.Id) has exited"
                        break
                    }
                }
                Write-Verbose -Message "$(Get-Date -Format G): waiting until $(Get-Date -Date $endTime -Format G) for PID $($process.Id) to find window title `"$windowTitle`" for $baseExe"
                Start-Sleep -Milliseconds $pollForWindowEveryMilliseconds
            }
        }
        while( $windowPid -le 0 -and [datetime]::Now -lt $endTime )

        if( $windowPid -gt 0 )
        {
            $process = Get-Process -Id $windowPid
        }

        if( -not $process -or -Not $process.MainWindowHandle )
        {
            Write-Warning "No main window handle for process $($process.id)"
            return
        }

        if( -Not $noMove )
        {
            ## if window is maximized, undo that first so positioning & resizing works ok - msrdc seems to ignore -WindowStyle Normal
            if( [user32]::IsZoomed( $process.MainWindowHandle ) -and -Not $fullScreen )
            {
                Write-Verbose -Message "Window is maximised so undoing"
                ## 1 is SW_NORMAL
                $unmaximiseResult = [user32]::ShowWindowAsync( $process.MainWindowHandle, 1 ) ; $lastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()

                if( -Not $unmaximiseResult )
                {
                    Write-Warning -Message "Failed ShowWindowAsync to unmaximise - $lastError"
                }
            }
            ## https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowpos

            [int]$flags = 0x4041 ## SWP_NOSIZE (leave size alone ) | SWP_SHOWWINDOW | SWP_ASYNCWINDOWPOS

            [int]$x = $chosenDisplay.ScreenBounds.x
            [int]$y = $chosenDisplay.ScreenBounds.y

            if( -Not [string]::IsNullOrEmpty( $xy ) )
            {
                ## we change the coordinates arelatively, not absolutely otherwise the window may not change window
                [string[]]$dimensions = @( $xy -split ':' )
                [int]$deltax = 0
                [int]$deltay = 0
                if( $dimensions.Count -ne 2 )
                {
                    Throw "Invalid parameter `"$xy`" specified - must be x:y"
                }
                if( $null -eq ( $deltax = $dimensions[0] -as [int] ) )
                {
                    Throw "Invalid x coordinate in `"$xy`""
                }
                if( $null -eq ( $deltay = $dimensions[1] -as [int] ) )
                {
                    Throw "Invalid y coordinate in `"$xy`""
                }
                $x += $deltax
                $y += $deltay
            }

            Write-Verbose -Message "SetWindowPos x=$x y=$y"
            $result = [user32]::SetWindowPos( $process.MainWindowHandle , [IntPtr]::Zero , $x , $y , $width , $height , $flags ) ; $lastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()

            if( -Not $result) 
            {
                Write-Warning -Message "Failed SetWindowPos - $lastError"
            }

            ## if msrdc and fill screen (so not percentage or x/y) we make it maximised so it fills that screen
            if( $fullScreen -or ( $usemsrdc -and [string]::IsNullOrEmpty( $percentage ) -and [string]::IsNullOrEmpty( $widthHeight ) ) )
            {
                ## if it has been moved then we may need to maximise it again
                [int]$cmdShow = 3 ## SHOWMAXIMIZED

                ## https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-showwindowasync
                $styleResult = [user32]::ShowWindowAsync( $process.MainWindowHandle, [int]$cmdShow ) ; $lastError = [ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()

                if( -Not $styleResult )
                {
                    Write-Warning -Message "Failed ShowWindowAsync - $lastError"
                }
            }
        }

        if( $tempRdpFile -and -Not $keepRdpFile )
        {
            Remove-Item -Path $tempRdpFile
            $tempRdpFile = $null
        }
        ## TODO add computer to MRU ?
    }
    catch
    {
        if( $rethrow )
        {
            Throw $_
        }
        else
        {
            Write-Error -Message $_
            return $false
        }
    }
    finally
    {
        if( $tempRdpFile -and -Not $keepRdpFile )
        {
            Remove-Item -Path $tempRdpFile -Force -ErrorAction SilentlyContinue
        }
    }

    return $true
}
 

Function New-WPFWindow( $inputXAML )
{
    $form = $NULL
    [xml]$XAML = $inputXAML -replace 'mc:Ignorable="d"' , '' -replace 'x:N' ,'N'  -replace '^<Win.*' , '<Window'
  
    if( $reader = New-Object -TypeName Xml.XmlNodeReader -ArgumentList $xaml )
    {
        try
        {
            if( $Form = [Windows.Markup.XamlReader]::Load( $reader ) )
            {
                $xaml.SelectNodes( '//*[@Name]' ) | . { Process `
                {
                    Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName( $_.Name ) -Scope Script
                }}
            }
        }
        catch
        {
            Write-Error "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .NET is installed.`n$($_.Exception.InnerException)"
            $form = $null
        }
    }

    $form ## return
}


Function Set-RemoteSessionProperties
{
    [CmdletBinding()]

    Param
    (
        ## relying on variables in parent scope being available here as this function was originally the main code body and later moved to a function to support the GUI
        [string]$connectTo
    )
    
    if( -Not $wpfcomboboxComputer.SelectedItem -and [string]::IsNullOrEmpty( $wpfcomboboxComputer.Text ) -and [string]::IsNullOrEmpty( $connectTo ) )
    {
        [void][Windows.MessageBox]::Show( 'No Computer Selected' , 'Select a computer or enter a name/address' , 'Ok' ,'Error' )
    }
    ## get primary monitor
    <#
    elseif( ( $WPFdatagridDisplays.SelectedItems.Count -ne 1 -and $activeDisplaysWithMonitors.Count -gt 1 ) -and -not $WPFchkboxPrimary.IsChecked )
    {
        [void][Windows.MessageBox]::Show( 'No Monitor Selected' , 'Select a Monitor' , 'Ok' ,'Error' )
    }
    #>
    elseif( $WPFradioPercentage.IsChecked -and [string]::IsNullOrEmpty( $WPFtxtboxScreenPercentage.Text ))
    {
        [void][Windows.MessageBox]::Show( 'No Screen Percentage Entered' , 'Enter a screen percentage' , 'Ok' ,'Error' )
    }
    elseif( $wpfradioWidthHeight.IsChecked -and [string]::IsNullOrEmpty( $wpftxtboxWidthHeight.Text ))
    {
        [void][Windows.MessageBox]::Show( 'No Width & Height Entered' , 'Enter width and height' , 'Ok' ,'Error' )
    }
    else
    {
        if( -Not ( $chosen = $WPFdatagridDisplays.SelectedItems ) )
        {
            $chosen = $WPFdatagridDisplays.items | Where-Object -Property ScreenPrimary -eq 'true' -ErrorAction SilentlyContinue
            ##$chosen = $WPFdatagridDisplays.Items[0] ## no monitor selected but only one monitor
        }
        if( [string]::IsNullOrEmpty( $connectTo ) )
        {
            $address = $wpfcomboboxComputer.Text
        }
        {
            $address = $connectTo
        }

        if( $wpfcomboboxComputer.SelectedIndex -lt 0 ) ## manually entered
        {
            [bool]$alreadyPresent = $false

            ForEach( $item in $wpfcomboboxComputer.Items )
            {
                if( $alreadyPresent = $item -ieq $address )
                {
                    break
                }
            }
            if( -not $alreadyPresent )
            {
                $wpfcomboboxComputer.Items.Insert( 0 , $address ) ## TODO should we resort it ? Need to check if already there
            }
        }

        $noMove = $wpfchkboxNoMove.IsChecked
        $usemsrdc = $WPFchkboxmsrdc.IsChecked

        ## clear from previous runs
        $fullScreen = $false
        $widthHeight = $null
        $percentage = $null
        $widthHeight = $null
        $xy = $null

        if( $wpfradioFullScreen.IsChecked )
        {
            $fullScreen = $true
        }
        elseif( $wpfradioPercentage.IsChecked )
        {
            $percentage = ($wpftxtboxScreenPercentage.Text -replace '%')
        }
        elseif( $wpfradioWidthHeight.IsChecked )
        {
            $widthHeight = $wpftxtboxWidthHeight.Text.Trim() -replace '[\&x\s,]' , ':' 
            $fullScreen = $fullScreen
        }
        elseif( $wpfradioFillScreen.IsChecked )
        {
            ## this is the default 
        }

        if( -Not [string]::IsNullOrEmpty( $wpftxtboxWindowPosition.Text ) )
        {
            $xy = $wpftxtboxWindowPosition.Text -replace ',' , ':'
        }

        $username = $wpftextboxUsername.Text
            
        $drivesToRedirect = $wpftxtboxDrivesToRedirect.Text

        if( $WPFchkboxPrimary.IsChecked )
        {
            $chosenDisplay = $activeDisplaysWithMonitors | Where-Object ScreenPrimary -eq $true
        }
        else
        {
            $chosenDisplay = $activeDisplaysWithMonitors | Where ScreenDeviceName -eq $chosen.ScreenDeviceName
        }
        if( -Not $chosenDisplay )
        {
            Write-Warning -Message "Failed to find device name $($chosen.ScreenDeviceName) in internal data"
        }
        else
        {
            New-RemoteSession
        }
    }
}

Function Set-WindowContent
{
    [CmdletBinding()]

    Param
    (
    )
    
    ## copy existing comment items so we can associate again where possible
    $itemsCopy = $null
    if( $WPFdatagridDisplays.Items -and $WPFdatagridDisplays.Items.Count -gt 0 )
    {
        $itemsCopy = New-Object -TypeName object[] -ArgumentList $WPFdatagridDisplays.Items.Count
        $WPFdatagridDisplays.Items.CopyTo( $itemsCopy , 0 )
    }

    $WPFdatagridDisplays.Clear()

    $datatable = New-Object -TypeName System.Data.DataTable

    ForEach( $property in ($activeDisplaysWithMonitors | Select-Object -Property $displayFields -First 1).Psobject.Properties )
    {
        if( $column = $Datatable.Columns.Add( $property.Name , [string] ) ) ##$property.TypeNameOfValue ) )
        {
            $column.ReadOnly = $true
        }
    }
    
    if( $column = $Datatable.Columns.Add( 'Comment' , [string] ) )
    {
        $column.ReadOnly = $false ## TODO persist to registry?
    }

    ForEach( $row in ( $activeDisplaysWithMonitors | Select-Object -Property $displayFields ))
    {
        ## check if previously had a comment and add if it did. ScreenDeviceName could be different as changes when docked/undocked/docked
        if( $itemsCopy -and $itemsCopy.Count -gt 0 )
        {
            ## we have to deal with a potential empty row
            [string]$comment = $itemsCopy | Where-Object { $_.PSobject -and $_.PSObject.Properties -and $_.PSObject.Properties[ 'ScreenPrimary' ] -and $_.ScreenPrimary -eq $row.ScreenPrimary -and $_.Width -eq $row.Width -and $_.Height -eq $row.Height `
                -and $_.MonitorManufacturerName -eq $row.MonitorManufacturerName -and $_.MonitorManufacturerCode -eq $row.MonitorManufacturerCode -and $_.MonitorModel -eq $row.MonitorModel } | Select-Object -ExpandProperty Comment
            if( -Not [string]::IsNullOrEmpty( $comment ) )
            {
                Add-Member -InputObject $row -MemberType NoteProperty -Name Comment -Value $comment
            }
        }
        [void]$datatable.Rows.Add( @( $row.PSObject.Properties | Select-Object -ExpandProperty Value ) )
    }
  
    if( -Not [string]::IsNullOrEmpty( $percentage ) )
    {
        $WPFradioPercentage.IsChecked = $true
        $WPFtxtboxScreenPercentage.Text = $percentage
    }
    elseif( $fullScreen )
    {
        $WPFradioFullScreen.IsChecked = $true
    }
    elseif( -not [string]::IsNullOrEmpty( $widthHeight ) )
    {
        $WPFradioWidthHeight.IsChecked = $true
        $WPFtxtboxWidthHeight = $widthHeight -replace '[\s\&:]' , 'x'
    }
    else
    {
        $WPFradioFillScreen.IsChecked = $true
    }

    $wpftextboxUsername.Text = $username

    $WPFdatagridDisplays.ItemsSource = $datatable.DefaultView
    ##$WPFdatagridDisplays.IsReadOnly = $false
    $WPFdatagridDisplays.CanUserSortColumns = $true
    
    if( $null -ne $wpfcomboboxComputer.Items -and $wpfcomboboxComputer.Items.Count -eq 0 )
    {
        $previouslyUsed = @( Get-ItemProperty -Path 'HKCU:\SOFTWARE\Guy Leech\mstsc wrapper' -Name Computers -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Computers )
        if( $null -ne $previouslyUsed -and $previouslyUsed.Count -gt 0 )
        {
            ForEach( $value in $previouslyUsed )
            {
                [void]$wpfcomboboxComputer.Items.Add( $value )
            }
        }
        else
        {
            $mru = Get-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\Terminal Server Client\Default' -Name MRU* -ErrorAction SilentlyContinue | Select-Object -Property MRU*
            if( $null -ne $mru )
            {
                ForEach( $value in ($mru.PSobject.properties | Select-Object -ExpandProperty Value | Sort-Object ))
                {
                    [void]$wpfcomboboxComputer.Items.Add( $value )
                }
            }
        }
    }

    $wpftxtboxDrivesToRedirect.Text = $drivesToRedirect
    $wpfchkboxNoMove.IsChecked = $noMove
    $wpfchkboxPrimary.IsChecked = $primary
    if( -Not [string]::IsNullOrEmpty( $xy ) )
    {
        $txtboxWindowPosition.Text = $xy -replace '[:]' , ','
    }
}

Function Get-DisplayInfo
{
    [CmdletBinding()]

    Param
    (
    )
    
    ## get display info for each display
             
    [array]$Screens = @( [system.windows.forms.screen]::AllScreens )

    Write-Verbose -Message "Got $($Screens.Count) screens"

    ## https://rakhesh.com/powershell/powershell-snippet-to-get-the-name-of-each-attached-monitor/
    [array]$monitors = @( Get-CimInstance -ClassName WmiMonitorID -Namespace root\wmi )

    try
    {
        $monitors = @( Get-CimInstance -ClassName WmiMonitorID -Namespace root\wmi )
    }
    catch
    {
        Write-Warning -Message "Exception getting WmiMonitorID : $_"
    }

    $chosenDisplay = $null

    [array]$activeDisplaysDevices = @( [Resolution.Displays]::GetDisplays( 0 ) | Where-Object StateFlags -match 'AttachedToDesktop' )

    ForEach( $activeDisplaysDevice in $activeDisplaysDevices )
    {
        $result = [pscustomobject]$activeDisplaysDevice

        if( $monitorInfo = [Resolution.Displays]::GetDisplayDeviceInfo( $activeDisplaysDevice.DeviceName ) )
        {
            ForEach( $property in $monitorInfo.PSObject.Properties )
            {   
                if( $property.Name -ine 'cb' )
                {
                    Add-Member -InputObject $result -MemberType NoteProperty -Name "Monitor$($property.Name)" -Value $property.Value

                    if( $property.Name -ieq 'DeviceId' )
                    {
                        if( $property.Value -match '^MONITOR\\([^\\]+)\\' )
                        {
                            [string]$manufacturerCode = $Matches[1]
                            if( $monitor = $monitors | Where-Object InstanceName -Match "^DISPLAY\\$manufacturerCode\\" )
                            {
                                [string]$manufacturerCode = [System.Text.Encoding]::ASCII.GetString( $monitor.ManufacturerName ).Trim([char]0)
                                [string]$manufacturer = $ManufacturerHash[ $manufacturerCode ] 
                                if( [string]::IsNullOrEmpty( $manufacturer ) )
                                {
                                    Write-Warning -Message "No monitor manufacturer for code $manufacturerCode"
                                    $manufacturer = $manufacturerCode
                                }
                                Add-Member -InputObject $result -NotePropertyMembers @{
                                    MonitorModel = $(if( [string]::IsNullOrEmpty( $monitor.UserFriendlyName ) ){ 'Generic/Unknown' } else { [System.Text.Encoding]::ASCII.GetString( $monitor.UserFriendlyName ) })
                                    MonitorManufacturerCode = $manufacturerCode
                                    MonitorManufacturerName = $manufacturer
                                    MonitorProductCodeId = [System.Text.Encoding]::ASCII.GetString( $monitor.ProductCodeID )
                                }
                            }
                        }
                    }
                }
            }
        }
        else
        {
            Write-Warning -Message "Failed to get monitor info for device $($activeDisplaysDevice.DeviceName)"
        }
        if( $screen = $screens | Where-Object DeviceName -ieq $activeDisplaysDevice.DeviceName )
        {
            ForEach( $property in $screen.PSObject.Properties )
            {   
                Add-Member -InputObject $result -MemberType NoteProperty -Name "Screen$($property.Name)" -Value $property.Value -Force
            }
        }
        else
        {
            Write-Warning -Message "Failed to get screen info for device $($activeDisplaysDevice.DeviceName)"
        }
        ## this gives us the actual resolution in dmPelsWidth & dmPelsHeight
        if( $activeDisplaysDevice.DeviceName -and ( $displaySettings = [Resolution.Displays]::GetCurrentDisplaySettings( $activeDisplaysDevice.DeviceName ) ) )
        {
            ForEach( $property in $displaysettings.PSObject.Properties )
            {   
                ## all properties start dm* so won't clash
                Add-Member -InputObject $result -MemberType NoteProperty -Name $property.Name -Value $property.Value -Force
            }
        }
        if( $result.PSObject.Properties[ 'cb' ] )
        {
            $result.PSObject.Properties.Remove( 'cb' )
        }
        $result
    }
}

## adapted from https://gist.github.com/mintsoft/22a5ae4cc68d3e51b2f2

$pinvokeCode = @" 
using System; 
using System.Runtime.InteropServices; 
using System.Collections.Generic;
namespace Resolution 
{ 
    [StructLayout(LayoutKind.Sequential)] 
    public struct DEVMODE1 
    { 
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] 
        public string dmDeviceName; 
        public short dmSpecVersion; 
        public short dmDriverVersion; 
        public short dmSize; 
        public short dmDriverExtra; 
        public int dmFields; 
        public short dmOrientation; 
        public short dmPaperSize; 
        public short dmPaperLength; 
        public short dmPaperWidth; 
        public short dmScale; 
        public short dmCopies; 
        public short dmDefaultSource; 
        public short dmPrintQuality; 
        public short dmColor; 
        public short dmDuplex; 
        public short dmYResolution; 
        public short dmTTOption; 
        public short dmCollate; 
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] 
        public string dmFormName; 
        public short dmLogPixels; 
        public short dmBitsPerPel; 
        public int dmPelsWidth; 
        public int dmPelsHeight; 
        public int dmDisplayFlags; 
        public int dmDisplayFrequency; 
        public int dmICMMethod; 
        public int dmICMIntent; 
        public int dmMediaType; 
        public int dmDitherType; 
        public int dmReserved1; 
        public int dmReserved2; 
        public int dmPanningWidth; 
        public int dmPanningHeight; 
    }; 
	
	[Flags()]
	public enum DisplayDeviceStateFlags : int
	{
		/// <summary>The device is part of the desktop.</summary>
		AttachedToDesktop = 0x1,
		MultiDriver = 0x2,
		/// <summary>The device is part of the desktop.</summary>
		PrimaryDevice = 0x4,
		/// <summary>Represents a pseudo device used to mirror application drawing for remoting or other purposes.</summary>
		MirroringDriver = 0x8,
		/// <summary>The device is VGA compatible.</summary>
		VGACompatible = 0x10,
		/// <summary>The device is removable; it cannot be the primary display.</summary>
		Removable = 0x20,
		/// <summary>The device has more display modes than its output devices support.</summary>
		ModesPruned = 0x8000000,
		Remote = 0x4000000,
		Disconnect = 0x2000000
	}
	[StructLayout(LayoutKind.Sequential, CharSet=CharSet.Ansi)]
	public struct DISPLAY_DEVICE 
	{
		  [MarshalAs(UnmanagedType.U4)]
		  public int cb;
		  [MarshalAs(UnmanagedType.ByValTStr, SizeConst=32)]
		  public string DeviceName;
		  [MarshalAs(UnmanagedType.ByValTStr, SizeConst=128)]
		  public string DeviceString;
		  [MarshalAs(UnmanagedType.U4)]
		  public DisplayDeviceStateFlags StateFlags;
		  [MarshalAs(UnmanagedType.ByValTStr, SizeConst=128)]
		  public string DeviceID;
		[MarshalAs(UnmanagedType.ByValTStr, SizeConst=128)]
		  public string DeviceKey;
	}
    public class User_32 
    { 
        [DllImport("user32.dll", SetLastError=true)] 
        public static extern int EnumDisplaySettings(string deviceName, int modeNum, ref DEVMODE1 devMode); 
        [DllImport("user32.dll", SetLastError=true)] 
        public static extern int ChangeDisplaySettings(ref DEVMODE1 devMode, int flags); 
		[DllImport("user32.dll", SetLastError=true)]
		public static extern bool EnumDisplayDevices(string lpDevice, uint iDevNum, ref DISPLAY_DEVICE lpDisplayDevice, uint dwFlags);
        public const int ENUM_CURRENT_SETTINGS = -1; 
        public const int CDS_UPDATEREGISTRY = 0x01; 
        public const int CDS_TEST = 0x02; 
        public const int DISP_CHANGE_SUCCESSFUL = 0; 
        public const int DISP_CHANGE_RESTART = 1; 
        public const int DISP_CHANGE_FAILED = -1; 
    } 
    public class Displays
    {
		public static IList<string> GetDisplayNames( )
		{
			var returnVals = new List<string>();
			for(var x=0U; x<1024; ++x)
			{
				DISPLAY_DEVICE outVar = new DISPLAY_DEVICE();
				outVar.cb = (short)Marshal.SizeOf(outVar);
				if(User_32.EnumDisplayDevices(null, x, ref outVar, 1U ))
				{
					returnVals.Add(outVar.DeviceName);
				}
			}
			return returnVals;
		}
		
        // added by Guy Leech in order to get all properties returned from EnumDisplayDevices()
		public static IList<object> GetDisplays( uint flags = 0 )
		{
			var returnVals = new List<object>();
			for(var x=0U; x<1024; ++x)
			{
				DISPLAY_DEVICE outVar = new DISPLAY_DEVICE();
				outVar.cb = (short)Marshal.SizeOf(outVar);
				if(User_32.EnumDisplayDevices(null, x, ref outVar, flags ))
				{
					returnVals.Add(outVar);
				}
			}
			return returnVals;
		}
		
        // added by Guy Leech in order to get properties for a specific device returned from EnumDisplayDevices()
		public static DISPLAY_DEVICE GetDisplayDeviceInfo( string deviceName , uint flags = 0 )
		{
			DISPLAY_DEVICE displayDevice = new DISPLAY_DEVICE() ;
            
            displayDevice.cb = (int)Marshal.SizeOf(displayDevice); 

			if( ! User_32.EnumDisplayDevices( deviceName , 0, ref displayDevice, flags ))
			{
				displayDevice.cb = 0 ;
			}

			return displayDevice ;
		}
		
		public static string GetCurrentResolution(string deviceName)
        {
            string returnValue = null;
            DEVMODE1 dm = GetDevMode1();
            if (0 != User_32.EnumDisplaySettings(deviceName, User_32.ENUM_CURRENT_SETTINGS, ref dm))
            {
                returnValue = dm.dmPelsWidth + "," + dm.dmPelsHeight;
            }
            return returnValue;
        }
		
        // added by Guy Leech in order to get all properties returned from EnumDisplaySettings()
		public static DEVMODE1 GetCurrentDisplaySettings(string deviceName)
        {
            DEVMODE1 dm = GetDevMode1();
            if (0 == User_32.EnumDisplaySettings(deviceName, User_32.ENUM_CURRENT_SETTINGS, ref dm))
            {
                dm.dmSize = 0 ; // denotes call failed
            }
            return dm;
        }

		public static IList<string> GetResolutions()
		{
			var displays = GetDisplayNames();
			var returnValue = new List<string>();
			foreach(var display in displays)
			{
				returnValue.Add(GetCurrentResolution(display));
			}
			return returnValue;
		}
		
        private static DEVMODE1 GetDevMode1() 
        { 
            DEVMODE1 dm = new DEVMODE1(); 
            dm.dmDeviceName = new String(new char[32]); 
            dm.dmFormName = new String(new char[32]); 
            dm.dmSize = (short)Marshal.SizeOf(dm); 
            return dm; 
        } 
    }
} 
"@

Function Add-AzureVMsToListView
{
    Param
    (
        [string]$filter ,
        [bool]$regex ,
        [bool]$allVMs
    )   
    # Keep HostPool column visibility in sync with the AVD checkbox state.
    Set-AzureHostPoolColumn -showHostPool ([bool]$WPFcheckBoxAzureAVD.IsChecked)

    [string]$powerState = '.'
    if( -Not $allVMs )
    {
        $powerState = 'Running'
    }

    Write-Verbose "$(Get-Date -Format G): Getting Azure VMs (all=$allVMs, filter=$filter, regex=$regex)"
    $script:azureLastGetAzVmCall = Get-Date

    # Fetch data synchronously but in a background job to avoid UI thread blocking
    $dataFetchScript = {
        Param(
            [hashtable]$Arguments
        )

        Import-Module -Name Az.Compute -Verbose:$false
        Import-Module -Name Az.Accounts -Verbose:$false
        if( $Arguments.includeAVD ) { Import-Module -Name Az.DesktopVirtualization -Verbose:$false }

        # Unpack arguments
        $powerState = $Arguments.powerState
        $filter = $Arguments.filter
        $regex = $Arguments.regex
        $includeAVD = $Arguments.includeAVD

        # Get VMs
        $azureError = $null
        $vms = @( Get-AZVM -ErrorVariable azureError -Status | 
                  Where-Object { $_.PowerState -match $powerState -and (( $regex -and $_.Name -match $filter ) -or ( -Not $regex -and ( [string]::IsNullOrEmpty( $filter ) -or $_.Name -like $filter ))) } | 
                  Sort-Object -Property Name )

        if( $azureError ) { throw "Error retrieving VMs: $($azureError | Select-Object -First 1)" }

        # Build disk SKU lookup (list API does not return StorageAccountType; must fetch from disks)
        [hashtable]$diskSkuById = @{}
        Get-AzDisk | ForEach-Object {
            if( -Not [string]::IsNullOrWhiteSpace( $_.Id ) )
            {
                $diskSkuById[ $_.Id.ToLowerInvariant() ] = $_.Sku.Name
            }
        }

        # Get tenant
        $tenant = Get-AZTenant
        $tenantDisplay = if( $null -ne $tenant ) { "{0} ({1})" -f $tenant.Name, $tenant.DefaultDomain } else { '' }

        # Get subscriptions
        $subscriptions = @( Get-AZSubscription )

        # Get AVD data if requested
        [hashtable]$sessionHosts = @{}
        [hashtable]$hostPoolIdsByName = @{}
        [hashtable]$hostPoolWorkspaceLookup = @{}

        if( $includeAVD -and $vms.Count -gt 0 )
        {
            $hostpools = @( Get-AZWVDHostPool )
            [hashtable]$workspaceByApplicationGroupId = @{}
            [array]$workspaces = @( Get-AZWVDWorkspace )
            [array]$applicationGroups = @( Get-AZWVDApplicationGroup )

            ForEach( $workspace in $workspaces )
            {
                [string]$workspaceFriendlyName = $workspace.FriendlyName
                if( [string]::IsNullOrWhiteSpace( $workspaceFriendlyName ) ) { $workspaceFriendlyName = $workspace.Name }

                ForEach( $applicationGroupReference in @( $workspace.ApplicationGroupReference ) )
                {
                    if( -Not [string]::IsNullOrWhiteSpace( $applicationGroupReference ) )
                    {
                        $workspaceByApplicationGroupId[ $applicationGroupReference.ToLowerInvariant() ] = @{
                            WorkspaceName = $workspace.Name
                            WorkspaceFriendlyName = $workspaceFriendlyName
                        }
                    }
                }
            }

            ForEach( $applicationGroup in $applicationGroups )
            {
                [string]$hostPoolId = $applicationGroup.HostPoolArmPath
                if( [string]::IsNullOrWhiteSpace( $hostPoolId ) -or [string]::IsNullOrWhiteSpace( $applicationGroup.Id ) ) { continue }

                $workspaceInfo = $workspaceByApplicationGroupId[ $applicationGroup.Id.ToLowerInvariant() ]
                if( $null -ne $workspaceInfo -and -Not $hostPoolWorkspaceLookup.ContainsKey( $hostPoolId.ToLowerInvariant() ) )
                {
                    $hostPoolWorkspaceLookup[ $hostPoolId.ToLowerInvariant() ] = $workspaceInfo
                }
            }

            ForEach( $hostpool in $hostpools )
            {
                if( -Not [string]::IsNullOrWhiteSpace( $hostpool.Name ) -and -Not [string]::IsNullOrWhiteSpace( $hostpool.Id ) )
                {
                    $hostPoolIdsByName[ $hostpool.Name ] = $hostpool.Id
                }

                Get-AZWVDSessionHost -HostPoolName $hostpool.Name -ResourceGroupName $hostpool.ResourceGroupName | 
                    Select-Object -Property *,@{n='HostPool';e={$hostpool.name}} | 
                    ForEach-Object {
                        [string]$sessionHostVmId = $_.VirtualMachineId
                        if( -Not [string]::IsNullOrWhiteSpace( $sessionHostVmId ) )
                        {
                            $sessionHostVmId = $sessionHostVmId.ToLowerInvariant()
                            if( -Not $sessionHosts.ContainsKey( $sessionHostVmId ) )
                            {
                                $sessionHosts[ $sessionHostVmId ] = $_
                            }
                        }
                    }
            }
        }

        # Build list items
        $items = @()
        ForEach( $vm in $vms )
        {
            [string]$subscriptionId = $vm.Id -replace '^/subscriptions/([^/]+)/.*$' , '$1' 
            
            # Calculate SessionHostInfo before creating object
            $sessionHostInfo = $null
            if( $includeAVD )
            {
                [string]$vmId = $vm.vmid
                if( -Not [string]::IsNullOrWhiteSpace( $vmId ) )
                {
                    $vmId = $vmId.ToLowerInvariant()
                }

                if( -Not [string]::IsNullOrWhiteSpace( $vmId ) -and ( $sessionHost = $sessionHosts[ $vmId ] ) )
                {
                    $sessionHostInfo = $sessionHost
                }
            }
            
            $item = [pscustomobject]@{ 
                Name = $vm.Name
                PowerState = $vm.PowerState
                Location = $vm.Location
                ResourceGroup = $vm.ResourceGroupName
                Size = $vm.HardwareProfile.VmSize
                OSDiskType = if( $null -ne $vm.StorageProfile.OsDisk.ManagedDisk.Id ) { $diskSkuById[ $vm.StorageProfile.OsDisk.ManagedDisk.Id.ToLowerInvariant() ] } else { 'Unmanaged' }
                Subscription = ($subscriptions | Where-Object Id -eq $subscriptionId | Select-Object -ExpandProperty Name)
                SubscriptionId = $subscriptionId
                Created = $vm.TimeCreated
                VMObject = $vm
                SessionHostInfo = $sessionHostInfo
            }

            $items += $item
        }

        return @{
            TenantDisplay = $tenantDisplay
            Items = $items
            HostPoolIdsByName = $hostPoolIdsByName
            HostPoolWorkspaceLookup = $hostPoolWorkspaceLookup
        }
    }

    try
    {
        $result = Invoke-AzureOperationWithTimeout -ScriptBlock $dataFetchScript -TimeoutSeconds 300 -ModulesToImport @() -Arguments @{
            powerState = $powerState
            filter = $filter
            regex = $regex
            includeAVD = $WPFcheckBoxAzureAVD.IsChecked
        }

        if( $null -eq $result -or $result.Count -eq 0 )
        {
            throw "No data returned from Azure VM fetch"
        }

        $wpftextBoxAzureTenant.Text = $result.TenantDisplay
        $hostPoolIdsByName = $result.HostPoolIdsByName
        $hostPoolWorkspaceLookup = $result.HostPoolWorkspaceLookup

        Write-Verbose "$(Get-Date -Format G): Got $($result.Items.Count) Azure VMs"
        $WPFlistViewAzureVMs.Items.Clear()

        if( $result.Items.Count -eq 0 )
        {
            Apply-AzureColumnFilters
            Update-AzureVMLabel
            return
        }

        ForEach( $item in $result.Items )
        {
            $displayItem = [pscustomobject]@{ 
                Name = $item.Name
                PowerState = $item.PowerState
                Location = $item.Location
                ResourceGroup = $item.ResourceGroup
                Size = $item.Size
                OSDiskType = $item.OSDiskType
                Subscription = $item.Subscription
                SubscriptionId = $item.SubscriptionId
                Created = $item.Created
            }

            if( $WPFcheckBoxAzureAVD.IsChecked -and $null -ne $item.SessionHostInfo )
            {
                $sessionHost = $item.SessionHostInfo
                [string]$hostPoolId = $hostPoolIdsByName[ $sessionHost.HostPool ]
                $workspaceInfo = $null
                if( -Not [string]::IsNullOrWhiteSpace( $hostPoolId ) )
                {
                    $workspaceInfo = $hostPoolWorkspaceLookup[ $hostPoolId.ToLowerInvariant() ]
                }

                Add-Member -InputObject $displayItem -NotePropertyMembers @{
                     'HostPool' = $sessionHost.HostPool
                     'Workspace Name' = if( $null -ne $workspaceInfo ) { $workspaceInfo.WorkspaceName } else { $null }
                     'Workspace Friendly Name' = if( $null -ne $workspaceInfo ) { $workspaceInfo.WorkspaceFriendlyName } else { $null }
                     'Sessions' = $sessionHost.Session
                     'AVD Status' = $sessionHost.Status
                     'Allow New Session' = $sessionHost.AllowNewSession
                     'Assigned User' = if( -Not [string]::IsNullOrWhiteSpace( $sessionHost.AssignedUser ) ) { $sessionHost.AssignedUser } else { $null }
                }
            }

            $WPFlistViewAzureVMs.Items.Add( $displayItem )
            Write-Verbose -Message "Added $($item.Name)"
        }
        Apply-AzureColumnFilters
    }
    catch
    {
        Write-Error "Failed to retrieve Azure VMs: $_"
        [void][Windows.MessageBox]::Show( $mainWindow , "Failed to retrieve Azure VMs: $($_.Exception.Message)" , 'Azure Error' , 'Ok' , 'Error' )
    }
}

Function Set-AzureHostPoolColumn
{
    Param
    (
        [bool]$showHostPool
    )

    $azureGridView = $WPFlistViewAzureVMs.View -as [System.Windows.Controls.GridView]
    if( $null -eq $azureGridView )
    {
        return
    }

    [array]$avdColumns = @(
        @{ Header = 'HostPool'   ; Binding = 'HostPool' } ,
        @{ Header = 'Workspace Name' ; Binding = 'Workspace Name' } ,
        @{ Header = 'Workspace Friendly Name' ; Binding = 'Workspace Friendly Name' } ,
        @{ Header = 'Sessions'   ; Binding = 'Sessions' } ,
        @{ Header = 'AVD Status' ; Binding = 'AVD Status' } ,
        @{ Header = 'Allow New Session' ; Binding = 'Allow New Session' } ,
        @{ Header = 'Assigned User' ; Binding = 'Assigned User' }
    )

    if( $showHostPool )
    {
        ForEach( $columnDefinition in $avdColumns )
        {
            $existingColumn = $azureGridView.Columns | Where-Object { $_.Header -eq $columnDefinition.Header } | Select-Object -First 1
            if( $null -eq $existingColumn )
            {
                $newColumn = New-Object -TypeName System.Windows.Controls.GridViewColumn
                $newColumn.Header = $columnDefinition.Header
                $newColumn.DisplayMemberBinding = New-Object -TypeName System.Windows.Data.Binding -ArgumentList $columnDefinition.Binding
                [void]$azureGridView.Columns.Add( $newColumn )
            }
        }
    }
    else
    {
        ForEach( $columnDefinition in $avdColumns )
        {
            $existingColumn = $azureGridView.Columns | Where-Object { $_.Header -eq $columnDefinition.Header } | Select-Object -First 1
            if( $null -ne $existingColumn )
            {
                [void]$azureGridView.Columns.Remove( $existingColumn )
            }
        }
    }

    Update-AzureColumnHeaders
}

Function Set-AzureSessionMenuState
{
    Param
    (
        [bool]$avdEnabled
    )

    if( $null -ne $WPFAzureSessionContextMenu )
    {
        $WPFAzureSessionContextMenu.IsEnabled = $avdEnabled
    }

    if( $null -ne $WPFAzureRunContextMenu )
    {
        $WPFAzureRunContextMenu.IsEnabled = $true
    }

    if( $null -ne $WPFAzureDeleteSessionHostContextMenu )
    {
        $WPFAzureDeleteSessionHostContextMenu.IsEnabled = $avdEnabled
    }

    if( $null -ne $WPFAzureDeleteSessionHostAndVMContextMenu )
    {
        $WPFAzureDeleteSessionHostAndVMContextMenu.IsEnabled = $avdEnabled
    }

    if( $null -ne $WPFAzureHostPoolContextMenu )
    {
        $WPFAzureHostPoolContextMenu.IsEnabled = $avdEnabled
    }

    if( $null -ne $WPFAzureAVDLogsContextMenu )
    {
        $WPFAzureAVDLogsContextMenu.IsEnabled = $avdEnabled
    }
}

Function Get-AzureRunCommandText
{
    if( -Not ( $textInputWindow = New-WPFWindow -inputXAML $textInputXAML ) )
    {
        return $null
    }

    if( [string]::IsNullOrWhiteSpace( $script:azureRunCommandLastText ) )
    {
        $persistedRunText = Get-ItemProperty -Path $configKey -Name 'AzureRunCommandText' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty 'AzureRunCommandText' -ErrorAction SilentlyContinue
        if( -Not [string]::IsNullOrWhiteSpace( $persistedRunText ) )
        {
            $script:azureRunCommandLastText = [string]$persistedRunText
        }
    }

    $WPFbtnInputTextOk.Add_Click({
        $_.Handled = $true
        $textInputWindow.DialogResult = $true
        $textInputWindow.Close()
    })

    [void]( $textInputWindow.Title = 'Run PowerShell Script' )
    [void]( $WPFlblInputTextLabel.Content = 'Enter the PowerShell script to run on the selected AVD VM(s)' )
    [void]( $WPFtextboxInputText.AcceptsReturn = $true )
    [void]( $WPFtextboxInputText.VerticalScrollBarVisibility = 'Auto' )
    [void]( $WPFtextboxInputText.TextWrapping = 'Wrap' )
    if( -Not [string]::IsNullOrWhiteSpace( $script:azureRunCommandLastText ) )
    {
        [void]( $WPFtextboxInputText.Text = $script:azureRunCommandLastText )
    }
    [void]$WPFtextboxInputText.Focus()
    [void]$WPFtextboxInputText.Select( $WPFtextboxInputText.Text.Length , 0 )

    if( -Not $textInputWindow.ShowDialog() )
    {
        return $null
    }

    [string]$scriptText = $WPFtextboxInputText.Text
    if( [string]::IsNullOrWhiteSpace( $scriptText ) )
    {
        return $null
    }

    $script:azureRunCommandLastText = $scriptText
    if( -Not ( Test-Path -Path $configKey ) )
    {
        $null = New-Item -Path $configKey -ItemType Key -Force -ErrorAction SilentlyContinue
    }
    if( Test-Path -Path $configKey )
    {
        $null = Set-ItemProperty -Path $configKey -Name 'AzureRunCommandText' -Value $script:azureRunCommandLastText -Force -ErrorAction SilentlyContinue
    }

    return $scriptText.Trim()
}

Function Show-AzureRunCommandOutputWindow
{
    Param
    (
        [Parameter(Mandatory)][string]$computerName ,
        [Parameter(Mandatory)][string]$scriptText ,
        [string]$outputText
    )

    $normalizeNewLines = {
        Param(
            [string]$text
        )

        if( $null -eq $text )
        {
            return ''
        }

        # WinForms multiline textbox renders consistently when line endings are CRLF.
        return ( $text -split "`r`n|`n|`r" ) -join "`r`n"
    }

    [string]$normalizedScriptText = & $normalizeNewLines $scriptText
    [string]$normalizedOutputText = & $normalizeNewLines $outputText

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Run Output - $computerName"
    $form.Size = New-Object System.Drawing.Size(900, 640)
    $form.StartPosition = 'CenterScreen'

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Dock = 'Fill'
    $textBox.Multiline = $true
    $textBox.ReadOnly = $true
    $textBox.ScrollBars = 'Both'
    $textBox.WordWrap = $false
    $textBox.Font = New-Object System.Drawing.Font( 'Consolas' , 9 )
    $textBox.Text = "Computer: $computerName`r`n`r`nCommand:`r`n$normalizedScriptText`r`n`r`nOutput:`r`n$(if( [string]::IsNullOrWhiteSpace( $normalizedOutputText ) ) { '<no output returned>' } else { $normalizedOutputText })"

    [void]$form.Controls.Add( $textBox )
    [void]$form.ShowDialog()
}

Function Convert-AzureRunCommandResultToText
{
    Param
    (
        $runCommandResult
    )

    if( $null -eq $runCommandResult )
    {
        return ''
    }

    $messages = New-Object System.Collections.Generic.List[string]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]'

    $addMessage = {
        Param(
            [string]$text
        )

        if( [string]::IsNullOrWhiteSpace( $text ) )
        {
            return
        }

        [string]$trimmedText = $text.Trim()
        if( -Not [string]::IsNullOrWhiteSpace( $trimmedText ) -and $seen.Add( $trimmedText ) )
        {
            [void]$messages.Add( $trimmedText )
        }
    }

    $extractFromStatusObject = {
        Param(
            $statusObject
        )

        $convertValueToText = {
            Param(
                $value
            )

            if( $null -eq $value )
            {
                return ''
            }

            if( ( $value -is [System.Collections.IEnumerable] ) -and -Not ( $value -is [string] ) )
            {
                return ( @( $value ) | ForEach-Object { if( $null -eq $_ ) { '' } else { [string]$_ } } | Where-Object { -Not [string]::IsNullOrWhiteSpace( $_ ) } ) -join "`r`n"
            }

            return [string]$value
        }

        if( $null -eq $statusObject )
        {
            return
        }

        if( $statusObject -is [string] )
        {
            & $addMessage $statusObject
            return
        }

        [string]$displayStatus = $null
        if( $statusObject.PSObject.Properties.Match( 'DisplayStatus' ).Count -gt 0 -and -Not [string]::IsNullOrWhiteSpace( $statusObject.DisplayStatus ) )
        {
            $displayStatus = $statusObject.DisplayStatus
        }
        elseif( $statusObject.PSObject.Properties.Match( 'Code' ).Count -gt 0 -and -Not [string]::IsNullOrWhiteSpace( $statusObject.Code ) )
        {
            $displayStatus = $statusObject.Code
        }

        [string]$messageText = $null
        if( $statusObject.PSObject.Properties.Match( 'Message' ).Count -gt 0 -and -Not [string]::IsNullOrWhiteSpace( $statusObject.Message ) )
        {
            $messageText = & $convertValueToText $statusObject.Message
        }
        elseif( $statusObject.PSObject.Properties.Match( 'Output' ).Count -gt 0 -and -Not [string]::IsNullOrWhiteSpace( $statusObject.Output ) )
        {
            $messageText = & $convertValueToText $statusObject.Output
        }

        if( -Not [string]::IsNullOrWhiteSpace( $messageText ) )
        {
            try
            {
                $jsonMessage = $messageText | ConvertFrom-Json -ErrorAction Stop
                if( $null -ne $jsonMessage )
                {
                    if( $jsonMessage.PSObject.Properties.Match( 'stdout' ).Count -gt 0 )
                    {
                        [string]$stdoutText = & $convertValueToText $jsonMessage.stdout
                        if( -Not [string]::IsNullOrWhiteSpace( $stdoutText ) )
                        {
                            & $addMessage "[stdout]`r`n$stdoutText"
                        }
                    }
                    if( $jsonMessage.PSObject.Properties.Match( 'stderr' ).Count -gt 0 )
                    {
                        [string]$stderrText = & $convertValueToText $jsonMessage.stderr
                        if( -Not [string]::IsNullOrWhiteSpace( $stderrText ) )
                        {
                            & $addMessage "[stderr]`r`n$stderrText"
                        }
                    }
                }
            }
            catch
            {
                & $addMessage $messageText
            }
        }

        if( -Not [string]::IsNullOrWhiteSpace( $displayStatus ) )
        {
            & $addMessage $displayStatus
        }
    }

    [array]$candidateCollections = @()
    if( $runCommandResult.PSObject.Properties.Match( 'Value' ).Count -gt 0 )
    {
        $candidateCollections += @( $runCommandResult.Value )
    }
    if( $runCommandResult.PSObject.Properties.Match( 'Output' ).Count -gt 0 )
    {
        $candidateCollections += @( $runCommandResult.Output )
    }

    if( ( $runCommandResult -is [System.Collections.IEnumerable] ) -and -Not ( $runCommandResult -is [string] ) )
    {
        $candidateCollections += @( $runCommandResult )
    }
    else
    {
        $candidateCollections += @( $runCommandResult )
    }

    ForEach( $collection in $candidateCollections )
    {
        if( $null -eq $collection )
        {
            continue
        }

        if( ( $collection -is [System.Collections.IEnumerable] ) -and -Not ( $collection -is [string] ) )
        {
            ForEach( $item in $collection )
            {
                & $extractFromStatusObject $item
            }
        }
        else
        {
            & $extractFromStatusObject $collection
        }
    }

    if( $messages.Count -eq 0 )
    {
        return ( $runCommandResult | Format-List * -Force | Out-String ).Trim()
    }

    ( $messages -join "`r`n`r`n" ).Trim()
}

Function Select-AzureSubscription
{
    Import-Module -Name Az.Accounts -Verbose:$false

    if( $null -eq ( Get-AzContext -ErrorAction SilentlyContinue ) )
    {
        $null = Connect-AzAccount
        if( $null -eq ( Get-AzContext -ErrorAction SilentlyContinue ) )
        {
            return
        }
    }

    [array]$subscriptions = @( Get-AZSubscription | Sort-Object -Property Name )

    if( $subscriptions.Count -eq 0 )
    {
        [void][Windows.MessageBox]::Show( $mainWindow , 'No Azure subscriptions found for the current account' , 'Azure Subscription' , 'Ok' ,'Information' )
        return
    }

    $selectionWindow = New-Object -TypeName System.Windows.Window
    $selectionWindow.Title = 'Select Azure Subscription'
    $selectionWindow.Width = 640
    $selectionWindow.Height = 420
    $selectionWindow.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterOwner

    if( $null -ne $mainWindow )
    {
        $selectionWindow.Owner = $mainWindow
    }

    $grid = New-Object -TypeName System.Windows.Controls.Grid
    $listRowDefinition = New-Object -TypeName System.Windows.Controls.RowDefinition
    $listRowDefinition.Height = New-Object -TypeName System.Windows.GridLength -ArgumentList 1,([System.Windows.GridUnitType]::Star)
    $null = $grid.RowDefinitions.Add( $listRowDefinition )
    $buttonRowDefinition = New-Object -TypeName System.Windows.Controls.RowDefinition
    $buttonRowDefinition.Height = [System.Windows.GridLength]::Auto
    $null = $grid.RowDefinitions.Add( $buttonRowDefinition )

    $listBox = New-Object -TypeName System.Windows.Controls.ListBox
    $listBox.Margin = New-Object -TypeName System.Windows.Thickness -ArgumentList 10
    $listBox.DisplayMemberPath = 'Name'
    $subscriptions | ForEach-Object { [void]$listBox.Items.Add( $_ ) }
    if( $null -ne $script:azureSelectedSubscription )
    {
        $existing = $subscriptions | Where-Object { $_.Id -eq $script:azureSelectedSubscription.Id } | Select-Object -First 1
        if( $null -ne $existing )
        {
            $listBox.SelectedItem = $existing
        }
    }
    if( $null -eq $listBox.SelectedItem -and $listBox.Items.Count -gt 0 )
    {
        $listBox.SelectedIndex = 0
    }
    [System.Windows.Controls.Grid]::SetRow( $listBox , 0 )
    [void]$grid.Children.Add( $listBox )

    $buttonPanel = New-Object -TypeName System.Windows.Controls.StackPanel
    $buttonPanel.Orientation = [System.Windows.Controls.Orientation]::Horizontal
    $buttonPanel.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Right
    $buttonPanel.Margin = New-Object -TypeName System.Windows.Thickness -ArgumentList 10

    $okButton = New-Object -TypeName System.Windows.Controls.Button
    $okButton.Content = 'OK'
    $okButton.MinWidth = 80
    $okButton.Margin = New-Object -TypeName System.Windows.Thickness -ArgumentList 0,0,8,0
    $okButton.IsDefault = $true
    $okButton.Add_Click({
        if( $null -ne $listBox.SelectedItem )
        {
            $selectionWindow.DialogResult = $true
            $selectionWindow.Close()
        }
    })

    $cancelButton = New-Object -TypeName System.Windows.Controls.Button
    $cancelButton.Content = 'Cancel'
    $cancelButton.MinWidth = 80
    $cancelButton.IsCancel = $true

    [void]$buttonPanel.Children.Add( $okButton )
    [void]$buttonPanel.Children.Add( $cancelButton )
    [System.Windows.Controls.Grid]::SetRow( $buttonPanel , 1 )
    [void]$grid.Children.Add( $buttonPanel )

    $listBox.Add_MouseDoubleClick({
        if( $null -ne $listBox.SelectedItem )
        {
            $selectionWindow.DialogResult = $true
            $selectionWindow.Close()
        }
    })

    $selectionWindow.Content = $grid
    if( -Not $selectionWindow.ShowDialog() )
    {
        return
    }

    $selectedSubscription = $listBox.SelectedItem
    if( $null -eq $selectedSubscription )
    {
        return
    }

    try
    {
        $null = Set-AzContext -SubscriptionId $selectedSubscription.Id -ErrorAction Stop
        $script:azureSelectedSubscription = $selectedSubscription
        Add-AzureVMsToListView -filter '' -regex $false -allVMs $WPFcheckBoxAzureAllVMs.IsChecked
    }
    catch
    {
        [void][Windows.MessageBox]::Show( $mainWindow , "Failed to set subscription context to $($selectedSubscription.Name)`n$($_.Exception.Message)" , 'Azure Subscription' , 'Ok' ,'Error' )
    }
}

Function Invoke-AzureOperationWithTimeout
{
    Param
    (
        [Parameter(Mandatory)][scriptblock]$ScriptBlock ,
        [int]$TimeoutSeconds = 120 ,
        [string[]]$ModulesToImport = @( 'Az.Accounts' , 'Az.Compute' ) ,
        [hashtable]$Arguments = @{}
    )

    [string]$jobName = "AzOp_{0}" -f ([guid]::NewGuid().ToString('N'))
    $job = $null
    $result = $null

    try
    {
        # Create import module script lines
        $importLines = $ModulesToImport | ForEach-Object { "Import-Module -Name $_ -Verbose:`$false" }

        # Combine imports with user scriptblock (no wrapper Param block - user's Param will be used)
        $fullScriptBlock = [scriptblock]::Create( ($importLines -join "`n") + "`n`n" + $ScriptBlock.ToString() )

        # Start job with argument passthrough
        $job = Start-Job -Name $jobName -ScriptBlock $fullScriptBlock -ArgumentList $Arguments

        # Wait for completion with timeout
        $null = Wait-Job -Job $job -Timeout $TimeoutSeconds

        if( $job.State -eq 'Running' )
        {
            Stop-Job -Job $job -ErrorAction SilentlyContinue | Out-Null
            throw "Azure operation timed out after $TimeoutSeconds seconds"
        }

        # Retrieve results and errors
        if( $job.ChildJobs.Count -gt 0 )
        {
            [array]$jobErrors = @( $job.ChildJobs | ForEach-Object { $_.Error } )
            if( $jobErrors.Count -gt 0 )
            {
                throw ( $jobErrors | Select-Object -First 1 ).Exception
            }
        }

        $result = Receive-Job -Job $job -ErrorAction Stop
        return $result
    }
    catch
    {
        Write-Error -Message "Azure operation failed: $($_.Exception.Message)" -ErrorAction Stop
    }
    finally
    {
        if( $null -ne $job )
        {
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
        }
    }
}

Function Connect-AzureAccountWithTimeout
{
    Param
    (
        [int]$timeoutSeconds = 120
    )

    Import-Module -Name Az.Accounts -Verbose:$false

    [string]$contextPath = Join-Path -Path $env:TEMP -ChildPath ( 'mstsc-sizer-auth-{0}.json' -f ([guid]::NewGuid().ToString()) )
    $authJob = $null

    try
    {
        $authJob = Start-Job -Name ( "Azure_Auth_{0}" -f ([guid]::NewGuid().ToString('N')) ) -ArgumentList $contextPath -ScriptBlock {
            Param
            (
                [string]$outputContextPath
            )

            Import-Module -Name Az.Accounts -Verbose:$false

            try
            {
                $connection = Connect-AzAccount -ErrorAction Stop
                if( $null -eq $connection )
                {
                    throw 'No authentication result returned by Connect-AzAccount.'
                }

                Save-AzContext -Path $outputContextPath -Force -ErrorAction Stop | Out-Null
                [pscustomobject]@{
                    Success = $true
                    ContextPath = $outputContextPath
                }
            }
            catch
            {
                [pscustomobject]@{
                    Success = $false
                    ContextPath = $outputContextPath
                    ErrorMessage = $_.Exception.Message
                }
            }
        }

        if( $null -eq ( Wait-Job -Job $authJob -Timeout $timeoutSeconds ) )
        {
            Stop-Job -Job $authJob -ErrorAction SilentlyContinue | Out-Null
            throw "Timed out waiting $timeoutSeconds seconds for Azure authentication."
        }

        [array]$authOutput = @( Receive-Job -Job $authJob -ErrorAction SilentlyContinue )
        $authResult = $authOutput | Where-Object { $_ -is [pscustomobject] -and $_.PSObject.Properties.Match( 'Success' ).Count -gt 0 } | Select-Object -Last 1

        if( $null -eq $authResult )
        {
            throw 'Azure authentication did not return a usable result.'
        }

        if( -Not $authResult.Success )
        {
            throw $(if( [string]::IsNullOrWhiteSpace( $authResult.ErrorMessage ) ) { 'Azure authentication failed.' } else { $authResult.ErrorMessage })
        }

        if( -Not ( Test-Path -Path $authResult.ContextPath ) )
        {
            throw 'Azure authentication completed but no context file was produced.'
        }

        Import-AzContext -Path $authResult.ContextPath -ErrorAction Stop | Out-Null
        return Get-AzContext -ErrorAction SilentlyContinue
    }
    finally
    {
        if( $null -ne $authJob )
        {
            Remove-Job -Job $authJob -Force -ErrorAction SilentlyContinue
        }

        if( Test-Path -Path $contextPath )
        {
            Remove-Item -Path $contextPath -Force -ErrorAction SilentlyContinue
        }
    }
}

Function Get-AzureAVDSessionHost
{
    Param
    (
        [Parameter(Mandatory)][pscustomobject]$selection
    )

    if( [string]::IsNullOrEmpty( $selection.HostPool ) )
    {
        throw "VM $($selection.Name) is not associated with an AVD host pool"
    }

    Import-Module -Name Az.DesktopVirtualization -Verbose:$false
    [array]$sessionHosts = @( Get-AZWVDSessionHost -HostPoolName $selection.HostPool -ResourceGroupName $selection.ResourceGroup -ErrorAction Stop )

    $sessionHost = $sessionHosts | Where-Object {
        [string]$sessionHostName = $_.Name -replace '^.*/' , ''
        [string]$sessionHostVmName = $sessionHostName -replace '\..*$' , ''
        $sessionHostVmName -ieq $selection.Name
    } | Select-Object -First 1

    if( $null -eq $sessionHost )
    {
        throw "Failed to find AVD session host for $($selection.Name) in host pool $($selection.HostPool)"
    }

    $sessionHost
}

Function Remove-AzureVMEntry
{
    Param
    (
        [Parameter(Mandatory)][pscustomobject]$selection
    )

    Import-Module -Name Az.Compute -Verbose:$false
    $actionResult = $null
    $actionResult = Remove-AzVM -Name $selection.Name -ResourceGroupName $selection.ResourceGroup -Force -Confirm:$false -ErrorAction Stop
    if( $null -ne $actionResult -and $actionResult.Status -ieq 'Failed' )
    {
        Write-Warning -Message "Delete VM failed for $($selection.Name): $($actionResult.Error)"
        [void][Windows.MessageBox]::Show( $mainWindow , "Delete failed for $($selection.Name)`n$($actionResult.Error)" , 'Azure Delete VM' , 'Ok' , 'Error' )
        return $false
    }
    else
    {
        Write-Host "Deleted VM: $($selection.Name)" -ForegroundColor Green
        return $true
    }
}

Function Get-AzureAVDUserSessions
{
    Param
    (
        [Parameter(Mandatory)][string]$sessionHostResourceId
    )

    $response = Invoke-AzRestMethod -Method GET -Path "$sessionHostResourceId/userSessions?api-version=$AVDAPIversion" -ErrorAction Stop
    if( [string]::IsNullOrEmpty( $response.Content ) )
    {
        return @()
    }

    [array]$sessions = @(( $response.Content | ConvertFrom-Json ).value)
    if( $null -eq $sessions )
    {
        return @()
    }

    $sessions
}

Function Invoke-AzureAVDRunCommandJson
{
    Param
    (
        [Parameter(Mandatory)][string]$vmName ,
        [Parameter(Mandatory)][string]$resourceGroupName ,
        [Parameter(Mandatory)][string]$scriptText
    )

    Import-Module -Name Az.Accounts -Verbose:$false
    Import-Module -Name Az.Compute -Verbose:$false

    $currentAzContext = Get-AzContext -ErrorAction Stop
    if( $null -eq $currentAzContext )
    {
        throw 'No active Azure context. Connect to Azure first.'
    }

    Write-Verbose "$(Get-Date -Format G): Invoking Azure Run Command for $vmName in resource group $resourceGroupName"
    $runCommandResult = Invoke-AzVMRunCommand -ResourceGroupName $resourceGroupName -Name $vmName -CommandId RunPowerShellScript -ScriptString $scriptText -ErrorAction Stop
    Write-Verbose "$(Get-Date -Format G): Azure Run Command completed for $vmName with status $($runCommandResult.Status) and code $($runCommandResult.Code)"
    [string]$outputText = Convert-AzureRunCommandResultToText -runCommandResult $runCommandResult
    if( [string]::IsNullOrWhiteSpace( $outputText ) )
    {
        throw "No output returned by Azure Run Command for $vmName"
    }

    Write-Verbose "Output text length is $($outputText.Length) characters for $vmName"
    [array]$base64Candidates = @()
    [array]$candidateLines = @( $outputText -split "`r`n|`n|`r" | Where-Object { $_ -match '^[A-Za-z0-9+/=]+$' -and $_.Length -ge 64 } )
    if( $candidateLines.Count -gt 0 )
    {
        $base64Candidates += ( $candidateLines -join '' )
    }

    [string]$condensedOutput = ( $outputText -replace '\s+' , '' )
    if( -Not [string]::IsNullOrWhiteSpace( $condensedOutput ) -and $condensedOutput -match '^[A-Za-z0-9+/=]+$' -and $condensedOutput.Length -ge 64 )
    {
        $base64Candidates += $condensedOutput
    }

    [array]$base64Candidates = @( $base64Candidates | Where-Object { -Not [string]::IsNullOrWhiteSpace( $_ ) } | Select-Object -Unique )
    ForEach( $base64Payload in $base64Candidates )
    {
        [string]$zipPath = Join-Path -Path $env:TEMP -ChildPath ( 'mstsc-sizer-process-payload-{0}.zip' -f ([guid]::NewGuid().ToString('N')) )
        [string]$extractPath = Join-Path -Path $env:TEMP -ChildPath ( 'mstsc-sizer-process-payload-{0}' -f ([guid]::NewGuid().ToString('N')) )

        try
        {
            [byte[]]$zipBytes = [System.Convert]::FromBase64String( $base64Payload )
            [System.IO.File]::WriteAllBytes( $zipPath , $zipBytes )

            if( Test-Path -Path $extractPath )
            {
                Remove-Item -Path $extractPath -Recurse -Force -ErrorAction SilentlyContinue
            }

            $null = New-Item -Path $extractPath -ItemType Directory -Force
            Expand-Archive -LiteralPath $zipPath -DestinationPath $extractPath -Force -ErrorAction Stop

            [string]$jsonFile = Join-Path -Path $extractPath -ChildPath 'payload.json'
            if( -Not ( Test-Path -Path $jsonFile ) )
            {
                $jsonFile = Get-ChildItem -Path $extractPath -Filter '*.json' -File -Recurse | Select-Object -First 1 -ExpandProperty FullName
            }

            if( [string]::IsNullOrWhiteSpace( $jsonFile ) -or -Not ( Test-Path -Path $jsonFile ) )
            {
                throw "Zip payload did not contain a JSON file for $vmName"
            }

            [string]$jsonPayload = Get-Content -Path $jsonFile -Raw -ErrorAction Stop
            if( [string]::IsNullOrWhiteSpace( $jsonPayload ) )
            {
                return @()
            }

            return ( $jsonPayload | ConvertFrom-Json -ErrorAction Stop )
        }
        catch
        {
            # Keep trying candidates; fallback to JSON marker parsing below.
        }
        finally
        {
            if( Test-Path -Path $zipPath )
            {
                Remove-Item -Path $zipPath -Force -ErrorAction SilentlyContinue
            }
            if( Test-Path -Path $extractPath )
            {
                Remove-Item -Path $extractPath -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }

    $jsonMatch = [regex]::Match( $outputText , '__MSTSCPROCESSJSON_BEGIN__\s*(?<json>[\s\S]*?)\s*__MSTSCPROCESSJSON_END__' )
    if( -Not $jsonMatch.Success )
    {
        throw "Failed to extract payload from Azure Run Command output for $vmName`n$outputText"
    }

    [string]$jsonPayload = $jsonMatch.Groups[ 'json' ].Value.Trim()
    if( [string]::IsNullOrWhiteSpace( $jsonPayload ) )
    {
        return @()
    }

    try
    {
        return ( $jsonPayload | ConvertFrom-Json -ErrorAction Stop )
    }
    catch
    {
        throw "Failed to parse JSON payload for $vmName : $($_.Exception.Message)"
    }
}

Function Get-AzureAVDSessionProcesses
{
    Param
    (
        [Parameter(Mandatory)][array]$sessionRows
    )

    [array]$results = @()

    ForEach( $sessionRow in $sessionRows )
    {
        if( $null -eq $sessionRow.SelectionObject )
        {
            continue
        }

        [int]$targetSessionId = 0
        if( -Not [int]::TryParse( [string]$sessionRow.'Session Id' , [ref]$targetSessionId ) )
        {
            continue
        }

        [string]$scriptToRun = @"
`$targetSessionId = $targetSessionId

`$userIndexByName = @{}
`$userList = New-Object 'System.Collections.Generic.List[string]'
`$pathIndexByValue = @{}
`$pathList = New-Object 'System.Collections.Generic.List[string]'
`$getUserIndex = {
    Param
    (
        [string]`$userName
    )

    if( [string]::IsNullOrWhiteSpace( `$userName ) )
    {
        return -1
    }

    if( -Not `$userIndexByName.ContainsKey( `$userName ) )
    {
        `$newIndex = `$userList.Count
        [void]`$userList.Add( `$userName )
        `$userIndexByName[ `$userName ] = `$newIndex
    }

    return [int]`$userIndexByName[ `$userName ]
}

`$getPathIndex = {
    Param
    (
        [string]`$pathValue
    )

    if( [string]::IsNullOrWhiteSpace( `$pathValue ) )
    {
        return -1
    }

    if( -Not `$pathIndexByValue.ContainsKey( `$pathValue ) )
    {
        `$newIndex = `$pathList.Count
        [void]`$pathList.Add( `$pathValue )
        `$pathIndexByValue[ `$pathValue ] = `$newIndex
    }

    return [int]`$pathIndexByValue[ `$pathValue ]
}

`$rows = ForEach( `$process in @( Get-Process -IncludeUserName -ErrorAction SilentlyContinue | Where-Object SessionId -eq `$targetSessionId | Sort-Object -Property WorkingSet64 -Descending ) )
{
    `$ownerName = [string]`$process.UserName
    `$startTimeText = `$null

    try
    {
        `$startTimeText = `$process.StartTime.ToString( 's' )
    }
    catch {}

    `$userIndex = & `$getUserIndex `$ownerName
    `$pathValue = `$null
    try
    {
        `$pathValue = [string]`$process.Path
    }
    catch {}
    `$pathIndex = & `$getPathIndex `$pathValue

    [pscustomobject]@{
        n = `$process.ProcessName
        p = `$process.Id
        c = [math]::Round( [double]`$( if( `$null -eq `$process.CPU ) { 0 } else { `$process.CPU } ) , 2 )
        w = [math]::Round( `$process.WorkingSet64 / 1MB , 2 )
        x = [int]`$pathIndex
        s = `$process.SessionId
        u = [int]`$userIndex
        t = `$startTimeText
    }
}

`$payload = [pscustomobject]@{
    l = @(`$userList)
    pl = @(`$pathList)
    d = @(`$rows)
}

`$json = ConvertTo-Json -InputObject `$payload -Depth 4 -Compress
if( [string]::IsNullOrWhiteSpace( `$json ) )
{
    `$json = '{"l":[],"pl":[],"d":[]}'
}

`$tempRoot = Join-Path -Path `$env:TEMP -ChildPath ( 'mstsc-sizer-payload-' + [guid]::NewGuid().ToString('N') )
`$jsonPath = Join-Path -Path `$tempRoot -ChildPath 'payload.json'
`$zipPath = Join-Path -Path `$tempRoot -ChildPath 'payload.zip'
try
{
    `$null = New-Item -Path `$tempRoot -ItemType Directory -Force
    Set-Content -Path `$jsonPath -Value `$json -Encoding UTF8 -NoNewline
    Compress-Archive -Path `$jsonPath -DestinationPath `$zipPath -CompressionLevel Optimal -Force
    [byte[]]`$zipBytes = [System.IO.File]::ReadAllBytes( `$zipPath )
    [string]`$base64 = [System.Convert]::ToBase64String( `$zipBytes )
    Write-Output `$base64
}
finally
{
    if( Test-Path -Path `$tempRoot )
    {
        Remove-Item -Path `$tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
"@

        try
        {
            $runPayload = Invoke-AzureAVDRunCommandJson -vmName $sessionRow.SelectionObject.Name -resourceGroupName $sessionRow.SelectionObject.ResourceGroup -scriptText $scriptToRun

            [array]$vmProcesses = @()
            [array]$userLookupList = @()
            [array]$pathLookupList = @()
            [hashtable]$userLookupMap = @{}
            if( $null -ne $runPayload -and $runPayload.PSObject.Properties.Match( 'd' ).Count -gt 0 )
            {
                $vmProcesses = @( $runPayload.d )

                if( $runPayload.PSObject.Properties.Match( 'l' ).Count -gt 0 -and $null -ne $runPayload.l )
                {
                    $userLookupList = @( $runPayload.l )
                }

                if( $runPayload.PSObject.Properties.Match( 'pl' ).Count -gt 0 -and $null -ne $runPayload.pl )
                {
                    $pathLookupList = @( $runPayload.pl )
                }

                if( $runPayload.PSObject.Properties.Match( 'u' ).Count -gt 0 -and $null -ne $runPayload.u )
                {
                    if( $runPayload.u -is [System.Collections.IDictionary] )
                    {
                        ForEach( $key in $runPayload.u.Keys )
                        {
                            $userLookupMap[ [string]$key ] = [string]$runPayload.u[ $key ]
                        }
                    }
                    else
                    {
                        ForEach( $property in $runPayload.u.PSObject.Properties )
                        {
                            $userLookupMap[ [string]$property.Name ] = [string]$property.Value
                        }
                    }
                }
            }
            else
            {
                $vmProcesses = @( $runPayload )
            }

            ForEach( $processInfo in $vmProcesses )
            {
                [bool]$isCompact = $processInfo.PSObject.Properties.Match( 'n' ).Count -gt 0
                [string]$userTokenOrName = if( $isCompact ) { [string]$processInfo.u } else { [string]$processInfo.User }
                [string]$expandedUser = $userTokenOrName

                if( $isCompact )
                {
                    [int]$userIndex = -1
                    if( $userLookupList.Count -gt 0 -and [int]::TryParse( $userTokenOrName , [ref]$userIndex ) -and $userIndex -ge 0 -and $userIndex -lt $userLookupList.Count )
                    {
                        $expandedUser = [string]$userLookupList[ $userIndex ]
                    }
                    elseif( -Not [string]::IsNullOrWhiteSpace( $userTokenOrName ) -and $userLookupMap.ContainsKey( $userTokenOrName ) )
                    {
                        $expandedUser = $userLookupMap[ $userTokenOrName ]
                    }
                }

                if( $isCompact -and ( [string]::IsNullOrWhiteSpace( $expandedUser ) -or $expandedUser -match '^-?\d+$' ) )
                {
                    $expandedUser = ''
                }

                [string]$expandedPath = if( $isCompact ) { [string]$processInfo.x } else { [string]$processInfo.Path }
                if( $isCompact )
                {
                    [int]$pathIndex = -1
                    if( $pathLookupList.Count -gt 0 -and [int]::TryParse( $expandedPath , [ref]$pathIndex ) -and $pathIndex -ge 0 -and $pathIndex -lt $pathLookupList.Count )
                    {
                        $expandedPath = [string]$pathLookupList[ $pathIndex ]
                    }
                    else
                    {
                        $expandedPath = ''
                    }
                }

                [string]$expandedStartTime = if( $isCompact ) { [string]$processInfo.t } else { [string]$processInfo.StartTime }
                if( [string]::IsNullOrWhiteSpace( $expandedStartTime ) )
                {
                    $expandedStartTime = 'n/a'
                }

                $results += [pscustomobject]@{
                    Name = if( $isCompact ) { [string]$processInfo.n } else { [string]$processInfo.Name }
                    PID = if( $isCompact ) { [int]$processInfo.p } else { [int]$processInfo.PID }
                    CPU = if( $isCompact ) { [double]$processInfo.c } else { [double]$processInfo.CPU }
                    'Working Set (MB)' = if( $isCompact ) { [double]$processInfo.w } else { [double]$processInfo.'Working Set (MB)' }
                    Path = [string]$expandedPath
                    SessionId = if( $isCompact ) { [int]$processInfo.s } else { [int]$processInfo.SessionId }
                    User = [string]$expandedUser
                    StartTime = [string]$expandedStartTime
                    VM = [string]$sessionRow.SelectionObject.Name
                    ResourceGroup = [string]$sessionRow.SelectionObject.ResourceGroup
                    HostPool = [string]$sessionRow.HostPool
                }
            }
        }
        catch
        {
            Write-Warning -Message "Failed to get process list for VM $($sessionRow.SelectionObject.Name) session $targetSessionId : $_"
            [void][Windows.MessageBox]::Show( $mainWindow , "Failed to get process list for VM $($sessionRow.SelectionObject.Name) session $targetSessionId`n$($_.Exception.Message)" , 'AVD Session Processes' , 'Ok' ,'Error' )
        }
    }

    $results
}

Function Invoke-AzureAVDKillSessionProcesses
{
    Param
    (
        [Parameter(Mandatory)][array]$processRows
    )

    [array]$killResults = @()
    [array]$groups = @( $processRows | Group-Object -Property VM , ResourceGroup , SessionId , User )

    ForEach( $group in $groups )
    {
        if( $null -eq $group.Group -or $group.Group.Count -eq 0 )
        {
            continue
        }

        [pscustomobject]$first = $group.Group[0]
        [string]$vmName = [string]$first.VM
        [string]$resourceGroupName = [string]$first.ResourceGroup
        [int]$targetSessionId = [int]$first.SessionId
        [string]$targetUser = [string]$first.User
        [string]$escapedUser = ( $targetUser -replace "'" , "''" )
        [string]$pidList = ( @( $group.Group | Select-Object -ExpandProperty PID -Unique ) | ForEach-Object { [int]$_ } ) -join ','

        [string]$scriptToRun = @"
`$targetSessionId = $targetSessionId
`$targetUser = '$escapedUser'
`$targetUserShort = if( [string]::IsNullOrWhiteSpace( `$targetUser ) ) { '' } elseif( `$targetUser -match '@' ) { ( `$targetUser -split '@' )[0] } elseif( `$targetUser -match '\\' ) { ( `$targetUser -split '\\' )[-1] } else { `$targetUser }
`$targetPids = @( $pidList )

`$results = ForEach( `$targetPid in `$targetPids )
{
    `$message = ''
    `$killed = `$false
    `$name = ''
    try
    {
        `$process = Get-Process -Id `$targetPid -ErrorAction Stop
        `$name = `$process.ProcessName
        if( `$process.SessionId -ne `$targetSessionId )
        {
            throw "Session mismatch: expected `$targetSessionId got `$(`$process.SessionId)"
        }

        `$ownerName = ''
        try
        {
            `$processWmi = Get-CimInstance -ClassName Win32_Process -Filter "ProcessId = `$targetPid" -ErrorAction Stop
            `$ownerResult = Invoke-CimMethod -InputObject `$processWmi -MethodName GetOwner -ErrorAction Stop
            if( `$ownerResult.ReturnValue -eq 0 -and -Not [string]::IsNullOrWhiteSpace( `$ownerResult.User ) )
            {
                `$ownerName = if( [string]::IsNullOrWhiteSpace( `$ownerResult.Domain ) ) { `$ownerResult.User } else { "`$(`$ownerResult.Domain)\`$(`$ownerResult.User)" }
            }
        }
        catch {}

        if( -Not [string]::IsNullOrWhiteSpace( `$targetUserShort ) -and -Not [string]::IsNullOrWhiteSpace( `$ownerName ) )
        {
            `$ownerShort = if( `$ownerName -match '\\' ) { ( `$ownerName -split '\\' )[-1] } else { `$ownerName }
            if( `$ownerShort -ine `$targetUserShort -and `$ownerName -ine `$targetUser )
            {
                throw "User mismatch: expected `$targetUser got `$ownerName"
            }
        }

        Stop-Process -Id `$targetPid -Force -ErrorAction Stop
        `$killed = `$true
        `$message = 'Killed'
    }
    catch
    {
        `$message = `$_.Exception.Message
    }

    [pscustomobject]@{
        PID = `$targetPid
        Name = `$name
        Killed = `$killed
        Message = `$message
    }
}

`$json = ConvertTo-Json -InputObject @(`$results) -Depth 4 -Compress
if( [string]::IsNullOrWhiteSpace( `$json ) )
{
    `$json = '[]'
}
Write-Output '__MSTSCPROCESSJSON_BEGIN__'
Write-Output `$json
Write-Output '__MSTSCPROCESSJSON_END__'
"@

        try
        {
            [array]$thisKillResults = @( Invoke-AzureAVDRunCommandJson -vmName $vmName -resourceGroupName $resourceGroupName -scriptText $scriptToRun )
            ForEach( $killResult in $thisKillResults )
            {
                $killResults += [pscustomobject]@{
                    VM = $vmName
                    ResourceGroup = $resourceGroupName
                    SessionId = $targetSessionId
                    User = $targetUser
                    PID = [int]$killResult.PID
                    Name = [string]$killResult.Name
                    Killed = [bool]$killResult.Killed
                    Message = [string]$killResult.Message
                }
            }
        }
        catch
        {
            Write-Warning -Message "Failed to kill process on VM $vmName : $_"
            $killResults += [pscustomobject]@{
                VM = $vmName
                ResourceGroup = $resourceGroupName
                SessionId = $targetSessionId
                User = $targetUser
                PID = 0
                Name = ''
                Killed = $false
                Message = $_.Exception.Message
            }
        }
    }

    $killResults
}

Function Show-AzureAVDSessionProcessListView
{
    Param
    (
        [Parameter(Mandatory)][array]$sessionRows
    )

    $mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
    [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
    [void]$mainWindow.Dispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Render,
        [System.Action]{ }
    )

    try
    {
        [array]$processRows = @( Get-AzureAVDSessionProcesses -sessionRows $sessionRows )
        if( $processRows.Count -eq 0 )
        {
            [void][Windows.MessageBox]::Show( $mainWindow , 'No matching processes found for the selected session(s)' , 'AVD Session Processes' , 'Ok' ,'Information' )
            return
        }

        $processForm = New-Object System.Windows.Forms.Form
        $processForm.Text = 'AVD Session Processes'
        $processForm.Size = New-Object System.Drawing.Size(1240, 680)
        $processForm.StartPosition = 'CenterScreen'
        $processForm.KeyPreview = $true
        $processForm.Add_KeyDown({
            Param(
                [Parameter(Mandatory)][Object]$sender,
                [Parameter(Mandatory)][System.Windows.Forms.KeyEventArgs]$keyInfo
            )

            if( $keyInfo.KeyCode -eq [System.Windows.Forms.Keys]::Escape )
            {
                $keyInfo.Handled = $true
                $keyInfo.SuppressKeyPress = $true
                $processForm.Close()
            }
        })

        $processListView = New-Object System.Windows.Forms.ListView
        $processListView.Dock = 'Fill'
        $processListView.View = [System.Windows.Forms.View]::Details
        $processListView.FullRowSelect = $true
        $processListView.GridLines = $true
        $processListView.MultiSelect = $true
        $processListView.HideSelection = $false

        [void]$processListView.Columns.Add( 'Name' , 220 )
        [void]$processListView.Columns.Add( 'PID' , 80 )
        [void]$processListView.Columns.Add( 'CPU (s)' , 90 )
        [void]$processListView.Columns.Add( 'Working Set (MB)' , 130 )
        [void]$processListView.Columns.Add( 'Path' , 360 )
        [void]$processListView.Columns.Add( 'Session Id' , 90 )
        [void]$processListView.Columns.Add( 'User' , 220 )
        [void]$processListView.Columns.Add( 'VM' , 170 )
        [void]$processListView.Columns.Add( 'Start Time' , 160 )

        [array]$processColumnMap = @(
            [pscustomobject]@{ Header = 'Name' ; Property = 'Name' } ,
            [pscustomobject]@{ Header = 'PID' ; Property = 'PID' } ,
            [pscustomobject]@{ Header = 'CPU (s)' ; Property = 'CPU' } ,
            [pscustomobject]@{ Header = 'Working Set (MB)' ; Property = 'Working Set (MB)' } ,
            [pscustomobject]@{ Header = 'Path' ; Property = 'Path' } ,
            [pscustomobject]@{ Header = 'Session Id' ; Property = 'SessionId' } ,
            [pscustomobject]@{ Header = 'User' ; Property = 'User' } ,
            [pscustomobject]@{ Header = 'VM' ; Property = 'VM' } ,
            [pscustomobject]@{ Header = 'Start Time' ; Property = 'StartTime' }
        )

        $processAllRows = @( $processRows )
        [hashtable]$processFilters = @{}
        [hashtable]$processViewState = @{
            SortProperty = 'Name'
            SortDescending = $false
            LastColumnIndex = 0
        }

        $getProcessColumnProperty = {
            Param
            (
                [int]$columnIndex
            )

            if( $columnIndex -lt 0 -or $columnIndex -ge $processColumnMap.Count )
            {
                return $null
            }

            return [string]$processColumnMap[ $columnIndex ].Property
        }

        $updateProcessColumnHeaders = {
            for( [int]$i = 0 ; $i -lt $processColumnMap.Count ; $i++ )
            {
                [string]$header = [string]$processColumnMap[ $i ].Header
                [string]$property = [string]$processColumnMap[ $i ].Property

                if( $processFilters.ContainsKey( $property ) -and -Not [string]::IsNullOrWhiteSpace( [string]$processFilters[ $property ] ) )
                {
                    $header += ' *'
                }

                if( [string]$processViewState.SortProperty -eq $property )
                {
                    $header += $( if( [bool]$processViewState.SortDescending ) { ' (desc)' } else { ' (asc)' } )
                }

                $processListView.Columns[ $i ].Text = $header
            }
        }

        $populateList = {
        Param
        (
            $rows
        )

        $processListView.BeginUpdate()
        try
        {
            $processListView.Items.Clear()
            ForEach( $row in $rows )
            {
                $item = New-Object System.Windows.Forms.ListViewItem( [string]$row.Name )
                [void]$item.SubItems.Add( [string]$row.PID )
                [void]$item.SubItems.Add( [string]([math]::Round( [double]$row.CPU , 2 )) )
                [void]$item.SubItems.Add( [string]([math]::Round( [double]$row.'Working Set (MB)' , 2 )) )
                [void]$item.SubItems.Add( [string]$row.Path )
                [void]$item.SubItems.Add( [string]$row.SessionId )
                [void]$item.SubItems.Add( [string]$row.User )
                [void]$item.SubItems.Add( [string]$row.VM )
                [void]$item.SubItems.Add( [string]$row.StartTime )
                $item.Tag = $row
                [void]$processListView.Items.Add( $item )
            }
        }
        finally
        {
            $processListView.EndUpdate()
        }
        }

        $getProcessSortValue = {
            Param
            (
                $row,
                [string]$property
            )

            switch( $property )
            {
                'PID'
                {
                    try { return [long]$row.PID } catch { return 0L }
                }
                'CPU'
                {
                    try { return [double]$row.CPU } catch { return 0.0 }
                }
                'Working Set (MB)'
                {
                    try { return [double]$row.'Working Set (MB)' } catch { return 0.0 }
                }
                'SessionId'
                {
                    try { return [long]$row.SessionId } catch { return 0L }
                }
                'StartTime'
                {
                    [string]$raw = [string]$row.StartTime
                    [datetime]$parsed = [datetime]::MinValue
                    if( [datetime]::TryParse( $raw , [ref]$parsed ) )
                    {
                        return $parsed
                    }

                    return $raw
                }
                default
                {
                    if( $null -eq $row.PSObject.Properties[ $property ] )
                    {
                        return ''
                    }

                    return [string]$row.( $property )
                }
            }
        }

        $applyProcessRows = {
            $rows = @( $processAllRows )

            if( $processFilters.Count -gt 0 )
            {
                $rows = @( $rows | Where-Object {
                    $thisRow = $_
                    ForEach( $filter in $processFilters.GetEnumerator() )
                    {
                        if( [string]::IsNullOrWhiteSpace( [string]$filter.Value ) )
                        {
                            continue
                        }

                        [string]$itemValue = ''
                        if( $thisRow.PSObject.Properties[ $filter.Key ] )
                        {
                            $itemValue = [string]$thisRow.( $filter.Key )
                        }

                        [string]$pattern = [string]$filter.Value
                        if( $pattern -notmatch '[\*\?\[]' )
                        {
                            $pattern = "*$pattern*"
                        }

                        if( $itemValue -notlike $pattern )
                        {
                            return $false
                        }
                    }

                    return $true
                })
            }

            [string]$sortProperty = [string]$processViewState.SortProperty
            [bool]$sortDescending = [bool]$processViewState.SortDescending

            if( -Not [string]::IsNullOrWhiteSpace( $sortProperty ) )
            {
                $rows = @( $rows | Sort-Object -Property @{ Expression = { & $getProcessSortValue $_ $sortProperty } } , @{ Expression = { & $getProcessSortValue $_ 'PID' } } -Descending:$sortDescending )
            }
            else
            {
                $rows = @( $rows | Sort-Object -Property Name )
            }

            & $populateList $rows
            & $updateProcessColumnHeaders
        }

        & $applyProcessRows

        $contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
        $menuItemRefresh = New-Object System.Windows.Forms.ToolStripMenuItem 'Refresh'
        $menuItemSetFilter = New-Object System.Windows.Forms.ToolStripMenuItem 'Set Filter...'
        $menuItemClearThisFilter = New-Object System.Windows.Forms.ToolStripMenuItem 'Clear This Column Filter'
        $menuItemClearAllFilters = New-Object System.Windows.Forms.ToolStripMenuItem 'Clear All Filters'
        $menuItemKill = New-Object System.Windows.Forms.ToolStripMenuItem 'Kill Selected Process(es)'
        [void]$contextMenu.Items.AddRange( @( $menuItemRefresh , $menuItemSetFilter , $menuItemClearThisFilter , $menuItemClearAllFilters , $menuItemKill ) )
        $processListView.ContextMenuStrip = $contextMenu

        $showFilterDialog = {
            $dialog = New-Object System.Windows.Forms.Form
            $dialog.Text = 'Process Column Filter'
            $dialog.StartPosition = 'CenterParent'
            $dialog.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
            $dialog.MinimizeBox = $false
            $dialog.MaximizeBox = $false
            $dialog.Size = New-Object System.Drawing.Size(480, 190)

            $labelColumn = New-Object System.Windows.Forms.Label
            $labelColumn.Text = 'Column'
            $labelColumn.Location = New-Object System.Drawing.Point(12, 16)
            $labelColumn.AutoSize = $true

            $comboColumn = New-Object System.Windows.Forms.ComboBox
            $comboColumn.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
            $comboColumn.Location = New-Object System.Drawing.Point(90, 12)
            $comboColumn.Size = New-Object System.Drawing.Size(360, 24)

            ForEach( $columnDefinition in $processColumnMap )
            {
                [void]$comboColumn.Items.Add( [string]$columnDefinition.Header )
            }
            if( [int]$processViewState.LastColumnIndex -ge 0 -and [int]$processViewState.LastColumnIndex -lt $comboColumn.Items.Count )
            {
                $comboColumn.SelectedIndex = [int]$processViewState.LastColumnIndex
            }
            else
            {
                $comboColumn.SelectedIndex = 0
            }

            $labelFilter = New-Object System.Windows.Forms.Label
            $labelFilter.Text = 'Filter (supports wildcard * ? )'
            $labelFilter.Location = New-Object System.Drawing.Point(12, 54)
            $labelFilter.AutoSize = $true

            $textFilter = New-Object System.Windows.Forms.TextBox
            $textFilter.Location = New-Object System.Drawing.Point(15, 76)
            $textFilter.Size = New-Object System.Drawing.Size(435, 24)

            $buttonSet = New-Object System.Windows.Forms.Button
            $buttonSet.Text = 'Set'
            $buttonSet.Location = New-Object System.Drawing.Point(190, 112)
            $buttonSet.Size = New-Object System.Drawing.Size(80, 28)
            $buttonSet.DialogResult = [System.Windows.Forms.DialogResult]::OK

            $buttonClear = New-Object System.Windows.Forms.Button
            $buttonClear.Text = 'Clear'
            $buttonClear.Location = New-Object System.Drawing.Point(280, 112)
            $buttonClear.Size = New-Object System.Drawing.Size(80, 28)
            $buttonClear.DialogResult = [System.Windows.Forms.DialogResult]::Retry

            $buttonCancel = New-Object System.Windows.Forms.Button
            $buttonCancel.Text = 'Cancel'
            $buttonCancel.Location = New-Object System.Drawing.Point(370, 112)
            $buttonCancel.Size = New-Object System.Drawing.Size(80, 28)
            $buttonCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

            $loadFilterText = {
                [string]$property = & $getProcessColumnProperty $comboColumn.SelectedIndex
                if( -Not [string]::IsNullOrWhiteSpace( $property ) -and $processFilters.ContainsKey( $property ) )
                {
                    $textFilter.Text = [string]$processFilters[ $property ]
                }
                else
                {
                    $textFilter.Text = ''
                }
                [void]$textFilter.SelectAll()
            }

            $comboColumn.Add_SelectedIndexChanged({ & $loadFilterText })
            & $loadFilterText

            [void]$dialog.Controls.AddRange( @( $labelColumn , $comboColumn , $labelFilter , $textFilter , $buttonSet , $buttonClear , $buttonCancel ) )
            $dialog.AcceptButton = $buttonSet
            $dialog.CancelButton = $buttonCancel

            $result = $dialog.ShowDialog( $processForm )
            if( $result -eq [System.Windows.Forms.DialogResult]::Cancel )
            {
                return
            }

            [int]$selectedColumn = $comboColumn.SelectedIndex
            if( $selectedColumn -lt 0 )
            {
                return
            }

            $processViewState.LastColumnIndex = $selectedColumn
            [string]$property = & $getProcessColumnProperty $selectedColumn
            if( [string]::IsNullOrWhiteSpace( $property ) )
            {
                return
            }

            if( $result -eq [System.Windows.Forms.DialogResult]::Retry )
            {
                if( $processFilters.ContainsKey( $property ) )
                {
                    $processFilters.Remove( $property )
                }
            }
            else
            {
                [string]$value = [string]$textFilter.Text
                if( [string]::IsNullOrWhiteSpace( $value ) )
                {
                    if( $processFilters.ContainsKey( $property ) )
                    {
                        $processFilters.Remove( $property )
                    }
                }
                else
                {
                    $processFilters[ $property ] = $value.Trim()
                }
            }

            & $applyProcessRows
        }

        $refreshFromSessions = {
            $processForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            try
            {
                $latest = @( Get-AzureAVDSessionProcesses -sessionRows $sessionRows )
                $processRows = $latest
                $processAllRows = @( $latest )
                & $applyProcessRows
            }
            finally
            {
                $processForm.Cursor = [System.Windows.Forms.Cursors]::Default
            }
        }

        $menuItemRefresh.Add_Click({ & $refreshFromSessions })
        $menuItemSetFilter.Add_Click({ & $showFilterDialog })
        $menuItemClearThisFilter.Add_Click({
            [string]$property = & $getProcessColumnProperty ([int]$processViewState.LastColumnIndex)
            if( -Not [string]::IsNullOrWhiteSpace( $property ) -and $processFilters.ContainsKey( $property ) )
            {
                $processFilters.Remove( $property )
                & $applyProcessRows
            }
        })
        $menuItemClearAllFilters.Add_Click({
            if( $processFilters.Count -gt 0 )
            {
                $processFilters.Clear()
                & $applyProcessRows
            }
        })

        $processListView.Add_ColumnClick({
            Param(
                [Parameter(Mandatory)][Object]$sourceControl,
                [Parameter(Mandatory)][System.Windows.Forms.ColumnClickEventArgs]$columnInfo
            )

            [int]$clickedColumn = [int]$columnInfo.Column
            $processViewState.LastColumnIndex = $clickedColumn
            [string]$property = & $getProcessColumnProperty $clickedColumn
            if( [string]::IsNullOrWhiteSpace( $property ) )
            {
                return
            }

            if( [string]$processViewState.SortProperty -eq $property )
            {
                $processViewState.SortDescending = -Not [bool]$processViewState.SortDescending
            }
            else
            {
                $processViewState.SortProperty = $property
                $processViewState.SortDescending = $false
            }

            & $applyProcessRows
        })

        $processListView.Add_MouseDown({
        Param(
            [Parameter(Mandatory)][Object]$sourceControl,
            [Parameter(Mandatory)][System.Windows.Forms.MouseEventArgs]$mouseInfo
        )

        if( $mouseInfo.Button -ne [System.Windows.Forms.MouseButtons]::Right )
        {
            return
        }

        $hitItem = $processListView.HitTest( $mouseInfo.Location ).Item
        if( $null -ne $hitItem -and -Not $hitItem.Selected )
        {
            ForEach( $listItem in $processListView.Items )
            {
                $listItem.Selected = $false
            }
            $hitItem.Selected = $true
            $hitItem.Focused = $true
        }
        })

        $menuItemKill.Add_Click({
        [array]$targets = @()
        ForEach( $selectedItem in $processListView.SelectedItems )
        {
            if( $null -ne $selectedItem.Tag )
            {
                $targets += $selectedItem.Tag
            }
        }

        if( $targets.Count -eq 0 )
        {
            [void][Windows.MessageBox]::Show( $mainWindow , 'No processes selected' , 'AVD Session Processes' , 'Ok' ,'Information' )
            return
        }

        [string]$prompt = if( $targets.Count -eq 1 )
        {
            "Kill process $($targets[0].Name) (PID $($targets[0].PID)) on $($targets[0].VM)?"
        }
        else
        {
            "Kill $($targets.Count) selected processes?"
        }

        if( [Windows.MessageBox]::Show( $mainWindow , $prompt , 'Confirm Process Kill' , 'YesNo' ,'Warning' ) -ine 'Yes' )
        {
            return
        }

        $processForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try
        {
            [array]$killResults = @( Invoke-AzureAVDKillSessionProcesses -processRows $targets )

            [hashtable]$killedByKey = @{}
            [array]$errors = @()
            ForEach( $killResult in $killResults )
            {
                if( $killResult.Killed )
                {
                    [string]$key = "{0}|{1}|{2}|{3}" -f $killResult.VM , $killResult.SessionId , $killResult.User , $killResult.PID
                    $killedByKey[ $key ] = $true
                }
                else
                {
                    $errors += $killResult
                }
            }

            [array]$survivors = @()
            ForEach( $row in $processAllRows )
            {
                [string]$key = "{0}|{1}|{2}|{3}" -f $row.VM , $row.SessionId , $row.User , $row.PID
                if( -Not $killedByKey.ContainsKey( $key ) )
                {
                    $survivors += $row
                }
            }
            $processAllRows = @( $survivors )

            & $applyProcessRows

            if( $errors.Count -gt 0 )
            {
                [string]$errorText = ( $errors | ForEach-Object { "VM=$($_.VM) Session=$($_.SessionId) PID=$($_.PID) $($_.Message)" } ) -join "`r`n"
                [void][Windows.MessageBox]::Show( $mainWindow , "Some process kills failed:`r`n$errorText" , 'AVD Session Processes' , 'Ok' ,'Warning' )
            }
        }
        finally
        {
            $processForm.Cursor = [System.Windows.Forms.Cursors]::Default
        }
        })

        $processForm.Controls.Add( $processListView )
        $processForm.Add_Shown({ 
            $processForm.Activate()
            $processListView.Focus()
        })
    }
    finally
    {
        $mainWindow.ClearValue( [System.Windows.FrameworkElement]::CursorProperty )
        [System.Windows.Input.Mouse]::OverrideCursor = $null
    }

    [void]$processForm.ShowDialog()
}

Function Get-AzureAVDSessionMessage
{
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'AVD Session Message'
    $form.Size = New-Object System.Drawing.Size(640, 320)
    $form.StartPosition = 'CenterScreen'

    $labelTitle = New-Object System.Windows.Forms.Label
    $labelTitle.Text = 'Message title'
    $labelTitle.AutoSize = $true
    $labelTitle.Location = New-Object System.Drawing.Point(12, 12)

    $textTitle = New-Object System.Windows.Forms.TextBox
    $textTitle.Location = New-Object System.Drawing.Point(12, 32)
    $textTitle.Size = New-Object System.Drawing.Size(600, 24)
    $textTitle.Text = 'Admin Message'

    $labelBody = New-Object System.Windows.Forms.Label
    $labelBody.Text = 'Message body'
    $labelBody.AutoSize = $true
    $labelBody.Location = New-Object System.Drawing.Point(12, 66)

    $textBody = New-Object System.Windows.Forms.TextBox
    $textBody.Location = New-Object System.Drawing.Point(12, 86)
    $textBody.Size = New-Object System.Drawing.Size(600, 150)
    $textBody.Multiline = $true
    $textBody.ScrollBars = 'Vertical'
    $textBody.Text = 'Please save your work and sign out.'

    $buttonOk = New-Object System.Windows.Forms.Button
    $buttonOk.Text = 'OK'
    $buttonOk.Size = New-Object System.Drawing.Size(90, 30)
    $buttonOk.Location = New-Object System.Drawing.Point(432, 246)
    $buttonOk.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $buttonCancel = New-Object System.Windows.Forms.Button
    $buttonCancel.Text = 'Cancel'
    $buttonCancel.Size = New-Object System.Drawing.Size(90, 30)
    $buttonCancel.Location = New-Object System.Drawing.Point(522, 246)
    $buttonCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    [void]$form.Controls.Add( $labelTitle )
    [void]$form.Controls.Add( $textTitle )
    [void]$form.Controls.Add( $labelBody )
    [void]$form.Controls.Add( $textBody )
    [void]$form.Controls.Add( $buttonOk )
    [void]$form.Controls.Add( $buttonCancel )

    $form.AcceptButton = $buttonOk
    $form.CancelButton = $buttonCancel

    if( $form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK )
    {
        return $null
    }

    [string]$messageTitle = $textTitle.Text.Trim()
    [string]$messageBody = $textBody.Text.Trim()
    if( [string]::IsNullOrEmpty( $messageTitle ) -or [string]::IsNullOrEmpty( $messageBody ) )
    {
        return $null
    }

    [pscustomobject]@{
        Title = $messageTitle
        Body = $messageBody
    }
}

Function Invoke-AzureAVDSessionOperation
{
    Param
    (
        [Parameter(Mandatory)][array]$userSessions ,
        [Parameter(Mandatory)][string]$operation ,
        [pscustomobject]$azureSessionMessage
    )

    [bool]$refreshAzureList = $false

    if( $operation -ieq 'Azure_MessageSession' -and $null -eq $azureSessionMessage )
    {
        $azureSessionMessage = Get-AzureAVDSessionMessage
        if( $null -eq $azureSessionMessage )
        {
            return [pscustomobject]@{ RefreshAzureList = $false ; AzureSessionMessage = $null }
        }
    }

    ForEach( $userSession in $userSessions )
    {
        if( $operation -ieq 'Azure_DisconnectSession' )
        {
            $null = Invoke-AzRestMethod -Method POST -Path "$($userSession.Id)/disconnect?api-version=$AVDAPIversion" -ErrorAction Stop
            $refreshAzureList = $true
        }
        elseif( $operation -ieq 'Azure_LogoffSession' )
        {
            $null = Invoke-AzRestMethod -Method DELETE -Path "$($userSession.Id)?api-version=$AVDAPIversion" -ErrorAction Stop
            $refreshAzureList = $true
        }
        elseif( $operation -ieq 'Azure_ForceLogoffSession' )
        {
            $null = Invoke-AzRestMethod -Method DELETE -Path "$($userSession.Id)?api-version=$AVDAPIversion&force=true" -ErrorAction Stop
            $refreshAzureList = $true
        }
        elseif( $operation -ieq 'Azure_MessageSession' )
        {
            [string]$payload = (@{ messageTitle = $azureSessionMessage.Title ; messageBody = $azureSessionMessage.Body } | ConvertTo-Json -Compress)
            $null = Invoke-AzRestMethod -Method POST -Path "$($userSession.Id)/sendMessage?api-version=$AVDAPIversion" -Payload $payload -ErrorAction Stop
        }
    }

    [pscustomobject]@{
        RefreshAzureList = $refreshAzureList
        AzureSessionMessage = $azureSessionMessage
    }
}

Function Confirm-AzureAVDSessionOperation
{
    Param
    (
        [Parameter(Mandatory)][array]$userSessions ,
        [Parameter(Mandatory)][string]$operation
    )

    if( $operation -notin @( 'Azure_DisconnectSession' , 'Azure_LogoffSession' , 'Azure_ForceLogoffSession' ) )
    {
        return $true
    }

    [string]$actionName = $operation -replace '^Azure_' , '' -replace 'Session$' , '' -creplace '([a-zA-Z])([A-Z])' , '$1 $2'
    [string]$prompt = if( $userSessions.Count -gt 1 )
    {
        "Are you sure you want to $actionName $($userSessions.Count) sessions?"
    }
    else
    {
        [string]$sessionUser = $userSessions[0].properties.userPrincipalName
        [string]$sessionName = $userSessions[0].Name -replace '^.*/' , ''
        "Are you sure you want to $actionName session $sessionName for $sessionUser?"
    }

    return [Windows.MessageBox]::Show( $mainWindow , $prompt , 'Confirm Session Operation' , 'YesNo' ,'Question' ) -ieq 'Yes'
}

Function Invoke-ExplorerStyleAllSelectedClick
{
    Param
    (
        [Parameter(Mandatory)][System.Windows.Controls.ListView]$listView,
        [Parameter(Mandatory)][System.Windows.Input.MouseButtonEventArgs]$mouseInfo
    )

    [int]$totalItems = $listView.Items.Count
    if( $totalItems -le 0 -or $listView.SelectedItems.Count -ne $totalItems )
    {
        return $false
    }

    $element = $mouseInfo.OriginalSource
    while( $element -and $element -isnot [System.Windows.Controls.ListViewItem] )
    {
        $element = [System.Windows.Media.VisualTreeHelper]::GetParent( $element )
    }

    if( $element -isnot [System.Windows.Controls.ListViewItem] -or $null -eq $element.DataContext -or $element.DataContext -is [System.Windows.Data.CollectionView] )
    {
        return $false
    }

    $clickedItem = $element.DataContext
    [bool]$ctrlDown = ( [System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Control ) -ne [System.Windows.Input.ModifierKeys]::None
    if( $ctrlDown )
    {
        [void]$listView.SelectedItems.Remove( $clickedItem )
    }
    else
    {
        $listView.SelectedItems.Clear()
        [void]$listView.SelectedItems.Add( $clickedItem )
        $listView.SelectedItem = $clickedItem
    }

    $mouseInfo.Handled = $true
    return $true
}

Function Show-AzureAVDSessionListView
{
    Param
    (
        [Parameter(Mandatory)][array]$selectedVMs
    )

    [array]$sessionRows = @()
    
    ForEach( $selection in $selectedVMs )
    {
        try
        {
            $sessionHost = Get-AzureAVDSessionHost -selection $selection
            [array]$userSessions = @( Get-AzureAVDUserSessions -sessionHostResourceId $sessionHost.Id )

            ForEach( $userSession in $userSessions )
            {
                [string]$sessionType = if( -Not [string]::IsNullOrWhiteSpace( $userSession.properties.applicationType ) ) { [string]$userSession.properties.applicationType } else { 'Unknown' }
                
                $sessionRows += [pscustomobject]@{
                    'Session Id' = $userSession.Name -replace '^.*/' , ''
                    User = $userSession.properties.userPrincipalName
                    Type = $sessionType
                    State = $userSession.properties.sessionState
                    'Create Time' = $userSession.properties.createTime
                    VM = $selection.Name
                    HostPool = $selection.HostPool
                    SelectionObject = $selection
                    SessionObject = $userSession
                }
            }
        }
        catch
        {
            Write-Warning -Message "Azure AVD session retrieval failed for $($selection.Name) : $_"
            [void][Windows.MessageBox]::Show( $mainWindow , "Azure AVD session retrieval failed for $($selection.Name)`n$($_.Exception.Message)" , 'Azure AVD Sessions' , 'Ok' ,'Error' )
        }
    }

    if( $sessionRows.Count -eq 0 )
    {
        [void][Windows.MessageBox]::Show( $mainWindow , 'No active user sessions found on the selected session hosts' , 'Azure AVD Sessions' , 'Ok' ,'Information' )
        return
    }

    $detailForm = New-Object System.Windows.Forms.Form
    $detailForm.Text = 'AVD Sessions'
    $detailForm.Size = New-Object System.Drawing.Size(980, 520)
    $detailForm.StartPosition = 'CenterScreen'
    $detailForm.KeyPreview = $true
    $detailForm.Add_KeyDown({
        Param
        (
            [Parameter(Mandatory)][Object]$sender,
            [Parameter(Mandatory)][System.Windows.Forms.KeyEventArgs]$keyInfo
        )

        if( $keyInfo.KeyCode -eq [System.Windows.Forms.Keys]::Escape )
        {
            $keyInfo.Handled = $true
            $keyInfo.SuppressKeyPress = $true
            $detailForm.Close()
        }
    })

    $sessionListView = New-Object System.Windows.Forms.ListView
    $sessionListView.Dock = 'Fill'
    $sessionListView.View = [System.Windows.Forms.View]::Details
    $sessionListView.FullRowSelect = $true
    $sessionListView.GridLines = $true
    $sessionListView.MultiSelect = $true
    $sessionListView.HideSelection = $false

    [pscustomobject]$sessionAllSelectedClickState = [pscustomobject]@{
        Apply = $false
        Ctrl = $false
        Item = $null
    }

    $sessionListView.Add_MouseDown({
        Param
        (
            [Parameter(Mandatory)][Object]$sourceControl,
            [Parameter(Mandatory)][System.Windows.Forms.MouseEventArgs]$mouseInfo
        )

        $sessionAllSelectedClickState.Apply = $false
        $sessionAllSelectedClickState.Ctrl = $false
        $sessionAllSelectedClickState.Item = $null

        if( $mouseInfo.Button -ne [System.Windows.Forms.MouseButtons]::Left )
        {
            return
        }

        [int]$totalItems = $sessionListView.Items.Count
        if( $totalItems -le 0 -or $sessionListView.SelectedItems.Count -ne $totalItems )
        {
            return
        }

        $hitItem = $sessionListView.HitTest( $mouseInfo.Location ).Item
        if( $null -eq $hitItem )
        {
            return
        }

        [bool]$ctrlDown = ( [System.Windows.Forms.Control]::ModifierKeys -band [System.Windows.Forms.Keys]::Control ) -eq [System.Windows.Forms.Keys]::Control
        $sessionAllSelectedClickState.Apply = $true
        $sessionAllSelectedClickState.Ctrl = $ctrlDown
        $sessionAllSelectedClickState.Item = $hitItem
    })

    $sessionListView.Add_MouseUp({
        Param
        (
            [Parameter(Mandatory)][Object]$sourceControl,
            [Parameter(Mandatory)][System.Windows.Forms.MouseEventArgs]$mouseInfo
        )

        if( $mouseInfo.Button -eq [System.Windows.Forms.MouseButtons]::Left -and $sessionAllSelectedClickState.Apply -and $null -ne $sessionAllSelectedClickState.Item )
        {
            if( $sessionAllSelectedClickState.Ctrl )
            {
                $sessionAllSelectedClickState.Item.Selected = $false
            }
            else
            {
                ForEach( $listItem in $sessionListView.Items )
                {
                    $listItem.Selected = $false
                }

                $sessionAllSelectedClickState.Item.Selected = $true
                $sessionAllSelectedClickState.Item.Focused = $true
            }
        }

        $sessionAllSelectedClickState.Apply = $false
        $sessionAllSelectedClickState.Ctrl = $false
        $sessionAllSelectedClickState.Item = $null
    })

    $sessionListView.Add_KeyDown({
        Param
        (
            [Parameter(Mandatory)][Object]$sourceControl,
            [Parameter(Mandatory)][System.Windows.Forms.KeyEventArgs]$keyInfo
        )
        if( $keyInfo.Control -and $keyInfo.KeyCode -eq [System.Windows.Forms.Keys]::A )
        {
            ForEach( $listItem in $sessionListView.Items )
            {
                $listItem.Selected = $true
            }
            $keyInfo.Handled = $true
            $keyInfo.SuppressKeyPress = $true
        }
    })

    [void]$sessionListView.Columns.Add( 'Session Id' , 120 )
    [void]$sessionListView.Columns.Add( 'User' , 260 )
    [void]$sessionListView.Columns.Add( 'Type' , 80 )
    [void]$sessionListView.Columns.Add( 'State' , 110 )
    [void]$sessionListView.Columns.Add( 'Create Time' , 180 )
    [void]$sessionListView.Columns.Add( 'VM' , 150 )
    [void]$sessionListView.Columns.Add( 'HostPool' , 150 )

    ForEach( $sessionRow in $sessionRows )
    {
        $item = New-Object System.Windows.Forms.ListViewItem( [string]$sessionRow.'Session Id' )
        [void]$item.SubItems.Add( [string]$sessionRow.User )
        [void]$item.SubItems.Add( [string]$sessionRow.Type )
        [void]$item.SubItems.Add( [string]$sessionRow.State )
        [void]$item.SubItems.Add( [string]$sessionRow.'Create Time' )
        [void]$item.SubItems.Add( [string]$sessionRow.VM )
        [void]$item.SubItems.Add( [string]$sessionRow.HostPool )
        $item.Tag = $sessionRow
        [void]$sessionListView.Items.Add( $item )
    }

    $contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $menuItemMessage = New-Object System.Windows.Forms.ToolStripMenuItem 'Message'
    $menuItemProcesses = New-Object System.Windows.Forms.ToolStripMenuItem 'Processes'
    $menuItemDisconnect = New-Object System.Windows.Forms.ToolStripMenuItem 'Disconnect'
    $menuItemLogoff = New-Object System.Windows.Forms.ToolStripMenuItem 'Logoff'
    $menuItemForceLogoff = New-Object System.Windows.Forms.ToolStripMenuItem 'Force Logoff'
    [void]$contextMenu.Items.AddRange( @( $menuItemMessage , $menuItemProcesses , $menuItemDisconnect , $menuItemLogoff , $menuItemForceLogoff ) )
    $sessionListView.ContextMenuStrip = $contextMenu

    [pscustomobject]$sessionMessage = $null
    [bool]$refreshAzureList = $false

    $invokeFromListView = {
        Param
        (
            [string]$sessionOperation
        )

        [array]$targetSessions = @()
        if( $sessionListView.SelectedItems.Count -gt 0 )
        {
            ForEach( $selectedItem in $sessionListView.SelectedItems )
            {
                if( $null -ne $selectedItem.Tag -and $null -ne $selectedItem.Tag.SessionObject )
                {
                    $targetSessions += $selectedItem.Tag.SessionObject
                }
            }
        }
        else
        {
            ForEach( $listItem in $sessionListView.Items )
            {
                if( $null -ne $listItem.Tag -and $null -ne $listItem.Tag.SessionObject )
                {
                    $targetSessions += $listItem.Tag.SessionObject
                }
            }
        }

        if( $targetSessions.Count -eq 0 )
        {
            return
        }

        if( -Not ( Confirm-AzureAVDSessionOperation -userSessions $targetSessions -operation $sessionOperation ) )
        {
            return
        }

        try
        {
            $result = Invoke-AzureAVDSessionOperation -userSessions $targetSessions -operation $sessionOperation -azureSessionMessage $sessionMessage
            $sessionMessage = $result.AzureSessionMessage
            if( $result.RefreshAzureList )
            {
                $refreshAzureList = $true
                [hashtable]$targetSessionIds = @{}
                ForEach( $targetSession in $targetSessions )
                {
                    $targetSessionIds[ $targetSession.Id ] = $true
                }

                for( [int]$index = $sessionListView.Items.Count - 1 ; $index -ge 0 ; $index-- )
                {
                    $listItem = $sessionListView.Items[ $index ]
                    if( $null -ne $listItem.Tag -and $null -ne $listItem.Tag.SessionObject -and $targetSessionIds.ContainsKey( $listItem.Tag.SessionObject.Id ) )
                    {
                        $sessionListView.Items.RemoveAt( $index )
                    }
                }
            }
        }
        catch
        {
            Write-Warning -Message "Azure AVD session operation failed : $_"
            [void][Windows.MessageBox]::Show( $mainWindow , "Azure AVD session operation failed`n$($_.Exception.Message)" , 'Azure AVD Sessions' , 'Ok' ,'Error' )
        }
    }

    $menuItemProcesses.Add_Click({
        [array]$targetRows = @()
        if( $sessionListView.SelectedItems.Count -gt 0 )
        {
            ForEach( $selectedItem in $sessionListView.SelectedItems )
            {
                if( $null -ne $selectedItem.Tag )
                {
                    $targetRows += $selectedItem.Tag
                }
            }
        }
        else
        {
            ForEach( $listItem in $sessionListView.Items )
            {
                if( $null -ne $listItem.Tag )
                {
                    $targetRows += $listItem.Tag
                }
            }
        }

        if( $targetRows.Count -eq 0 )
        {
            return
        }

        Show-AzureAVDSessionProcessListView -sessionRows $targetRows
    })

    $menuItemMessage.Add_Click({ & $invokeFromListView 'Azure_MessageSession' })
    $menuItemDisconnect.Add_Click({ & $invokeFromListView 'Azure_DisconnectSession' })
    $menuItemLogoff.Add_Click({ & $invokeFromListView 'Azure_LogoffSession' })
    $menuItemForceLogoff.Add_Click({ & $invokeFromListView 'Azure_ForceLogoffSession' })

    $sessionListView.Add_MouseDoubleClick({
        Param(
            [Parameter(Mandatory)][Object]$sourceControl,
            [Parameter(Mandatory)][System.Windows.Forms.MouseEventArgs]$mouseInfo
        )

        $hitItem = $sessionListView.HitTest( $mouseInfo.Location ).Item
        if( $null -eq $hitItem -or $null -eq $hitItem.Tag )
        {
            return
        }

        [array]$targetRows = @( $hitItem.Tag )
        Show-AzureAVDSessionProcessListView -sessionRows $targetRows
    })

    $detailForm.Controls.Add( $sessionListView )
    $detailForm.Add_Shown({ 
        $detailForm.Activate()
        $sessionListView.Focus()
    })
    [void]$detailForm.ShowDialog()

    if( $refreshAzureList )
    {
        Add-AzureVMsToListView -filter '' -regex $false -allVMs $WPFcheckBoxAzureAllVMs.IsChecked
    }
}

Function Add-VMwareVMsToListView
{
    Param
    (
        [string]$filter ,
        [bool]$regex
    )
    $vmwareError = $null
    $mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
    $script:vms = @( Get-VM -ErrorVariable vmwareError -norecursion | Where-Object { $_.PowerState -ieq 'PoweredOn' -and (( $regex -and $_.Name -match $filter ) -or ( -Not $regex -and ( [string]::IsNullOrEmpty( $filter ) -or $_.Name -like $filter ))) } | Sort-Object -Property Name | Select-Object -ExpandProperty Guest )
    $mainWindow.ClearValue([System.Windows.FrameworkElement]::CursorProperty) 
    if( $vmwareError )
    {
        [void][Windows.MessageBox]::Show( $vmwareError , 'VMware Error' , 'Ok' ,'Error' )
    }
    Write-Verbose -Message "Got $($vms.Count) powered on VMware VMs"
    $WPFlistViewVMwareVMs.Items.Clear()

    ForEach( $vm in $vms )
    {
        $WPFlistViewVMwareVMs.Items.Add( [pscustomobject]@{ Name = $vm.VmName  } )
        Write-Verbose -Message "Added $($vm.VmName)"
    }
    $WPFlabelVMwareVMs.Content = "$($WPFlistViewVMwareVMs.Items.Count) VMs"
}

Function Add-HyperVVMsToListView
{
    Param
    (
        [string]$hyperVhost ,
        [string]$filter ,
        [bool]$regex ,
        [bool]$allVMs
    )
    $hyperVError = $null
    [string]$powerState = '.'
    if( -Not $allVMs )
    {
        $powerState = 'Running'
    }
    ## module qualifying in case clash with VMware PowerCLI. Deal with multiple hosts
    $mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
    $script:vms = @( Hyper-V\Get-VM -ErrorVariable hyperVError -ComputerName ($hyperVhost -split ',') | Where-Object { $_.State -match $powerState -and (( $regex -and $_.Name -match $filter ) -or ( -Not $regex -and ( [string]::IsNullOrEmpty( $filter ) -or $_.Name -like $filter ))) } | Sort-Object -Property Name )
    $mainWindow.ClearValue([System.Windows.FrameworkElement]::CursorProperty)
    if( $hyperVError )
    {
        [void][Windows.MessageBox]::Show( $hyperVError , "Hyper-V Error from $hyperVhost" , 'Ok' ,'Error' )
    }
    Write-Verbose -Message "Got $($vms.Count) powered on Hyper-V VMs"
    $WPFlistViewHyperVVMs.Items.Clear()
    ForEach( $vm in $vms )
    {
        $WPFlistViewHyperVVMs.Items.Add( [pscustomobject]@{ Name = $vm.Name ; PowerState = $vm.State } ) ## value comes from what is in Binding property for the grid view column
    }
    $WPFlabelHyperVVMs.Content = "$($WPFlistViewHyperVVMs.Items.Count) VMs"
}

## TODO make this generic for detecting which hypervisor is connected to and use that 

Function Start-RemoteSessionFromHypervisor
{
    Param
    (
        [Parameter(Mandatory=$true)]
        [ValidateSet('VMware','Hyper-V','AD')]
        [string]$hypervisorType ,
        [switch]$console
    )
    [int]$loopIterations = 0
    [string]$answer = $null
    $listView = $WPFlistViewVMwareVMs
    $radioButton = $WPFradioButtonVMwareConnectByIP
    $rdpportTextbox = $WPFtextBoxVMwareRDPPort
    if( $hypervisorType -ieq 'Hyper-V' )
    {
        $listView = $WPFlistViewHyperVVMs
        $radioButton = $WPFradioButtonHyperVConnectByIP
        $rdpportTextbox = $WPFtextBoxHyperVRDPPort
    }
    elseif( $hypervisorType -ieq 'AD' )
    {
        $listView = $WPFlistViewAD
        $radioButton = $null
        $rdpportTextbox = $WPFtextBoxADRDPPort
    }
    if( $listview.SelectedIndex -ge 0 )
    {
        if( $hypervisorType -ine 'AD' -and ( $null -eq $script:vms -or $listView.SelectedIndex -gt $script:vms.Count ) )
        {
            Write-Error -Message "Internal error : selected grid view index $($listView.SelectedIndex) greater than $(($script:vms|Measure-Object).Count)"
            return
        }
        ForEach( $selection in $listView.selectedItems )
        {
            $loopIterations++
            if( $hypervisorType -ieq 'Hyper-V' )
            {
                $vm = $script:vms | Where-Object Name -eq $selection.Name
            }
            elseif( $hypervisorType -ieq 'AD' )
            {
                $vm = $selection.Name ## not a VM but this was hacked in later :-)
            }
            else
            {
                $vm = $script:vms | Where-Object VMName -eq $selection.Name
            }
            if( $null -eq $vm )
            {
                Write-Warning "Could not find VM for selected item $($selection.Name) out of $($script:vms.Count)"
                continue
            }
            if( $listview.selectedItems.Count -gt 1 -and ( [string]::IsNullOrEmpty( $answer ) -or $answer -eq 'No' ))
            {
                [string]$buttons = 'YesNoCancel'
                [string]$prompt = "Are you sure you want to connect to $($listview.selectedItems.Count - $loopIterations + 1) VMs?`nYes for all , No for $($selection.Name) only or Cancel for none"
                ## make it modal to the main window
                $answer = [Windows.MessageBox]::Show( $mainWindow , $prompt , 'Confirm Multiple Connections' , $buttons ,'Question' )
                if( $answer -ieq 'cancel' )
                {
                    return
                }
            }
            $address = $null
            
            if( $console )
            {
                if( $hypervisorType -ieq 'Hyper-V' )
                {
                    ## TODO see if there is already a running process with these arguments (ish) and offer to activate that instead
                    $consoleProcess = $null
                    [hashtable]$consoleArguments = @{
                        FilePath = 'vmconnect.exe' 
                        ArgumentList = @( $WPFtextBoxHyperVHost.Text , "`"$($vm.Name)`"" , '-G' , $vm.Id )
                        PassThru = $true
                    }
                    $consoleProcess = Start-Process @consoleArguments
                    if( $null -eq $consoleProcess )
                    {
                        [void][Windows.MessageBox]::Show( "Failed to run $($consoleArguments[ 'filepath']) $($consoleArguments[ 'argumentlist' ] -join ' ') :`r`n$(Error[0])" , 'Hypervisor Console Error' , 'Ok' ,'Error' )
                    }

                }
                else
                {
                    ## TODO launch VMware console - web or local app?
                }
            }
            elseif( $radioButton -and $radioButton.IsChecked )
            {
                ## TODO do we allow IPv6 ?
                if( $hypervisorType -ieq 'vmware' )
                {
                    $address = $vm.IPaddress | Where-Object { $_ -match '^\d+\.' -and $_ -ne '127.0.0.1' -and $_ -ne '::1' -and $_ -notmatch '^169\.254\.' }
                }
                else
                {
                    $address = Get-VMNetworkAdapter -VM $vm | Select-Object -ExpandProperty IPAddresses | Where-Object { $_ -match '^\d+\.' -and $_ -ne '127.0.0.1' -and $_ -ne '::1' -and $_ -notmatch '^169\.254\.' }
                }
                if( $null -eq $address )
                {
                    [void][Windows.MessageBox]::Show( "No IP address for $($vm.VmName)" , 'Hypervisor Error' , 'Ok' ,'Error' )
                }
                elseif( $address -is [array] -and $address.Count -gt 1 )
                {
                    [void][Windows.MessageBox]::Show( "$($address.Count) IP addresses for $($vm.VmName)" , 'Hypervisor Error' , 'Ok' ,'Error' )
                    ## TODO do we ask them to select one? Try in turn?
                    $address = $null
                }
            }
            elseif( $hypervisorType -ieq 'Hyper-V' )
            {
                $address = $vm.Name
            }
            elseif( $hypervisorType -ieq 'AD' )
            {
                $address = $vm
            }
            else ## VMware
            {
                $address = $vm.Hostname
            }
            if( $address )
            {
                if( -Not [string]::IsNullOrEmpty( $rdpportTextbox.Text ) )
                {
                    if( $rdpportTextbox.Text -notmatch '^\d+$' )
                    {
                        [void][Windows.MessageBox]::Show( "Port `"$($rdpportTextbox.Text)`" is invalid" , 'Hypervisor Error' , 'Ok' ,'Error' )
                        $address = $null
                    }
                    else
                    {
                        $address = "$($address):$($rdpportTextbox.Text)"
                    }
                }
                Write-Verbose -Message "Connecting to VM $address"
                if( $address )
                {
                    ## put into main computer list if not already there
                    if( -Not $wpfcomboboxComputer.Items.Contains( $address ) )
                    {
                        $wpfcomboboxComputer.Items.Add( $address )
                    }
                    Set-RemoteSessionProperties -connectTo $address
                    
                    [bool]$alreadyPresent = $false

                    ForEach( $item in $wpfcomboboxComputer.Items )
                    {
                        if( $alreadyPresent = $item -ieq $address )
                        {
                            break
                        }
                    }
                    if( -not $alreadyPresent )
                    {
                        $wpfcomboboxComputer.Items.Insert( 0 , $address ) ## TODO should we resort it ? Need to check if already there
                    }
                }
                else
                {
                    Write-Warning -Message "No address for $hypervisorType VM $($listView.SelectedItem)"
                }
            }
        }
    }
    else
    {
        [void][Windows.MessageBox]::Show( "No VM selected" , "$hypervisorType Error" , 'Ok' ,'Error' )
    }
}
    
Function Process-Action
{
    Param
    (
        $GUIobject , 
        [string]$Operation 
        ##$context  ,
        ##$thisObject
    )

    $thisObject = $_

    $thisObject.Handled = $true

    Write-Verbose -Message "Process-Action $operation "

    if( $GUIobject )
    {  
        [array]$selectedVMs = @( $GUIobject.selectedItems )
        if( $null -eq $selectedVMs -or $selectedVMs.Count -eq 0 )
        {
            Write-Verbose -Message "No items selected"
            [Windows.MessageBox]::Show( $mainWindow , 'No VMs selected' , 'Error' , 'OK' ,'Exclamation' )
            return
        }

        if( $operation -ieq 'Azure_DetailSession' )
        {
            Show-AzureAVDSessionListView -selectedVMs $selectedVMs
            return
        }

        [hashtable]$hypervParameters = @{}
        if( -Not [string]::IsNullOrEmpty( $WPFtextBoxHyperVHost.Text ) )
        {
            $hypervParameters.Add( 'ComputerName' , $WPFtextBoxHyperVHost.Text.Trim() )
        }
        [string]$answer = $null
        [int]$loopIterations = 0
        [hashtable]$clipboardParameters = @{}
        [bool]$refreshAzureList = $false
        [pscustomobject]$azureSessionMessage = $null
        [string]$azureRunCommandText = $null
        [hashtable]$async = @{}
        if( $asyncActions )
        {
            $async.Add( 'AsJob' , $true )
        }
        else
        {
            $async.Add( 'PassThru' , $true )
        }

        try
        { 
            if( $operation -ieq 'Azure_RunOn' )
            {
                $azureRunCommandText = Get-AzureRunCommandText
                if( [string]::IsNullOrWhiteSpace( $azureRunCommandText ) )
                {
                    return
                }

                # Set busy cursor after the modal run-command dialog closes so it is visible during job execution.
                $mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
                [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
                [void]$mainWindow.Dispatcher.Invoke(
                    [System.Windows.Threading.DispatcherPriority]::Render,
                    [System.Action]{ }
                )
            }
            else
            {
                $mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
            }

            [array]$jobs = @( ForEach( $selection in $selectedVMs )
            {
                $loopIterations++
                
                if( $operation -ieq 'DeleteComputer' ) ## not Hyper-V context
                {
                    Write-Verbose -Message "Deleting computer $selection"
                    $GUIobject.Items.RemoveAt( $selection )
                    continue
                }

                if( $Operation -match 'PowerOn|Detail|Resume|Clipboard|TakeSnapshot|MessageSession|OpenInPortal|ChangeDiskType|EditTags|ChangeHostPoolSize|HostPoolDetail|HostPoolActivityLogs|VMActivityLogs|AVDLogs|AppGroups|((Manage|Revert|Delete).*Snapshot)' ) ## don't need to prompt or will prompt with more information later
                {
                    $answer = 'yes'
                }
                elseif( [string]::IsNullOrEmpty( $answer ) -or $answer -eq 'No' )
                {
                    [string]$buttons = 'YesNo'
                    [string]$prompt = $(if( $selectedVMs.Count -gt 1 )
                    {
                        ## action may be prefixed with the hypervisor eg HyperV or Azure so remove that since on that specific tab anyway
                        "Are you sure you want to $($operation -replace '^[^_]*_' -creplace '([a-zA-Z])([A-Z])' , '$1 $2') $($selectedVMs.Count - $loopIterations + 1) VMs?`nYes for All , No for $($selection.Name) Only or Cancel for None"
                        $buttons = 'YesNoCancel'
                    }
                    else
                    {
                        "Are you sure you want to $($operation -replace '^[^_]*_' -creplace '([a-zA-Z])([A-Z])' , '$1 $2') $($selection.Name)?"
                    })
                    ## make it modal to the main window
                    $answer = [Windows.MessageBox]::Show( $mainWindow , $prompt , 'Confirm Power Operation' , $buttons ,'Question' )
                }
                if( [string]::IsNullOrEmpty( $answer ) -or $answer -ieq 'cancel' )
                {
                    $answer = $null
                    break
                }
                ## for a single VM selection, 'No' means skip; for multiple VMs, 'No' means do this one and re-prompt for the next
                if( $answer -ieq 'No' -and $selectedVMs.Count -le 1 )
                {
                    break
                }
                ## else if $answer = no then we are just performing on this VM and will prompt again next time round this loop           
                if( $operation -ieq 'NameToClipboard' )
                {
                    $selection.Name | Set-Clipboard @clipboardParameters
                }
                elseif( $operation -ieq 'Azure_PowerOn' )
                {
                    $actionStatus = $null
                    $actionStatus = Start-AZVM -Name $selection.Name -ResourceGroupName $selection.ResourceGroup -Confirm:$false -Nowait
                }
                elseif( $operation -ieq 'Azure_Shutdown' )
                {
                    $actionStatus = $null
                    $actionStatus = Stop-AZVM -Name $selection.Name -ResourceGroupName $selection.ResourceGroup -Confirm:$false -Nowait -Force
                }
                elseif( $operation -ieq 'Azure_Poweroff' )
                {
                    $actionStatus = $null
                    $actionStatus = Stop-AZVM -Name $selection.Name -ResourceGroupName $selection.ResourceGroup -Confirm:$false -Nowait -Force -SkipShutdown
                }
                elseif( $operation -ieq 'Azure_Restart' )
                {
                    $actionStatus = $null
                    $actionStatus = Restart-AzVM -Name $selection.Name -ResourceGroupName $selection.ResourceGroup -Confirm:$false -NoWait
                }
                elseif( $operation -ieq 'Azure_Hibernate' )
                {
                    $actionStatus = $null
                    $actionError = $null
                    $actionStatus = Stop-AZVM -Name $selection.Name -ResourceGroupName $selection.ResourceGroup -Confirm:$false -Nowait -Hibernate -Force -ErrorVariable actionError
                    if( $null -eq $actionStatus )
                    {
                        Write-Warning -Message "Azure hibernate failed for $($selection.Name)"
                        [void][Windows.MessageBox]::Show( $mainWindow , "Azure hibernate failed for $($selection.Name)`n$($actionError)" , 'Azure Power Operation' , 'Ok' ,'Error' )
                    }
                }
                elseif( $operation -ieq 'Azure_OpenInPortal' )
                {
                    [string]$subscriptionId = $selection.SubscriptionId
                    if( [string]::IsNullOrWhiteSpace( $subscriptionId ) )
                    {
                        Write-Warning -Message "No subscription id available for VM $($selection.Name)"
                        [void][Windows.MessageBox]::Show( $mainWindow , "Unable to determine subscription for $($selection.Name)" , 'Azure Portal' , 'Ok' ,'Error' )
                        continue
                    }

                    [string]$portalUrl = "https://portal.azure.com/#resource/subscriptions/$([uri]::EscapeDataString($subscriptionId))/resourceGroups/$([uri]::EscapeDataString($selection.ResourceGroup))/providers/Microsoft.Compute/virtualMachines/$([uri]::EscapeDataString($selection.Name))/overview"
                    Start-Process -FilePath $portalUrl -Verb Open
                }
                elseif( $operation -ieq 'Azure_ChangeDiskType' )
                {
                    try
                    {
                        Import-Module -Name Az.Compute -Verbose:$false

                        # Check the VM is deallocated
                        if( $selection.PowerState -notmatch 'deallocat' )
                        {
                            [void][Windows.MessageBox]::Show( $mainWindow , "$($selection.Name) must be deallocated before changing disk type.`nCurrent power state: $($selection.PowerState)" , 'Change Disk Type' , 'Ok' , 'Warning' )
                            continue
                        }

                        [hashtable]$diskTypeFriendlyNames = @{
                            'Standard_LRS'    = 'Standard HDD LRS'
                            'StandardSSD_LRS' = 'Standard SSD LRS'
                            'Premium_LRS'     = 'Premium SSD LRS'
                            'Premium_ZRS'     = 'Premium SSD ZRS'
                            'StandardSSD_ZRS' = 'Standard SSD ZRS'
                            'UltraSSD_LRS'    = 'Ultra SSD LRS'
                            'PremiumV2_LRS'   = 'Premium SSD v2 LRS'
                        }

                        $vmFull = $null
                        $vmFull = Get-AzVM -Name $selection.Name -ResourceGroupName $selection.ResourceGroup -ErrorAction Stop

                        [string]$osDiskId   = $vmFull.StorageProfile.OsDisk.ManagedDisk.Id
                        [string]$osDiskName = $vmFull.StorageProfile.OsDisk.Name
                        [string]$osDiskRG   = ( $osDiskId -replace '^.*resourceGroups/([^/]+)/.*$' , '$1' )

                        $osDiskResource = $null
                        $osDiskResource = Get-AzDisk -ResourceGroupName $osDiskRG -DiskName $osDiskName -ErrorAction Stop

                        [string]$currentSkuName = $osDiskResource.Sku.Name
                        [string]$currentFriendly = if( $diskTypeFriendlyNames.ContainsKey( $currentSkuName ) ) { $diskTypeFriendlyNames[ $currentSkuName ] } else { $currentSkuName }

                        # Build list of available types excluding the current one
                        [array]$availableTypes = @( $diskTypeFriendlyNames.GetEnumerator() | Where-Object { $_.Key -ne $currentSkuName } | Sort-Object Value | Select-Object -ExpandProperty Value )

                        if( $changeDiskWindow = New-WPFWindow -inputXAML $comboSelectXAML )
                        {
                            $WPFbtnComboOK.Add_Click({
                                $_.Handled = $true
                                $changeDiskWindow.DialogResult = $true
                                $changeDiskWindow.Close()
                            })
                            $changeDiskWindow.Title   = "Change Disk Type - $($selection.Name)"
                            $WPFlblComboHeader.Content        = "VM: $($selection.Name)"
                            $WPFlblComboCurrentValue.Content  = "Current OS disk type: $currentFriendly"
                            $WPFlblComboLabel.Content         = "New disk type:"
                            ForEach( $t in $availableTypes ) { [void]$WPFcomboBoxSelect.Items.Add( $t ) }
                            $WPFcomboBoxSelect.SelectedIndex  = 0

                            if( $changeDiskWindow.ShowDialog() )
                            {
                                [string]$selectedFriendly = $WPFcomboBoxSelect.SelectedItem
                                [string]$newSkuName = ( $diskTypeFriendlyNames.GetEnumerator() | Where-Object { $_.Value -eq $selectedFriendly } | Select-Object -First 1 -ExpandProperty Key )

                                if( -Not [string]::IsNullOrWhiteSpace( $newSkuName ) )
                                {
                                    # Update OS disk
                                    $diskUpdate = New-AzDiskUpdateConfig -SkuName $newSkuName
                                    $updateResult = $null
                                    $updateResult = Update-AzDisk -ResourceGroupName $osDiskRG -DiskName $osDiskName -DiskUpdate $diskUpdate -ErrorAction Stop

                                    # Update any data disks
                                    [array]$dataDisks = @( $vmFull.StorageProfile.DataDisks | Where-Object { $null -ne $_.ManagedDisk } )
                                    ForEach( $dataDisk in $dataDisks )
                                    {
                                        [string]$ddId   = $dataDisk.ManagedDisk.Id
                                        [string]$ddName = $dataDisk.Name
                                        [string]$ddRG   = ( $ddId -replace '^.*resourceGroups/([^/]+)/.*$' , '$1' )
                                        try
                                        {
                                            $ddUpdate = New-AzDiskUpdateConfig -SkuName $newSkuName
                                            $null = Update-AzDisk -ResourceGroupName $ddRG -DiskName $ddName -DiskUpdate $ddUpdate -ErrorAction Stop
                                        }
                                        catch
                                        {
                                            Write-Warning -Message "Failed to update data disk $ddName for $($selection.Name): $($_.Exception.Message)"
                                        }
                                    }

                                    Write-Host "Changed disk type for $($selection.Name) from $currentFriendly to $selectedFriendly" -ForegroundColor Green
                                }
                            }
                        }
                    }
                    catch
                    {
                        Write-Warning -Message "Change disk type error for $($selection.Name): $($_.Exception.Message)"
                        [void][Windows.MessageBox]::Show( $mainWindow , "Change disk type failed for $($selection.Name)`n$($_.Exception.Message)" , 'Change Disk Type' , 'Ok' , 'Error' )
                    }
                }
                elseif( $operation -ieq 'Azure_ChangeHostPoolSize' )
                {
                    try
                    {
                        Import-Module -Name Az.DesktopVirtualization -Verbose:$false

                        # Validate: all selected VMs must be in the same host pool
                        [array]$distinctHostPools = @( $selectedVMs | Where-Object { -not [string]::IsNullOrWhiteSpace( $_.HostPool ) } | Select-Object -ExpandProperty HostPool -Unique )

                        if( $distinctHostPools.Count -eq 0 )
                        {
                            [void][Windows.MessageBox]::Show( $mainWindow , 'None of the selected VMs have an associated AVD host pool.' , 'Change Host Pool Size' , 'Ok' , 'Warning' )
                        }
                        elseif( $distinctHostPools.Count -gt 1 )
                        {
                            [void][Windows.MessageBox]::Show( $mainWindow , "Selected VMs span $($distinctHostPools.Count) host pools: $($distinctHostPools -join ', ').`nPlease select VMs from a single host pool only." , 'Change Host Pool Size' , 'Ok' , 'Warning' )
                        }
                        else
                        {
                            [string]$hostPoolName = $distinctHostPools[0]

                            # Look up host pool resource ID directly (not relying on the cached lookup table)
                            $hostPoolObject = $null
                            $hostPoolObject = Get-AzWvdHostPool -ErrorAction SilentlyContinue | Where-Object { $_.Name -ieq $hostPoolName } | Select-Object -First 1
                            [string]$hostPoolId = if( $null -ne $hostPoolObject ) { $hostPoolObject.Id } else { $null }

                            if( [string]::IsNullOrWhiteSpace( $hostPoolId ) )
                            {
                                [void][Windows.MessageBox]::Show( $mainWindow , "Could not resolve resource ID for host pool '$hostPoolName'." , 'Change Host Pool Size' , 'Ok' , 'Error' )
                            }
                            else
                            {
                                [string]$shmPath = "$hostPoolId/sessionHostManagements/default?api-version=$sessionHostMgmtAPIVersion"
                                $shmResponse = $null
                                $shmResponse  = Invoke-AzRestMethod -Method GET -Path $shmPath -ErrorAction Stop

                                if( $shmResponse.StatusCode -notin @( 200 , 201 ) )
                                {
                                    [string]$errMsg = $shmResponse.Content
                                    try { $errMsg = ( $shmResponse.Content | ConvertFrom-Json ).error.message } catch {}
                                    [void][Windows.MessageBox]::Show( $mainWindow , "Host pool '$hostPoolName' does not use Session Host Configuration or the API is unavailable.`nStatus: $($shmResponse.StatusCode)`n$errMsg" , 'Change Host Pool Size' , 'Ok' , 'Warning' )
                                }
                                else
                                {
                                    $shmJson = $shmResponse.Content | ConvertFrom-Json

                                    if( $null -eq $shmJson.properties -or $null -eq $shmJson.properties.provisioning )
                                    {
                                        [void][Windows.MessageBox]::Show( $mainWindow , "Host pool '$hostPoolName' does not have a session host provisioning configuration.`nSession Host Configuration must be enabled on the host pool to use this feature." , 'Change Host Pool Size' , 'Ok' , 'Warning' )
                                    }
                                    else
                                    {
                                        [int]$currentCount = [int]$shmJson.properties.provisioning.instanceCount

                                        if( $hostPoolSizeWindow = New-WPFWindow -inputXAML $hostPoolSizeXAML )
                                        {
                                            $hostPoolSizeWindow.Owner = $mainWindow
                                            $WPFlblHostPoolSizeHeader.Content   = "Host Pool: $hostPoolName"
                                            $WPFlblHostPoolCurrentSize.Content  = "Current instance count: $currentCount"
                                            $WPFtxtHostPoolNewSize.Text         = ''
                                            $WPFlblHostPoolSizeError.Content    = ''
                                            $WPFbtnHostPoolSizeOK.IsEnabled     = $false

                                            $WPFtxtHostPoolNewSize.Add_TextChanged({
                                                [string]$txt = $WPFtxtHostPoolNewSize.Text.Trim()
                                                $WPFbtnHostPoolSizeOK.IsEnabled     = $false
                                                $WPFlblHostPoolSizeError.Content    = ''
                                                if( $txt -match '^\d+$' )
                                                {
                                                    [int]$v = [int]$txt
                                                    if( $v -le 0 )
                                                    {
                                                        $WPFlblHostPoolSizeError.Content = 'Value must be a positive integer.'
                                                    }
                                                    elseif( $v -le $currentCount )
                                                    {
                                                        $WPFlblHostPoolSizeError.Content = "New value must be greater than the current count ($currentCount)."
                                                    }
                                                    else
                                                    {
                                                        $WPFbtnHostPoolSizeOK.IsEnabled = $true
                                                    }
                                                }
                                                elseif( $txt.Length -gt 0 )
                                                {
                                                    $WPFlblHostPoolSizeError.Content = 'Please enter a positive whole number.'
                                                }
                                            })

                                            $WPFbtnHostPoolSizeOK.Add_Click({
                                                $_.Handled = $true
                                                $hostPoolSizeWindow.DialogResult = $true
                                                $hostPoolSizeWindow.Close()
                                            })

                                            if( $hostPoolSizeWindow.ShowDialog() )
                                            {
                                                [int]$newCount      = [int]$WPFtxtHostPoolNewSize.Text.Trim()
                                                [string]$confirmMsg = "Grow host pool '$hostPoolName' from $currentCount to $newCount instance(s)?`n`nThis will provision $($newCount - $currentCount) additional session host(s)."

                                                if( [Windows.MessageBox]::Show( $mainWindow , $confirmMsg , 'Confirm Host Pool Resize' , 'YesNo' , 'Question' ) -ieq 'Yes' )
                                                {
                                                    $shmJson.properties.provisioning.instanceCount = $newCount
                                                    [string]$putBody   = $shmJson | ConvertTo-Json -Depth 10 -Compress
                                                    $putResponse       = $null
                                                    $putResponse       = Invoke-AzRestMethod -Method PUT -Path $shmPath -Payload $putBody -ErrorAction Stop

                                                    if( $putResponse.StatusCode -in @( 200 , 201 , 202 ) )
                                                    {
                                                        Write-Host "Host pool '$hostPoolName' resize initiated: $currentCount -> $newCount" -ForegroundColor Green
                                                        [void][Windows.MessageBox]::Show( $mainWindow , "Host pool '$hostPoolName' resize from $currentCount to $newCount instance(s) initiated successfully." , 'Change Host Pool Size' , 'Ok' , 'Information' )
                                                    }
                                                    else
                                                    {
                                                        [string]$errDetail = $putResponse.Content
                                                        try { $errDetail = ( $putResponse.Content | ConvertFrom-Json ).error.message } catch {}
                                                        Write-Warning -Message "Host pool '$hostPoolName' resize failed: $errDetail"
                                                        [void][Windows.MessageBox]::Show( $mainWindow , "Host pool resize failed for '$hostPoolName'.`nStatus: $($putResponse.StatusCode)`n$errDetail" , 'Change Host Pool Size' , 'Ok' , 'Error' )
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch
                    {
                        Write-Warning -Message "Change host pool size error: $($_.Exception.Message)"
                        [void][Windows.MessageBox]::Show( $mainWindow , "Change host pool size failed.`n$($_.Exception.Message)" , 'Change Host Pool Size' , 'Ok' , 'Error' )
                    }
                    break  # operation applies to the whole selection, not per VM
                }
                elseif( $operation -ieq 'Azure_HostPoolDetail' )
                {
                    try
                    {
                        Import-Module -Name Az.DesktopVirtualization -Verbose:$false

                        [string]$hostPoolName = $selection.HostPool
                        if( [string]::IsNullOrWhiteSpace( $hostPoolName ) )
                        {
                            [void][Windows.MessageBox]::Show( $mainWindow , "VM '$($selection.Name)' is not associated with an AVD host pool." , 'Host Pool Detail' , 'Ok' , 'Warning' )
                        }
                        else
                        {
                            $hostPoolObject = $null
                            $hostPoolObject = Get-AzWvdHostPool -ErrorAction SilentlyContinue | Where-Object { $_.Name -ieq $hostPoolName } | Select-Object -First 1

                            if( $null -eq $hostPoolObject )
                            {
                                [void][Windows.MessageBox]::Show( $mainWindow , "Could not find host pool '$hostPoolName'." , 'Host Pool Detail' , 'Ok' , 'Error' )
                            }
                            else
                            {
                                [string]$hostPoolRg = ($hostPoolObject.Id -split '/')[4]

                                [array]$sessionHosts        = @( Get-AzWvdSessionHost -HostPoolName $hostPoolName -ResourceGroupName $hostPoolRg -ErrorAction SilentlyContinue )
                                [int]$totalSessionHosts     = $sessionHosts.Count
                                [int]$availableSessionHosts = @( $sessionHosts | Where-Object { $_.Status -ieq 'Available' -and $_.AllowNewSession -eq $true } ).Count
                                [int]$totalSessions         = [int]( $sessionHosts | Measure-Object -Property Session -Sum ).Sum
                                [int]$maxSessions           = $hostPoolObject.MaxSessionLimit * $totalSessionHosts
                                [int]$availableSessions     = [Math]::Max( 0 , $maxSessions - $totalSessions )

                                [array]$appGroups          = @( Get-AzWvdApplicationGroup -ErrorAction SilentlyContinue | Where-Object { $_.HostPoolArmPath -ieq $hostPoolObject.Id } )
                                [int]$appGroupTotal        = $appGroups.Count
                                [int]$appGroupRemoteApp    = @( $appGroups | Where-Object { $_.ApplicationGroupType -ieq 'RemoteApp' } ).Count

                                [hashtable]$appGroupIdSet  = @{}
                                foreach( $ag in $appGroups ) { if( $ag.Id ) { $appGroupIdSet[ $ag.Id.ToLowerInvariant() ] = $true } }
                                [array]$linkedWorkspaces   = @(
                                    Get-AzWvdWorkspace -ErrorAction SilentlyContinue | Where-Object {
                                        $_.ApplicationGroupReference | Where-Object { $appGroupIdSet.ContainsKey( $_.ToLowerInvariant() ) }
                                    } | ForEach-Object { if( -not [string]::IsNullOrWhiteSpace( $_.FriendlyName ) ) { $_.FriendlyName } else { $_.Name } }
                                )
                                [string]$workspaceList     = if( $linkedWorkspaces.Count -gt 0 ) { $linkedWorkspaces -join '; ' } else { '(none)' }

                                [array]$rawProps = @(
                                    @{ Property = 'Name'                         ; Value = $hostPoolObject.Name }
                                    @{ Property = 'Location'                     ; Value = $hostPoolObject.Location }
                                    @{ Property = 'Resource Group'               ; Value = $hostPoolRg }
                                    @{ Property = 'Friendly Name'                ; Value = $hostPoolObject.FriendlyName }
                                    @{ Property = 'Description'                  ; Value = $hostPoolObject.Description }
                                    @{ Property = 'Host Pool Type'               ; Value = $hostPoolObject.HostPoolType }
                                    @{ Property = 'Load Balancer Type'           ; Value = $hostPoolObject.LoadBalancerType }
                                    @{ Property = 'Max Session Limit'            ; Value = $hostPoolObject.MaxSessionLimit }
                                    @{ Property = 'Session Hosts (Total)'        ; Value = $totalSessionHosts }
                                    @{ Property = 'Session Hosts (Available)'    ; Value = $availableSessionHosts }
                                    @{ Property = 'Current Sessions'             ; Value = $totalSessions }
                                    @{ Property = 'Available Sessions'           ; Value = $availableSessions }
                                    @{ Property = 'Application Groups (Total)'   ; Value = $appGroupTotal }
                                    @{ Property = 'Application Groups (RemoteApp)'; Value = $appGroupRemoteApp }
                                    @{ Property = 'Linked Workspaces'            ; Value = $workspaceList }
                                    @{ Property = 'Personal Desktop Assignment'  ; Value = $hostPoolObject.PersonalDesktopAssignmentType }
                                    @{ Property = 'Preferred App Group Type'     ; Value = $hostPoolObject.PreferredAppGroupType }
                                    @{ Property = 'Validation Environment'       ; Value = $hostPoolObject.ValidationEnvironment }
                                    @{ Property = 'Start VM on Connect'          ; Value = $hostPoolObject.StartVMOnConnect }
                                    @{ Property = 'Registration Expiry'          ; Value = if( $hostPoolObject.RegistrationInfoExpirationTime ) { $hostPoolObject.RegistrationInfoExpirationTime.ToString('g') } else { '' } }
                                    @{ Property = 'Custom RDP Property'          ; Value = $hostPoolObject.CustomRdpProperty }
                                    @{ Property = 'SSO Client ID'                ; Value = $hostPoolObject.SsoClientId }
                                    @{ Property = 'Tags'                         ; Value = if( $hostPoolObject.Tag ) { ($hostPoolObject.Tag.Keys | Sort-Object | ForEach-Object { "$_=$($hostPoolObject.Tag[$_])" }) -join '; ' } else { '' } }
                                    @{ Property = 'Resource ID'                  ; Value = $hostPoolObject.Id }
                                )

                                if( $hostPoolDetailWindow = New-WPFWindow -inputXAML $hostPoolDetailXAML )
                                {
                                    $hostPoolDetailWindow.Owner = $mainWindow
                                    $WPFlblHostPoolDetailHeader.Content = "Host Pool: $hostPoolName"

                                    $dt = [System.Data.DataTable]::new()
                                    [void]$dt.Columns.Add( 'Property' )
                                    [void]$dt.Columns.Add( 'Value' )
                                    foreach( $p in ( $rawProps | Where-Object { -not [string]::IsNullOrWhiteSpace( $_.Value ) } ) )
                                    {
                                        $dr = $dt.NewRow()
                                        $dr['Property'] = $p.Property
                                        $dr['Value']    = [string]$p.Value
                                        [void]$dt.Rows.Add( $dr )
                                    }
                                    $WPFdgHostPoolDetail.ItemsSource = $dt.DefaultView
                                    [void]$hostPoolDetailWindow.ShowDialog()
                                }
                            }
                        }
                    }
                    catch
                    {
                        Write-Warning -Message "Host Pool Detail error: $($_.Exception.Message)"
                        [void][Windows.MessageBox]::Show( $mainWindow , "Host Pool Detail failed.`n$($_.Exception.Message)" , 'Host Pool Detail' , 'Ok' , 'Error' )
                    }
                    break  # single selection operation
                }
                elseif( $operation -ieq 'Azure_HostPoolActivityLogs' )
                {
                    try
                    {
                        Import-Module -Name Az.DesktopVirtualization -Verbose:$false
                        Import-Module -Name Az.Monitor               -Verbose:$false

                        [string]$hostPoolName = $selection.HostPool
                        if( [string]::IsNullOrWhiteSpace( $hostPoolName ) )
                        {
                            [void][Windows.MessageBox]::Show( $mainWindow , "VM '$($selection.Name)' is not associated with an AVD host pool." , 'Host Pool Activity Logs' , 'Ok' , 'Warning' )
                        }
                        else
                        {
                            $hostPoolObject = $null
                            $hostPoolObject = Get-AzWvdHostPool -ErrorAction SilentlyContinue | Where-Object { $_.Name -ieq $hostPoolName } | Select-Object -First 1

                            if( $null -eq $hostPoolObject )
                            {
                                [void][Windows.MessageBox]::Show( $mainWindow , "Could not find host pool '$hostPoolName'." , 'Host Pool Activity Logs' , 'Ok' , 'Error' )
                            }
                            else
                            {
                                [string]$capturedHostPoolId   = $hostPoolObject.Id
                                [string]$capturedHostPoolName = $hostPoolName

                                if( $hostPoolActLogsWindow = New-WPFWindow -inputXAML $hostPoolActivityLogsXAML )
                                {
                                    $hostPoolActLogsWindow.Owner = $mainWindow
                                    $WPFlblHostPoolActLogsHeader.Content = "Activity Logs: $capturedHostPoolName"

                                    $WPFbtnHostPoolActLogsRetrieve.Add_Click({
                                        [string]$valStr = $WPFtxtHostPoolActLogsValue.Text.Trim()
                                        if( $valStr -notmatch '^\d+(\.\d+)?$' -or [double]$valStr -le 0 )
                                        {
                                            $WPFlblHostPoolActLogsStatus.Foreground = [System.Windows.Media.Brushes]::Red
                                            $WPFlblHostPoolActLogsStatus.Content    = 'Please enter a positive number (e.g. 2 or 1.5).'
                                            return
                                        }

                                        [double]$timeVal   = [double]$valStr
                                        [TimeSpan]$span    = if( $WPFradHostPoolActLogsHours.IsChecked ) { [TimeSpan]::FromHours( $timeVal ) } else { [TimeSpan]::FromDays( $timeVal ) }
                                        [datetime]$startDT = (Get-Date) - $span

                                        $WPFlblHostPoolActLogsStatus.Foreground = [System.Windows.Media.Brushes]::Gray
                                        $WPFlblHostPoolActLogsStatus.Content    = 'Retrieving...'
                                        $WPFdgHostPoolActLogs.ItemsSource       = $null
                                        $hostPoolActLogsWindow.Dispatcher.Invoke( [System.Windows.Threading.DispatcherPriority]::Background , [action]{} )

                                        try
                                        {
                                            [array]$logEntries = @( Get-AzActivityLog -ResourceId $capturedHostPoolId -StartTime $startDT -EndTime (Get-Date) -WarningAction SilentlyContinue -ErrorAction Stop | Sort-Object -Property EventTimestamp )

                                            if( $logEntries.Count -eq 0 )
                                            {
                                                $WPFlblHostPoolActLogsStatus.Content = 'No activity log entries found for the specified time range.'
                                            }
                                            else
                                            {
                                                $dt = [System.Data.DataTable]::new()
                                                [void]$dt.Columns.Add( 'Time' )
                                                [void]$dt.Columns.Add( 'Caller' )
                                                [void]$dt.Columns.Add( 'Operation' )
                                                [void]$dt.Columns.Add( 'Status' )
                                                [void]$dt.Columns.Add( 'Level' )
                                                [void]$dt.Columns.Add( 'ResourceType' )
                                                [void]$dt.Columns.Add( 'Description' )
                                                # Resolve a field that may be a plain string, a LocalizableString, or null
                                                $resolveActVal = {
                                                    param( $v )
                                                    if( $null -eq $v )    { return '' }
                                                    if( $v -is [string] ) { return $v }
                                                    $lp = $v.PSObject.Properties.Item( 'LocalizedValue' )
                                                    if( $null -ne $lp -and -not [string]::IsNullOrEmpty( $lp.Value ) ) { return [string]$lp.Value }
                                                    $vp = $v.PSObject.Properties.Item( 'Value' )
                                                    if( $null -ne $vp -and -not [string]::IsNullOrEmpty( [string]$vp.Value ) ) { return [string]$vp.Value }
                                                    return ''
                                                }
                                                foreach( $entry in $logEntries )
                                                {
                                                    $dr = $dt.NewRow()
                                                    $dr['Time']         = $entry.EventTimestamp.ToLocalTime().ToString('g')
                                                    $dr['Caller']       = [string]$entry.Caller
                                                    $dr['Operation']    = & $resolveActVal $entry.OperationName
                                                    $dr['Status']       = & $resolveActVal $entry.Status
                                                    $dr['Level']        = [string]$entry.Level
                                                    $dr['ResourceType'] = & $resolveActVal $entry.ResourceType
                                                    $dr['Description']  = [string]$entry.Description
                                                    [void]$dt.Rows.Add( $dr )
                                                }
                                                $WPFdgHostPoolActLogs.ItemsSource    = $dt.DefaultView
                                                $WPFlblHostPoolActLogsStatus.Content = "$($logEntries.Count) entries  (oldest first)"
                                            }
                                        }
                                        catch
                                        {
                                            $WPFlblHostPoolActLogsStatus.Foreground = [System.Windows.Media.Brushes]::Red
                                            $WPFlblHostPoolActLogsStatus.Content    = "Retrieval failed: $($_.Exception.Message)"
                                        }
                                    })

                                    [void]$hostPoolActLogsWindow.ShowDialog()
                                }
                            }
                        }
                    }
                    catch
                    {
                        Write-Warning -Message "Host Pool Activity Logs error: $($_.Exception.Message)"
                        [void][Windows.MessageBox]::Show( $mainWindow , "Host Pool Activity Logs failed.`n$($_.Exception.Message)" , 'Host Pool Activity Logs' , 'Ok' , 'Error' )
                    }
                    break  # single selection operation
                }
                elseif( $operation -ieq 'Azure_AVDLogs' )
                {
                    try
                    {
                        Import-Module -Name Az.DesktopVirtualization -Verbose:$false
                        Import-Module -Name Az.Monitor               -Verbose:$false
                        Import-Module -Name Az.OperationalInsights   -Verbose:$false

                        [string]$hostPoolName = $selection.HostPool
                        if( [string]::IsNullOrWhiteSpace( $hostPoolName ) )
                        {
                            [void][Windows.MessageBox]::Show( $mainWindow , "VM '$($selection.Name)' is not associated with an AVD host pool." , 'AVD Logs' , 'Ok' , 'Warning' )
                        }
                        else
                        {
                            # Resolve host pool ARM resource ID
                            $hostPoolObject = $null
                            $hostPoolObject = Get-AzWvdHostPool -ErrorAction SilentlyContinue | Where-Object { $_.Name -ieq $hostPoolName } | Select-Object -First 1

                            if( $null -eq $hostPoolObject )
                            {
                                [void][Windows.MessageBox]::Show( $mainWindow , "Could not find host pool '$hostPoolName'." , 'AVD Logs' , 'Ok' , 'Error' )
                            }
                            else
                            {
                                # Find Log Analytics workspace via host pool diagnostic settings
                                [string]$workspaceResourceId = $null
                                $diagSettings = $null
                                $diagSettings = Get-AzDiagnosticSetting -ResourceId $hostPoolObject.Id -ErrorAction SilentlyContinue
                                foreach( $ds in $diagSettings )
                                {
                                    if( -not [string]::IsNullOrWhiteSpace( $ds.WorkspaceId ) )
                                    {
                                        $workspaceResourceId = $ds.WorkspaceId
                                        break
                                    }
                                }

                                if( [string]::IsNullOrWhiteSpace( $workspaceResourceId ) )
                                {
                                    [void][Windows.MessageBox]::Show( $mainWindow , "No Log Analytics workspace is configured in the diagnostic settings for host pool '$hostPoolName'.`nEnable diagnostics on the host pool and direct logs to a Log Analytics workspace first." , 'AVD Logs' , 'Ok' , 'Warning' )
                                }
                                else
                                {
                                    [string]$wsName = ($workspaceResourceId -split '/')[-1]
                                    [string]$wsRg   = ($workspaceResourceId -split '/')[4]
                                    $wsObject = $null
                                    $wsObject = Get-AzOperationalInsightsWorkspace -ResourceGroupName $wsRg -Name $wsName -ErrorAction Stop
                                    [string]$capturedWsId   = $wsObject.CustomerId.ToString()
                                    [string]$capturedVMName = $selection.Name

                                    if( $avdLogsWindow = New-WPFWindow -inputXAML $avdLogsXAML )
                                    {
                                        $avdLogsWindow.Owner = $mainWindow
                                        $WPFlblAVDLogsHeader.Content = "AVD Logs: $capturedVMName   (Host Pool: $hostPoolName)"

                                        $WPFbtnAVDLogsRetrieve.Add_Click({
                                            [string]$valStr = $WPFtxtAVDLogsValue.Text.Trim()
                                            if( $valStr -notmatch '^\d+$' -or [int]$valStr -le 0 )
                                            {
                                                $WPFlblAVDLogsStatus.Foreground = [System.Windows.Media.Brushes]::Red
                                                $WPFlblAVDLogsStatus.Content    = 'Please enter a positive whole number.'
                                                return
                                            }

                                            [int]$timeVal    = [int]$valStr
                                            [string]$unit    = if( $WPFradAVDLogsHours.IsChecked ) { 'h' } else { 'd' }
                                            [string]$agoExpr = "ago(${timeVal}${unit})"

                                            $WPFlblAVDLogsStatus.Foreground = [System.Windows.Media.Brushes]::Gray
                                            $WPFlblAVDLogsStatus.Content    = 'Retrieving...'
                                            $WPFdgAVDLogs.ItemsSource       = $null
                                            $avdLogsWindow.Dispatcher.Invoke( [System.Windows.Threading.DispatcherPriority]::Background , [action]{} )

                                            [string]$kql = @"
union isfuzzy=true WVDConnections, WVDErrors, WVDCheckpoints, WVDManagement, WVDHostRegistrations, WVDAgentHealthStatus
| where TimeGenerated >= $agoExpr
| where SessionHostName has `"$capturedVMName`" or SessionHost has `"$capturedVMName`" or _ResourceId has `"$capturedVMName`"
| sort by TimeGenerated asc
| project-away TenantId, SourceSystem, _ResourceId, _SubscriptionId, ResourceId, _IsBillable, Computer, MG
"@

                                            try
                                            {
                                                $qResult = Invoke-AzOperationalInsightsQuery -WorkspaceId $capturedWsId -Query $kql -ErrorAction Stop
                                                [array]$rows = @( $qResult.Results )
                                                if( $rows.Count -eq 0 )
                                                {
                                                    $WPFlblAVDLogsStatus.Content = 'No log entries found for the specified time range.'
                                                }
                                                else
                                                {
                                                    # Convert PSObjects to DataTable for reliable DataGrid binding
                                                    $dt = [System.Data.DataTable]::new()
                                                    foreach( $prop in $rows[0].PSObject.Properties )
                                                    {
                                                        [void]$dt.Columns.Add( $prop.Name )
                                                    }
                                                    foreach( $row in $rows )
                                                    {
                                                        $dr = $dt.NewRow()
                                                        foreach( $prop in $row.PSObject.Properties )
                                                        {
                                                            $dr[ $prop.Name ] = if( $null -ne $prop.Value ) { $prop.Value } else { [DBNull]::Value }
                                                        }
                                                        [void]$dt.Rows.Add( $dr )
                                                    }
                                                    $WPFdgAVDLogs.ItemsSource    = $dt.DefaultView
                                                    $WPFlblAVDLogsStatus.Content = "$($rows.Count) log entries  (oldest first)"
                                                }
                                            }
                                            catch
                                            {
                                                $WPFlblAVDLogsStatus.Foreground = [System.Windows.Media.Brushes]::Red
                                                $WPFlblAVDLogsStatus.Content    = "Query failed: $($_.Exception.Message)"
                                            }
                                        })

                                        [void]$avdLogsWindow.ShowDialog()
                                    }
                                }
                            }
                        }
                    }
                    catch
                    {
                        Write-Warning -Message "AVD Logs error: $($_.Exception.Message)"
                        [void][Windows.MessageBox]::Show( $mainWindow , "AVD Logs failed.`n$($_.Exception.Message)" , 'AVD Logs' , 'Ok' , 'Error' )
                    }
                    break  # single VM operation
                }
                elseif( $operation -ieq 'Azure_AppGroups' )
                {
                    try
                    {
                        Import-Module -Name Az.DesktopVirtualization -Verbose:$false
                        Import-Module -Name Az.Resources            -Verbose:$false

                        [string]$hostPoolName = $selection.HostPool
                        if( [string]::IsNullOrWhiteSpace( $hostPoolName ) )
                        {
                            [void][Windows.MessageBox]::Show( $mainWindow , "VM '$($selection.Name)' is not associated with an AVD host pool." , 'Application Groups' , 'Ok' , 'Warning' )
                        }
                        else
                        {
                            $hostPoolObject = Get-AzWvdHostPool -ErrorAction SilentlyContinue | Where-Object { $_.Name -ieq $hostPoolName } | Select-Object -First 1

                            if( $null -eq $hostPoolObject )
                            {
                                [void][Windows.MessageBox]::Show( $mainWindow , "Could not find host pool '$hostPoolName'." , 'Application Groups' , 'Ok' , 'Error' )
                            }
                            else
                            {
                                [array]$appGroups = @( Get-AzWvdApplicationGroup -ErrorAction SilentlyContinue | Where-Object { $_.HostPoolArmPath -ieq $hostPoolObject.Id } )

                                # Build workspace lookup
                                [hashtable]$wsLookup = @{}
                                Get-AzWvdWorkspace -ErrorAction SilentlyContinue | ForEach-Object {
                                    [string]$wsName = if( -not [string]::IsNullOrWhiteSpace( $_.FriendlyName ) ) { $_.FriendlyName } else { $_.Name }
                                    foreach( $ref in @( $_.ApplicationGroupReference ) )
                                    {
                                        if( -not [string]::IsNullOrWhiteSpace( $ref ) )
                                        {
                                            $wsLookup[ $ref.ToLowerInvariant() ] = $wsName
                                        }
                                    }
                                }

                                $dtApps = [System.Data.DataTable]::new()
                                [void]$dtApps.Columns.Add( 'AppGroup' )
                                [void]$dtApps.Columns.Add( 'GroupType' )
                                [void]$dtApps.Columns.Add( 'Workspace' )
                                [void]$dtApps.Columns.Add( 'AppName' )
                                [void]$dtApps.Columns.Add( 'DisplayName' )
                                [void]$dtApps.Columns.Add( 'Description' )
                                [void]$dtApps.Columns.Add( 'AppType' )
                                [void]$dtApps.Columns.Add( 'FilePath' )

                                $dtAssign = [System.Data.DataTable]::new()
                                [void]$dtAssign.Columns.Add( 'AppGroup' )
                                [void]$dtAssign.Columns.Add( 'GroupType' )
                                [void]$dtAssign.Columns.Add( 'Workspace' )
                                [void]$dtAssign.Columns.Add( 'Principal' )
                                [void]$dtAssign.Columns.Add( 'PrincipalType' )
                                [void]$dtAssign.Columns.Add( 'Role' )

                                foreach( $ag in $appGroups )
                                {
                                    [string]$agRg        = ($ag.Id -split '/')[4]
                                    [string]$agType      = $ag.ApplicationGroupType
                                    [string]$agWorkspace = $wsLookup[ $ag.Id.ToLowerInvariant() ]
                                    if( [string]::IsNullOrWhiteSpace( $agWorkspace ) ) { $agWorkspace = '' }

                                    # Applications
                                    if( $agType -ieq 'RemoteApp' )
                                    {
                                        [array]$apps = @( Get-AzWvdApplication -ResourceGroupName $agRg -ApplicationGroupName $ag.Name -ErrorAction SilentlyContinue )
                                        if( $apps.Count -eq 0 )
                                        {
                                            $dr = $dtApps.NewRow()
                                            $dr['AppGroup']    = $ag.Name
                                            $dr['GroupType']   = $agType
                                            $dr['Workspace']   = $agWorkspace
                                            $dr['AppName']     = '(no applications)'
                                            $dr['DisplayName'] = ''
                                            $dr['Description'] = ''
                                            $dr['AppType']     = ''
                                            $dr['FilePath']    = ''
                                            [void]$dtApps.Rows.Add( $dr )
                                        }
                                        else
                                        {
                                            foreach( $app in $apps )
                                            {
                                                $dr = $dtApps.NewRow()
                                                $dr['AppGroup']    = $ag.Name
                                                $dr['GroupType']   = $agType
                                                $dr['Workspace']   = $agWorkspace
                                                $dr['AppName']     = $app.Name
                                                $dr['DisplayName'] = [string]$app.FriendlyName
                                                $dr['Description'] = [string]$app.Description
                                                $dr['AppType']     = [string]$app.AppAliaPath
                                                $dr['FilePath']    = if( -not [string]::IsNullOrWhiteSpace( $app.FilePath ) ) { $app.FilePath } elseif( -not [string]::IsNullOrWhiteSpace( $app.MsixPackageApplicationId ) ) { $app.MsixPackageApplicationId } else { '' }
                                                [void]$dtApps.Rows.Add( $dr )
                                            }
                                        }
                                    }
                                    else
                                    {
                                        # Desktop app group - show single entry
                                        $dr = $dtApps.NewRow()
                                        $dr['AppGroup']    = $ag.Name
                                        $dr['GroupType']   = $agType
                                        $dr['Workspace']   = $agWorkspace
                                        $dr['AppName']     = '(Full Desktop)'
                                        $dr['DisplayName'] = [string]$ag.FriendlyName
                                        $dr['Description'] = [string]$ag.Description
                                        $dr['AppType']     = ''
                                        $dr['FilePath']    = ''
                                        [void]$dtApps.Rows.Add( $dr )
                                    }

                                    # Assignments (role assignments on the app group resource)
                                    [array]$roleAssignments = @( Get-AzRoleAssignment -Scope $ag.Id -ErrorAction SilentlyContinue | Where-Object { $_.Scope -ieq $ag.Id } )
                                    if( $roleAssignments.Count -eq 0 )
                                    {
                                        $dr = $dtAssign.NewRow()
                                        $dr['AppGroup']      = $ag.Name
                                        $dr['GroupType']     = $agType
                                        $dr['Workspace']     = $agWorkspace
                                        $dr['Principal']     = '(no assignments)'
                                        $dr['PrincipalType'] = ''
                                        $dr['Role']          = ''
                                        [void]$dtAssign.Rows.Add( $dr )
                                    }
                                    else
                                    {
                                        foreach( $ra in $roleAssignments )
                                        {
                                            $dr = $dtAssign.NewRow()
                                            $dr['AppGroup']      = $ag.Name
                                            $dr['GroupType']     = $agType
                                            $dr['Workspace']     = $agWorkspace
                                            $dr['Principal']     = if( -not [string]::IsNullOrWhiteSpace( $ra.DisplayName ) ) { $ra.DisplayName } else { $ra.SignInName }
                                            $dr['PrincipalType'] = [string]$ra.ObjectType
                                            $dr['Role']          = [string]$ra.RoleDefinitionName
                                            [void]$dtAssign.Rows.Add( $dr )
                                        }
                                    }
                                }

                                if( $appGroupsWindow = New-WPFWindow -inputXAML $appGroupsXAML )
                                {
                                    $appGroupsWindow.Owner = $mainWindow
                                    $WPFlblAppGroupsHeader.Content  = "Application Groups: $hostPoolName  ($($appGroups.Count) group(s))"
                                    $WPFlblAppGroupsStatus.Content  = "$($dtApps.Rows.Count) application(s) across $($appGroups.Count) group(s)  |  $($dtAssign.Rows.Count) assignment(s)"
                                    $WPFdgAppGroupsApplications.ItemsSource = $dtApps.DefaultView
                                    $WPFdgAppGroupsAssignments.ItemsSource  = $dtAssign.DefaultView
                                    [void]$appGroupsWindow.ShowDialog()
                                }
                            }
                        }
                    }
                    catch
                    {
                        Write-Warning -Message "Application Groups error: $($_.Exception.Message)"
                        [void][Windows.MessageBox]::Show( $mainWindow , "Application Groups failed.`n$($_.Exception.Message)" , 'Application Groups' , 'Ok' , 'Error' )
                    }
                    break  # single selection operation
                }
                elseif( $operation -ieq 'Azure_VMActivityLogs' )
                {
                    try
                    {
                        Import-Module -Name Az.Monitor -Verbose:$false

                        if( [string]::IsNullOrWhiteSpace( $selection.SubscriptionId ) )
                        {
                            [void][Windows.MessageBox]::Show( $mainWindow , "Unable to determine subscription for VM '$($selection.Name)'." , 'VM Activity Logs' , 'Ok' , 'Warning' )
                        }
                        else
                        {
                            [string]$capturedVMResourceId = "/subscriptions/$($selection.SubscriptionId)/resourceGroups/$($selection.ResourceGroup)/providers/Microsoft.Compute/virtualMachines/$($selection.Name)"
                            [string]$capturedVMName       = $selection.Name

                            if( $vmActLogsWindow = New-WPFWindow -inputXAML $vmActivityLogsXAML )
                            {
                                $vmActLogsWindow.Owner = $mainWindow
                                $WPFlblVMActLogsHeader.Content = "Activity Logs: $capturedVMName"

                                $WPFbtnVMActLogsRetrieve.Add_Click({
                                    [string]$valStr = $WPFtxtVMActLogsValue.Text.Trim()
                                    if( $valStr -notmatch '^\d+(\.\d+)?$' -or [double]$valStr -le 0 )
                                    {
                                        $WPFlblVMActLogsStatus.Foreground = [System.Windows.Media.Brushes]::Red
                                        $WPFlblVMActLogsStatus.Content    = 'Please enter a positive number (e.g. 2 or 1.5).'
                                        return
                                    }

                                    [double]$timeVal   = [double]$valStr
                                    [TimeSpan]$span    = if( $WPFradVMActLogsHours.IsChecked ) { [TimeSpan]::FromHours( $timeVal ) } else { [TimeSpan]::FromDays( $timeVal ) }
                                    [datetime]$startDT = (Get-Date) - $span

                                    $WPFlblVMActLogsStatus.Foreground = [System.Windows.Media.Brushes]::Gray
                                    $WPFlblVMActLogsStatus.Content    = 'Retrieving...'
                                    $WPFdgVMActLogs.ItemsSource       = $null
                                    $vmActLogsWindow.Dispatcher.Invoke( [System.Windows.Threading.DispatcherPriority]::Background , [action]{} )

                                    try
                                    {
                                        [array]$logEntries = @( Get-AzActivityLog -ResourceId $capturedVMResourceId -StartTime $startDT -EndTime (Get-Date) -WarningAction SilentlyContinue -ErrorAction Stop | Sort-Object -Property EventTimestamp )

                                        if( $logEntries.Count -eq 0 )
                                        {
                                            $WPFlblVMActLogsStatus.Content = 'No activity log entries found for the specified time range.'
                                        }
                                        else
                                        {
                                            $dt = [System.Data.DataTable]::new()
                                            [void]$dt.Columns.Add( 'Time' )
                                            [void]$dt.Columns.Add( 'Caller' )
                                            [void]$dt.Columns.Add( 'Operation' )
                                            [void]$dt.Columns.Add( 'Status' )
                                            [void]$dt.Columns.Add( 'Level' )
                                            [void]$dt.Columns.Add( 'ResourceType' )
                                            [void]$dt.Columns.Add( 'Description' )
                                            # Resolve a field that may be a plain string, a LocalizableString, or null
                                            $resolveActVal = {
                                                param( $v )
                                                if( $null -eq $v )    { return '' }
                                                if( $v -is [string] ) { return $v }
                                                $lp = $v.PSObject.Properties.Item( 'LocalizedValue' )
                                                if( $null -ne $lp -and -not [string]::IsNullOrEmpty( $lp.Value ) ) { return [string]$lp.Value }
                                                $vp = $v.PSObject.Properties.Item( 'Value' )
                                                if( $null -ne $vp -and -not [string]::IsNullOrEmpty( [string]$vp.Value ) ) { return [string]$vp.Value }
                                                return ''
                                            }
                                            foreach( $entry in $logEntries )
                                            {
                                                $dr = $dt.NewRow()
                                                $dr['Time']         = $entry.EventTimestamp.ToLocalTime().ToString('g')
                                                $dr['Caller']       = [string]$entry.Caller
                                                $dr['Operation']    = & $resolveActVal $entry.OperationName
                                                $dr['Status']       = & $resolveActVal $entry.Status
                                                $dr['Level']        = [string]$entry.Level
                                                $dr['ResourceType'] = & $resolveActVal $entry.ResourceType
                                                $dr['Description']  = [string]$entry.Description
                                                [void]$dt.Rows.Add( $dr )
                                            }
                                            $WPFdgVMActLogs.ItemsSource    = $dt.DefaultView
                                            $WPFlblVMActLogsStatus.Content = "$($logEntries.Count) entries  (oldest first)"
                                        }
                                    }
                                    catch
                                    {
                                        $WPFlblVMActLogsStatus.Foreground = [System.Windows.Media.Brushes]::Red
                                        $WPFlblVMActLogsStatus.Content    = "Retrieval failed: $($_.Exception.Message)"
                                    }
                                })

                                [void]$vmActLogsWindow.ShowDialog()
                            }
                        }
                    }
                    catch
                    {
                        Write-Warning -Message "VM Activity Logs error: $($_.Exception.Message)"
                        [void][Windows.MessageBox]::Show( $mainWindow , "VM Activity Logs failed.`n$($_.Exception.Message)" , 'VM Activity Logs' , 'Ok' , 'Error' )
                    }
                    break  # single VM operation
                }
                elseif( $operation -ieq 'Azure_EditTags' )
                {
                    try
                    {
                        Import-Module -Name Az.Compute   -Verbose:$false
                        Import-Module -Name Az.Resources -Verbose:$false

                        $vmFull = $null
                        $vmFull = Get-AzVM -Name $selection.Name -ResourceGroupName $selection.ResourceGroup -ErrorAction Stop

                        # Snapshot of original tags (sorted)
                        [hashtable]$originalTags = @{}
                        if( $null -ne $vmFull.Tags )
                        {
                            foreach( $kv in $vmFull.Tags.GetEnumerator() )
                            {
                                $originalTags[ $kv.Key ] = $kv.Value
                            }
                        }

                        if( $tagsWindow = New-WPFWindow -inputXAML $azureTagsEditorXAML )
                        {
                            $tagsWindow.Owner = $mainWindow
                            $WPFlblTagsHeader.Content = "Tags for VM: $($selection.Name)   ($($originalTags.Count) tag(s))"

                            # List tracking each row's data
                            $tagRows = [System.Collections.Generic.List[hashtable]]::new()

                            # ── helper function: add one tag row to the ItemsControl ──────────────
                            function Add-AzureTagRow
                            {
                                Param(
                                    [string]$tagName  = '' ,
                                    [string]$tagValue = '' ,
                                    [bool]$isNew      = $false
                                )

                                $rowGrid = New-Object -TypeName System.Windows.Controls.Grid
                                $rowGrid.Margin = [System.Windows.Thickness]::new( 0 , 2 , 0 , 2 )

                                $c0 = New-Object System.Windows.Controls.ColumnDefinition ; $c0.Width = [System.Windows.GridLength]::new( 210 )
                                $c1 = New-Object System.Windows.Controls.ColumnDefinition ; $c1.Width = [System.Windows.GridLength]::new( 1 , [System.Windows.GridUnitType]::Star )
                                $c2 = New-Object System.Windows.Controls.ColumnDefinition ; $c2.Width = [System.Windows.GridLength]::new( 32 )
                                [void]$rowGrid.ColumnDefinitions.Add( $c0 )
                                [void]$rowGrid.ColumnDefinitions.Add( $c1 )
                                [void]$rowGrid.ColumnDefinitions.Add( $c2 )

                                # Name: Label for existing, TextBox for new
                                if( $isNew )
                                {
                                    $nameCtrl = New-Object -TypeName System.Windows.Controls.TextBox
                                    $nameCtrl.Text = $tagName
                                    $nameCtrl.Height = 24
                                    $nameCtrl.VerticalContentAlignment = [System.Windows.VerticalAlignment]::Center
                                    $nameCtrl.Margin = [System.Windows.Thickness]::new( 0 , 0 , 4 , 0 )
                                    $nameCtrl.ToolTip = 'Tag name (must be unique)'
                                }
                                else
                                {
                                    $nameCtrl = New-Object -TypeName System.Windows.Controls.TextBlock
                                    $nameCtrl.Text = $tagName
                                    $nameCtrl.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
                                    $nameCtrl.Margin = [System.Windows.Thickness]::new( 4 , 0 , 4 , 0 )
                                    $nameCtrl.ToolTip = $tagName
                                }
                                [System.Windows.Controls.Grid]::SetColumn( $nameCtrl , 0 )
                                [void]$rowGrid.Children.Add( $nameCtrl )

                                # Value TextBox
                                $valueTB = New-Object -TypeName System.Windows.Controls.TextBox
                                $valueTB.Text = $tagValue
                                $valueTB.Height = 24
                                $valueTB.VerticalContentAlignment = [System.Windows.VerticalAlignment]::Center
                                $valueTB.Margin = [System.Windows.Thickness]::new( 0 , 0 , 4 , 0 )
                                [System.Windows.Controls.Grid]::SetColumn( $valueTB , 1 )
                                [void]$rowGrid.Children.Add( $valueTB )

                                # Delete (X) button
                                $delBtn = New-Object -TypeName System.Windows.Controls.Button
                                $delBtn.Content = 'X'
                                $delBtn.Foreground = [System.Windows.Media.Brushes]::Red
                                $delBtn.FontWeight = [System.Windows.FontWeights]::Bold
                                $delBtn.Padding = [System.Windows.Thickness]::new( 2 )
                                $delBtn.Height = 24
                                $delBtn.ToolTip = 'Delete this tag'
                                [System.Windows.Controls.Grid]::SetColumn( $delBtn , 2 )
                                [void]$rowGrid.Children.Add( $delBtn )

                                $rowData = @{
                                    Name      = $tagName
                                    IsNew     = $isNew
                                    IsDeleted = $false
                                    RowGrid   = $rowGrid
                                    ValueTB   = $valueTB
                                    NameCtrl  = $nameCtrl  # TextBox if new, TextBlock if existing
                                }
                                [void]$tagRows.Add( $rowData )
                                [void]$WPFtagsItemsControl.Items.Add( $rowGrid )

                                # Capture by value for the click closure
                                $capturedRow  = $rowData
                                $capturedGrid = $rowGrid
                                $delBtn.Add_Click( {
                                    $capturedRow.IsDeleted = $true
                                    [void]$WPFtagsItemsControl.Items.Remove( $capturedGrid )
                                }.GetNewClosure() )
                            }

                            # Add existing tags (sorted alphabetically)
                            foreach( $kv in $originalTags.GetEnumerator() | Sort-Object -Property Key )
                            {
                                Add-AzureTagRow -tagName $kv.Key -tagValue $kv.Value -isNew $false
                            }

                            # ── Add Tag button handler ────────────────────────────────────────────
                            $WPFbtnAddTag.Add_Click( {
                                $newName  = $WPFtxtNewTagName.Text.Trim()
                                $newValue = $WPFtxtNewTagValue.Text

                                if( [string]::IsNullOrWhiteSpace( $newName ) )
                                {
                                    [void][Windows.MessageBox]::Show( $tagsWindow , 'Tag name cannot be empty.' , 'Add Tag' , 'Ok' , 'Warning' )
                                    return
                                }

                                # Collect all current non-deleted names
                                [array]$currentNames = @(
                                    $tagRows | Where-Object { -not $_.IsDeleted } | ForEach-Object {
                                        if( $_.IsNew ) { $_.NameCtrl.Text.Trim() } else { $_.Name }
                                    }
                                )

                                if( $currentNames -icontains $newName )
                                {
                                    [void][Windows.MessageBox]::Show( $tagsWindow , "Tag '$newName' already exists." , 'Add Tag' , 'Ok' , 'Warning' )
                                    return
                                }

                                Add-AzureTagRow -tagName $newName -tagValue $newValue -isNew $true
                                $WPFtxtNewTagName.Text  = ''
                                $WPFtxtNewTagValue.Text = ''
                                $WPFtxtNewTagName.Focus() | Out-Null
                            } )

                            # ── OK button handler ─────────────────────────────────────────────────
                            $WPFbtnTagsOK.Add_Click( {
                                $_.Handled = $true
                                $tagsWindow.DialogResult = $true
                                $tagsWindow.Close()
                            } )

                            if( $tagsWindow.ShowDialog() )
                            {
                                # Build new tags hashtable from current UI state
                                [hashtable]$newTags = @{}
                                [System.Collections.Generic.List[string]]$duplicateCheck = New-Object -TypeName System.Collections.Generic.List[string]
                                [bool]$hasDuplicate = $false

                                foreach( $row in $tagRows )
                                {
                                    if( $row.IsDeleted ) { continue }
                                    [string]$n = if( $row.IsNew ) { $row.NameCtrl.Text.Trim() } else { $row.Name }
                                    [string]$v = $row.ValueTB.Text

                                    if( [string]::IsNullOrWhiteSpace( $n ) ) { continue }

                                    if( $duplicateCheck -icontains $n )
                                    {
                                        [void][Windows.MessageBox]::Show( $mainWindow , "Duplicate tag name found: '$n'. Please fix before saving." , 'Edit Tags' , 'Ok' , 'Warning' )
                                        $hasDuplicate = $true
                                        break
                                    }
                                    [void]$duplicateCheck.Add( $n )
                                    $newTags[ $n ] = $v
                                }

                                if( -not $hasDuplicate )
                                {
                                    # Determine what changed
                                    [System.Collections.Generic.List[string]]$changeLines = New-Object -TypeName System.Collections.Generic.List[string]

                                    foreach( $key in $originalTags.Keys )
                                    {
                                        if( -not $newTags.ContainsKey( $key ) )
                                        {
                                            [void]$changeLines.Add( "Delete:  $key" )
                                        }
                                    }
                                    foreach( $kv in $newTags.GetEnumerator() )
                                    {
                                        if( -not $originalTags.ContainsKey( $kv.Key ) )
                                        {
                                            [void]$changeLines.Add( "Add:     $($kv.Key) = '$($kv.Value)'" )
                                        }
                                        elseif( $originalTags[ $kv.Key ] -cne $kv.Value )
                                        {
                                            [void]$changeLines.Add( "Modify:  $($kv.Key)  '$($originalTags[$kv.Key])' -> '$($kv.Value)'" )
                                        }
                                    }

                                    if( $changeLines.Count -eq 0 )
                                    {
                                        [void][Windows.MessageBox]::Show( $mainWindow , "No tag changes detected for $($selection.Name)." , 'Edit Tags' , 'Ok' , 'Information' )
                                    }
                                    else
                                    {
                                        [string]$summary = "$($changeLines.Count) change(s) for '$($selection.Name)':`n`n" + ( $changeLines -join "`n" ) + "`n`nApply these changes?"
                                        if( [Windows.MessageBox]::Show( $mainWindow , $summary , 'Confirm Tag Changes' , 'YesNo' , 'Question' ) -ieq 'Yes' )
                                        {
                                            $null = Update-AzTag -ResourceId $vmFull.Id -Tag $newTags -Operation Replace -ErrorAction Stop
                                            Write-Host "Tags updated for $($selection.Name): $($changeLines.Count) change(s)" -ForegroundColor Green
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch
                    {
                        Write-Warning -Message "Edit tags error for $($selection.Name): $($_.Exception.Message)"
                        [void][Windows.MessageBox]::Show( $mainWindow , "Edit tags failed for $($selection.Name)`n$($_.Exception.Message)" , 'Edit Azure Tags' , 'Ok' , 'Error' )
                    }
                }
                elseif( $operation -ieq 'Azure_Delete' )
                {
                    try
                    {
                        if( Remove-AzureVMEntry -selection $selection )
                        {
                            $refreshAzureList = $true
                        }
                    }
                    catch
                    {
                        Write-Warning -Message "Delete VM error for $($selection.Name): $($_.Exception.Message)"
                        [void][Windows.MessageBox]::Show( $mainWindow , "Delete failed for $($selection.Name)`n$($_.Exception.Message)" , 'Azure Delete VM' , 'Ok' , 'Error' )
                    }
                }
                elseif( $operation -ieq 'Azure_DeleteSessionHost' -or $operation -ieq 'Azure_DeleteSessionHostAndVM' )
                {
                    try
                    {
                        if( [string]::IsNullOrEmpty( $selection.HostPool ) )
                        {
                            [void][Windows.MessageBox]::Show( $mainWindow , "$($selection.Name) is not an AVD session host" , 'Delete Session Host' , 'Ok' , 'Warning' )
                            continue
                        }

                        $sessionHost = Get-AzureAVDSessionHost -selection $selection
                        $actonResult = Invoke-AzRestMethod -Method DELETE -Path "$($sessionHost.Id)?api-version=$AVDAPIversion" -ErrorAction Stop
                        Write-Host "Removed session host: $($selection.Name) from host pool $($selection.HostPool)" -ForegroundColor Green

                        if( $operation -ieq 'Azure_DeleteSessionHostAndVM' )
                        {
                            [void]( Remove-AzureVMEntry -selection $selection )
                        }
                        $refreshAzureList = $true
                    }
                    catch
                    {
                        Write-Warning -Message "Delete session host error for $($selection.Name): $($_.Exception.Message)"
                        [void][Windows.MessageBox]::Show( $mainWindow , "Delete session host failed for $($selection.Name)`n$($_.Exception.Message)" , 'Delete Session Host' , 'Ok' , 'Error' )
                    }
                }
                elseif( $operation -ieq 'Azure_Detail' )
                {
                    try
                    {
                        Import-Module -Name Az.Compute    -Verbose:$false
                        Import-Module -Name Az.Network    -Verbose:$false
                        Import-Module -Name Az.Resources  -Verbose:$false

                        $vm       = Get-AzVM -Name $selection.Name -ResourceGroupName $selection.ResourceGroup -ErrorAction Stop
                        $vmStatus = Get-AzVM -Name $selection.Name -ResourceGroupName $selection.ResourceGroup -Status -ErrorAction SilentlyContinue

                        $detailsBuilder = New-Object -TypeName System.Text.StringBuilder

                        # ── General ──────────────────────────────────────────────────────
                        [void]$detailsBuilder.AppendLine( '=== General ===' )
                        [void]$detailsBuilder.AppendLine( "VM Name         : $($vm.Name)" )
                        [void]$detailsBuilder.AppendLine( "Resource Group  : $($vm.ResourceGroupName)" )
                        [void]$detailsBuilder.AppendLine( "Location        : $($vm.Location)" )
                        [void]$detailsBuilder.AppendLine( "VM ID           : $($vm.VmId)" )
                        [void]$detailsBuilder.AppendLine( "Subscription    : $($selection.Subscription)" )
                        [void]$detailsBuilder.AppendLine( "VM Size         : $($vm.HardwareProfile.VmSize)" )
                        [void]$detailsBuilder.AppendLine( "Time Created    : $($vm.TimeCreated)" )
                        if( $null -ne $vmStatus )
                        {
                            [string]$powerState = ( $vmStatus.Statuses | Where-Object Code -match '^PowerState/' | Select-Object -First 1 -ExpandProperty DisplayStatus )
                            [string]$provState  = ( $vmStatus.Statuses | Where-Object Code -match '^ProvisioningState/' | Select-Object -First 1 -ExpandProperty DisplayStatus )
                            [void]$detailsBuilder.AppendLine( "Power State     : $powerState" )
                            [void]$detailsBuilder.AppendLine( "Prov. State     : $provState" )
                        }

                        # ── Tags ─────────────────────────────────────────────────────────
                        [void]$detailsBuilder.AppendLine( '' )
                        [void]$detailsBuilder.AppendLine( '=== Tags ===' )
                        if( $null -ne $vm.Tags -and $vm.Tags.Count -gt 0 )
                        {
                            ForEach( $tag in $vm.Tags.GetEnumerator() | Sort-Object -Property Key )
                            {
                                [void]$detailsBuilder.AppendLine( "  $($tag.Key) = $($tag.Value)" )
                            }
                        }
                        else
                        {
                            [void]$detailsBuilder.AppendLine( '  (none)' )
                        }

                        # ── Source Image ─────────────────────────────────────────────────
                        [void]$detailsBuilder.AppendLine( '' )
                        [void]$detailsBuilder.AppendLine( '=== Source Image ===' )
                        $imageRef = $vm.StorageProfile.ImageReference
                        if( $null -ne $imageRef )
                        {
                            if( -Not [string]::IsNullOrWhiteSpace( $imageRef.Id ) )
                            {
                                [void]$detailsBuilder.AppendLine( "  Image ID        : $($imageRef.Id)" )
                            }
                            else
                            {
                                [void]$detailsBuilder.AppendLine( "  Publisher       : $($imageRef.Publisher)" )
                                [void]$detailsBuilder.AppendLine( "  Offer           : $($imageRef.Offer)" )
                                [void]$detailsBuilder.AppendLine( "  SKU             : $($imageRef.Sku)" )
                                [void]$detailsBuilder.AppendLine( "  Version         : $($imageRef.Version)" )
                                [void]$detailsBuilder.AppendLine( "  Exact Version   : $($imageRef.ExactVersion)" )
                            }
                        }
                        else
                        {
                            [void]$detailsBuilder.AppendLine( '  (no image reference)' )
                        }
                        $osDisk = $vm.StorageProfile.OsDisk
                        if( $null -ne $osDisk )
                        {
                            [void]$detailsBuilder.AppendLine( "  OS Disk Name    : $($osDisk.Name)" )
                            [void]$detailsBuilder.AppendLine( "  OS Type         : $($osDisk.OsType)" )
                            [void]$detailsBuilder.AppendLine( "  Caching         : $($osDisk.Caching)" )
                            [void]$detailsBuilder.AppendLine( "  Create Option   : $($osDisk.CreateOption)" )
                            if( $null -ne $osDisk.ManagedDisk )
                            {
                                [void]$detailsBuilder.AppendLine( "  Managed Disk ID : $($osDisk.ManagedDisk.Id)" )
                                $resolvedDisk = $null
                                try
                                {
                                    [string]$diskRg   = ($osDisk.ManagedDisk.Id -split '/')[4]
                                    [string]$diskName = ($osDisk.ManagedDisk.Id -split '/')[-1]
                                    $resolvedDisk = Get-AzDisk -ResourceGroupName $diskRg -DiskName $diskName -ErrorAction Stop
                                }
                                catch {}
                                [string]$diskSku    = if( $null -ne $resolvedDisk -and -not [string]::IsNullOrWhiteSpace( $resolvedDisk.Sku.Name ) ) { $resolvedDisk.Sku.Name } else { [string]$osDisk.ManagedDisk.StorageAccountType }
                                [int]$resolvedSize  = if( $null -ne $resolvedDisk -and $resolvedDisk.DiskSizeGB -gt 0 ) { $resolvedDisk.DiskSizeGB } else { [int]$osDisk.DiskSizeGB }
                                if( -not [string]::IsNullOrWhiteSpace( $diskSku ) )  { [void]$detailsBuilder.AppendLine( "  Disk SKU        : $diskSku" ) }
                                if( $resolvedSize -gt 0 )                             { [void]$detailsBuilder.AppendLine( "  Disk Size GB    : $resolvedSize" ) }
                            }
                            elseif( $null -ne $osDisk.DiskSizeGB -and $osDisk.DiskSizeGB -gt 0 )
                            {
                                [void]$detailsBuilder.AppendLine( "  Disk Size GB    : $($osDisk.DiskSizeGB)" )
                            }
                        }

                        # ── Data Disks ───────────────────────────────────────────────────
                        [void]$detailsBuilder.AppendLine( '' )
                        [void]$detailsBuilder.AppendLine( '=== Data Disks ===' )
                        [array]$dataDisks = @( $vm.StorageProfile.DataDisks )
                        if( $dataDisks.Count -gt 0 )
                        {
                            ForEach( $disk in $dataDisks )
                            {
                                [void]$detailsBuilder.AppendLine( "  LUN $($disk.Lun): $($disk.Name)" )
                                [void]$detailsBuilder.AppendLine( "    Size GB       : $($disk.DiskSizeGB)" )
                                [void]$detailsBuilder.AppendLine( "    Caching       : $($disk.Caching)" )
                                [void]$detailsBuilder.AppendLine( "    Create Option : $($disk.CreateOption)" )
                                if( $null -ne $disk.ManagedDisk )
                                {
                                    [void]$detailsBuilder.AppendLine( "    Storage Type  : $($disk.ManagedDisk.StorageAccountType)" )
                                    [void]$detailsBuilder.AppendLine( "    Disk ID       : $($disk.ManagedDisk.Id)" )
                                }
                                [void]$detailsBuilder.AppendLine( '' )
                            }
                        }
                        else
                        {
                            [void]$detailsBuilder.AppendLine( '  (no data disks)' )
                        }

                        # ── Networking ───────────────────────────────────────────────────
                        [void]$detailsBuilder.AppendLine( '' )
                        [void]$detailsBuilder.AppendLine( '=== Networking ===' )
                        ForEach( $nicRef in @( $vm.NetworkProfile.NetworkInterfaces ) )
                        {
                            [string]$nicName    = ( $nicRef.Id -split '/' | Select-Object -Last 1 )
                            [string]$nicRG      = ( $nicRef.Id -replace '^.*resourceGroups/([^/]+)/.*$','$1' )
                            [void]$detailsBuilder.AppendLine( "  NIC: $nicName" )
                            [void]$detailsBuilder.AppendLine( "    Primary       : $($nicRef.Primary)" )
                            $nicDetail = $null
                            try { $nicDetail = Get-AzNetworkInterface -Name $nicName -ResourceGroupName $nicRG -ErrorAction Stop } catch {}
                            if( $null -ne $nicDetail )
                            {
                                [void]$detailsBuilder.AppendLine( "    MAC Address   : $($nicDetail.MacAddress)" )
                                [void]$detailsBuilder.AppendLine( "    DNS Servers   : $($nicDetail.DnsSettings.DnsServers -join ', ')" )
                                [void]$detailsBuilder.AppendLine( "    Enable Accel. Networking: $($nicDetail.EnableAcceleratedNetworking)" )
                                ForEach( $ipConfig in @( $nicDetail.IpConfigurations ) )
                                {
                                    [void]$detailsBuilder.AppendLine( "    IP Config     : $($ipConfig.Name)" )
                                    [void]$detailsBuilder.AppendLine( "      Private IP  : $($ipConfig.PrivateIpAddress)  ($($ipConfig.PrivateIpAllocationMethod))" )
                                    if( $null -ne $ipConfig.PublicIpAddress )
                                    {
                                        [string]$pipName = ( $ipConfig.PublicIpAddress.Id -split '/' | Select-Object -Last 1 )
                                        [string]$pipRG   = ( $ipConfig.PublicIpAddress.Id -replace '^.*resourceGroups/([^/]+)/.*$','$1' )
                                        $pip = $null
                                        try { $pip = Get-AzPublicIpAddress -Name $pipName -ResourceGroupName $pipRG -ErrorAction Stop } catch {}
                                        [string]$publicIp = if( $null -ne $pip ) { "$($pip.IpAddress)  ($($pip.PublicIpAllocationMethod))" } else { $pipName }
                                        [void]$detailsBuilder.AppendLine( "      Public IP   : $publicIp" )
                                    }
                                    if( $null -ne $ipConfig.Subnet )
                                    {
                                        [string]$subnetId = $ipConfig.Subnet.Id
                                        [string]$vnetName = ( $subnetId -split '/' | Select-Object -Index 8 )
                                        [string]$subnetName = ( $subnetId -split '/' | Select-Object -Last 1 )
                                        [void]$detailsBuilder.AppendLine( "      VNet        : $vnetName" )
                                        [void]$detailsBuilder.AppendLine( "      Subnet      : $subnetName" )
                                    }
                                    if( $null -ne $ipConfig.LoadBalancerBackendAddressPools -and $ipConfig.LoadBalancerBackendAddressPools.Count -gt 0 )
                                    {
                                        [void]$detailsBuilder.AppendLine( "      LB Pools    : $($ipConfig.LoadBalancerBackendAddressPools.Count)" )
                                    }
                                }
                                if( $null -ne $nicDetail.NetworkSecurityGroup )
                                {
                                    [string]$nsgName = ( $nicDetail.NetworkSecurityGroup.Id -split '/' | Select-Object -Last 1 )
                                    [void]$detailsBuilder.AppendLine( "    NSG           : $nsgName" )
                                }
                            }
                            [void]$detailsBuilder.AppendLine( '' )
                        }

                        # ── Security ─────────────────────────────────────────────────────
                        [void]$detailsBuilder.AppendLine( '=== Security ===' )
                        $secProfile = $vm.SecurityProfile
                        if( $null -ne $secProfile )
                        {
                            [void]$detailsBuilder.AppendLine( "  Security Type          : $($secProfile.SecurityType)" )
                            if( $null -ne $secProfile.UefiSettings )
                            {
                                [void]$detailsBuilder.AppendLine( "  Secure Boot Enabled    : $($secProfile.UefiSettings.SecureBootEnabled)" )
                                [void]$detailsBuilder.AppendLine( "  vTPM Enabled           : $($secProfile.UefiSettings.VTpmEnabled)" )
                            }
                            if( $null -ne $secProfile.EncryptionAtHost )
                            {
                                [void]$detailsBuilder.AppendLine( "  Encryption At Host     : $($secProfile.EncryptionAtHost)" )
                            }
                        }
                        else
                        {
                            [void]$detailsBuilder.AppendLine( '  (standard / no Trusted Launch profile)' )
                        }

                        # Disk encryption
                        [array]$diskEncStatuses = @()
                        if( $null -ne $vmStatus -and $vmStatus.PSObject.Properties.Match( 'Disks' ).Count -gt 0 )
                        {
                            ForEach( $d in @( $vmStatus.Disks ) )
                            {
                                if( $null -ne $d -and $d.PSObject.Properties.Match( 'EncryptionSettings' ).Count -gt 0 -and $null -ne $d.EncryptionSettings )
                                {
                                    $diskEncStatuses += "  Disk $($d.Name): enabled=$($d.EncryptionSettings.Enabled)"
                                }
                            }
                        }
                        if( $diskEncStatuses.Count -gt 0 )
                        {
                            [void]$detailsBuilder.AppendLine( '  Disk Encryption:' )
                            $diskEncStatuses | ForEach-Object { [void]$detailsBuilder.AppendLine( $_ ) }
                        }
                        else
                        {
                            [void]$detailsBuilder.AppendLine( '  Disk Encryption        : not reported / not enabled' )
                        }

                        # ── Auto-Shutdown ─────────────────────────────────────────────────
                        [void]$detailsBuilder.AppendLine( '' )
                        [void]$detailsBuilder.AppendLine( '=== Auto-Shutdown ===' )
                        try
                        {
                            $autoShutdown = $null
                            $subscriptionId = $selection.SubscriptionId
                            if( [string]::IsNullOrWhiteSpace( $subscriptionId ) )
                            {
                                $subscriptionId = (Get-AzContext).Subscription.Id
                            }
                            $autoShutdownUri = "/subscriptions/$subscriptionId/resourceGroups/$($selection.ResourceGroup)/providers/microsoft.devtestlab/schedules/shutdown-computevm-$($selection.Name)"
                            $autoShutdown = Get-AzResource -ResourceId $autoShutdownUri -ApiVersion '2018-09-15' -ErrorAction SilentlyContinue
                            if( $null -ne $autoShutdown )
                            {
                                [void]$detailsBuilder.AppendLine( "  Status          : $($autoShutdown.Properties.status)" )
                                [void]$detailsBuilder.AppendLine( "  Daily Time      : $($autoShutdown.Properties.dailyRecurrence.time)" )
                                [void]$detailsBuilder.AppendLine( "  Time Zone       : $($autoShutdown.Properties.timeZoneId)" )
                                if( $null -ne $autoShutdown.Properties.notificationSettings )
                                {
                                    [void]$detailsBuilder.AppendLine( "  Notify Status   : $($autoShutdown.Properties.notificationSettings.status)" )
                                    [void]$detailsBuilder.AppendLine( "  Notify Email    : $($autoShutdown.Properties.notificationSettings.emailRecipient)" )
                                    [void]$detailsBuilder.AppendLine( "  Notify Mins     : $($autoShutdown.Properties.notificationSettings.timeInMinutes)" )
                                }
                            }
                            else
                            {
                                [void]$detailsBuilder.AppendLine( '  (not configured)' )
                            }
                        }
                        catch
                        {
                            [void]$detailsBuilder.AppendLine( "  Error retrieving auto-shutdown: $($_.Exception.Message)" )
                        }

                        # ── AVD Detail (only when AVD checkbox ticked) ────────────────────
                        if( $WPFcheckBoxAzureAVD.IsChecked )
                        {
                            [void]$detailsBuilder.AppendLine( '' )
                            [void]$detailsBuilder.AppendLine( '=== AVD ===' )
                            try
                            {
                                Import-Module -Name Az.DesktopVirtualization -Verbose:$false

                                # Find the session host for this VM across all host pools
                                [string]$vmId = $vm.VmId
                                $matchedSessionHost = $null
                                $matchedHostPool    = $null

                                ForEach( $hp in @( Get-AzWvdHostPool ) )
                                {
                                    $sh = Get-AzWvdSessionHost -HostPoolName $hp.Name -ResourceGroupName $hp.ResourceGroupName -ErrorAction SilentlyContinue |
                                          Where-Object { $_.VirtualMachineId -ieq $vmId } |
                                          Select-Object -First 1
                                    if( $null -ne $sh )
                                    {
                                        $matchedSessionHost = $sh
                                        $matchedHostPool    = $hp
                                        break
                                    }
                                }

                                if( $null -ne $matchedSessionHost )
                                {
                                    [void]$detailsBuilder.AppendLine( "  Host Pool         : $($matchedHostPool.Name)" )
                                    [void]$detailsBuilder.AppendLine( "  Host Pool Type    : $($matchedHostPool.HostPoolType)" )
                                    [void]$detailsBuilder.AppendLine( "  Load Balancer     : $($matchedHostPool.LoadBalancerType)" )
                                    [void]$detailsBuilder.AppendLine( "  Max Session Limit : $($matchedHostPool.MaxSessionLimit)" )
                                    [void]$detailsBuilder.AppendLine( "  Validation Env    : $($matchedHostPool.ValidationEnvironment)" )
                                    [void]$detailsBuilder.AppendLine( "  Start VM On Conn  : $($matchedHostPool.StartVMOnConnect)" )
                                    [void]$detailsBuilder.AppendLine( '' )
                                    [void]$detailsBuilder.AppendLine( "  Session Host      : $($matchedSessionHost.Name -replace '^.*/','')") 
                                    [void]$detailsBuilder.AppendLine( "  SH Status         : $($matchedSessionHost.Status)" )
                                    [void]$detailsBuilder.AppendLine( "  Update State      : $($matchedSessionHost.UpdateState)" )
                                    [void]$detailsBuilder.AppendLine( "  Last Heartbeat    : $($matchedSessionHost.LastHeartBeat)" )
                                    [void]$detailsBuilder.AppendLine( "  Agent Version     : $($matchedSessionHost.AgentVersion)" )
                                    [void]$detailsBuilder.AppendLine( "  Allow New Session : $($matchedSessionHost.AllowNewSession)" )
                                    [void]$detailsBuilder.AppendLine( "  Sessions          : $($matchedSessionHost.Session)" )
                                    [void]$detailsBuilder.AppendLine( "  Assigned User     : $($matchedSessionHost.AssignedUser)" )
                                    [void]$detailsBuilder.AppendLine( "  Domain Join Type  : $($matchedSessionHost.DomainName)" )
                                    if( -Not [string]::IsNullOrWhiteSpace( $matchedSessionHost.UpdateErrorMessage ) )
                                    {
                                        [void]$detailsBuilder.AppendLine( "  Update Error      : $($matchedSessionHost.UpdateErrorMessage)" )
                                    }

                                    # Application groups referencing this host pool
                                    [void]$detailsBuilder.AppendLine( '' )
                                    [void]$detailsBuilder.AppendLine( '  Application Groups:' )
                                    [array]$appGroups = @( Get-AzWvdApplicationGroup | Where-Object { $_.HostPoolArmPath -ieq $matchedHostPool.Id } )
                                    if( $appGroups.Count -gt 0 )
                                    {
                                        ForEach( $ag in $appGroups )
                                        {
                                            [void]$detailsBuilder.AppendLine( "    $($ag.Name)  [$($ag.ApplicationGroupType)]" )
                                        }
                                    }
                                    else
                                    {
                                        [void]$detailsBuilder.AppendLine( '    (none found)' )
                                    }
                                }
                                else
                                {
                                    [void]$detailsBuilder.AppendLine( '  VM is not registered as a session host in any host pool.' )
                                }
                            }
                            catch
                            {
                                [void]$detailsBuilder.AppendLine( "  Error retrieving AVD detail: $($_.Exception.Message)" )
                            }
                        }

                        Show-AzureRunCommandOutputWindow -computerName $selection.Name -scriptText 'VM Detail' -outputText $detailsBuilder.ToString().Trim()
                    }
                    catch
                    {
                        [void][Windows.MessageBox]::Show( $mainWindow , "Failed to retrieve detail for $($selection.Name)`n$($_.Exception.Message)" , 'Azure VM Detail' , 'Ok' ,'Error' )
                    }
                }
                elseif( $operation -ieq 'Azure_DetailExtensionsApplications' )
                {
                    try
                    {
                        Import-Module -Name Az.Compute -Verbose:$false

                        $vm = Get-AzVM -Name $selection.Name -ResourceGroupName $selection.ResourceGroup -ErrorAction Stop
                        $vmStatus = Get-AzVM -Name $selection.Name -ResourceGroupName $selection.ResourceGroup -Status -ErrorAction SilentlyContinue

                        [array]$extensions = @()
                        [string]$extensionsError = $null
                        try
                        {
                            $extensions = @( Get-AzVMExtension -VMName $selection.Name -ResourceGroupName $selection.ResourceGroup -ErrorAction Stop )
                        }
                        catch
                        {
                            $extensionsError = $_.Exception.Message
                        }

                        [hashtable]$extensionStatusByName = @{}
                        if( $null -ne $vmStatus -and $vmStatus.PSObject.Properties.Match( 'Extensions' ).Count -gt 0 )
                        {
                            ForEach( $statusExtension in @( $vmStatus.Extensions ) )
                            {
                                if( $null -ne $statusExtension -and -Not [string]::IsNullOrWhiteSpace( $statusExtension.Name ) )
                                {
                                    $extensionStatusByName[ $statusExtension.Name ] = $statusExtension
                                }
                            }
                        }

                        $detailsBuilder = New-Object -TypeName System.Text.StringBuilder
                        [void]$detailsBuilder.AppendLine( "VM: $($selection.Name)" )
                        [void]$detailsBuilder.AppendLine( "Resource Group: $($selection.ResourceGroup)" )
                        [void]$detailsBuilder.AppendLine( '' )

                        [void]$detailsBuilder.AppendLine( 'Extensions' )
                        [void]$detailsBuilder.AppendLine( '----------' )
                        if( $extensions.Count -gt 0 )
                        {
                            ForEach( $extension in $extensions )
                            {
                                [void]$detailsBuilder.AppendLine( "Name: $($extension.Name)" )
                                [void]$detailsBuilder.AppendLine( "Publisher: $($extension.Publisher)" )
                                [void]$detailsBuilder.AppendLine( "Type: $($extension.ExtensionType)" )
                                [void]$detailsBuilder.AppendLine( "Version: $($extension.TypeHandlerVersion)" )
                                if( $extension.PSObject.Properties.Match( 'ProvisioningState' ).Count -gt 0 )
                                {
                                    [void]$detailsBuilder.AppendLine( "Provisioning State: $($extension.ProvisioningState)" )
                                }
                                if( $extension.PSObject.Properties.Match( 'EnableAutomaticUpgrade' ).Count -gt 0 )
                                {
                                    [void]$detailsBuilder.AppendLine( "Enable Automatic Upgrade: $($extension.EnableAutomaticUpgrade)" )
                                }
                                if( $extension.PSObject.Properties.Match( 'AutoUpgradeMinorVersion' ).Count -gt 0 )
                                {
                                    [void]$detailsBuilder.AppendLine( "Auto Upgrade Minor Version: $($extension.AutoUpgradeMinorVersion)" )
                                }

                                if( $extensionStatusByName.ContainsKey( $extension.Name ) )
                                {
                                    $statusDetails = $extensionStatusByName[ $extension.Name ]
                                    if( $statusDetails.PSObject.Properties.Match( 'Statuses' ).Count -gt 0 -and $null -ne $statusDetails.Statuses )
                                    {
                                        [array]$statusTexts = @( ForEach( $statusEntry in $statusDetails.Statuses )
                                        {
                                            [string]$display = $null
                                            if( $statusEntry.PSObject.Properties.Match( 'DisplayStatus' ).Count -gt 0 )
                                            {
                                                $display = $statusEntry.DisplayStatus
                                            }
                                            if( [string]::IsNullOrWhiteSpace( $display ) -and $statusEntry.PSObject.Properties.Match( 'Code' ).Count -gt 0 )
                                            {
                                                $display = $statusEntry.Code
                                            }
                                            if( -Not [string]::IsNullOrWhiteSpace( $display ) )
                                            {
                                                $display
                                            }
                                        } )

                                        if( $statusTexts.Count -gt 0 )
                                        {
                                            [void]$detailsBuilder.AppendLine( "Status: $($statusTexts -join ' | ')" )
                                        }
                                    }
                                }

                                [void]$detailsBuilder.AppendLine( '' )
                            }
                        }
                        else
                        {
                            [void]$detailsBuilder.AppendLine( 'No extensions found.' )
                            if( -Not [string]::IsNullOrWhiteSpace( $extensionsError ) )
                            {
                                [void]$detailsBuilder.AppendLine( "Error retrieving extensions: $extensionsError" )
                            }
                            [void]$detailsBuilder.AppendLine( '' )
                        }

                        [void]$detailsBuilder.AppendLine( 'Applications' )
                        [void]$detailsBuilder.AppendLine( '------------' )
                        [array]$galleryApplications = @()
                        if( $null -ne $vm -and $vm.PSObject.Properties.Match( 'ApplicationProfile' ).Count -gt 0 -and $null -ne $vm.ApplicationProfile )
                        {
                            if( $vm.ApplicationProfile.PSObject.Properties.Match( 'GalleryApplications' ).Count -gt 0 -and $null -ne $vm.ApplicationProfile.GalleryApplications )
                            {
                                $galleryApplications = @( $vm.ApplicationProfile.GalleryApplications )
                            }
                        }

                        if( $galleryApplications.Count -gt 0 )
                        {
                            ForEach( $application in $galleryApplications )
                            {
                                [void]$detailsBuilder.AppendLine( "Package Reference: $($application.PackageReferenceId)" )
                                if( $application.PSObject.Properties.Match( 'ConfigurationReference' ).Count -gt 0 -and -Not [string]::IsNullOrWhiteSpace( $application.ConfigurationReference ) )
                                {
                                    [void]$detailsBuilder.AppendLine( "Configuration Reference: $($application.ConfigurationReference)" )
                                }
                                if( $application.PSObject.Properties.Match( 'Order' ).Count -gt 0 )
                                {
                                    [void]$detailsBuilder.AppendLine( "Order: $($application.Order)" )
                                }
                                if( $application.PSObject.Properties.Match( 'Tag' ).Count -gt 0 -and -Not [string]::IsNullOrWhiteSpace( $application.Tag ) )
                                {
                                    [void]$detailsBuilder.AppendLine( "Tag: $($application.Tag)" )
                                }
                                [void]$detailsBuilder.AppendLine( '' )
                            }
                        }
                        else
                        {
                            [void]$detailsBuilder.AppendLine( 'No gallery applications assigned.' )
                        }

                        Show-AzureRunCommandOutputWindow -computerName $selection.Name -scriptText 'Extensions + Applications' -outputText $detailsBuilder.ToString().Trim()
                    }
                    catch
                    {
                        [void][Windows.MessageBox]::Show( $mainWindow , "Failed to retrieve extensions and applications for $($selection.Name)`n$($_.Exception.Message)" , 'Azure VM Details' , 'Ok' ,'Error' )
                    }
                }
                elseif( $operation -ieq 'Azure_RunOn' )
                {
                    try
                    {
                        Import-Module -Name Az.Accounts -Verbose:$false
                        Import-Module -Name Az.Compute -Verbose:$false

                        $currentAzContext = Get-AzContext -ErrorAction Stop
                        if( $null -eq $currentAzContext )
                        {
                            throw 'No active Azure context. Connect to Azure first.'
                        }

                        [string]$contextPath = Join-Path -Path $env:TEMP -ChildPath ( 'mstsc-sizer-azcontext-{0}.json' -f ([guid]::NewGuid().ToString()) )
                        Save-AzContext -Path $contextPath -Force -ErrorAction Stop | Out-Null
                        Write-Verbose -Message "Azure_Run preparing command for VM '$($selection.Name)' in resource group '$($selection.ResourceGroup)'"
                        Write-Verbose -Message "Azure_Run script text: $azureRunCommandText"

                        try
                        {
                            $job = Start-Job -Name "Azure_Run_$($selection.Name)" -ArgumentList $selection.Name, $selection.ResourceGroup, $azureRunCommandText, $contextPath, $currentAzContext.Subscription.Id -ScriptBlock {
                                Param(
                                    [string]$vmName,
                                    [string]$resourceGroupName,
                                    [string]$scriptToRun,
                                    [string]$azContextPath,
                                    [string]$subscriptionId
                                )

                                Import-Module -Name Az.Accounts -Verbose:$false
                                Import-Module -Name Az.Compute -Verbose:$false
                                Import-AzContext -Path $azContextPath -ErrorAction Stop | Out-Null
                                if( -Not [string]::IsNullOrWhiteSpace( $subscriptionId ) )
                                {
                                    Set-AzContext -SubscriptionId $subscriptionId -ErrorAction Stop | Out-Null
                                }

                                Write-Verbose -Message "Azure_Run job executing on VM '$vmName' in resource group '$resourceGroupName' with script: $scriptToRun" -Verbose

                                $result = Invoke-AzVMRunCommand -ResourceGroupName $resourceGroupName -Name $vmName -CommandId RunPowerShellScript -ScriptString $scriptToRun -ErrorAction Stop
                                $outputParts = New-Object System.Collections.Generic.List[string]

                                $addOutputPart = {
                                    Param(
                                        [string]$text
                                    )

                                    if( -Not [string]::IsNullOrWhiteSpace( $text ) )
                                    {
                                        [void]$outputParts.Add( $text.TrimEnd() )
                                    }
                                }

                                $extractStatusMessage = {
                                    Param(
                                        $status
                                    )

                                    $convertValueToText = {
                                        Param(
                                            $value
                                        )

                                        if( $null -eq $value )
                                        {
                                            return ''
                                        }

                                        if( ( $value -is [System.Collections.IEnumerable] ) -and -Not ( $value -is [string] ) )
                                        {
                                            return ( @( $value ) | ForEach-Object { if( $null -eq $_ ) { '' } else { [string]$_ } } | Where-Object { -Not [string]::IsNullOrWhiteSpace( $_ ) } ) -join "`r`n"
                                        }

                                        return [string]$value
                                    }

                                    if( $null -eq $status )
                                    {
                                        return
                                    }

                                    [string]$messageText = $null
                                    if( $status.PSObject.Properties.Match( 'Message' ).Count -gt 0 )
                                    {
                                        $messageText = & $convertValueToText $status.Message
                                    }
                                    elseif( $status.PSObject.Properties.Match( 'Output' ).Count -gt 0 )
                                    {
                                        $messageText = & $convertValueToText $status.Output
                                    }

                                    if( [string]::IsNullOrWhiteSpace( $messageText ) )
                                    {
                                        return
                                    }

                                    try
                                    {
                                        $jsonMessage = $messageText | ConvertFrom-Json -ErrorAction Stop
                                        if( $null -ne $jsonMessage )
                                        {
                                            [string]$stdoutText = $null
                                            if( $jsonMessage.PSObject.Properties.Match( 'stdout' ).Count -gt 0 )
                                            {
                                                $stdoutText = & $convertValueToText $jsonMessage.stdout
                                            }
                                            if( -Not [string]::IsNullOrWhiteSpace( $stdoutText ) )
                                            {
                                                & $addOutputPart ( "[stdout]`r`n$stdoutText" )
                                            }

                                            [string]$stderrText = $null
                                            if( $jsonMessage.PSObject.Properties.Match( 'stderr' ).Count -gt 0 )
                                            {
                                                $stderrText = & $convertValueToText $jsonMessage.stderr
                                            }
                                            if( -Not [string]::IsNullOrWhiteSpace( $stderrText ) )
                                            {
                                                & $addOutputPart ( "[stderr]`r`n$stderrText" )
                                            }

                                            [string]$summaryMessageText = $null
                                            if( $jsonMessage.PSObject.Properties.Match( 'message' ).Count -gt 0 )
                                            {
                                                $summaryMessageText = & $convertValueToText $jsonMessage.message
                                            }
                                            if( -Not [string]::IsNullOrWhiteSpace( $summaryMessageText ) )
                                            {
                                                & $addOutputPart ( "[message]`r`n$summaryMessageText" )
                                            }
                                            return
                                        }
                                    }
                                    catch
                                    {
                                        ## not JSON output, use as-is
                                    }

                                    & $addOutputPart $messageText
                                }

                                ForEach( $status in @( $result.Value ) )
                                {
                                    & $extractStatusMessage $status
                                }
                                ForEach( $status in @( $result.Output ) )
                                {
                                    & $extractStatusMessage $status
                                }

                                [string]$resultText = ( $outputParts -join "`r`n`r`n" ).Trim()
                                if( [string]::IsNullOrWhiteSpace( $resultText ) )
                                {
                                    $resultText = ( $result | ConvertTo-Json -Depth 12 )
                                }

                                [pscustomobject]@{
                                    Name = $vmName
                                    OutputText = $resultText
                                    Result = $result
                                }
                            }

                            if( $null -ne $job )
                            {
                                $job | Add-Member -NotePropertyName ContextPath -NotePropertyValue $contextPath -Force
                                $job
                                continue
                            }
                        }
                        catch
                        {
                            if( Test-Path -Path $contextPath )
                            {
                                Remove-Item -Path $contextPath -Force -ErrorAction SilentlyContinue
                            }
                            throw
                        }
                    }
                    catch
                    {
                        Write-Warning -Message "Azure run command failed to start for $($selection.Name) : $_"
                        [void][Windows.MessageBox]::Show( $mainWindow , "Azure run command failed to start for $($selection.Name)`n$($_.Exception.Message)" , 'Azure Run Command' , 'Ok' ,'Error' )
                    }
                }
                elseif( $operation -imatch '^Azure_.*Session$' )
                {
                    try
                    {
                        $sessionHost = $null
                        $sessionHost = Get-AzureAVDSessionHost -selection $selection

                        if( $operation -ieq 'Azure_DrainModeOnSession' -or $operation -ieq 'Azure_DrainModeOffSession' )
                        {
                            [bool]$allowNewSession = $operation -ieq 'Azure_DrainModeOffSession'
                            [string]$payload = (@{ properties = @{ allowNewSession = $allowNewSession } } | ConvertTo-Json -Compress)
                            $null = Invoke-AzRestMethod -Method PATCH -Path "$($sessionHost.Id)?api-version=$AVDAPIversion" -Payload $payload -ErrorAction Stop
                            $refreshAzureList = $true
                            continue
                        }


                        [array]$userSessions = @( Get-AzureAVDUserSessions -sessionHostResourceId $sessionHost.Id )
                        if( $null -eq $userSessions -or $userSessions.Count -eq 0 )
                        {
                            [void][Windows.MessageBox]::Show( $mainWindow , "No active user sessions found on $($selection.Name)" , 'Azure AVD Sessions' , 'Ok' ,'Information' )
                            continue
                        }
                        if( $operation -ieq 'Azure_MessageSession' )
                        {
                            if( $null -eq $azureSessionMessage )
                            {
                                $azureSessionMessage = Get-AzureAVDSessionMessage
                                if( $null -eq $azureSessionMessage )
                                {
                                    break
                                }
                            }
                        }

                        ForEach( $userSession in $userSessions )
                        {
                            if( $operation -ieq 'Azure_DisconnectSession' )
                            {
                                $null = Invoke-AzRestMethod -Method POST -Path "$($userSession.Id)/disconnect?api-version=$AVDAPIversion" -ErrorAction Stop
                                $refreshAzureList = $true
                            }
                            elseif( $operation -ieq 'Azure_LogoffSession' )
                            {
                                $null = Invoke-AzRestMethod -Method DELETE -Path "$($userSession.Id)?api-version=$AVDAPIversion" -ErrorAction Stop
                                $refreshAzureList = $true
                            }
                            elseif( $operation -ieq 'Azure_ForceLogoffSession' )
                            {
                                $null = Invoke-AzRestMethod -Method DELETE -Path "$($userSession.Id)?api-version=$AVDAPIversion&force=true" -ErrorAction Stop
                                $refreshAzureList = $true
                            }
                            elseif( $operation -ieq 'Azure_MessageSession' )
                            {
                                [string]$payload = (@{ messageTitle = $azureSessionMessage.Title ; messageBody = $azureSessionMessage.Body } | ConvertTo-Json -Compress)
                                $null = Invoke-AzRestMethod -Method POST -Path "$($userSession.Id)/sendMessage?api-version=$AVDAPIversion" -Payload $payload -ErrorAction Stop
                            }
                        }
                    }
                    catch
                    {
                        Write-Warning -Message "Azure AVD session operation failed for $($selection.Name) : $_"
                        [void][Windows.MessageBox]::Show( $mainWindow , "Azure AVD session operation failed for $($selection.Name)`n$($_.Exception.Message)" , 'Azure AVD Sessions' , 'Ok' ,'Error' )
                    }
                }
                elseif( $operation -ieq 'HyperV_PowerOn' )
                {
                    Hyper-V\Start-VM -VMName $selection.Name @hypervParameters @async
                }
                elseif( $operation -ieq 'HyperV_Shutdown' )
                {
                    Hyper-V\Stop-VM -VMName $selection.Name  @hypervParameters @async
                }
                elseif( $operation -ieq 'HyperV_PowerOff' )
                {
                    Hyper-V\Stop-VM -VMName $selection.Name -TurnOff @hypervParameters -Force -Confirm:$false @async
                }
                elseif( $operation -ieq 'HyperV_Restart' )
                {
                    Hyper-V\Restart-VM -VMName $selection.Name  @hypervParameters -Force -Confirm:$false @async
                }
                elseif( $operation -ieq 'HyperV_Resume' )
                {
                    Hyper-V\Resume-VM -VMName $selection.Name  @hypervParameters -Asjob @async
                }
                elseif( $operation -ieq 'HyperV_Suspend' )
                {
                    Hyper-V\Suspend-VM -VMName $selection.Name @hypervParameters @async
                }
                elseif( $operation -ieq 'HyperV_RevertLatestSnapshot' -or $operation -ieq 'HyperV_DeleteLatestSnapshot')
                {
                    $latestCheckPoint = $null
                    $latestCheckPoint = Get-VMCheckpoint -VMName $selection.Name @hypervParameters | Sort-Object -Property CreationTime -Descending | Select-Object -First 1
                    if( $null -ne $latestCheckPoint )
                    {
                        if( [Windows.MessageBox]::Show( $mainWindow , "$($latestCheckPoint.CreationTime.ToString('g')) : $($latestCheckPoint.Name)",
                            "$(if( $operation -match 'delete' ) { 'Delete' } else { 'Restore' }) Snapshot on $($selection.Name)" , 'YesNoCancel' ,'Question' ) -ieq 'yes' )
                        {
                            ## do not need hypervparameters as cannot user -computername with a snapshot object as that contains the remote details
                            if( $operation -match 'delete' )
                            {
                                Hyper-V\Remove-VMCheckpoint  -VMSnapshot $latestCheckPoint -Confirm:$false -Passthru
                            }
                            else
                            {
                                Hyper-V\Restore-VMCheckpoint -VMSnapshot $latestCheckPoint -Confirm:$false -Passthr
                            }
                        }
                    }
                    else
                    {
                        Write-Warning -Message "No snapshots found for $($selection.Name)"
                        ## TODO No snapshots message
                    }
                }         
                elseif( $operation -ieq 'HyperV_TakeSnapshot' )
                {        
                    if( $textInputWindow = New-WPFWindow -inputXAML $textInputXAML )
                    {
                        $WPFbtnInputTextOk.Add_Click({ 
                            $_.Handled = $true
                            $textInputWindow.DialogResult = $true 
                            $textInputWindow.Close()  })
                        $textInputWindow.Title = "New Snapshot"
                        $WPFlblInputTextLabel.Content = "Enter Snapshot Name"
                        if( $vm -and $vm.State -ieq 'Running' )
                        {
                            $response = [Windows.MessageBox]::Show( $mainWindow , "Uptime $($vm.Uptime)", "Shutdown $($vm.Name) first" , 'YesNoCancel' ,'Question' )
                            if( $response -ieq 'Cancel' )
                            {
                                return
                            }
                            elseif( $response -ieq 'yes' )
                            {
                                Hyper-V\Stop-VM -VM $VM -TurnOff ## synchronous
                            }
                        }
                        if( $textInputWindow.ShowDialog() )
                        {
                            [hashtable]$snapshotParameters = @{}
                            if( $WPFtextboxInputText.Text.Length )
                            {
                                $snapshotParameters.Add( 'SnapshotName' , $WPFtextboxInputText.Text.Trim() )
                            }
                            
                            Hyper-V\Checkpoint-VM -VMName $selection.Name -Passthru @hypervParameters @snapshotParameters
                            $checkpointStatus = $?
                            $response = [Windows.MessageBox]::Show( $mainWindow , "Snapshot $(if( $checkpointStatus ) { 'succeeded' } else { 'failed' })", "Start $($vm.Name) after restore" , 'YesNo' ,'Question' )
                            if( $response -ieq 'yes' )
                            {
                                Hyper-V\Start-VM -VMName $selection.Name @hypervParameters @async
                            }
                        }
                    }
                }
                elseif( $operation -ieq 'HyperV_Rename' )
                {        
                    if( $textInputWindow = New-WPFWindow -inputXAML $textInputXAML )
                    {
                        $WPFbtnInputTextOk.Add_Click({ 
                            $_.Handled = $true
                            $textInputWindow.DialogResult = $true 
                            $textInputWindow.Close()  })
                        $textInputWindow.Title = "Rename $($selection.Name)"
                        $WPFlblInputTextLabel.Content = "Enter New VM Name"
                        if( $textInputWindow.ShowDialog() )
                        {
                            [string]$newname = $WPFtextboxInputText.Text.Trim().Trim('"')
                            if( [string]::IsNullOrEmpty( $newname ) )
                            {
                                Write-Error "New name `"$newname`" too short"
                            }
                            else
                            {
                                if( $newname -ieq $selection.Name )
                                {
                                    Write-Error "New name $newname is the same"
                                }
                                elseif( $null -ne ($existingVM = Hyper-V\Get-VM -Name $newname -ErrorAction SilentlyContinue @hypervParameters ) )
                                {
                                    Write-Error "VM $newname already exists"
                                }
                                else
                                {
                                    Hyper-V\Rename-VM -VM $vm -Passthru -NewName $newname
                                }
                            }
                        }
                    }
                }
                elseif( $operation -ieq 'HyperV_ManageSnapshot' )
                {
                    Show-SnapShotWindow -vm $selection.Name
                }
                elseif( $operation -ieq 'HyperV_Save' )
                {
                    Hyper-V\Save-VM -VMName $selection.Name @hypervParameters @async
                }
                elseif( $operation -ieq 'HyperV_Delete' -or $operation -ieq 'HyperV_DeleteIncludingDisks' )
                {
                    $disks = $null
                    if( $operation -ieq 'HyperV_DeleteIncludingDisks' )
                    {
                        $disks = @( Hyper-V\Get-VMHardDiskDrive -VMName $selection.Name @hypervParameters )
                        Write-Verbose -Message "Got $($disks.Count) disks for VM $($disks.VMName)"
                    }
                    $removal = $null
                    $removal = Hyper-V\Remove-VM -VMName $selection.Name -Passthru -Force -Confirm:$false @hypervParameters
                    if( $? -and $null -ne $removal )
                    {
                        ForEach( $disk in $disks )
                        {
                            Write-Verbose -Message "Deleting disk $($disk.Path)"
                            ## Could be remote so we use WMI with the CIM session in the disks object
                            $file = $null
                            ## needs backslashes escaping
                            $file = Get-CimInstance -ClassName cim_datafile -Filter "Name = '$($disk.Path -replace '\\' , '\\')'" -CimSession $disk.CimSession
                            if( $null -ne $file )
                            {
                                Remove-CimInstance -InputObject $file -CimSession $disk.CimSession -Confirm:$false
                            }
                            else
                            {
                                Write-Warning -Message "Failed to get file for disk $($disk.Path)"
                            }
                        }
                    }
                    else
                    {
                        Write-Verbose -Message "Not deleting disks for $($selection.Name) as deleting VM errored"
                    }
                }
                elseif( $operation -ieq 'HyperV_Detail' )
                {
                    $details = $null
                    $details = Hyper-V\Get-VM -VMName $selection.Name @hypervParameters
                    if( $null -ne $details )
                    {
                        [array]$hardDrives = @( Get-VMHardDiskDrive -VM $details )
                        [string[]]$diskDetails = @( ForEach( $disk in $hardDrives )
                        {
                            [string]$size = ''
                            $file = $null
                            $file = Get-CimInstance -ClassName cim_datafile -Filter "Name = '$($disk.Path -replace '\\' , '\\')'" -CimSession $disk.CimSession
                            if( $null -ne $file )
                            {
                                $size = "$([math]::Round( $file.filesize / 1GB , 1 ))"
                            }
                            "$($disk.Path) ($($size) GB)"
                        })
                        [array]$snapshots = @( Get-VMSnapshot -VM $details | Sort-Object -Property CreationTime )
                        [array]$NICs = @( Get-VMNetworkAdapter -VM $details )
                        $form = New-Object System.Windows.Forms.Form
                        $form.Text = $selection.Name
                        $form.Size = New-Object System.Drawing.Size(800, 400)
                        $form.StartPosition = "CenterScreen"

                        $listView = New-Object System.Windows.Forms.ListView
                        $listView.View = 'Details'
                        $listView.FullRowSelect = $true
                        $listView.GridLines = $true
                        $listView.Dock = 'Fill'
                        $listView.Columns.Add("Setting", 180)
                        $listView.Columns.Add("Value", 600 )

                        $data = @(
                            @{ Setting = "Notes"; Value = $details.Notes }
                            @{ Setting = "State"; Value = $details.State.ToString() }
                            @{ Setting = "vCPU"; Value = $details.ProcessorCount }
                            @{ Setting = "Resource Metering Enabled"; Value = $details.ResourceMeteringEnabled.ToString() }
                            @{ Setting = "Uptime"; Value = $details.Uptime.ToString() }
                            @{ Setting = "Version"; Value = $details.Version.ToString() }
                            @{ Setting = "Memory Startup MB"; Value = $details.MemoryStartup / 1MB }
                            @{ Setting = "Memory Assigned MB"; Value = $details.MemoryAssigned / 1MB }
                            @{ Setting = "Memory Minimum MB"; Value = $details.MemoryMinimum / 1MB }
                            @{ Setting = "Memory Maximum MB"; Value = $details.MemoryMaximum / 1MB }
                            @{ Setting = "Dynamic Memory Enabled"; Value = $details.DynamicMemoryEnabled.ToString() }
                            @{ Setting = "Hard Drives"; Value = $diskDetails -join "`n" }
                            @{ Setting = "NICs"; Value = $NICs.Count }                        
                            @{ Setting = "IP Addresses"; Value = ( $NICs | Select-Object -ExpandProperty IPAddresses -ErrorAction SilentlyContinue | Where-Object { $_ -notmatch ':' } ) -join ' , ' }
                            @{ Setting = "DNS Resolved"; Value = ( Resolve-DnsName -Name $details.Name -Type A | Select-Object -ExpandProperty IPAddress -ErrorAction SilentlyContinue | Where-Object { $_ -notmatch ':' } ) -join ' , ' }
                            @{ Setting = "Snapshots"; Value = $snapshots.Count }
                            @{ Setting = "Created"; Value = $details.CreationTime.ToString('G') }
                            @{ Setting = "Oldest Snapshot"; Value = $(if( $null -ne $snapshots -and $snapshots.Count -gt 0 ) { "$($snapshots[0].CreationTime.ToString('G')) ($($snapshots[0].Name) ($($snapshots[0].Notes)))" })}
                            @{ Setting = "Latest Snapshot"; Value = $(if( $null -ne $snapshots -and $snapshots.Count -gt 0 ) { "$($snapshots[-1].CreationTime.ToString('G')) ($($snapshots[-1].Name) ($($snapshots[-1].Notes)))" })}  
                        )

                        foreach ($item in $data)
                        {
                        try {
                            $listItem = New-Object System.Windows.Forms.ListViewItem($item.Setting)
                            $listItem.SubItems.Add($item.Value)
                            $null = $listView.Items.Add($listItem)
                        } catch {
                            Write-Warning "Exception adding $($item.value) to list view: $_"
                        }
                        }

                        $form.Controls.Add($listView)
                        $form.Add_Shown({ $form.Activate() })
                        [void]$form.ShowDialog()
                    }
                    else
                    {
                        ## TODO error dialogue
                    }
                }
                elseif( $Operation -ieq 'HyperV_EnableNestedVirtualisation' )
                {
                    if( ( $vm = Hyper-V\Get-VM -VMName $selection.Name @hypervParameters) -and $vm.State -ieq 'Running' )
                    {
                        $response = [Windows.MessageBox]::Show( $mainWindow , "VM cannot be running`nUptime $($vm.Uptime)", "Shutdown $($vm.Name) first" , 'YesNoCancel' ,'Question' )
                        if( $response -ieq 'Cancel' )
                        {
                            return
                        }
                        elseif( $response -ieq 'yes' )
                        {
                            Hyper-V\Stop-VM -VM $VM -TurnOff ## syncrhonous
                        }
                    }
                    Set-VMProcessor -VM $vm -ExposeVirtualizationExtensions $true
                    $configChangeStatus = $?
                    $response = [Windows.MessageBox]::Show( $mainWindow , "Change $(if( $configChangeStatus ) { 'succeeded' } else { 'failed' })", "Start $($vm.Name)" , 'YesNo' ,'Question' )
                    if( $response -ieq 'yes' )
                    {
                        Hyper-V\Start-VM -VM $VM @hypervParameters @async
                    }
                }
                elseif( $operation -ieq 'HyperV_EnableResourceMetering' )
                {
                    Hyper-V\Enable-VMResourceMetering -VMName $selection.Name @hypervParameters
                }
                elseif( $operation -ieq 'HyperV_DisableResourceMetering' )
                {
                    Hyper-V\Disable-VMResourceMetering -VMName $selection.Name @hypervParameters
                }
                elseif( $operation -ieq 'HyperV_DisConnectNIC' )
                {
                    ## TODO what if more than one NIC?
                    Hyper-V\Disconnect-VMNetworkAdapter -VMName $selection.Name @hypervParameters -Passthru
                }
                elseif( $operation -imatch 'HyperV_ConnectNIC*' )
                {
                    [string]$switchType = $Operation -replace '^HyperV_ConnectNIC'
                    ## TODO need to get virtual switch names and if more than one prompt for the required one
                    [array]$switches = @( Hyper-V\Get-VMSwitch -SwitchType $switchType @hypervParameters )
                    if( $switches.Count -gt 1 )
                    {
                        Write-Warning "VM $($selecion.Name) has $($switches.Count) NICs which isn't yet implemented sorry"
                    }
                    Hyper-V\Connect-VMNetworkAdapter -VMName $selection.Name -SwitchName $switches[ 0 ].Name @hypervParameters @async
                }
                elseif( $operation -ieq 'HyperV_RunOn' )
                {
                    ## TODO bring output processing code for this from VMware GUI
                    if( -Not [string]::IsNullOrEmpty( $WPFtextBoxHyperVHost.Text ) -and $WPFtextBoxHyperVHost -ine 'localhost' )
                    {
                        if( $null -eq $script:remoteSession )
                        {
                            $remoteError = $null
                            $script:remoteSession = New-PSSession -ComputerName $WPFtextBoxHyperVHost.Text -ErrorVariable remoteError
                            if( $null -eq $script:remoteSession )
                            {
                                $null = [Windows.MessageBox]::Show( $mainWindow , "Failed to remote to $($WPFtextBoxHyperVHost.Text)`n$remoteError" , 'Ok' ,'Error' )
                                return
                            }
                        }
                    }
                    if( $null -eq $script:credentials )
                    {
                        $script:credentials = Get-Credential -Message "Credentials for running in VM"
                        if( $null -eq $script:credentials )
                        {
                            return
                        }
                    }
                    ##TODO  get command and parameters
                    $commandError = $null
                    [string]$vmName = $selection.Name
                    [scriptblock]$commandToRun = { Invoke-Command -VMName $using:vmname -Credential $using:credentials -ScriptBlock { hostname.exe }}
                    [hashtable]$remoteParameters = @{}
                    if( $null -ne $script:remoteSession )
                    {
                        $remoteParameters.Add( 'Session' , $script:remoteSession )
                    }
                    $job = Invoke-Command @remoteParameters -ScriptBlock $commandToRun -ErrorVariable commandError ## -AsJob
                    $status = $?
                    
                    ## TODO spawn commands async and then gather results later

                    ## need to figure how we detect credentials didn't work so we clear them so get prompted again
                }
                else
                {
                    Write-Warning -Message "Unimplemented operation $Operation"
                }
                $clipboardParameters[ 'Append' ] = $true
            })

            if( $operation -ieq 'Azure_RunOn' )
            {
                [array]$azureRunJobs = @( $jobs | Where-Object { $null -ne $_ } )
                if( $azureRunJobs.Count -gt 0 )
                {
                    try
                    {
                        [int]$effectiveAzureRunJobTimeoutSeconds = $azureRunJobTimeoutSeconds
                        [hashtable]$jobStartTimes = @{}
                        [System.Collections.ArrayList]$pendingAzureRunJobs = [System.Collections.ArrayList]::new()
                        ForEach( $job in $azureRunJobs )
                        {
                            $null = $pendingAzureRunJobs.Add( $job )
                            $jobStartTimes[ $job.Id ] = Get-Date
                        }

                        while( $pendingAzureRunJobs.Count -gt 0 )
                        {
                            $completedJob = Wait-Job -Job @( $pendingAzureRunJobs ) -Any -Timeout 2
                            if( $null -ne $completedJob )
                            {
                                ForEach( $job in @( $completedJob ) )
                                {
                                    try
                                    {
                                        [array]$jobErrors = @()
                                        [array]$receivedOutput = @( Receive-Job -Job $job -Keep -ErrorVariable +jobErrors )
                                        $jobResult = $receivedOutput | Where-Object { $_ -is [pscustomobject] -and $_.PSObject.Properties.Match( 'Result' ).Count -gt 0 } | Select-Object -First 1

                                        [string]$vmName = $job.Name -replace '^Azure_Run_' , ''
                                        if( $null -ne $jobResult -and $jobResult.PSObject.Properties.Match( 'Name' ).Count -gt 0 -and -Not [string]::IsNullOrWhiteSpace( $jobResult.Name ) )
                                        {
                                            $vmName = $jobResult.Name
                                        }

                                        [string]$outputText = $null
                                        if( $jobErrors.Count -gt 0 )
                                        {
                                            $outputText = ( $jobErrors | ForEach-Object { $_.ToString() } ) -join "`r`n"
                                        }
                                        elseif( $null -ne $jobResult )
                                        {
                                            if( $jobResult.PSObject.Properties.Match( 'OutputText' ).Count -gt 0 -and -Not [string]::IsNullOrWhiteSpace( $jobResult.OutputText ) )
                                            {
                                                $outputText = [string]$jobResult.OutputText
                                            }
                                            else
                                            {
                                                $outputText = Convert-AzureRunCommandResultToText -runCommandResult $jobResult.Result
                                            }
                                        }
                                        else
                                        {
                                            $outputText = 'No output was returned by the Azure run command job.'
                                        }

                                        Show-AzureRunCommandOutputWindow -computerName $vmName -scriptText $azureRunCommandText -outputText $outputText
                                    }
                                    catch
                                    {
                                        [void][Windows.MessageBox]::Show( $mainWindow , "Azure run command failed for $($job.Name -replace '^Azure_Run_' , '')`n$($_.Exception.Message)" , 'Azure Run Command' , 'Ok' ,'Error' )
                                    }
                                    finally
                                    {
                                        if( $job.PSObject.Properties.Match( 'ContextPath' ).Count -gt 0 -and ( Test-Path -Path $job.ContextPath ) )
                                        {
                                            Remove-Item -Path $job.ContextPath -Force -ErrorAction SilentlyContinue
                                        }
                                        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
                                        [void]$pendingAzureRunJobs.Remove( $job )
                                    }
                                }
                            }

                            [datetime]$now = Get-Date
                            ForEach( $pendingJob in @( $pendingAzureRunJobs ) )
                            {
                                [timespan]$elapsed = $now - $jobStartTimes[ $pendingJob.Id ]
                                if( $elapsed.TotalSeconds -ge $effectiveAzureRunJobTimeoutSeconds )
                                {
                                    try
                                    {
                                        Stop-Job -Job $pendingJob -ErrorAction SilentlyContinue | Out-Null
                                        [void][Windows.MessageBox]::Show( $mainWindow , "Azure run command timed out for $($pendingJob.Name -replace '^Azure_Run_' , '') after $effectiveAzureRunJobTimeoutSeconds seconds." , 'Azure Run Command' , 'Ok' ,'Error' )
                                    }
                                    finally
                                    {
                                        if( $pendingJob.PSObject.Properties.Match( 'ContextPath' ).Count -gt 0 -and ( Test-Path -Path $pendingJob.ContextPath ) )
                                        {
                                            Remove-Item -Path $pendingJob.ContextPath -Force -ErrorAction SilentlyContinue
                                        }
                                        Remove-Job -Job $pendingJob -Force -ErrorAction SilentlyContinue
                                        [void]$pendingAzureRunJobs.Remove( $pendingJob )
                                    }
                                }
                            }
                        }
                    }
                    finally
                    {
                        Get-Job | Where-Object { $_.Name -like 'Azure_Run_*' } | Remove-Job -Force -ErrorAction SilentlyContinue
                    }
                }

                return
            }

            if( $null -ne $jobs -and $jobs.count -gt 0 )
            {
                $jobs | Write-Verbose
            }
            if( $refreshAzureList -and $GUIobject -eq $WPFlistViewAzureVMs )
            {
                Add-AzureVMsToListView -filter '' -regex $false -allVMs $WPFcheckBoxAzureAllVMs.IsChecked
            }
        }
        catch
        {
            Write-Error "$_" ## nothing here should be fatal
        }
        finally
        {
            [System.Windows.Input.Mouse]::OverrideCursor = $null
            $mainWindow.ClearValue([System.Windows.FrameworkElement]::CursorProperty) 
        }
    }
    else
    {
        Write-Warning "No GUI object passed to Process-Action for action $operation"
    }
}

Function Search-AD
{
    Param
    (
        [string]$domainController ,
        [string]$searchFor ,
        [string]$searchType
    )

    if( $searchType -ieq 'OU' )
    {
        $searchTerm = "(&(objectClass=organizationalUnit)(ou=$searchFor))"
    }
    elseif( $searchType -ieq 'group' )
    {
        $searchTerm = "(&(objectClass=group)(name=$searchFor))"
    }
    elseif( $searchType -ieq 'computer' )
    {
        $searchTerm = "(&(objectClass=computer)(name=$searchFor))"
    }
    else
    {
        Write-Error "Unexpected search type $searchType"
        return
    }

    Write-Verbose -Message "Search term $searchTerm"

    $objects = $null
    $objects = (New-Object System.DirectoryServices.DirectorySearcher $searchTerm).FindAll()
    if( $null -eq $objects -or $objects.Count -eq 0 )
    {
        [void][Windows.MessageBox]::Show( "Found no AD $searchType for `"$searchFor`"" , 'AD Searcher' , 'Ok' ,'Information' )
        return
    }
    Write-Verbose "Got $($objects.Count) $searchType for $searchFor"

    [System.Collections.Generic.List[object]]$items = @() 

    ## if OU, get computers in the matched OU(s)
    ## if group, get computers in the matched group(s)
    if( $searchType -ieq 'group' )
    {
        ## TODO but these into a chooser rather than include all
        ForEach( $group in $objects )
        {
            $groupEntry = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList $group.Path ## don't cache as could change
            $members = $groupEntry.Properties["member"]
            ForEach( $member in $members )
            {
                $entry = New-Object -TypeName System.DirectoryServices.DirectoryEntry -argumentList "LDAP://$member" ## TODO could cache this
    
                if ( $entry -and $entry.SchemaClassName -eq 'computer')
                {
                    $object = [pscustomobject]@{ Name = $entry.Properties['name'][0] ; Container = $group.Properties['name'][0] }
                    if( $items -notcontains $object )
                    {
                        $items.Add( $object )
                    }
                }
                elseif( $entry -and $entry.SchemaClassName -eq 'group' -and $WPFcheckBoxADRecurse.IsChecked )
                {
                    ## TODO recurse this group - have a mechanism to prevent infinite recursion
                }
            }
        }
    }
    elseif( $searchType -ieq 'OU' )
    {
        ## TODO but these into a chooser rather than include all
        ForEach( $OU in $objects )
        {
            $searcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher
            $searcher.SearchRoot = $OU.Path
            if( $WPFcheckBoxADRecurse.IsChecked )
            {
                $searcher.SearchScope = [System.DirectoryServices.SearchScope]::Subtree
            }
            else
            {
                $searcher.SearchScope = [System.DirectoryServices.SearchScope]::OneLevel
            }
            $searcher.Filter = "(objectCategory=computer)"
            [void]$searcher.PropertiesToLoad.AddRange( @("name","distinguishedName","operatingSystem"))
            ForEach( $entry in $searcher.FindAll() )
            {
                $object = [pscustomobject]@{ Name = $entry.Properties['name'][0] ; Container = $OU.Properties['distinguishedname'][0] }
                if( $items -notcontains $object )
                {
                    $items.Add( $object )
                }
            }
        }
    }
    else ## computers
    {
        ForEach( $computer in $objects )
        {
            $items.Add( [pscustomobject]@{ Name = $computer.Properties['name'][0] ; Container = $computer.Properties['distinguishedname'][0] -replace '^CN=[^,]+,' } )
        }
    }

    Write-Verbose -Message "Got $($items.Count) items"

    $WPFlistViewAD.Items.Clear()
    if( $items.Count -gt 0 )
    {
        ForEach( $item in $items )
        {
            $WPFlistViewAD.Items.Add( $item ) ## value comes from what is in Binding property for the grid view column
        }
    }
    $WPFlabelADComputers.Content = "$($WPFlistViewAD.Items.Count) computers"
 

    $null = 42
}

#endregion pre-main


if( [string]::IsNullOrEmpty( $tempFolder ) )
{
    Throw "No temp folder"
}
if( -Not ( Test-Path -Path $tempFolder -PathType Container ) -and -Not ( New-Item -Path $tempFolder -ItemType Directory -Force ) )
{
    Throw "Failed to create temp folder $tempFolder"
}
         
<#
if( $usemsrdc -and [string]::IsNullOrEmpty( $address ) )
{
    Throw "Must specify computer to connect to via -address when using msrdc mode"
}
#>

try
{
    Add-Type -TypeDefinition $pinvokeCode
}
catch
{
    ## hopefully because already loaded
}

Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase,System.Windows.Forms,System.Drawing

$script:activeDisplaysWithMonitors = @( Get-DisplayInfo )

if( $showDisplays )
{
    $activeDisplaysWithMonitors
    exit 0
}
if( $showManufacturerCodes )
{
    $ManufacturerHash.GetEnumerator() | Select-Object -Property @{n='Manufacturer';e={$_.Value}},@{n='Code';e={$_.Name}} | Sort-Object -Property Manufacturer
    exit 0
}

## don't need it yet but need it before we start the GUI

[string]$windowTypes = @'
    using System;
    using System.Runtime.InteropServices;
    
    [StructLayout(LayoutKind.Sequential)]

    public struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    public static class user32
    {
        [DllImport("user32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow); 
            
        [DllImport("user32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow); 
            
        [DllImport("user32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsIconic(IntPtr hWnd); 
        
        [DllImport("user32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsZoomed(IntPtr hWnd); 
        
        [DllImport("user32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindowVisible(IntPtr hWnd); 
        
        [DllImport("user32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindowUnicode(IntPtr hWnd); 
        
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
   
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetWindowRect(IntPtr hWnd, string lpString);
   
        [DllImport("user32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetWindowPos( IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    }
'@ 

try
{
    Add-Type -TypeDefinition $windowTypes
}
catch
{
    ## hopefully because we already have it
}

## https://www.linkedin.com/pulse/fun-powershell-finding-suspicious-cmd-processes-britton-manahan/

$TypeDef = @"

using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Api
{

 public class WinStruct
 {
   public string WinTitle {get; set; }
   public int WinHwnd { get; set; }
   public int PID { get; set; }
 }

 public class ApiDef
 {
   private delegate bool CallBackPtr(int hwnd, int lParam);
   private static CallBackPtr callBackPtr = Callback;
   private static List<WinStruct> _WinStructList = new List<WinStruct>();

   [DllImport("User32.dll")]
   [return: MarshalAs(UnmanagedType.Bool)]
   private static extern bool EnumWindows(CallBackPtr lpEnumFunc, IntPtr lParam);

   [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
   static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
   
   [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
   public static extern int SetWindowText(IntPtr hWnd, string lpString );
   
   [DllImport("user32.dll")]
   static extern bool IsWindowVisible(IntPtr hWnd);
   
   [DllImport("user32.dll")]
   public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
 
   private static bool Callback(int hWnd, int pid)
   {
        if( IsWindowVisible( (IntPtr)hWnd ) )
        {
            int ipid = 0 ;
            GetWindowThreadProcessId( (IntPtr)hWnd , out ipid );
            if( ipid == pid || pid < 0 ) // -1 will return all windows
            {
                StringBuilder sb = new StringBuilder(256);
                int res = GetWindowText((IntPtr)hWnd, sb, 256);
                _WinStructList.Add( new WinStruct { WinHwnd = hWnd, WinTitle = sb.ToString() , PID = ipid  });
            }
        }
        return true;
   }   

   public static List<WinStruct> GetWindows( int pid )
   {
      _WinStructList = new List<WinStruct>();
      EnumWindows(callBackPtr, (IntPtr)pid );
      return _WinStructList;
   }

 }
}
"@

try
{
    Add-Type -TypeDefinition $TypeDef -ErrorAction Stop
}
catch
{
    ## hopefully because we already have it
}

if( $primary )
{
    $chosenDisplay = $activeDisplaysWithMonitors | Where-Object ScreenPrimary
}
elseif( $PSBoundParameters.ContainsKey( 'displayModel' ) )
{
    if( -Not ( $chosenDisplay = $activeDisplaysWithMonitors | Where-Object MonitorModel -IMatch $displayModel ) )
    {
        Throw "No displays for model `"$displayModel`" out of $(($activeDisplaysWithMonitors | Select-Object -ExpandProperty MonitorModel) -join ',')"
    }
    elseif( $chosenDisplay -is [array] -and $chosenDisplay.Count -gt 1 )
    {
        Throw "Multiple monitors for model `"$displayModel`""
    }
}
elseif( $PSBoundParameters[ 'displayManufacturer' ] )
{
    [string]$displayManufacturerCode = 'NONE'

    $displayManufacturerCodes = $ManufacturerHash.GetEnumerator() | Where-Object Value -match $displayManufacturer | Select-Object -ExpandProperty Name
    if( -Not $displayManufacturerCodes )
    {
        Throw "No monitor manufacturer found matching `"$displayManufacturer`""
    }
    elseif( $displayManufacturerCodes -is [array] -and $displayManufacturerCodes.Count -gt 1 )
    {
        Throw "Found ($displayManufacturerCodes.Count) manufacturer codes matching `"$displayManufacturer`" - use code instead - $($displayManufacturerCodes -join ' or ')"
    }
    else
    {
        $displayManufacturerCode = $displayManufacturerCodesG
    }
  
    if( -Not ( $chosenDisplay = $activeDisplaysWithMonitors | Where-Object MonitorManufacturerName -ieq $displayManufacturerCode ) )
    {
        Throw "No displays for manufacturer code `"$displayManufacturerCode`" ($displayManufacturer) out of $(($activeDisplaysWithMonitors | Select-Object -ExpandProperty MonitorManufacturerName) -join ',')"
    }
    elseif( $chosenDisplay -is [array] -and $chosenDisplay.Count -gt 1 )
    {
        Throw "Multiple monitors for manufacturer code `"$displayManufacturerCode`" ($displayManufacturer) - try model number?"
    }
}
elseif( $PSBoundParameters[ 'displayManufacturerCode' ] )
{
    if( -Not ( $chosenDisplay = $activeDisplaysWithMonitors | Where-Object MonitorManufacturerName -ieq $displayManufacturerCode ) )
    {
        Throw "No displays for manufacturer code `"$displayManufacturerCode`" out of $(($activeDisplaysWithMonitors | Select-Object -ExpandProperty MonitorManufacturerName) -join ',')"
    }
    elseif( $chosenDisplay -is [array] -and $chosenDisplay.Count -gt 1 )
    {
        Throw "Multiple monitors for manufacturer code `"$displayManufacturerCode`" - try model number?"
    }
}
else ## if not passed displayNumber or displaymanufacturer , display a GUI with the choices
{
    [array]$displayFields = @( 'ScreenPrimary','ScreenDeviceName',@{n='Width';e={$_.dmPelswidth}},@{n='Height';e={$_.dmPelsHeight}},'MonitorManufacturerName','MonitorManufacturerCode','MonitorModel' )
    if( $usegridviewpicker )
    {
        if( -Not ( $chosen = $activeDisplaysWithMonitors | Select-Object -Property $displayFields | Out-GridView -Title "Select monitor for $exe" -PassThru ) )
        {
            Throw "Please select a monitor"
        }
    }
    else
    {
        if( -Not ( $mainWindow = New-WPFWindow -inputXAML $mainwindowXAML ) )
        {
            Throw 'Failed to create WPF from XAML'
        }
        
        Set-WindowContent
        
        $wpfbtnRefresh.Add_Click({
            $_.Handled = $true
            Write-Verbose "Refresh clicked"
            $script:activeDisplaysWithMonitors = @( Get-DisplayInfo )
            Set-WindowContent
        })

        $wpfbtnLaunch.IsDefault = $true
        $wpfbtnLaunch.Add_Click({
            $_.Handled = $true
            Write-Verbose "Launch clicked"
            Set-RemoteSessionProperties
        })
        
        $WPFbtnLaunchMstscOptions.Add_Click({
            $_.Handled = $true
            Write-Verbose "Launch on mstsc options clicked"
            Set-RemoteSessionProperties
        })

        $WPFbtnLaunchOtherOptions.Add_Click({
            $_.Handled = $true
            Write-Verbose "Launch on other options clicked"
            Set-RemoteSessionProperties
        })
        
        $WPFbtnLaunchVMwareOptions.Add_Click({
            $_.Handled = $true
            Write-Verbose "Launch on VMware clicked"
            Start-RemoteSessionFromHypervisor -hypervisorType VMware
        })

        $WPFbtnLaunchHyperVOptions.Add_Click({
            $_.Handled = $true
            Write-Verbose "Launch on Hyper-V clicked"
            Start-RemoteSessionFromHypervisor -hypervisorType Hyper-V
        })
        
        $WPFbtnLaunchADOptions.Add_Click({
            $_.Handled = $true
            Write-Verbose "Launch on AD clicked"
            Start-RemoteSessionFromHypervisor -hypervisorType AD
        })
        
        $WPFbtnLaunchHyperVConsole.Add_Click({
            $_.Handled = $true
            Write-Verbose "Launch Console on Hyper-V clicked"
            Start-RemoteSessionFromHypervisor -hypervisorType Hyper-V -console
        })
        <#
        $WPFbtnConnectHyperV.Add_Click({
            $_.Handled = $true
            Write-Verbose "Connect on Hyper-V clicked"
            Add-HyperVVMsToListView -filter $WPFtextBoxHyperVFilter.Text -regex $WPFcheckBoxHyperVRegEx.IsChecked -hyperVhost $WPFtextBoxHyperVHost.Text -all $WPFcheckBoxHyperVAllVMs.IsChecked
        })
        #>
        $WPFlistViewVMwareVMs.add_PreviewMouseDoubleClick({
            param($sender, $eventArgs)
            Write-Verbose "Launch on VMware list item double clicked"
            
            # Get the item under the mouse cursor
            $position = $eventArgs.GetPosition($sender)
            $hitTestResult = [System.Windows.Media.VisualTreeHelper]::HitTest($sender, $position)
            
            if ($hitTestResult) {
                # Walk up the visual tree to find the ListViewItem
                $element = $hitTestResult.VisualHit
                while ($element -and $element -isnot [System.Windows.Controls.ListViewItem]) {
                    $element = [System.Windows.Media.VisualTreeHelper]::GetParent($element)
                }
                
                if ($element -and $element.DataContext) {
                    # Select the double-clicked item
                    $sender.SelectedItems.Clear()
                    $sender.SelectedItems.Add($element.DataContext)
                    $sender.SelectedItem = $element.DataContext
                    
                    Write-Verbose "Selected item for launch: $($element.DataContext.Name)"
                    Start-RemoteSessionFromHypervisor -hypervisorType VMware
                }
            }
            $eventArgs.Handled = $true
        })

        $WPFlistViewVMwareVMs.Add_PreviewMouseLeftButtonDown({
            param($sourceControl, $mouseInfo)
            [void]( Invoke-ExplorerStyleAllSelectedClick -listView $WPFlistViewVMwareVMs -mouseInfo $mouseInfo )
        })
        
        $WPFlistViewHyperVVMs.Add_PreviewMouseLeftButtonDown({
            param($sender, $eventArguments)

            $script:leftButtonClickedTime = [datetime]::Now
             
            $dataContext = $null
            $element = $eventArguments.OriginalSource
    
            # Walk up the visual tree looking for an element with DataContext
            while ($element -and $null -eq $dataContext)
            {
                if ($element.DataContext -and $element.DataContext -isnot [System.Windows.Data.CollectionView])
                {
                    $dataContext = $element.DataContext
                    break
                }
                $element = $element.Parent
            }
    
            $script:targetItemData = $dataContext
            [void]( Invoke-ExplorerStyleAllSelectedClick -listView $WPFlistViewHyperVVMs -mouseInfo $eventArguments )
        })
        
        $WPFlistViewHyperVVMs.Add_PreviewMouseLeftButtonUp({
            param($sender, $eventArguments )

            $now = [datetime]::Now
            $leftButtonClickedDuration = New-TimeSpan
            if( $null -ne $script:leftButtonClickedTime )
            {
                $leftButtonClickedDuration = $now - $script:leftButtonClickedTime
                if( $leftButtonClickedDuration.TotalSeconds -gt $longPressSeconds )
                {
                    Write-Verbose "Long press, unselected $($WPFlistViewHyperVVMs.SelectedItems.Count) items"
                    $WPFlistViewHyperVVMs.SelectedItems.Clear()
                    $foundItem = $null
                    if( $null -ne $script:targetItemData )
                    {
                         foreach ($item in $WPFlistViewHyperVVMs.Items)
                         {
                            if ($item -eq $script:targetItemData)
                            {
                                $WPFlistViewHyperVVMs.SelectedItems.Add( ( $foundItem = $item ))
                                Write-Verbose "Long press, selecting $item"
                                break
                            }
                        }
                    }
                    if( $null -ne $foundItem )
                    {
                        
                        $modifiers = [System.Windows.Input.Keyboard]::Modifiers
                        ## if ( ( $ctrlDown = [System.Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::LeftCtrl)) -or [System.Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::LeftAlt ) ) )
                        if( ( $ctrlDown = ( $modifiers -band [System.Windows.Input.ModifierKeys]::Control ) -ne [System.Windows.Input.ModifierKeys]::None ) `
                            -or ( $modifiers -band [System.Windows.Input.ModifierKeys]::Alt ) -ne [System.Windows.Input.ModifierKeys]::None )
                        {
                            Write-Verbose "Alt or control ($ctrlDown) down so launching"
                        
                            # Force immediate UI update otherwise previously selected items will show whilst session launching
                            $WPFlistViewHyperVVMs.UpdateLayout()
                            [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Render, [System.Action]{})
           
                            Start-RemoteSessionFromHypervisor -hypervisorType Hyper-V -console:$ctrlDown
                        }
                    }
                }
            }
            Write-Verbose -Message "$([datetime]::Now.ToString('G')) left mouse up, duration $($leftButtonClickedDuration.TotalMilliseconds)"
        })

        $WPFlistViewHyperVVMs.add_PreviewMouseDoubleClick({
            param($sender, $eventArgs)
            Write-Verbose "Launch on Hyper-V list item double clicked"
            
            # Get the item under the mouse cursor
            $position = $eventArgs.GetPosition($sender)
            $hitTestResult = [System.Windows.Media.VisualTreeHelper]::HitTest($sender, $position)
            
            if ($hitTestResult) {
                # Walk up the visual tree to find the ListViewItem
                $element = $hitTestResult.VisualHit
                while ($element -and $element -isnot [System.Windows.Controls.ListViewItem]) {
                    $element = [System.Windows.Media.VisualTreeHelper]::GetParent($element)
                }
                
                if ($element -and $element.DataContext) {
                    # Select the double-clicked item
                    $sender.SelectedItems.Clear()
                    $sender.SelectedItems.Add($element.DataContext)
                    $sender.SelectedItem = $element.DataContext
                    
                    Write-Verbose "Selected item for launch: $($element.DataContext.Name)"
                    Start-RemoteSessionFromHypervisor -hypervisorType Hyper-V
                }
            }
            $eventArgs.Handled = $true
        })

        $WPFbuttonVMwareConnect.Add_Click({
            $_.Handled = $true
            Write-Verbose "VMware Connect clicked"
            if( -Not [string]::IsNullOrEmpty( $WPFtextBoxVMwarevCenter.Text ) )
            {
                Import-Module -Name VMware.VimAutomation.Core
                $script:vmwareConnection = Connect-VIServer -Server $WPFtextBoxVMwarevCenter.Text -Force
                if( -Not $script:vmwareConnection )
                {
                    [void][Windows.MessageBox]::Show( "Failed to connect to $($WPFtextBoxVMwarevCenter.Text)" , 'VMware Error' , 'Ok' ,'Error' )
                }
            }
            Add-VMwareVMsToListView -filter $wpfTextBoxVMwareFilter.Text -regex $WPFcheckBoxVMwareRegEx.IsChecked
        })
        
        $WPFbuttonADSearch.Add_Click({
            [string]$searchType = 'OU'
            if( $wpfradioButtonADTypeGroup.IsChecked )
            {
                $searchType = 'group'
            }
            elseif( $WPFradioButtonADTypeComputer.IsChecked )
            {
                $searchType = 'computer'
            }
            $_.Handled = $true
            Write-Verbose "AD Search clicked - search type $searchType"
            Search-AD -domainController $WPFtextBoxDomainController -searchFor $WPFtextBoxADFilter.Text -searchType $searchType
        })

        $WPFlistViewAD.Add_PreviewMouseLeftButtonDown({
            param($sourceControl, $mouseInfo)
            [void]( Invoke-ExplorerStyleAllSelectedClick -listView $WPFlistViewAD -mouseInfo $mouseInfo )
        })
        
        <#
        $WPFbtnConnectHyperV.Add_Click({
            $_.Handled = $true
            Write-Verbose "Hyper-V Connect clicked"
            $hyperVhost = $WPFtextBoxHyperVHost.Text.Trim() -replace '"'
            if( [string]::IsNullOrEmpty( $hyperVhost ) )
            {
                $hyperVhost = 'localhost'
            }
            Import-Module -Name Hyper-V
            
            Add-HyperVVMsToListView -hyperVhost $hyperVhost -filter $wpfTextBoxHyperVFilter.Text -regex $WPFcheckBoxHyperVRegEx.IsChecked -all $WPFcheckBoxHyperVAllVMs.IsChecked
        })
        #>

        if( -Not [string]::IsNullOrEmpty( $hypervHost ) )
        {
            $WPFtextBoxHyperVHost.Text = $hypervHost
        }
        $WPFbuttonVMwareDisconnect.Add_Click({
            $_.Handled = $true
            Write-Verbose "VMware Disconnect clicked"
            if( $script:vmwareConnection )
            {
                Import-Module -Name VMware.VimAutomation.Core
                $disconnection = Disconnect-VIServer -Force -Confirm:$false
                $script:vmwareConnection = $null
                $WPFlistViewVMwareVMs.Items.Clear()
            }
            else
            {
                [void][Windows.MessageBox]::Show( "Not connected" , 'VMware Error' , 'Ok' ,'Error' )
            }
        })

        $WPFbuttonAzureApplyFilter.Add_Click({
            $_.Handled = $true
            Write-Verbose "Azure Refresh clicked"
            
            $WPFbuttonAzureApplyFilter.IsEnabled = $false
            $mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
            [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
            
            try
            {
                Add-AzureVMsToListView -filter '' -regex $false -allVMs $WPFcheckBoxAzureAllVMs.IsChecked
                Update-AzureVMLabel
            }
            finally
            {
                $WPFbuttonAzureApplyFilter.IsEnabled = $true
                [System.Windows.Input.Mouse]::OverrideCursor = $null
                $mainWindow.ClearValue([System.Windows.FrameworkElement]::CursorProperty)
            }
        })

        $WPFlistViewAzureVMs.AddHandler( [System.Windows.Controls.GridViewColumnHeader]::ClickEvent , [System.Windows.RoutedEventHandler]{
            param( $sourceControl , $routedEventArgs )

            Sort-Columns -control $sourceControl -eventArgs $routedEventArgs
        })

        $WPFlistViewAzureVMs.AddHandler( [System.Windows.Controls.GridViewColumnHeader]::MouseRightButtonUpEvent , [System.Windows.Input.MouseButtonEventHandler]{
            param( $sourceControl , $mouseEventArgs )

            $columnHeader = Get-GridViewColumnHeader -eventArgs $mouseEventArgs
            if( $null -ne $columnHeader )
            {
                Set-AzureColumnFilter -columnHeader $columnHeader
                $mouseEventArgs.Handled = $true
            }
        })

        $WPFcheckBoxAzureAVD.Add_Checked({
            Set-AzureHostPoolColumn -showHostPool $true
            Set-AzureSessionMenuState -avdEnabled $true
        })

        $WPFcheckBoxAzureAVD.Add_Unchecked({
            Set-AzureHostPoolColumn -showHostPool $false
            Set-AzureSessionMenuState -avdEnabled $false
        })

        Set-AzureSessionMenuState -avdEnabled ([bool]$WPFcheckBoxAzureAVD.IsChecked)
        
        $WPFbuttonAzureConnect.Add_Click({
            $_.Handled = $true
            Write-Verbose "Azure Authenticate clicked"
            try
            {
                $connection = $null
                $mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
                $connection = Connect-AzureAccountWithTimeout -timeoutSeconds $azureAuthenticateTimeoutSeconds
                if( $null -ne $connection )
                {
                    Select-AzureSubscription
                }
            }
            catch
            {
                [void][Windows.MessageBox]::Show( $mainWindow , "Azure authentication failed`n$($_.Exception.Message)" , 'Azure Authentication' , 'Ok' ,'Error' )
            }
            finally
            {
                $mainWindow.ClearValue([System.Windows.FrameworkElement]::CursorProperty)
            }
        })

        $WPFbuttonAzureSubscription.Add_Click({
            $_.Handled = $true
            Write-Verbose "Azure Subscription clicked"
            $mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
            Select-AzureSubscription
            $mainWindow.ClearValue([System.Windows.FrameworkElement]::CursorProperty)
        })
        
        $WPFbuttonVMwareApplyFilter.Add_Click({
            $_.Handled = $true
            Write-Verbose "VMware Apply Filter clicked"
            Add-VMwareVMsToListView -filter $wpfTextBoxVMwareFilter.Text -regex $WPFcheckBoxVMwareRegEx.IsChecked
        })
        
        $WPFbuttonHyperVApplyFilter.Add_Click({
            $_.Handled = $true
            Write-Verbose "Hyper-V Apply Filter clicked"
            if( [string]::IsNullOrWhiteSpace( $WPFtextBoxHyperVHost.Text ) )
            {
                [void][Windows.MessageBox]::Show( 'No Hyper-V host specified' , 'Hyper-V Host Required' , 'Ok' , 'Error' )
                return
            }
            Add-HyperVVMsToListView -filter $WPFtextBoxHyperVFilter.Text -regex $WPFcheckBoxHyperVRegEx.IsChecked -hyperVhost $WPFtextBoxHyperVHost.Text -all $WPFcheckBoxHyperVAllVMs.IsChecked
        })

        $WPFbuttonHyperVClearFilter.Add_Click({
            $_.Handled = $true
            Write-Verbose "Hyper-V Clear Filter clicked"
            $WPFtextBoxHyperVFilter.Text = ''
            $WPFcheckBoxHyperVAllVMs.IsChecked = $false
            Add-HyperVVMsToListView -filter $WPFtextBoxHyperVFilter.Text -regex $WPFcheckBoxHyperVRegEx.IsChecked -hyperVhost $WPFtextBoxHyperVHost.Text -all $WPFcheckBoxHyperVAllVMs.IsChecked
        })

        $WPFbuttonAzureClearFilter.Add_Click({
            $_.Handled = $true
            Write-Verbose "Azure Clear Filter clicked"
            $WPFbuttonAzureClearFilter.IsEnabled = $false
            $mainWindow.Cursor = [System.Windows.Input.Cursors]::Wait
            [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
            try
            {
                $WPFcheckBoxAzureAllVMs.IsChecked = $false
                $script:azureColumnFilters.Clear()
                ##$WPFcheckBoxAzureAVD.IsChecked = $false
                Add-AzureVMsToListView -filter '' -regex $false -allVMs $WPFcheckBoxAzureAllVMs.IsChecked
            }
            finally
            {
                $WPFbuttonAzureClearFilter.IsEnabled = $true
                [System.Windows.Input.Mouse]::OverrideCursor = $null
                $mainWindow.ClearValue([System.Windows.FrameworkElement]::CursorProperty)
            }
        })
        $WPFcheckBoxAzureAVD.IsChecked = $avd

        $WPFAzurePowerOnContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewAzureVMs -Operation 'Azure_PowerOn' })
        $WPFAzurePowerOffContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewAzureVMs -Operation 'Azure_PowerOff' })
        $WPFAzureShutdownContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewAzureVMs -Operation 'Azure_ShutDown' })
        $WPFAzureHibernateContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewAzureVMs -Operation 'Azure_Hibernate' })
        $WPFAzureRestartContextMenu.Add_Click(  { Process-Action -GUIobject $WPFlistViewAzureVMs -Operation 'Azure_Restart' })
        $WPFAzureRunContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewAzureVMs -Operation 'Azure_RunOn' })
        $WPFAzureDetailContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewAzureVMs -Operation 'Azure_Detail' })
        $WPFAzureExtensionsApplicationsContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewAzureVMs -Operation 'Azure_DetailExtensionsApplications' })
        $WPFAzureDetailSessionContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewAzureVMs -Operation 'Azure_DetailSession' })
        $WPFAzureMessageSessionContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewAzureVMs -Operation 'Azure_MessageSession' })
        $WPFAzureDisconnectSessionContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewAzureVMs -Operation 'Azure_DisconnectSession' })
        $WPFAzureLogoffSessionContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewAzureVMs -Operation 'Azure_LogoffSession' })
        $WPFAzureForceLogoffSessionContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewAzureVMs -Operation 'Azure_ForceLogoffSession' })
        $WPFAzureDrainModeOnSessionContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewAzureVMs -Operation 'Azure_DrainModeOnSession' })
        $WPFAzureDrainModeOffSessionContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewAzureVMs -Operation 'Azure_DrainModeOffSession' })
        $WPFAzureOpenInPortalContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewAzureVMs -Operation 'Azure_OpenInPortal' })
        $WPFAzureDeleteContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewAzureVMs -Operation 'Azure_Delete' })
        $WPFAzureDeleteSessionHostContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewAzureVMs -Operation 'Azure_DeleteSessionHost' })
        $WPFAzureDeleteSessionHostAndVMContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewAzureVMs -Operation 'Azure_DeleteSessionHostAndVM' })
        $WPFAzureChangeDiskTypeContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewAzureVMs -Operation 'Azure_ChangeDiskType' })
        $WPFAzureEditTagsContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewAzureVMs -Operation 'Azure_EditTags' })
        $WPFAzureVMActivityLogsContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewAzureVMs -Operation 'Azure_VMActivityLogs' })
        $WPFAzureChangeHostPoolSizeContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewAzureVMs -Operation 'Azure_ChangeHostPoolSize' })
        $WPFAzureHostPoolDetailContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewAzureVMs -Operation 'Azure_HostPoolDetail' })
        $WPFAzureHostPoolActivityLogsContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewAzureVMs -Operation 'Azure_HostPoolActivityLogs' })
        $WPFAzureAppGroupsContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewAzureVMs -Operation 'Azure_AppGroups' })
        $WPFAzureAVDLogsContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewAzureVMs -Operation 'Azure_AVDLogs' })
        $WPFAzureNameToClipboard.Add_Click( { Process-Action -GUIobject $WPFlistViewAzureVMs -Operation 'NameToClipboard' })

        $WPFlistViewAzureVMs.Add_PreviewMouseLeftButtonDown({
            Param
            (
              [Parameter(Mandatory)][Object]$sourceControl,
              [Parameter(Mandatory)][System.Windows.Input.MouseButtonEventArgs]$mouseInfo
            )
            [void]( Invoke-ExplorerStyleAllSelectedClick -listView $WPFlistViewAzureVMs -mouseInfo $mouseInfo )
        })

        $WPFlistViewAzureVMs.Add_ContextMenuOpening({
            Param
            (
                [Parameter(Mandatory)][Object]$sender,
                [Parameter(Mandatory)][System.Windows.Controls.ContextMenuEventArgs]$eventArgs
            )
            if( $sender.SelectedItems.Count -eq 0 )
            {
                $hitPos = [System.Windows.Input.Mouse]::GetPosition( $sender )
                $hit    = [System.Windows.Media.VisualTreeHelper]::HitTest( $sender , $hitPos )
                if( $null -ne $hit )
                {
                    $element = $hit.VisualHit
                    while( $null -ne $element -and $element -isnot [System.Windows.Controls.ListViewItem] )
                    {
                        $element = [System.Windows.Media.VisualTreeHelper]::GetParent( $element )
                    }
                    if( $null -ne $element -and $null -ne $element.DataContext )
                    {
                        $sender.SelectedItem = $element.DataContext
                    }
                }
            }
        })

        $WPFlistViewAzureVMs.Add_PreviewMouseDoubleClick({
            param($sender, $eventArgs)

            if( -Not [bool]$WPFcheckBoxAzureAVD.IsChecked )
            {
                return
            }

            Write-Verbose "Azure sessions view list item double clicked"

            $position = $eventArgs.GetPosition($sender)
            $hitTestResult = [System.Windows.Media.VisualTreeHelper]::HitTest($sender, $position)

            if( $hitTestResult )
            {
                $element = $hitTestResult.VisualHit
                while( $element -and $element -isnot [System.Windows.Controls.ListViewItem] )
                {
                    $element = [System.Windows.Media.VisualTreeHelper]::GetParent($element)
                }

                if( $element -and $element.DataContext )
                {
                    $sender.SelectedItems.Clear()
                    $sender.SelectedItems.Add($element.DataContext)
                    $sender.SelectedItem = $element.DataContext

                    Process-Action -GUIobject $WPFlistViewAzureVMs -Operation 'Azure_DetailSession'
                    $eventArgs.Handled = $true
                }
            }
        })

        $azureListViewSelectAllCommand = New-Object System.Windows.Input.CommandBinding( [System.Windows.Input.ApplicationCommands]::SelectAll )
        $azureListViewSelectAllCommand.Add_CanExecute({
            Param
            (
                [Parameter(Mandatory)][Object]$sender,
                [Parameter(Mandatory)][System.Windows.Input.CanExecuteRoutedEventArgs]$event
            )

            if( $WPFlistViewAzureVMs.IsKeyboardFocusWithin )
            {
                $event.CanExecute = $WPFlistViewAzureVMs.Items.Count -gt 0
                $event.Handled = $true
            }
        })
        $azureListViewSelectAllCommand.Add_Executed({
            Param
            (
                [Parameter(Mandatory)][Object]$sender,
                [Parameter(Mandatory)][System.Windows.Input.ExecutedRoutedEventArgs]$event
            )

            if( $WPFlistViewAzureVMs.Items.Count -gt 0 )
            {
                [void]$WPFlistViewAzureVMs.Focus()
                $WPFlistViewAzureVMs.SelectAll()
            }
            $event.Handled = $true
        })
        [void]$WPFlistViewAzureVMs.CommandBindings.Add( $azureListViewSelectAllCommand )
        [void]$WPFlistViewAzureVMs.InputBindings.Add(
            ( New-Object System.Windows.Input.KeyBinding( [System.Windows.Input.ApplicationCommands]::SelectAll , [System.Windows.Input.Key]::A , [System.Windows.Input.ModifierKeys]::Control ) )
        )

        $WPFlistViewAzureVMs.Add_PreviewKeyDown({
            Param
            (
                            [Parameter(Mandatory)][Object]$sourceControl,
                            [Parameter(Mandatory)][Windows.Input.KeyEventArgs]$keyInfo
            )
            [bool]$ctrlDown = ( [Windows.Input.Keyboard]::Modifiers -band [Windows.Input.ModifierKeys]::Control ) -ne [Windows.Input.ModifierKeys]::None
            [bool]$isAKey = $keyInfo.Key -eq [Windows.Input.Key]::A -or $keyInfo.SystemKey -eq [Windows.Input.Key]::A

            if( $keyInfo -and $ctrlDown -and $isAKey )
            {
                if( $WPFlistViewAzureVMs.Items.Count -gt 0 )
                {
                    [void]$WPFlistViewAzureVMs.Focus()
                    $WPFlistViewAzureVMs.SelectAll()
                }
                $keyInfo.Handled = $true
            }
        })
        
        $WPFdatagridDisplays.add_MouseDoubleClick({
            $_.Handled = $true
            Write-Verbose "Grid item double clicked"
            $script:activeDisplaysWithMonitors = @( Get-DisplayInfo )
            Set-RemoteSessionProperties
        })

        $WPFbuttonHyperBuyMeACoffee.Add_Click({
            $_.Handled = $true
            [string]$beggingURL = 'https://www.buymeacoffee.com/guyrleech'
            Write-Verbose "Buy Me A Coffee clicked - opening $beggingURL"
            Start-Process -FilePath $beggingURL -Verb Open
        })

        $WPFbuttonBuyMeACoffee.Add_Click({
            $_.Handled = $true
            [string]$beggingURL = 'https://www.buymeacoffee.com/guyrleech'
            Write-Verbose "Buy Me A Coffee clicked - opening $beggingURL"
            Start-Process -FilePath $beggingURL -Verb Open
        })

        ## so enter key can launch rather than move to next grid line
        $WPFdatagridDisplays.add_PreviewKeyDown({
            Param
            (
              [Parameter(Mandatory)][Object]$sender,
              [Parameter(Mandatory)][Windows.Input.KeyEventArgs]$event
            )
            if( $event -and $event.Key -ieq 'return' )
            {
                $_.Handled = $true
                $script:activeDisplaysWithMonitors = @( Get-DisplayInfo )
                Set-RemoteSessionProperties
            }    
        })

        $WPFtxtboxWidthHeight.add_GotFocus({
            $_.Handled = $true
            $WPFradioWidthHeight.IsChecked = $true
        })
 
        $WPFtxtboxScreenPercentage.add_GotFocus({
            $_.Handled = $true
            $WPFradioPercentage.IsChecked = $true
        })
        
        if( $rdpoptions = Get-ItemProperty -Path $configKey -Name 'RDPOptions' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty 'RDPOptions'  )
        {
            $WPFtxtBoxOtherOptions.Text = $rdpoptions
        }
        
        $WPFdeleteComputersContextMenu.Add_Click( { Process-Action -GUIobject $WPFcomboboxComputer -Operation 'DeleteComputer' -Context $_  -thisObject $this } )

        $mainWindow.add_KeyDown({
            Param
            (
              [Parameter(Mandatory)][Object]$sender,
              [Parameter(Mandatory)][Windows.Input.KeyEventArgs]$event
            )
            if( $event -and $event.Key -ieq 'F5' )
            {
                $_.Handled = $true
                $script:activeDisplaysWithMonitors = @( Get-DisplayInfo )
                Set-WindowContent
            }    
        })

        $WPFchkboxRdpSigning.Add_Checked({
            $_.Handled = $true
            $WPFcomboboxSigningCert.IsEnabled = $WPFcomboboxSigningCert.Items.Count -gt 0
        })

        $WPFchkboxRdpSigning.Add_Unchecked({
            $_.Handled = $true
            $WPFcomboboxSigningCert.IsEnabled = $false
        })

        $mainWindow.Add_Loaded({
            $_.Handled = $true
            if( $_.Source -and $_.Source.WindowState -ieq 'Minimized' )
            {
                $_.Source.WindowState = 'Normal'
            }

            if( $avd -and $null -ne $WPFtabControl -and $null -ne $WPFtabAzure )
            {
                $WPFtabControl.SelectedItem = $WPFtabAzure
            }

            ## Populate code signing certificates for RDP file signing
            [array]$codeSigningCerts = @( Get-ChildItem -Path Cert:\CurrentUser\My -CodeSigningCert -ErrorAction SilentlyContinue | Where-Object { $_.NotAfter -gt (Get-Date) } )
            if( $codeSigningCerts.Count -gt 0 )
            {
                ForEach( $cert in $codeSigningCerts )
                {
                    [string]$displayName = if( -Not [string]::IsNullOrEmpty( $cert.FriendlyName ) ) { $cert.FriendlyName } else { $cert.Subject -replace '^CN=' }
                    $displayName += " (exp: $($cert.NotAfter.ToString('yyyy-MM-dd')))"
                    $comboItem = New-Object System.Windows.Controls.ComboBoxItem
                    $comboItem.Content = $displayName
                    $comboItem.Tag = $cert.Thumbprint
                    [void]$WPFcomboboxSigningCert.Items.Add( $comboItem )
                }
                $WPFcomboboxSigningCert.SelectedIndex = 0
            }
            else
            {
                $WPFchkboxRdpSigning.IsEnabled = $false
                $WPFchkboxRdpSigning.ToolTip = 'No valid code signing certificates found in Cert:\CurrentUser\My'
            }
        })
        
        <# neither of these actually worked so changed hyperv host textbox to be first in tab

        ## make Hyper-V host text box the item with the insert cursor
        $wpftextBoxHyperVHost.Add_IsVisibleChanged({
            Param( $sender , $args )

            Write-Verbose -Message "Add_IsVisibleChanged: visible $($wpftextBoxHyperVHost.IsVisible)"
            if( $wpftextBoxHyperVHost.IsVisible )
            {
                $wpftextBoxHyperVHost.Focus()
                [System.Windows.Input.Keyboard]::Focus( $wpftextBoxHyperVHost )
            }
        })

        $wpftabControl.Add_SelectionChanged({
            Param( $sender , $args )

            Write-Verbose -Message "Add_SelectionChanged: to $($WPFtabControl.SelectedItem.Header)"
            if( $WPFtabControl.SelectedItem.Header -ieq 'Hyper-V' )
            {
                $wpftextBoxHyperVHost.Focus()
                [System.Windows.Input.Keyboard]::Focus( $wpftextBoxHyperVHost )
            }
        })
        #>

        $WPFCoffeeImage.Source = Convert-Base64ToImageSource -base64 $buyMeACoffee
        $WPFCoffeeImage2.Source = Convert-Base64ToImageSource -base64 $buyMeACoffee
        $WPFCoffeeImageAzure.Source = Convert-Base64ToImageSource -base64 $buyMeACoffee
        $WPFHyperVPowerOnContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_PowerOn' })
        $WPFHyperVPowerOffContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_PowerOff' })
        $WPFHyperVShutdownContextMenu.Add_Click( { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_ShutDown' })
        $WPFHyperVRestartContextMenu.Add_Click(  { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_Restart' })
        
        $WPFHyperVRunContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_RunOn' })
        $WPFHyperVDetailContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_Detail' })
        $WPFHyperVDeleteContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_Delete' })
        $WPFHyperVDeleteAllContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_DeleteIncludingDisks' })
        
        $WPFHyperVEjectCDContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_EjectCD' })
        $WPFHyperVMountCDContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_MountCD' })
        $WPFHyperVNameToClipboard.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'NameToClipboard' })
        $WPFHyperVSaveContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_Save' })
        $WPFHyperVNewVMFromTemplateContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_New' })
        $WPFHyperVNewVMContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_NewFromTemplate' })
        $WPFHyperVReconfigureMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_Reconfigure' })
        $WPFHyperVConnectNICInternalContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_ConnectNICInternal' })
        $WPFHyperVConnectNICExternalContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_ConnectNICExternal' })
        $WPFHyperVConnectNICPrivateContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_ConnectNICPrivate' })
        $WPFHyperVDisconnectNICContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_DisconnectNIC' })
        $WPFHyperVRenameMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_Rename' })
        $WPFHyperVSuspendContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_Suspend' })
        
        $WPFHyperVResumeContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_Resume' })
        $WPFHyperVEnableNestedVirtualisationContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_EnableNestedVirtualisation' })
        $WPFHyperVEnableResourceMeteringContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_EnableResourceMetering' })
        $WPFHyperVDisableResourceMeteringContextMenu.Add_Click(   { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_DisableResourceMetering' })
        $WPFHyperVTakeSnapshotContextMenu.Add_Click(  { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_TakeSnapshot' })
        $WPFHyperVManageSnapshotContextMenu.Add_Click(  { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_ManageSnapshot' })
        $WPFHyperVRevertLatestSnapshotContextMenu.Add_Click(  { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_RevertLatestSnapshot' })
        $WPFHyperVDeleteLatestSnapshotContextMenu.Add_Click(  { Process-Action -GUIobject $WPFlistViewHyperVVMs -Operation 'HyperV_DeleteLatestSnapshot' })

        $wpfchkboxDoNotApply.IsChecked = (-Not $useOtherOptions ) ## don't want it on by default as passes blank password to mstsc giving failed logon

        [string]$msrdcExe = Get-Msrdc
        $WPFchkboxmsrdc.IsEnabled = -Not [string]::IsNullOrEmpty( $msrdcExe )
        $WPFchkboxmsrdc.IsChecked = $usemsrdc -or -Not [string]::IsNullOrEmpty($msrdcExe ) 
        
        $guiResult = $mainWindow.ShowDialog()

        if( -Not [string]::IsNullOrEmpty( $WPFtxtBoxOtherOptions.Text ) -and -Not $wpfchkboxDoNotSave.IsChecked )
        {
            if( -Not ( Get-ChildItem -Path $configKey -ErrorAction SilentlyContinue ) -and -Not (New-Item -Path $configKey -ItemType Key -Force) )
            {
                Write-Warning -Message "Failed to create `"$configKey`""
            }
            if( -Not ( Set-ItemProperty -Path $configKey -Name 'RDPOptions' -Value $WPFtxtBoxOtherOptions.Text -PassThru -Force -Type MultiString))
            {
                Write-Warning -Message "Problem writing RDP options to `"$configKey`""
            }
        }

        ## persist computers to registry
        if( -Not (Test-Path -Path $configKey ) )
        {
            if( -Not ( New-Item -Path $configKey -ItemType Key -Force ) )
            {
                Write-Warning -Message "Problem creating $configKey"
            }
        }
        Set-ItemProperty -Path $configKey -Name Computers -Value ([string[]]@( $wpfcomboboxComputer.Items.GetEnumerator() | Sort-Object -Unique )) -Force

        exit $guiResult
    }

    if( $chosenDisplay -is [array] -and $chosenDisplay.Count -gt 1 )
    {
        Throw "Spanning monitors not yet supported"
    }
    $chosenDisplay = $activeDisplaysWithMonitors | Where ScreenDeviceName -eq $chosen.ScreenDeviceName
    if( -Not $chosenDisplay )
    {
        Throw "Failed to find device name $($chosen.ScreenDeviceName) in internal data"
    }
}

New-RemoteSession -rethrow

# SIG # Begin signature block
# MIIkkgYJKoZIhvcNAQcCoIIkgzCCJH8CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU+/YUx8BUMKjwgqTPkS+sbVjO
# 1tSggh9gMIIFfTCCA2WgAwIBAgIQAdazdTZfIM2RHdcv5fmTZDANBgkqhkiG9w0B
# AQsFADBaMQswCQYDVQQGEwJMVjEZMBcGA1UEChMQRW5WZXJzIEdyb3VwIFNJQTEw
# MC4GA1UEAxMnR29HZXRTU0wgRzQgQ1MgUlNBNDA5NiBTSEEyNTYgMjAyMiBDQS0x
# MB4XDTI1MDcyMTAwMDAwMFoXDTI2MDcyMDIzNTk1OVowcTELMAkGA1UEBhMCR0Ix
# EjAQBgNVBAcTCVdha2VmaWVsZDEmMCQGA1UEChMdU2VjdXJlIFBsYXRmb3JtIFNv
# bHV0aW9ucyBMdGQxJjAkBgNVBAMTHVNlY3VyZSBQbGF0Zm9ybSBTb2x1dGlvbnMg
# THRkMHYwEAYHKoZIzj0CAQYFK4EEACIDYgAERFbrIQcZmiw2ScrP4eHxhzHvoBGn
# AnE3GpY3vjU5CpVG6JLtPgXTQz8aLW7IdGhx7x4cJ3a6y+3/6Q+OX+VVFSiuRd60
# GO22Y2eoMcBmvwc7hWbEYTtdjEzAu82sMmkAo4IB1DCCAdAwHwYDVR0jBBgwFoAU
# yfwQ71DIy2t/vQhE7zpik+1bXpowHQYDVR0OBBYEFPu1ucNQfJlsl2iXm5HJzCZH
# EgfnMD4GA1UdIAQ3MDUwMwYGZ4EMAQQBMCkwJwYIKwYBBQUHAgEWG2h0dHA6Ly93
# d3cuZGlnaWNlcnQuY29tL0NQUzAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYI
# KwYBBQUHAwMwgZcGA1UdHwSBjzCBjDBEoEKgQIY+aHR0cDovL2NybDMuZGlnaWNl
# cnQuY29tL0dvR2V0U1NMRzRDU1JTQTQwOTZTSEEyNTYyMDIyQ0EtMS5jcmwwRKBC
# oECGPmh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9Hb0dldFNTTEc0Q1NSU0E0MDk2
# U0hBMjU2MjAyMkNBLTEuY3JsMIGDBggrBgEFBQcBAQR3MHUwJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBNBggrBgEFBQcwAoZBaHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0dvR2V0U1NMRzRDU1JTQTQwOTZTSEEyNTYyMDIy
# Q0EtMS5jcnQwCQYDVR0TBAIwADANBgkqhkiG9w0BAQsFAAOCAgEACx2bHTrjHk/3
# tX7HUH4SM/9sEtfVFpRtZcV4nmDpjwI7tAwOSGYXk4DLVjJFJveoFjbYsZ8vquZ1
# /HJM7rg+O3rCNmzOBvUXFVSjdL3S2R7+kF2ROR7dqk1/BNW6n3o7Q3BmNGqjo1WH
# jov6PfAbEffCLZI1jT98RNqChMesWMmQS+nf8xwdskne4XZOFX5h/a00X7QLAJ+S
# /bOptiC0SvEEa5FCWPUcV7ML0MtoDc3HIPnmMMuYLy586eJHbE5XlfEsmWUNk3Kf
# hxzxsXpAdTSDOeb5Qm/aHGMOY+56Gnt/zxfrv2bfxPnKKZtXPjA47tm89RHpal8b
# lbCAkVfYpKSe0BFPi8FIk+zXvpoAZkNyCMm/HUMEdMbtRP7CqFmYz0YWuiS3uuUW
# qAZ1zl+n1kIJT8eOu6o01EKS8ShijHUI0vixibiNvwTFgRyX3Yc/9xkfV1Wgzli4
# ZPgoZI6FwYBdrhRF0or+CzYIoENUfUYqI7pBM5kkXuSytFD3SXIeSPx14NZSRTzk
# cdOSJWtLkjLrIrIKzzb5eXxLn/gxmJdssB7GUKZHik+cB0OUCRKHEysBj34hnvXa
# zuQ6DKLOQFy+cZ6z4f2kAeFyq7bWUxctPmF61FkmGvb9q6e3AMLg7JnfYC6EM31u
# 42oGx38b5i0NAiUzvWOAbCWTC+G44pgwggWNMIIEdaADAgECAhAOmxiO+dAt5+/b
# UOIIQBhaMA0GCSqGSIb3DQEBDAUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxE
# aWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNVBAMT
# G0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0yMjA4MDEwMDAwMDBaFw0z
# MTExMDkyMzU5NTlaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJ
# bmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0
# IFRydXN0ZWQgUm9vdCBHNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIB
# AL/mkHNo3rvkXUo8MCIwaTPswqclLskhPfKK2FnC4SmnPVirdprNrnsbhA3EMB/z
# G6Q4FutWxpdtHauyefLKEdLkX9YFPFIPUh/GnhWlfr6fqVcWWVVyr2iTcMKyunWZ
# anMylNEQRBAu34LzB4TmdDttceItDBvuINXJIB1jKS3O7F5OyJP4IWGbNOsFxl7s
# Wxq868nPzaw0QF+xembud8hIqGZXV59UWI4MK7dPpzDZVu7Ke13jrclPXuU15zHL
# 2pNe3I6PgNq2kZhAkHnDeMe2scS1ahg4AxCN2NQ3pC4FfYj1gj4QkXCrVYJBMtfb
# BHMqbpEBfCFM1LyuGwN1XXhm2ToxRJozQL8I11pJpMLmqaBn3aQnvKFPObURWBf3
# JFxGj2T3wWmIdph2PVldQnaHiZdpekjw4KISG2aadMreSx7nDmOu5tTvkpI6nj3c
# AORFJYm2mkQZK37AlLTSYW3rM9nF30sEAMx9HJXDj/chsrIRt7t/8tWMcCxBYKqx
# YxhElRp2Yn72gLD76GSmM9GJB+G9t+ZDpBi4pncB4Q+UDCEdslQpJYls5Q5SUUd0
# viastkF13nqsX40/ybzTQRESW+UQUOsxxcpyFiIJ33xMdT9j7CFfxCBRa2+xq4aL
# T8LWRV+dIPyhHsXAj6KxfgommfXkaS+YHS312amyHeUbAgMBAAGjggE6MIIBNjAP
# BgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBTs1+OC0nFdZEzfLmc/57qYrhwPTzAf
# BgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzAOBgNVHQ8BAf8EBAMCAYYw
# eQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2Vy
# dC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9E
# aWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwRQYDVR0fBD4wPDA6oDigNoY0aHR0
# cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNy
# bDARBgNVHSAECjAIMAYGBFUdIAAwDQYJKoZIhvcNAQEMBQADggEBAHCgv0NcVec4
# X6CjdBs9thbX979XB72arKGHLOyFXqkauyL4hxppVCLtpIh3bb0aFPQTSnovLbc4
# 7/T/gLn4offyct4kvFIDyE7QKt76LVbP+fT3rDB6mouyXtTP0UNEm0Mh65ZyoUi0
# mcudT6cGAxN3J0TU53/oWajwvy8LpunyNDzs9wPHh6jSTEAZNUZqaVSwuKFWjuyk
# 1T3osdz9HNj0d1pcVIxv76FQPfx2CWiEn2/K2yCNNWAcAgPLILCsWKAOQGPFmCLB
# sln1VWvPJ6tsds5vIy30fnFqI2si/xK4VC0nftg62fC2h5b9W9FcrBjDTZ9ztwGp
# n1eqXijiuZQwggahMIIEiaADAgECAhAHhD2tAcEVwnTuQacoIkZ5MA0GCSqGSIb3
# DQEBCwUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAX
# BgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRydXN0
# ZWQgUm9vdCBHNDAeFw0yMjA2MjMwMDAwMDBaFw0zMjA2MjIyMzU5NTlaMFoxCzAJ
# BgNVBAYTAkxWMRkwFwYDVQQKExBFblZlcnMgR3JvdXAgU0lBMTAwLgYDVQQDEydH
# b0dldFNTTCBHNCBDUyBSU0E0MDk2IFNIQTI1NiAyMDIyIENBLTEwggIiMA0GCSqG
# SIb3DQEBAQUAA4ICDwAwggIKAoICAQCtHvQHskNmiqJndyWVCqX4FtYp5FfJLO9S
# h0BuwXuvBeNYt21xf8h/pLJ/7YzeKcNq9z4zEhecqtD0xhbvSB8ksBAfWBMZO0NL
# fOT0j7WyNuD7rv+ZFza+mxIQ79s1dCiwUMwGonaoDK7mqZfDpKEExR6UyKBh3aat
# T73U2Imx/x+fYTmQFq+N8FrLs6Fh6YEGWJTgsxyw1fAChCfgtEcZkdtcgK7quqsk
# HtW6PJ9l5VNJ7T3WXpznsOOxrz3qx0CzWjwK8+3Kv2X6piWvd8YRfAOycSrT4/PM
# 0cHLFc5xs/4m/ek4FCnYSem43doFftBxZBQkHKoPW3Bt6VIrhVIwvO7hrUjhchJJ
# ZYdSld3bANDviJ5/ToP7ENv97U9MtKFvmC5dzd1p4HxFR0p5wWmYQbW+y3RFm0np
# 6H9m57MUMNp0ysmdJjb0f7+dVLX3OEBUb6H+r1LRLZT/xEOTuwOxGg2S4w25KGL9
# SCBUW4nkBljPHeJToU+THt0P8ZQf4B9IFlGxtLK0g3uOAnwSFgKtmNjhkTl8caLA
# QwbgEINCqrhc0b6k2Z8+QwgVAL0nIuzM9ckKP8xtIcWg85L3/l0cTkHQde+jKGDG
# 2CdxBHtflLIUtwqD7JA2uCxWlIzRNgwT0kH2en0+QV8KziSGaqO2r06kwboq2/xy
# 4e98CEfSYwIDAQABo4IBWTCCAVUwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4E
# FgQUyfwQ71DIy2t/vQhE7zpik+1bXpowHwYDVR0jBBgwFoAU7NfjgtJxXWRM3y5n
# P+e6mK4cD08wDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHcG
# CCsGAQUFBwEBBGswaTAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQu
# Y29tMEEGCCsGAQUFBzAChjVodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGln
# aUNlcnRUcnVzdGVkUm9vdEc0LmNydDBDBgNVHR8EPDA6MDigNqA0hjJodHRwOi8v
# Y3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNybDAcBgNV
# HSAEFTATMAcGBWeBDAEDMAgGBmeBDAEEATANBgkqhkiG9w0BAQsFAAOCAgEAC9sK
# 17IdmKTCUatEs7+yewhJnJ4tyrLwNEnfl6HrG8Pm7HZ0b+5Jc+GGqJT8kRc7mihu
# VrdsYNHdicueDL9imhtCusI/rUmjwhtflp+XgLkmgLGrmsEho1b+lGiRp7LC/10d
# i8SAOilDkHj5Zx142xRvBrrWj9eOdSGHwYubAsEd6CDojwcaVz9pfXMzYO3kc0O6
# PXg1TkcgkYlCUAuDHuk/sZx68W0FVj1P2iMh+VUq9lL1puroAydoeWVUh/+cMXeq
# fgpBqlAW+r8ma5F6yKL0stVQH8vYb1ES0mJSIPyIfkIjC1V0pbZS3p0QWsKaafEo
# r8fLfLNfSxntVI/ugut0+6ekluPWRpEXH+JAiNdRjbLbZchCREe3/Xl0YlwkA+eQ
# VJfM0A7XiuFtY/mOpK2AN+E25t5mQYFhpdxZX5LTDKWgDnb+A6QnEt4iNyukcLaJ
# uS8IPgPz0E2ALZLt3Rqs+lXifK/GwnNIWQNbf7FmLDB9ph8i8dvsR1hsjc2KPEW4
# bAsbvLcz8hN1zE1/QbOV92vDGoFjwZOi2koQ+UyEh0e8jDFHAKJeTI+p8EPE/mqv
# ojLFAnt31yXIA2tjt0ERtsjkhBNmZY6SEOfnIoOwvyqavLPya1Ut3/2cOFLuNQ8Q
# l6HaZsNQErnnzn+ZEAaUTkPZaeVyoHIkODECLzkwgga0MIIEnKADAgECAhANx6xX
# Bf8hmS5AQyIMOkmGMA0GCSqGSIb3DQEBCwUAMGIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAf
# BgNVBAMTGERpZ2lDZXJ0IFRydXN0ZWQgUm9vdCBHNDAeFw0yNTA1MDcwMDAwMDBa
# Fw0zODAxMTQyMzU5NTlaMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2Vy
# dCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBp
# bmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwggIiMA0GCSqGSIb3DQEBAQUAA4IC
# DwAwggIKAoICAQC0eDHTCphBcr48RsAcrHXbo0ZodLRRF51NrY0NlLWZloMsVO1D
# ahGPNRcybEKq+RuwOnPhof6pvF4uGjwjqNjfEvUi6wuim5bap+0lgloM2zX4kftn
# 5B1IpYzTqpyFQ/4Bt0mAxAHeHYNnQxqXmRinvuNgxVBdJkf77S2uPoCj7GH8BLux
# BG5AvftBdsOECS1UkxBvMgEdgkFiDNYiOTx4OtiFcMSkqTtF2hfQz3zQSku2Ws3I
# fDReb6e3mmdglTcaarps0wjUjsZvkgFkriK9tUKJm/s80FiocSk1VYLZlDwFt+cV
# FBURJg6zMUjZa/zbCclF83bRVFLeGkuAhHiGPMvSGmhgaTzVyhYn4p0+8y9oHRaQ
# T/aofEnS5xLrfxnGpTXiUOeSLsJygoLPp66bkDX1ZlAeSpQl92QOMeRxykvq6gby
# lsXQskBBBnGy3tW/AMOMCZIVNSaz7BX8VtYGqLt9MmeOreGPRdtBx3yGOP+rx3rK
# WDEJlIqLXvJWnY0v5ydPpOjL6s36czwzsucuoKs7Yk/ehb//Wx+5kMqIMRvUBDx6
# z1ev+7psNOdgJMoiwOrUG2ZdSoQbU2rMkpLiQ6bGRinZbI4OLu9BMIFm1UUl9Vne
# Ps6BaaeEWvjJSjNm2qA+sdFUeEY0qVjPKOWug/G6X5uAiynM7Bu2ayBjUwIDAQAB
# o4IBXTCCAVkwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQU729TSunkBnx6
# yuKQVvYv1Ensy04wHwYDVR0jBBgwFoAU7NfjgtJxXWRM3y5nP+e6mK4cD08wDgYD
# VR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMIMHcGCCsGAQUFBwEBBGsw
# aTAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEEGCCsGAQUF
# BzAChjVodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVk
# Um9vdEc0LmNydDBDBgNVHR8EPDA6MDigNqA0hjJodHRwOi8vY3JsMy5kaWdpY2Vy
# dC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBABfO+xaAHP4HPRF2
# cTC9vgvItTSmf83Qh8WIGjB/T8ObXAZz8OjuhUxjaaFdleMM0lBryPTQM2qEJPe3
# 6zwbSI/mS83afsl3YTj+IQhQE7jU/kXjjytJgnn0hvrV6hqWGd3rLAUt6vJy9lMD
# PjTLxLgXf9r5nWMQwr8Myb9rEVKChHyfpzee5kH0F8HABBgr0UdqirZ7bowe9Vj2
# AIMD8liyrukZ2iA/wdG2th9y1IsA0QF8dTXqvcnTmpfeQh35k5zOCPmSNq1UH410
# ANVko43+Cdmu4y81hjajV/gxdEkMx1NKU4uHQcKfZxAvBAKqMVuqte69M9J6A47O
# vgRaPs+2ykgcGV00TYr2Lr3ty9qIijanrUR3anzEwlvzZiiyfTPjLbnFRsjsYg39
# OlV8cipDoq7+qNNjqFzeGxcytL5TTLL4ZaoBdqbhOhZ3ZRDUphPvSRmMThi0vw9v
# ODRzW6AxnJll38F0cuJG7uEBYTptMSbhdhGQDpOXgpIUsWTjd6xpR6oaQf/DJbg3
# s6KCLPAlZ66RzIg9sC+NJpud/v4+7RWsWCiKi9EOLLHfMR2ZyJ/+xhCx9yHbxtl5
# TPau1j/1MIDpMPx0LckTetiSuEtQvLsNz3Qbp7wGWqbIiOWCnb5WqxL3/BAPvIXK
# UjPSxyZsq8WhbaM2tszWkPZPubdcMIIG7TCCBNWgAwIBAgIQCoDvGEuN8QWC0cR2
# p5V0aDANBgkqhkiG9w0BAQsFADBpMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGln
# aUNlcnQsIEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0IFRydXN0ZWQgRzQgVGltZVN0
# YW1waW5nIFJTQTQwOTYgU0hBMjU2IDIwMjUgQ0ExMB4XDTI1MDYwNDAwMDAwMFoX
# DTM2MDkwMzIzNTk1OVowYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0
# LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBTSEEyNTYgUlNBNDA5NiBUaW1lc3Rh
# bXAgUmVzcG9uZGVyIDIwMjUgMTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoC
# ggIBANBGrC0Sxp7Q6q5gVrMrV7pvUf+GcAoB38o3zBlCMGMyqJnfFNZx+wvA69HF
# TBdwbHwBSOeLpvPnZ8ZN+vo8dE2/pPvOx/Vj8TchTySA2R4QKpVD7dvNZh6wW2R6
# kSu9RJt/4QhguSssp3qome7MrxVyfQO9sMx6ZAWjFDYOzDi8SOhPUWlLnh00Cll8
# pjrUcCV3K3E0zz09ldQ//nBZZREr4h/GI6Dxb2UoyrN0ijtUDVHRXdmncOOMA3Co
# B/iUSROUINDT98oksouTMYFOnHoRh6+86Ltc5zjPKHW5KqCvpSduSwhwUmotuQhc
# g9tw2YD3w6ySSSu+3qU8DD+nigNJFmt6LAHvH3KSuNLoZLc1Hf2JNMVL4Q1Opbyb
# pMe46YceNA0LfNsnqcnpJeItK/DhKbPxTTuGoX7wJNdoRORVbPR1VVnDuSeHVZlc
# 4seAO+6d2sC26/PQPdP51ho1zBp+xUIZkpSFA8vWdoUoHLWnqWU3dCCyFG1roSrg
# HjSHlq8xymLnjCbSLZ49kPmk8iyyizNDIXj//cOgrY7rlRyTlaCCfw7aSUROwnu7
# zER6EaJ+AliL7ojTdS5PWPsWeupWs7NpChUk555K096V1hE0yZIXe+giAwW00aHz
# rDchIc2bQhpp0IoKRR7YufAkprxMiXAJQ1XCmnCfgPf8+3mnAgMBAAGjggGVMIIB
# kTAMBgNVHRMBAf8EAjAAMB0GA1UdDgQWBBTkO/zyMe39/dfzkXFjGVBDz2GM6DAf
# BgNVHSMEGDAWgBTvb1NK6eQGfHrK4pBW9i/USezLTjAOBgNVHQ8BAf8EBAMCB4Aw
# FgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwgZUGCCsGAQUFBwEBBIGIMIGFMCQGCCsG
# AQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wXQYIKwYBBQUHMAKGUWh0
# dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNydDBfBgNVHR8EWDBWMFSgUqBQ
# hk5odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRUaW1l
# U3RhbXBpbmdSU0E0MDk2U0hBMjU2MjAyNUNBMS5jcmwwIAYDVR0gBBkwFzAIBgZn
# gQwBBAIwCwYJYIZIAYb9bAcBMA0GCSqGSIb3DQEBCwUAA4ICAQBlKq3xHCcEua5g
# QezRCESeY0ByIfjk9iJP2zWLpQq1b4URGnwWBdEZD9gBq9fNaNmFj6Eh8/YmRDfx
# T7C0k8FUFqNh+tshgb4O6Lgjg8K8elC4+oWCqnU/ML9lFfim8/9yJmZSe2F8AQ/U
# dKFOtj7YMTmqPO9mzskgiC3QYIUP2S3HQvHG1FDu+WUqW4daIqToXFE/JQ/EABgf
# ZXLWU0ziTN6R3ygQBHMUBaB5bdrPbF6MRYs03h4obEMnxYOX8VBRKe1uNnzQVTeL
# ni2nHkX/QqvXnNb+YkDFkxUGtMTaiLR9wjxUxu2hECZpqyU1d0IbX6Wq8/gVutDo
# jBIFeRlqAcuEVT0cKsb+zJNEsuEB7O7/cuvTQasnM9AWcIQfVjnzrvwiCZ85EE8L
# UkqRhoS3Y50OHgaY7T/lwd6UArb+BOVAkg2oOvol/DJgddJ35XTxfUlQ+8Hggt8l
# 2Yv7roancJIFcbojBcxlRcGG0LIhp6GvReQGgMgYxQbV1S3CrWqZzBt1R9xJgKf4
# 7CdxVRd/ndUlQ05oxYy2zRWVFjF7mcr4C34Mj3ocCVccAvlKV9jEnstrniLvUxxV
# ZE/rptb7IRE2lskKPIJgbaP5t2nGj/ULLi49xTcBZU8atufk+EMF/cWuiC7POGT7
# 5qaL6vdCvHlshtjdNXOCIUjsarfNZzGCBJwwggSYAgEBMG4wWjELMAkGA1UEBhMC
# TFYxGTAXBgNVBAoTEEVuVmVycyBHcm91cCBTSUExMDAuBgNVBAMTJ0dvR2V0U1NM
# IEc0IENTIFJTQTQwOTYgU0hBMjU2IDIwMjIgQ0EtMQIQAdazdTZfIM2RHdcv5fmT
# ZDAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG
# 9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIB
# FTAjBgkqhkiG9w0BCQQxFgQU4AyWChlkiDTl3x889S2vJt+Q66MwCwYHKoZIzj0C
# AQUABGcwZQIxAKJhtK70THEo30uQfR+7H6h0ez5vqR6U00sPFr1WLzQ7AtKs3yrD
# M7zdDGYUrAMKzQIwV1Tpuf+Frww9QogtnMIbRBjmLG53y0rA2M667cTrBUSi+Xfp
# e5K45R8tic9+S3duoYIDJjCCAyIGCSqGSIb3DQEJBjGCAxMwggMPAgEBMH0waTEL
# MAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMUEwPwYDVQQDEzhE
# aWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGluZyBSU0E0MDk2IFNIQTI1NiAy
# MDI1IENBMQIQCoDvGEuN8QWC0cR2p5V0aDANBglghkgBZQMEAgEFAKBpMBgGCSqG
# SIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI2MDcxNzE2NTIw
# NVowLwYJKoZIhvcNAQkEMSIEIEXZTLkHu4GjHxwXNtUsii/a/b+f11tX7v8M/PLU
# bsVHMA0GCSqGSIb3DQEBAQUABIICAMCMz9Z3fG+Q1h0Dlm8wC/p+OK4cJRr11atR
# w9WvXyGHmDFm9ZRBUc1DLzllZoqgcSK14ELNHyvCI8aogzSww79Xg00NFxUXGv2N
# jR18LsVIECO6N8eLEvNYkK7HglnUOM567EOovlPZnfGynemk4O1Nuo+2Hdnrx3/u
# /BGLVxthsG4XWwEequXJuv6qcv893fUAsrZZn0McAeAJimzhWQS5BZFvsZc3bWS3
# d+JrCxa/FbKA4OefjV+CjBOcczaB38+FL2IFuTzGmzwSf82jwSbIxEnP0/pxeSBP
# hJxLDawlb1+oSg+XEEQOUQRy8cyf7T4qvAMWfelXFvxvCN883inlzaZU5bMMEpKE
# bpsY3rRQA3fc2tGcXrRzNIT0Hsnpw05x8es+H5ttcDUgKqQlqc5rPpLFw2epe6Ru
# +zTnO/Q+aYektgPRzIkqGljfmIN/WcjeeNL/+k2Gpbq7T45QzY2qPLHZ9KGsaqgr
# NsQgnwKuNLFQsdtourKwkjbVFrLZaTafYznQo7IVief95ky2yTSO1cmjSE9jMKGC
# dhQadmhiRreZGHNWTDDe7kDXngyb6N2xwr95dX8GV776S3H4wjRvRGnJeyagvhuZ
# 8yX07UrnrV+QM69KGC7aOaW7I8FufCr48WmHEg21MQugqfbyIaaKGNKh+9h8ouOW
# Yqjr8TlJ
# SIG # End signature block
