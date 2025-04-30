<#
.SYNOPSIS
    WPF GUI with a button that 5 seconds after clicking will report what is under the cursor

.PARAMETER waitTimeSeconds
    How long to wait in seconds from clicking to when the window under the cursor is identified
    
.PARAMETER fontSize
    Fontsize in points to use in the output window
 
.PARAMETER outputWidth
    Width to make the output window (helps to avoid wrapping)

.NOTES
    Modification History:

    2024/09/26  @guyrleech  Script born
    2024/10/04  @guyrleech  Added dialogue at cursor position for displaying info
#>

[CmdletBinding()]

Param
(
    [decimal]$waitTimeSeconds = 5 ,
    [int]$fontSize = 20 ,
    [int]$outputWidth = 400
)

## these are used by the script we get the function Get-DirectRelativeProcessDetails from so define them here rather than remove in the code so easier to add updates from that script
[bool]$signing = $true
[bool]$file = $true
[bool]$noOwner = $false
[bool]$noIndent = $false
[string]$unknownProcessName = '<UNKNOWN>'
[int]$indentMultiplier = 1
[string]$indenter = ' '
[System.Collections.Generic.List[object]]$openedWindows = @()

# Load required WPF assemblies
Add-Type -AssemblyName PresentationFramework

# Define XAML for the WPF window
$mainwindowXAML = @"
<Window x:Class="WindowIdentifier.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        mc:Ignorable="d"
        Title="Window Identifier" Height="200" Width="300" Topmost="True">
    <Grid>
        <Label Name="DragLabel" Content="Click" HorizontalAlignment="Center" VerticalAlignment="Center" 
               Background="LightBlue" Padding="20" AllowDrop="True"/>
    </Grid>
</Window>
"@

[string]$infowindowXAML = @'
<Window x:Class="WindowIdentifier.InfoWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:TextViewer"
        mc:Ignorable="d"
        Title="Text Viewer" Height="450" Width="1400" Background="Black">
    <Grid>
        <RichTextBox x:Name="richtextboxMain" HorizontalAlignment="Left" Margin="0" VerticalAlignment="Top" IsReadOnly="False" HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto" BorderThickness="0" FontFamily="Consolas" FontSize="14" Foreground="#FD7609" Background="Black">
            <FlowDocument>
                <Paragraph x:Name="paragraph">
                    <Run x:Name="run"/>
                </Paragraph>
            </FlowDocument>
        </RichTextBox>
    </Grid>
</Window>
'@

Function New-GUI( $inputXAML )
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

Add-Type -AssemblyName PresentationCore , PresentationFramework , System.Windows.Forms

if( -Not ( $mainwindow = New-GUI -inputXAML $mainwindowXAML ) )
{
    Throw 'Failed to create WPF from XAML'
}

## Borrowed from http://stackoverflow.com/a/15846912 and adapted
Add-Type @'
using System;
using System.Text;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace PInvoke.Win32
{
    public static class UserInput
    {  
        [DllImport("user32.dll", SetLastError=true)]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        public static extern IntPtr GetTopWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

        public enum GetWindow_Cmd : uint {
            GW_HWNDFIRST = 0,
            GW_HWNDLAST = 1,
            GW_HWNDNEXT = 2,
            GW_HWNDPREV = 3,
            GW_OWNER = 4,
            GW_CHILD = 5,
            GW_ENABLEDPOPUP = 6
        }
        [DllImport("user32.dll")]
        public static extern bool GetCursorPos(out POINT lpPoint);
        
        [DllImport("user32.dll")]
        public static extern IntPtr WindowFromPoint(POINT Point);

        public struct POINT {
            public int X;
            public int Y;
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Auto)]
        public static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(IntPtr sClassName, String sAppName);
    
        [DllImport("user32.dll")]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
        
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();
        
        [DllImport("user32.dll", SetLastError=true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
 
        [DllImport("user32.dll", SetLastError=false)]
        private static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

        [StructLayout(LayoutKind.Sequential)]
        private struct LASTINPUTINFO
        {
            public uint cbSize;
            public int dwTime;
        }
        public static DateTime LastInput
        {
            get
            {
                DateTime bootTime = DateTime.UtcNow.AddMilliseconds(-Environment.TickCount);
                DateTime lastInput = bootTime.AddMilliseconds(LastInputTicks);
                return lastInput;
            }
        }
        public static TimeSpan IdleTime
        {
            get
            {
                return DateTime.UtcNow.Subtract(LastInput);
            }
        }
        public static int LastInputTicks
        {
            get
            {
                LASTINPUTINFO lii = new LASTINPUTINFO();
                lii.cbSize = (uint)Marshal.SizeOf(typeof(LASTINPUTINFO));
                GetLastInputInfo(ref lii);
                return lii.dwTime;
            }
        }
    }
}
'@

Function Get-DirectRelativeProcessDetails
{
    Param
    (
        [int]$id ,
        [int]$level = 0 ,
        [datetime]$created ,
        [bool]$children = $false ,
        [switch]$recurse ,
        [switch]$quiet ,
        [switch]$firstCall
    )
    ##Write-Verbose -Message "Get-DirectRelativeProcessDetails pid $id level $level"
    $processDetail = Get-CimInstance -ClassName win32_process -Filter "ProcessId = '$id'" -Verbose:$false

    ## guard against pid re-use (do not need to check pid created after child process since could not exist before with same pid although can't guarantee that pid hasn't been reused since unless we check process auditing/sysmon)
    if( $null -ne $processDetail -and ( $null -eq $created -or ( -not $children -and $processDetail.CreationDate -le $created ) -or $children ) )
    {
    <#
        ## * means any session, -1 means session script is running in any other positive value is session id it process must be running in
        if( $firstCall )
        {
            if( $script:sessionIdAsInt -lt 0 ) ## session for script only
            {
                if( $processDetail.SessionId -ne $script:thisSessionId )
                {
                    $processDetail = $null
                }
            }
            elseif( $script:sessionIdAsInt -ne $processDetail.SessionId ) ## session id passed so check process is in this session
            {
                $processDetail = $null
            }
        }
    #>
        if( $null -ne $processDetail -and $null -ne $processDetail.ParentProcessId -and $processDetail.ParentProcessId -gt 0 )
        {
            if( $recurse )
            {
                if( $children )
                {
                    $script:processes | Where-Object ParentProcessId -eq $id -PipelineVariable childProcess | ForEach-Object `
                    {
                        Get-DirectRelativeProcessDetails -id $childProcess.ProcessId -level ($level - 1) -recurse -children $true -created $processDetail.CreationDate -quiet:$quiet
                    }
                }
                if( $firstCall -or -not $children ) ## getting parents
                {
                    Get-DirectRelativeProcessDetails -id $processDetail.ParentProcessId -level ($level + 1) -children $false -recurse  -created $processDetail.CreationDate -quiet:$quiet
                }
            }

            ## don't just look up svchost.exe as could be a service with it's own exe
            [string]$service = ( Get-CimInstance -ClassName win32_service -Filter "ProcessId = '$id'" -ErrorAction SilentlyContinue -Verbose:$false | Select-Object -ExpandProperty Name) -join '/'

            $owner = $null
            if( -Not $noOwner )
            {
                if( -Not $processDetail.PSObject.Properties[ 'Owner' ] )
                {
                    $ownerDetail = Invoke-CimMethod -InputObject $processDetail -MethodName GetOwner -ErrorAction SilentlyContinue -Verbose:$false
                    if( $null -ne $ownerDetail -and $ownerDetail.ReturnValue -eq 0 )
                    {
                        $owner = "$($ownerDetail.Domain)\$($ownerDetail.User)"
                    }

                    Add-Member -InputObject $processDetail -MemberType NoteProperty -Name Owner -Value $owner
                }
                else
                {
                    $owner = $processDetail.owner
                }
            }
            
            ## clone the process detail since may be used by another process being analysed and could be at a different level in that
            ## clone() method not available in PS 7.x
            $clone = [CimInstance]::new( $processDetail )

            Add-Member -InputObject $clone -NotePropertyMembers @{
                Owner   = $owner
                Service = $service
                Level   = $level
                '-'     = $(if( $firstCall ) { '*' } else {''})
            }

            if( $signing )
            {
                $signingDetail = $null
                if( -Not [string]::IsNullOrEmpty( $processDetail.Path ) )
                {
                    $signingDetail = Get-AuthenticodeSignature -FilePath $processDetail.Path
                }
                Add-Member -InputObject $clone -MemberType NoteProperty -Name Signing -Value $signingDetail
            }
            if( $file )
            {
                $fileInfo = $null
                if( -Not [string]::IsNullOrEmpty( $processDetail.Path ) )
                {
                    $fileInfo = Get-ItemProperty -Path $processDetail.Path
                }
                Add-Member -InputObject $clone -MemberType NoteProperty -Name FileInfo -Value $fileInfo
            }
            $clone ## return
        }
        ## else no parent or excluded based on session id
    }
    elseif( $firstCall ) ## only warn on first call
    {
        if( -not $quiet )
        {
            Write-Warning "No process found for id $id"
        }
    }
    elseif( -not $quiet )
    {
        ## TODO search process auditing/sysmon ?
        $emptyResult = [CimInstance]::new( 'Win32_Process' , 'root/cimv2' )
        Add-Member -InputObject $emptyResult -NotePropertyMembers @{
            Name = $unknownProcessName
            ProcessId = $id
            Level = $level
        }
        if( $signing )
        {
            Add-Member -InputObject $emptyResult -MemberType NoteProperty -Name Signing -Value $null
        }
        if( $file )
        {
            Add-Member -InputObject $emptyResult -MemberType NoteProperty -Name File -Value $null
        }
        $emptyResult ## return
    }
}

Function Get-WindowDetails
{
    Param
    (
        $window ,
        [int]$x ,
        [int]$y
    )
    [uint32]$windowPid = 0

    if( -Not $PSBoundParameters[ 'window' ] )
    {
        $window = [PInvoke.Win32.UserInput]::GetForegroundWindow()
    }
    if( $null -ne $window )
    {
        $string = New-Object -TypeName System.Text.Stringbuilder
        if ([PInvoke.Win32.UserInput]::GetWindowThreadProcessId( $window , [ref]$windowPid ) -gt 0 )
        {
            $length = [PInvoke.Win32.UserInput]::GetWindowTextLength( $window )
            $string = New-Object -TypeName System.Text.Stringbuilder -ArgumentList ($length + 1)
            if ($length -gt 0)
            {
                if( [PInvoke.Win32.UserInput]::GetWindowText( $window , $string , ( $length + 1 )) -eq 0 )
                {
                    $lastError =[ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
                    Write-Warning -Message "Failed to get window text for window $window, length is $length - error $lastError"
                }
            }
            else
            {
                $lastError =[ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
                if( $lastError.NativeErrorCode -eq 0 )
                {
                    Write-Output -InputObject "Window $window has no title"
                }
                else
                {
                    Write-Warning -Message "Failed to get windows text length for window $window - error $lastError"
                }
            }
        }
        else
        {
            $lastError =[ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
            Write-Warning -Message "Failed to get process id for window $window - error $lastError"
        }
        
        $process = Get-Process -id $windowPid -Verbose:$false ## win32_process doesn't give us window title
        [array]$processHierarchy = @(
            [array]$result = @( Get-DirectRelativeProcessDetails -id $windowPid -recurse -firstCall | Sort-Object -Property Level -Descending )
            ## now we know how many levels we can indent so the topmost process has no ident - no point waiting for all results as some may not have still existing parents so don't know what level in relation to other processes
            if( -not $noIndent -and $null -ne $result -and $result.Count -gt 1 )
            {
                $levelRange = $result | Measure-Object -Maximum -Minimum -Property Level
                ForEach( $item in $result )
                {
                    Add-Member -InputObject $item -MemberType NoteProperty -Name IndentedName -Value ("$($indenter * ($levelRange.Maximum - $item.level) * $indentMultiplier)$($item.name)")
                }
            }
            elseif( $null -ne $properties -and $properties.Count -gt 0 ) ## not indenting
            {
                $properties[ 0 ] = 'Name'
            }
            $result
        )
        [string]$text = ( "Title: `"{0}`", Text: `"{1}`"" -f $process.MainWindowTitle , $string.ToString() )
        $text = $text , ( $processHierarchy | Select-Object -Property 'IndentedName' , 'ProcessId' , 'ParentProcessId' , 'Sessionid' , '-' , 'Owner' , 'CreationDate' , 'Level' , 'Service' , @{ name = 'Signature' ; expression = { if( $null -ne $_.signing ) { $_.signing.Status }}} , 'CommandLine'  | Format-Table -AutoSize | Out-String -Width $outputWidth ) -join "`r`n"
        Write-Verbose $text
        if( $infoWindow = New-GUI -inputXAML $infowindowXAML )
        {
            $foregroundColour = [System.Windows.Media.Brushes]::Red
            ## assess how "safe" the process is
            $theProcess = [array]::IndexOf( $processHierarchy.Level , 0 )
            if( $theProcess -ge 0 )
            {
                if( $processHierarchy[ $theProcess ].signing -and $processHierarchy[ $theProcess ].Signing.Status -eq 'valid' )
                {
                    ## TODO check parents are all signed too
                    $foregroundColour  = [System.Windows.Media.Brushes]::Green
                }
                if( $processHierarchy[ $theProcess ].Name -match '^(iexplore|chrome|firefox|msedge)\.exe$' ) ## browser
                {
                    $foregroundColour = [System.Windows.Media.Brushes]::Orange
                }
            }
            ## else can't find selected process which shouldn't happen
            $WPFrichtextboxMain.Foreground = $foregroundColour
            $WPFrichtextboxMain.FontSize = $fontSize
            $WPFrichtextboxMain.AppendText( $text )
            ## put top left of dialogue where cursor is
            $infoWindow.Top = $y
            $infoWindow.Left = $x
            $infoWindow.add_Closing({
                $_.Cancel = $false
            })
            
            $openedWindows.Add( $infoWindow )
            $infoWindow.Show() ## don't want to be blocking but need to store it so we close when main window is closed
        }
    }
    else
    {
        $lastError =[ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
        Write-Warning -Message "Failed to get foreground window - error $lastError"
    }
}

$WPFDragLabel.Add_MouseLeftButtonDown({
    param($sender, $e)
    Write-Verbose -Message "Mouse left button down"
    
    ## run timer every second so we can have count down in dialogue
    $timer.Interval = [TimeSpan]::FromSeconds( 1 )
    $timer.Start()
    $script:secondsLeft = $waitTimeSeconds
    $WPFDragLabel.Content = $script:secondsLeft
})

$timer = New-Object System.Windows.Threading.DispatcherTimer

# Define the action to perform when the timer elapses
$timer.Add_Tick({
    $script:secondsLeft--
    if( $script:secondsLeft -ge 0 )
    {
        $WPFDragLabel.Content = $script:secondsLeft
    }
    else
    {
        $timer.Stop() # Stop the timer if you only want it to run once
    
        $WPFDragLabel.Content = 'Click'
        $window = $null
        $point = New-Object 'PInvoke.Win32.UserInput+POINT'
        $result = [PInvoke.Win32.UserInput]::GetCursorPos( [ref]$point )
        if( $result )
        {
            Write-Verbose "at $($point.X) $($point.Y)"
            $window = [PInvoke.Win32.UserInput]::WindowFromPoint( $point )
        }
        else
        {
            Write-Warning "Failed to get cursor position"
        }
        if( $null -ne $window )
        {
            Get-WindowDetails -window $window -X $point.X -Y $point.Y
        }
        else
        {
            Get-WindowDetails -X $point.X -Y $point.Y
        }
    }
})

# Show the WPF window
$returned = $mainwindow.ShowDialog()

ForEach( $window in $openedWindows )
{
    try
    {
        $window.Close()
    }
    catch
    {
        ## not much we can do but maybe because already been closed
    }
}
