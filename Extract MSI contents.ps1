<#
.SYNOPSIS
    Take an MSI file and extract its contents to a specified folder using a WPF GUI.
.DESCRIPTION
    Designed to be used in Explorer send to context menu, this script provides a simple WPF GUI that allows users to select an MSI file and extract its contents to a specified folder.   
.PARAMETER MSIfile
    Path to the MSI file to extract
.PARAMETER OutputFolder
    The folder where the MSI contents will be extracted. If not specified, the script will prompt for a folder using a GUI.
.PARAMETER install
    If specified, the script will add a right-click context menu entry in Windows Explorer for MSI files, allowing you to run this script directly from Explorer.

.NOTES
    Modification History:
    2025/06/26  @guyrleech  Script born
#>

[CmdletBinding(DefaultParameterSetName='Extract')]

Param
(
    [Parameter(Mandatory = $true,Position = 0 ,ParameterSetName = 'Extract')]
    [string]$MSIfile ,
    [Parameter(Mandatory = $false,ParameterSetName = 'Extract')]
    [string]$OutputFolder ,
    [Parameter(Mandatory = $true,ParameterSetName = 'Install')]
    [switch]$install
)

[string]$scriptPath = & { $MyInvocation.ScriptName}

if( $install )
{
    $menuName = 'ExtractMSI'
    $menuText = 'Extract MSI Contents'
    $command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -MSIfile `"%1`""

    New-Item -Path "Registry::HKEY_CLASSES_ROOT\msi.package\shell\$menuName" -Force | Out-Null
    Set-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\msi.package\shell\$menuName" -Name '(Default)' -Value $menuText
    New-Item -Path "Registry::HKEY_CLASSES_ROOT\msi.package\shell\$menuName\command" -Force | Out-Null
    Set-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\msi.package\shell\$menuName\command" -Name '(Default)' -Value $command
    exit 0
}

[console]::Title = "Extracting from $MSIfile at $([datetime]::Now.ToString('G'))"
[string]$persistentKey = "HKCU:\Software\Guy Leech"
$MSIfile = $MSIfile.Trim('"').Trim()
if (-not (Test-Path -Path $MSIfile -PathType Leaf)) {
    Write-Error "The specified MSI file does not exist: $MSIfile"
    exit 1
}
if( $MSIfile -notmatch '\.msi$' ) {
    Write-Error "The specified file is not an MSI file: $MSIfile"
    exit 1
}

# XAML definition for the GUI
[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="MSI Contents Extractor" Height="240" Width="500"
    WindowStartupLocation="CenterScreen"
    ResizeMode="CanMinimize">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/> <!-- Browse button -->
            <RowDefinition Height="Auto"/> <!-- Label -->
            <RowDefinition Height="Auto"/> <!-- TextBox -->
            <RowDefinition Height="*"/>    <!-- Spacer/expanding area -->
            <RowDefinition Height="Auto"/> <!-- Buttons -->
        </Grid.RowDefinitions>
        
        <Button Name="BrowseButton" Grid.Row="0" Content="Browse for Folder" 
                Height="30" Margin="0,0,0,10" FontSize="12"/>
        <Label Grid.Row="1" Content="Selected Folder:" FontWeight="Bold" Margin="0,0,0,5"/>
        <TextBox Name="PathTextBox" Grid.Row="2" Height="25" IsReadOnly="False" 
                 Background="LightGray" Margin="0,0,0,15" VerticalContentAlignment="Center"/>
        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,0,0,10">
            <Button Name="OkButton" Content="OK" Height="35" Width="80"
                    FontSize="14" FontWeight="Bold" IsEnabled="False" Margin="0,0,10,0"/>
            <Button Name="CancelButton" Content="Cancel" Height="35" Width="80"
                    FontSize="14" FontWeight="Bold"/>
        </StackPanel>
    </Grid>
</Window>
"@

if( [string]::IsNullOrEmpty($OutputFolder) )
{
    [string]$previousOutputFolder = Get-ItemPropertyValue -Path $persistentKey -Name 'LastExtractedFolder' -ErrorAction SilentlyContinue
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName System.Windows.Forms

    # Create the WPF window
    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)

    # Get references to controls
    $browseButton = $window.FindName("BrowseButton")
    $pathTextBox = $window.FindName("PathTextBox")
    $okButton = $window.FindName("OkButton")
    $cancelButton = $window.FindName("CancelButton")


    # Variable to store selected folder path
    $selectedPath = ""

    if( -not [string]::IsNullOrEmpty($previousOutputFolder) -and (Test-Path -Path $previousOutputFolder -PathType Container)) {
        $selectedPath = $previousOutputFolder
        $pathTextBox.Text = $selectedPath
        $okButton.IsEnabled = $true
    } else {
        $pathTextBox.Text = ""
        $okButton.IsEnabled = $false
    }

    $window.Add_PreviewKeyDown({
            param($keySender, $e)
            if ($e.Key -eq 'Enter' -or $e.Key -eq 'Return') {
                if ($okButton.IsEnabled) {
                    $okButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                    $e.Handled = $true
                }
            } elseif ($e.Key -eq 'Escape') {
                $cancelButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                $e.Handled = $true
            }
        })

    $browseButton.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "Select a folder"
        $folderBrowser.ShowNewFolderButton = $true
        
        # Set initial directory to user's Documents folder
        if( -not [string]::IsNullOrEmpty($previousOutputFolder) -and (Test-Path -Path $previousOutputFolder -PathType Container)) {
            $folderBrowser.SelectedPath = $previousOutputFolder
        } else {
            $folderBrowser.SelectedPath = [Environment]::GetFolderPath("MyDocuments")
        }
        
        $result = $folderBrowser.ShowDialog()
        
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $script:selectedPath = $folderBrowser.SelectedPath
            $pathTextBox.Text = $script:selectedPath
            $okButton.IsEnabled = $true
            
            Write-Host "Folder selected: $script:selectedPath" -ForegroundColor Green
        }
    })

    # OK button click event
    $okButton.Add_Click({
        if (-not [string]::IsNullOrEmpty($script:selectedPath))
        {
            if( Test-Path -Path $Script:selectedPath -PathType Container )
            {
                $contents = $null
                $contents = @( Get-ChildItem -Path $script:selectedPath -ErrorAction SilentlyContinue)
                if( $null -ne $contents -and $contents.Count -gt 0 )
                {
                    $message = "The folder '$script:selectedPath' is not empty. Do you want to continue?"
                    $caption = "Confirm Folder Usage"
                    $result = [System.Windows.MessageBox]::Show($message, $caption, [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
                    if( $result -eq [System.Windows.MessageBoxResult]::No )
                    {
                        return
                    }
                }
            }
            $window.DialogResult = $true
            $window.Close()
        }
    })

    $pathTextBox.Add_TextChanged({
    $currentPath = $pathTextBox.Text
    ## path doesn't have to exist
    if ( -not [string]::IsNullOrWhiteSpace($currentPath) -and ( $currentPath -match '^[a-z]:\\' -or $currentPath -match '^\\\\[^\\]+\\\w+' )) {
        $okButton.IsEnabled = $true
        $script:selectedPath = $currentPath
    } else {
        $okButton.IsEnabled = $false
    }
})
    # Cancel button click event
    $cancelButton.Add_Click({
        $window.DialogResult = $false
        $window.Close()
    })

    # Show the window
    $returned = $window.ShowDialog() 
    if( $returned )
    {
        $OutputFolder = $script:selectedPath
    }
}

if( -Not [string]::IsNullOrEmpty($OutputFolder) )
{    
    $process = $null
    $process = Start-Process -FilePath msiexec.exe -ArgumentList @( 
        '/a' , 
        "`"$msiFile`"" , 
        '/qn' ,
        "TARGETDIR=`"$OutputFolder`"" ) -wait -PassThru
    if( $null -eq $process )
    {
        Write-Error "Failed to start msiexec.exe"
        exit 1
    }
    elseif( $process.ExitCode -ne 0 )
    {
        Write-Error "msiexec.exe failed with exit code $($process.ExitCode)"
        exit 1
    }
    else
    {
        $result = [System.Windows.MessageBox]::Show("Open $outputFolder ?", "Contents Extracted OK", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Information)
        if( $result -eq [System.Windows.MessageBoxResult]::Yes )
        {
            Start-Process -FilePath $OutputFolder -Verb Open
        }
        if( -Not [string]::IsNullOrEmpty($persistentKey) )
        {           
            if( -Not ( Test-Path -Path $persistentKey -PathType Container ) )
            {
                New-Item -Path $persistentKey -Force | Out-Null
            }
            Set-ItemProperty -Path $persistentKey -Name 'LastExtractedFolder' -Value $OutputFolder -Force
        }
    }
}