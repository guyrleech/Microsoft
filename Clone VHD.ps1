<#

Written by Guy Leech (c) 2017

NO WARRANTY SUPPLIED. USE AT YOUR OWN RISK.

Revision History:

  21/11/15   1.00   Initial release
  04/05/17   1.10   Added -install, -uninstall, -hide and -notes options
  06/05/17   1.11   Added -startAction
  10/05/17   1.12   Now installs correctly if full path to script not specified
#>

<#
.SYNOPSIS

Copy an existing Hyper-V VHDor VHDX file to a new VHD file or create a linked clone disk from it and create a virtual machine using the new disk.

.DESCRIPTION

Either use in GUI mode, via shell integration in the registry (HKEY_CLASSES_ROOT\Windows.VhdFile), or from the command line. The memory, generation, number of processors and virtual switch may be specified otherwise defaults will be used.
it will elevate itself if not run with sufficient privileges.

.PARAMETER SourceVHD

The name of the source VHD to copy including the full path. Must not be in use.

.PARAMETER Name

The name of the new virtual machine. Must not exist in Hyper-V already or the script will terminate.

.PARAMETER Notes

Optional notes for the new VM

.PARAMETER Memory

The memory to allocate to the new VM in MB.

.PARAMETER Processors

The number of processors to allocate to the new VM.

.PARAMETER Folder

The destination folder where the new VHD will be copied to.

.PARAMETER NumberOfVMs

The number of VMs to create. The default is 1. Where more than 1 is specified, the VM name and corresponding VHD will be suffixed with " - <count>"

.PARAMETER vSwitch

The name of the virtual network switch to connect to the new virtual machine. Specify <None> for none.

.PARAMETER Generation

Defaults to 1 which is compatible with older Hyper-V versions. The only other allowed value is 2 which is newer features requiring Server 2012 or Windows 8 x64. Once set it cannot be changed.

.PARAMETER ShowUI

Will show the user interface otherwise just use the command line

.PARAMETER Boot

Boot the VM after creation. By default it won't be booted.

.PARAMETER Force

By default if the destination VHD already exists the script will fail but specifying -force will cause the VHD to be overwritten. Also used to overwrite the registry value during installation if the name already exists

.PARAMETER Hide

The parent PowerShell window will be hidden. Note that this will reduce the visbility of errors if they occur. Do not use with -wait

.PARAMETER Folder

The destination folder where the VHD will be copied to. VHD must not exist already unless the -Force option is specified to overwrite the existing VHD file in that folder.

.PARAMETER LinkedClone

With this specified, a new differencing VHD/VHDX will be created and linked to the source disk rather than copying the whole source disk.
This reduces the amount of disk space required and greatly speeds up the cloning but the source disk must always be available and unmodified.
Without it specified, a full copy of the source VHD will be made which will then be completely independent of the souce disk.

.PARAMETER AlreadyElevated

Only specify this if you do not want the script to try and elevate itself if it detects that it does not have administrative rights.

.PARAMETER Wait

Waits for <Enter> to be pressed before exiting the PowerShell window so that errors and verbose output can be inspected if required. Do not use with -hide

.PARAMETER Install

Install as right click option in explorer for vhd/vhdx files where the string specified is the name of the right click option in the menu

.PARAMETER Uninstall

Uninstalls the right click explorer integration

.PARAMETER registryKey

The registry key in which to install or uninstall the integration. Do not change this unless you need to!

.EXAMPLE

& .\Clone VHD.ps1 -SourceVHD c:\Master-Disks\Server2012.vhdx -Name "Server 2012 VM" -Memory 1024 -Processors 2 -VSwitch "External Virtual Switch" -LinkedClone -Boot -Folder "c:\Hyper-V Disks" -Verbose -Quiet

This will create a new differencing disk in the "c:\Hyper-V Disks" folder, create a new virtual machine using this disk with 1GB of memory and 2 processors and then start it.

.EXAMPLE

& .\Clone VHD.ps1 -SourceVHD c:\Master-Disks\Server2012.vhdx -ShowUI

This will show the user interface for cloning of the specified virtual disk.

.EXAMPLE

& .\Clone VHD.ps1 -Install "Clone to VM" -Generation 2 -Memory 1024 -Folder c:\Hyper-V -LinkedClone -Boot -Hide

This will install a right click option for vhd and vhdx files in explorer which will run the user interface with a default of 1GB of memory, generation 2, the disk will be a linked clone in folder c:\Hyper-V and it will be started after creation

.NOTES

If you create linked clones then ensure that the parent disk does not get modified or deleted

#>

[CmdletBinding()]

Param
(
    [Parameter(ParameterSetName='NotInstall',Mandatory=$true,HelpMessage='The source virtual disk to clone')]
	[string]$SourceVHD,
    [string]$Name ,
    [int]$Memory = 512 , # MB
    [int]$Processors = 1 ,
    [int]$numberOfVMs = 1 ,
	[ValidateSet(1,2)]
    [int]$generation = 1 ,
    [string]$Folder ,
    [string]$notes ,
    [string]$vSwitch ,
    [switch]$ShowUI ,
    [switch]$Force ,
    [switch]$AlreadyElevated , 
    [switch]$boot  ,
    [switch]$LinkedClone ,
    [switch]$Quiet  ,
    [switch]$Wait ,
    [switch]$hide ,
    [Parameter(ParameterSetName=’Install’,Mandatory=$true,HelpMessage='Name of the explorer right click menu option to add')]
    [string]$Install ,
    [string]$registryKey = 'HKCR:\Windows.VhdFile\shell' ,
    [Parameter(ParameterSetName=’Uninstall’,Mandatory=$true,HelpMessage='Uninstall explorer right click menu option')]
    [switch]$Uninstall ,
	[ValidateSet('Nothing','Start','StartIfRunning')]
    [string]$startAction = 'Nothing'
)

Function CreateNew-VM
{
    Param
    (
	    [string]$sourceVHD,
        [string]$VMName ,
        [int]$Memory ,
        [int]$Processors ,
        [int]$numberOfVMs  ,
        [int]$generation ,
        [string]$vSwitch ,
        [string]$notes ,
        [string]$Folder ,
        [switch]$boot ,
        [switch]$force ,
        [switch]$linkedClone ,
        [switch]$quiet ,
        [string]$startAction = 'Nothing'
    )
    
    Write-Verbose "Creating `"$VMName`" from `"$sourceVHD`" with $Processors CPUs and $Memory memory, switch $vSwitch in `"$Folder`" linkedClone $linkedClone start $boot startAction `"$startAction`" force $force number $numberOfVMs generation $generation"
    [int]$status = 0
    [void]$error.Clear()

    ## Figure out what type of disk it is so we keep the same type
    $extension = [System.IO.Path]::GetExtension( $sourceVHD ).ToLower()

    if( $extension -ne ".vhd" -and $extension -ne ".vhdx" )
    {
        if( ! $quiet )
        {
            [void][System.Windows.Forms.MessageBox]::Show("Unknown virtual disk format `"$extension`"" , 'Disk Cloning Error' , [System.Windows.Forms.MessageBoxButtons]::OK , [System.Windows.Forms.MessageBoxIcon]::Error )
        }
        return 87
    }

    for( $counter = 1 ; $counter -le $numberOfVMs ; $counter++ )
    {
        $thisVMName = $VMName
        if( $numberOfVMs -gt 1 )
        {
            $thisVMName += " - $($counter)"
        }
        $thisDisk = "$Folder\$thisVMName$extension"
        $activity = "Creating VM #$counter/$numberOfVMs"

        if( Get-VM -Name $thisVMName -EA SilentlyContinue )
        {
            if( ! $quiet )
            {
                [void][System.Windows.Forms.MessageBox]::Show("VM `"$thisVMName`" already exists" , 'Disk Cloning Error' , [System.Windows.Forms.MessageBoxButtons]::OK , [System.Windows.Forms.MessageBoxIcon]::Error )
            }
            return 80
        }
        elseif( ( Test-Path -Path $thisDisk ) -And -Not $force )
        {
            if( ! $quiet )
            {
                [void][System.Windows.Forms.MessageBox]::Show("Disk `"$thisDisk`" already exists in destination", 'Disk Cloning Error' , [System.Windows.Forms.MessageBoxButtons]::OK , [System.Windows.Forms.MessageBoxIcon]::Error )
            }
            return 80
        }
        else
        {
            Write-Progress -id 1 -Activity $activity -Status 'Copying disk' -PercentComplete 50
            try
            {
                if( $linkedClone )
                {
                    $null = New-VHD -Path $thisDisk -ParentPath $sourceVHD -Differencing -EA Stop
                }
                else
                {
                    Copy-Item $sourceVHD $thisDisk -EA Stop
                    ## May be read-only if source was so flick it off
                    Set-ItemProperty -Path $thisDisk -Name IsReadOnly -Value $false
                }
            }
            catch
            {
                $status = $?
                if( ! $quiet )
                {
                     [void][System.Windows.Forms.MessageBox]::Show("Failed to copy `"$sourceVHD`" to `"$thisDisk`"`n$($error[0].Exception.Message)", 'Disk Cloning Error' , [System.Windows.Forms.MessageBoxButtons]::OK , [System.Windows.Forms.MessageBoxIcon]::Error )
                }
                return $status
            }

            Write-Progress -id 1 -Activity $activity -Status 'Creating VM'   -PercentComplete 75
            ## Now create the VM
            $switchargs = @{}

            if( ! [String]::IsNullOrEmpty($vSwitch) -And $vSwitch -ne '<None>' )
            {
                $switchargs = @{ 'switch' = "$vswitch" }
            }
            if( ( $vm = New-VM -Name $thisVMName -VHDPath $thisDisk -MemoryStartupBytes ($memory * 1024 * 1024) -Generation $generation @switchargs))
            {
                if( $Processors -gt 1 )
                {
                    Write-Progress -id 1 -Activity $activity -Status 'Adding processors' -PercentComplete 85
                    if( ! ( $vm | Set-VM -ProcessorCount $Processors ) -And ( (Get-VM -Name $thisVMName).ProcessorCount -ne $Processors ))
                    {
                        $status = $?
                        $failed = $true
                        if( ! $quiet )
                        {
                            [void][System.Windows.Forms.MessageBox]::Show("Failed to set processors to $Processors`n$($error[0].Exception.Message)", 'Disk Cloning Error' , [System.Windows.Forms.MessageBoxButtons]::OK , [System.Windows.Forms.MessageBoxIcon]::Error )
                        }
                    }
                }
                if( ! [string]::IsNullOrEmpty( $notes ) )
                {
                    Write-Progress -id 1 -Activity $activity -Status 'Adding notes' -PercentComplete 90
                    if( ! ( $vm | Set-VM -Notes $notes ) -And ( (Get-VM -Name $thisVMName).Notes -ne $notes ))
                    {
                        $status = $?
                        $failed = $true
                        if( ! $quiet )
                        {
                            [void][System.Windows.Forms.MessageBox]::Show("Failed to set notes to `"$notes`"`n$($error[0].Exception.Message)", 'Disk Cloning Error' , [System.Windows.Forms.MessageBoxButtons]::OK , [System.Windows.Forms.MessageBoxIcon]::Error )
                        }
                    }
                }
                $vm | Set-VM -AutomaticStartAction $startAction
                if( ! $failed )
                {
                    if( $boot )
                    {
                        Write-Progress -id 1 -Activity $activity -Status 'Starting VM' -PercentComplete 95
                        if( ! ( Start-VM -Name $thisVMName ) -And ( (Get-VM -Name $thisVMName).State -eq 'Off' ) )
                        {
                            $status = $?
                            if( ! $quiet )
                            {
                              [void][System.Windows.Forms.MessageBox]::Show("Failed to start `"$thisVMName`"`n$($error[0].Exception.Message)", 'Disk Cloning Error' , [System.Windows.Forms.MessageBoxButtons]::OK , [System.Windows.Forms.MessageBoxIcon]::Error )
                            }
                        }
                    }
                    Write-Progress -id 1 -Activity $activity -Status 'VM created' -PercentComplete 100
                    if( ! $quiet )
                    {
                        [void][System.Windows.Forms.MessageBox]::Show("Disk $counter/$numberOfVMs cloned to VM OK", "Disk Cloning" , [System.Windows.Forms.MessageBoxButtons]::OK , [System.Windows.Forms.MessageBoxIcon]::Information )
                    }
                }
            }
            else
            {
                $status = $?
                if( ! $quiet )
                {
                     [void][System.Windows.Forms.MessageBox]::Show("Failed to create `"$thisVMName`"`n$($error[0].Exception.Message)", 'Disk Cloning Error' , [System.Windows.Forms.MessageBoxButtons]::OK , [System.Windows.Forms.MessageBoxIcon]::Error )
                }
            }
        }
    }
    ## Need to return status so caller can take action if required
    return $status
}
## Also need this for message boxes
[void][reflection.assembly]::LoadWithPartialName( 'System.Windows.Forms' )

## Check we have cmdlets we need
try
{
    $null = Get-Command New-VM -ErrorAction Stop
    $null = Get-Command Get-VM -ErrorAction Stop
    if( $linkedClone )
    {
        $null = Get-Command New-VHD -Path -ErrorAction Stop
    }
    if( $boot )
    {
        $null = Get-Command Start-VM -Path -ErrorAction Stop
    }
    if( $Processors -gt 1 -or ! [string]::IsNullOrEmpty( $notes ) )
    {
        $null = Get-Command Set-VM -Path -ErrorAction Stop
    }
}
catch
{
    [void][System.Windows.Forms.MessageBox]::Show('Hyper-V cmdlets missing', 'Disk Cloning Error' , [System.Windows.Forms.MessageBoxButtons]::OK , [System.Windows.Forms.MessageBoxIcon]::Error )
    Exit 10
}

if( -Not $alreadyElevated )
{
    ## Needs to run elevated which it won't from right click in Explorer if UAC Is enabled and not running as an admin (which is sensible!)
    # Get the ID and security principal of the current user account
    $myWindowsPrincipal = new-object System.Security.Principal.WindowsPrincipal([System.Security.Principal.WindowsIdentity]::GetCurrent())
 
    # Get the security principal for the Administrator role
    $adminRole = [System.Security.Principal.WindowsBuiltInRole]::Administrator
 
    # Check to see if we are currently running "as Administrator"
    if ( ! ($myWindowsPrincipal.IsInRole($adminRole))  )
     {
       # We are not running "as Administrator" - so relaunch as administrator
   
       # Create a new process object that starts PowerShell - replace " with ' else gets stripped out and spaces muck things up
       [string]$Arguments = $script:MyInvocation.Line.Replace( "`"" , "'" ) + " -alreadyElevated"
       if( $hide )
       {
            $Arguments = '-WindowStyle Hidden ' + $Arguments
       }
       $wait = $true
       if( $ShowUI )
       {
            $wait = $false
       }
       $child = Start-Process -Wait:$wait -PassThru -FilePath powershell.exe -Verb Runas -ArgumentList $Arguments

       if( ! $? -or ! $child )
       {
            [void][System.Windows.Forms.MessageBox]::Show( 'Failed to relaunch script elevated', 'Disk Cloning Error' , [System.Windows.Forms.MessageBoxButtons]::OK , [System.Windows.Forms.MessageBoxIcon]::Error )
            Exit 1
       }
       else
       {
            Exit $child.ExitCode
       }
    }
}

[int]$exitCode = 0

if( ! [string]::IsNullOrEmpty( $Install ) -or $Uninstall )
{
    [string]$baseKey = ($registryKey -split ':')[0]
    [string]$ourscriptFullPath = & { $myInvocation.ScriptName } 
    [string]$ourscript = Split-Path $ourscriptFullPath   -Leaf
    $drive = $null
    if( ! ( Test-Path ( $baseKey + ':' ) -ErrorAction SilentlyContinue ) )
    {
        [string]$root = switch( $baseKey )
        {
            'HKCR' { 'HKEY_CLASSES_ROOT' }
            'HKU'  { 'HKEY_USERS' }
            default { 
                [void][System.Windows.Forms.MessageBox]::Show("Unknown registry root $baseKey", 'Disk Cloning Error' , [System.Windows.Forms.MessageBoxButtons]::OK , [System.Windows.Forms.MessageBoxIcon]::Error )
                Exit 11
            }
        }
        $drive = New-PSDrive -Name $baseKey -PSProvider Registry -Root $root
        if( ! $? -or ! $drive )
        {
            [void][System.Windows.Forms.MessageBox]::Show("Failed to create registry drive $root", 'Disk Cloning Error' , [System.Windows.Forms.MessageBoxButtons]::OK , [System.Windows.Forms.MessageBoxIcon]::Error )
            Exit 12
        }
    }
    if( $Uninstall )
    {
        ## Find our script in the registry
        [bool]$removed = $false
        [string]$found = $null
        ## Need to get all keys (should only be one) before we start deleting otherwise the enumeration errors
        $keys = Get-ChildItem -Path $registryKey -Recurse | ?{ ( Get-ItemProperty -Path $_.PSPath -Name '(Default)' -ErrorAction SilentlyContinue ) -match $ourscript } 
        if( $keys )
        {
            $keys |  %{
            $_.PSParentPath
                Remove-Item -Path $_.PSParentPath -Recurse -Force 
                if( $? )
                {
                    $removed = $true
                }
                $found += ($_.PSParentPath -split '::')[1]           
            }
        }
        if( ! $removed )
        {
            [string]$message = ""
            if( $found )
            {
                $message = "Failed to remove key `"$found`""
            }
            else
            {
                $message = "Failed to find script `"$ourscript`" in key `"$registryKey`""
            }
            [void][System.Windows.Forms.MessageBox]::Show( $message ,  'Disk Cloning Error' , [System.Windows.Forms.MessageBoxButtons]::OK , [System.Windows.Forms.MessageBoxIcon]::Error )
            Exit 16
        }
              
    }
    else ## install
    {
        [string]$ourKey = $registryKey + '\' + $Install + '\command'
        if( Test-Path $ourKey )
        {
            if( ! $force )
            {
                [void][System.Windows.Forms.MessageBox]::Show("Registry key $ourKey already exists - use -force to overwrite", 'Disk Cloning Error' , [System.Windows.Forms.MessageBoxButtons]::OK , [System.Windows.Forms.MessageBoxIcon]::Error )
                Exit 13
            }
        }
        elseif( ! ( New-Item -Path ( Split-Path -Path $ourKey -Parent ) -Name ( Split-Path -Path $ourKey -Leaf ) -Force ) )
        {
            [void][System.Windows.Forms.MessageBox]::Show("Failed to create registry key $ourKey", 'Disk Cloning Error' , [System.Windows.Forms.MessageBoxButtons]::OK , [System.Windows.Forms.MessageBoxIcon]::Error )
            Exit 14
        }
        ## Get script arguments but remove -install "argument". Also need to change the script to the absolute path as relative may not work
        [string]$matcher = "(?<script>.*$($ourScript -replace '\.' , '\.'))'?(?<before>.*)(?<install>-install[\s\w'`"]+)(?<after>.*$)"
        if( $script:MyInvocation.Line.Replace( "`"" , "'" ) -match $matcher )
        {
            ## Escape double quotes and remove -AlreadyElevated if present 
            [string]$newValue = 'powershell.exe -ExecutionPolicy RemoteSigned'
            if( $hide )
            {
                $newValue += ' -WindowStyle Hidden '
            }
            $newValue += " -Command `"& '"+ $ourscriptFullPath + "'" + ( ( $Matches[ 'before' ] + ' ' + $Matches[ 'After' ] ) -replace '"' , "\`"" -replace '-Al[a-z]*' , '' -replace '-For[a-z]*' , '' ) + ' -sourceVHD \"%1\""'
            if( $newValue -notmatch '-ShowUI' )
            {
                $newValue += ' -ShowUI'
            }
            Set-ItemProperty -Path $ourKey -name '(Default)' -Value $newValue -Force
            if( ! $? )
            {
                [void][System.Windows.Forms.MessageBox]::Show("Failed to set registry value in $ourKey", 'Disk Cloning Error' , [System.Windows.Forms.MessageBoxButtons]::OK , [System.Windows.Forms.MessageBoxIcon]::Error )
                Exit 15
            }
        }
        else
        {
            [void][System.Windows.Forms.MessageBox]::Show("Failed to parse command line", 'Disk Cloning Error' , [System.Windows.Forms.MessageBoxButtons]::OK , [System.Windows.Forms.MessageBoxIcon]::Error )
            Exit 16
        }
    }
    if( $drive )
    {
        Remove-PSDrive $drive
    }
    return
}

if( $showUI )
{
    [int]$y = 20

    $form = New-Object Windows.Forms.Form
    $form.Size = New-Object System.Drawing.Size(350,100)
    $form.MaximumSize = New-Object System.Drawing.Size(350,1200)
    $form.AutoSize = $true
    $form.text = "Hyper-V Disk Cloner"
    
    $label0 = New-Object Windows.Forms.Label
    $label0.Location = New-Object Drawing.Point 50,$y
    $label0.Size = New-Object Drawing.Point 250,30
    $label0.Font = New-Object System.Drawing.Font($label0.Font, [System.Drawing.FontStyle]::Bold);
    $label0.text = "Cloning '" + (Split-Path $sourceVHD -Leaf) + "'"

    $label1 = New-Object Windows.Forms.Label
    $label1.Location = New-Object Drawing.Point 50,$($y+=30;$y)
    $label1.Size = New-Object Drawing.Point 250,30 
    $label1.text = "New &Machine Name"

    $textfield1 = New-Object Windows.Forms.TextBox
    $textfield1.Location = New-Object Drawing.Point 50,$($y+=30;$y)
    $textfield1.Size = New-Object Drawing.Point 210,15
    
    $label11 = New-Object Windows.Forms.Label
    $label11.Location = New-Object Drawing.Point 50,$($y+=30;$y)
    $label11.Size = New-Object Drawing.Point 250,30 
    $label11.text = "N&otes"

    $textfield11 = New-Object Windows.Forms.TextBox
    $textfield11.Location = New-Object Drawing.Point 50,$($y+=30;$y)
    $textfield11.Size = New-Object Drawing.Point 210,15

    $label2 = New-Object Windows.Forms.Label
    $label2.Location = New-Object Drawing.Point 50,$($y+=30;$y)
    $label2.Size = New-Object Drawing.Point 250,30 
    $label2.text = "&RAM (MB)"

    $textfield2 = New-Object Windows.Forms.TextBox
    $textfield2.Location = New-Object Drawing.Point 50,$($y+=30;$y)
    $textfield2.Size = New-Object Drawing.Point 50,15
    $textfield2.Text = $Memory

    $label3 = New-Object Windows.Forms.Label
    $label3.Location = New-Object Drawing.Point 50,$($y+=30;$y)
    $label3.Size = New-Object Drawing.Point 250,30 
    $label3.text = "&Processors"

    $textfield3 = New-Object Windows.Forms.TextBox
    $textfield3.Location = New-Object Drawing.Point 50,$($y+=30;$y)
    $textfield3.Size = New-Object Drawing.Point 50,15 
    $textfield3.Text = $Processors

    $label4 = New-Object Windows.Forms.Label
    $label4.Location = New-Object Drawing.Point 50,$($y+=30;$y)
    $label4.Size = New-Object Drawing.Point 250,30 
    $label4.text = "&Disk Destination Folder"

    $textfield4 = New-Object Windows.Forms.TextBox
    $textfield4.Location = New-Object Drawing.Point 50,$($y+=30;$y)
    $textfield4.Size = New-Object Drawing.Point 210,15 
    $textfield4.Text = $Folder

    $button = New-Object Windows.Forms.Button
    $button.text = "..."
    $button.Size = New-Object System.Drawing.Size(25,25)
    $button.Location = New-Object Drawing.Point 270,$y ## Don't increment as want on same line

    $button.add_click(
    {
        $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
 
        [void]$FolderBrowser.ShowDialog()

        $textfield4.text = $FolderBrowser.SelectedPath
    })

    $form.add_FormClosing(
    {
        ##Write-Verbose "Closing size is $($form.Size)"
    })

    $label5 = New-Object Windows.Forms.Label
    $label5.Location = New-Object Drawing.Point 50,$($y+=30;$y)
    $label5.Size = New-Object Drawing.Point 250,30 
    $label5.text = "&Virtual Switch:"
    
    $objComboBox = New-Object System.Windows.Forms.ComboBox
    $objComboBox.Location = New-Object System.Drawing.Size(50,$($y+=30;$y)) 
    $objComboBox.Size = New-Object System.Drawing.Size(210,200) 
    $objComboBox.Height = 80
    
    [int]$selectedIndex = 0
    [int]$index = 0

    Get-VMSwitch | %{
        [void] $objComboBox.Items.Add($_.Name) 
        if( $_.Name -match $vSwitch )
        {
            $selectedIndex = $index 
        }
        $index++ 
    }

    [void]$objComboBox.Items.Add("<None>")
    
    $objComboBox.SelectedIndex = $selectedIndex
    
    $label6 = New-Object Windows.Forms.Label
    $label6.Location = New-Object Drawing.Point 50,$($y+=30;$y)
    $label6.Size = New-Object Drawing.Point 250,30 
    $label6.text = "&Number of VMs"

    $textfield6 = New-Object Windows.Forms.TextBox
    $textfield6.Location = New-Object Drawing.Point 50,$($y+=30;$y)
    $textfield6.Size = New-Object Drawing.Point 50,15 
    $textfield6.Text = $NumberOfVMs 

    ## Generation
    $label7 = New-Object Windows.Forms.Label
    $label7.Location = New-Object Drawing.Point 50,$($y+=30;$y)
    $label7.Size = New-Object Drawing.Point 250,30 
    $label7.text = "&Generation"
    
    $objComboBox2 = New-Object System.Windows.Forms.ComboBox
    $objComboBox2.Location = New-Object System.Drawing.Size(50,$($y+=30;$y)) 
    $objComboBox2.Size = New-Object System.Drawing.Size(50,200) 
    $objComboBox2.Height = 80

    [void]$objComboBox2.Items.Add("1")
    [void]$objComboBox2.Items.Add("2")
    
    switch( $generation )
    {
        1 { $objComboBox2.SelectedIndex = 0 }
        2 { $objComboBox2.SelectedIndex = 1 }
        default { [void][System.Windows.Forms.MessageBox]::Show("Generation can only be 1 or 2", 'Disk Cloning Error' , [System.Windows.Forms.MessageBoxButtons]::OK , [System.Windows.Forms.MessageBoxIcon]::Error ); Exit 87 }
    }

    ## Boot Action
    $label77 = New-Object Windows.Forms.Label
    $label77.Location = New-Object Drawing.Point 50,$($y+=30;$y)
    $label77.Size = New-Object Drawing.Point 250,30 
    $label77.text = "Automatic Start Action"
    
    $objComboBox22 = New-Object System.Windows.Forms.ComboBox
    $objComboBox22.Location = New-Object System.Drawing.Size(50,$($y+=30;$y)) 
    $objComboBox22.Size = New-Object System.Drawing.Size(120,200) 
    $objComboBox22.Height = 80

    [void]$objComboBox22.Items.Add("Nothing")
    [void]$objComboBox22.Items.Add("Start")
    [void]$objComboBox22.Items.Add("Start if Running")
    
    switch( $startAction )
    {
        'Nothing' { $objComboBox22.SelectedIndex = 0 }
        'Start' { $objComboBox22.SelectedIndex = 1 }
        'StartIfRunning' { $objComboBox22.SelectedIndex = 2 }
    }

    ## Tick boxes
    $checkbox1 = new-object System.Windows.Forms.checkbox
    $checkbox1.Location = new-object System.Drawing.Size(50,$($y+=30;$y))
    $checkbox1.Size = new-object System.Drawing.Size(250,50)
    $checkbox1.Text = "&Start VM after creation"
    $checkbox1.Checked = $boot
    
    $checkbox2 = new-object System.Windows.Forms.checkbox
    $checkbox2.Location = new-object System.Drawing.Size(50,$($y+=40;$y))
    $checkbox2.Size = new-object System.Drawing.Size(250,50)
    $checkbox2.Text = "&Linked clone"
    $checkbox2.Checked = $linkedClone
    
    $button2 = New-Object Windows.Forms.Button
    $button2.text = "&Create"
    $button2.Location = New-Object Drawing.Point 140,$($y+=40;$y)
    
    $button2.add_click(
    {
        [bool]$valid = $true

        if( [String]::IsNullOrEmpty( $textfield1.Text ) )
        {
            $valid = $false
            [void][System.Windows.Forms.MessageBox]::Show("Must specify a valid machine name", 'Disk Cloning Error' , [System.Windows.Forms.MessageBoxButtons]::OK , [System.Windows.Forms.MessageBoxIcon]::Error )
        }
        ## Could do more validation of numeric fields if desired
        if( $valid )
        {
            $button2.Text = 'Creating ...'
            $cursor = $form.Cursor
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $exitcode = CreateNew-VM -sourceVHD $sourceVHD -VMName $textfield1.Text -Notes $textfield11.Text -Memory ( $textfield2.Text -as [Int]) -Processors $textfield3.Text -Folder $textfield4.Text -vSwitch $objComboBox.SelectedItem -Generation $objComboBox2.SelectedItem -startAction $objComboBox22.SelectedItem -boot:$checkbox1.Checked -force:$force -linkedClone:$checkbox2.Checked -Quiet:$quiet -Number $textField6.Text -Wait:$wait
            $form.Cursor = $cursor
            $form.Close()
        }
    })
    
    $form.Size = New-Object System.Drawing.Size(350,$($y+=60;$y)) ## Allow for Create button and space below

    # Ensure add in the correct order to ensure tab order maintained
    $form.controls.add($label0)
    $form.controls.add($button2)
    $form.controls.add($label1)
    $form.controls.add($textfield1)
    $form.controls.add($label11)
    $form.controls.add($textfield11)
    $form.controls.add($label2)
    $form.controls.add($textfield2)
    $form.controls.add($label3)
    $form.controls.add($textfield3)
    $form.controls.add($label4)
    $form.controls.add($textfield4)
    $form.controls.add($button)
    $form.Controls.Add($label5)  
    $form.Controls.Add($objComboBox) 
    $form.controls.add($label6)
    $form.controls.add($textfield6)
    $form.Controls.Add($label7)  
    $form.Controls.Add($objComboBox2) 
    $form.Controls.Add($label77)  
    $form.Controls.Add($objComboBox22) 
    $form.Controls.Add($checkbox1)   
    $form.Controls.Add($checkbox2)  

    $form.KeyPreview = $True
    $form.Add_KeyDown({
        if ($_.KeyCode -eq 'Enter') 
        {
            $button2.PerformClick()
        }
    })
    $form.Add_KeyDown({
        if ($_.KeyCode -eq 'Escape') 
        {
            $form.Close()
        }
     })

    ## Set focus on first text field
    $form.Add_Shown( {
        $form.Activate() ; $textfield1.Focus()
    })

    # DISPLAY DIALOG
    [void]$form.ShowDialog()
}
else
{
    if( [String]::IsNullOrEmpty( $Name ) )
    {
        Write-Error 'Must specify a name via the -name option'
        $exitcode = 87 ##The parameter is incorrect.
    }
    elseif( [String]::IsNullOrEmpty( $Folder ) )
    {
        Write-Error 'Must specify a destination folder via the -folder option'
        $exitcode = 87 ##The parameter is incorrect.
    }
    else
    {
        $exitCode = CreateNew-VM -sourceVHD $sourceVHD -VMName $Name -Notes $notes -Memory $Memory -Processors $Processors -Folder $Folder -vSwitch $vSwitch -start:$start -force:$force -linkedClone:$linkedClone -Quiet:$quiet -Number $NumberOfVMs -generation $generation -wait:$wait
    }
}

if( $Wait )
{
    ## So any errors/verbose output can be read before exit
    Read-Host 'Hit enter to quit - '
}

Exit $exitCode
