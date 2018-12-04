<#
    List/add/remove autorun entries in start menu or registry for user running script or all users

    @guyrleech, 2018
#>

<#
.SYNOPSIS

List/add/remove autorun entries, so items run at logon, in start menu or registry, for user running the script or all users.
Can also work on the RunOnce key instead if required.

.DESCRIPTION

Can remove named entries or search entries for a string such as the executable name.
You will be prompted before anything is removed unless you specify -confirm:$false on the command line.

.PARAMETER allusers

By default the script operates on the invoking user's start menu or HKCU. Using this option changes the behaviour to operate on the all users start menu and HKLM.
To remove entries the script must be run with adequate rights.

.PARAMETER registry

By default the script operates on the start menu. Using this option changes the behaviour to operate on HKCU or HKLM depdending if the -allusers options is specified too.

.PARAMETER wow6432node

When in registry mode on 64 bit system, operate on the 32 bit entries instead of the 64 bit ones.

.PARAMETER run

The command line to run including any parameters. If there are spaces in the command then encase in double quote characters.
Actually a regular expression so partial matches are allowed when used with -list or -remove.

.PARAMETER name

The name of the registry value or shortcut to be operated on. Actually a regular expression so partial matches are allowed when used with -list or -remove.

.PARAMETER remove

Remove the entries that match the -name or -run options specified.

.PARAMETER list

Show the autoruns entries in the file system or registry. Use with -name to show a specific one or with -run to match some or all of a command line.

.PARAMETER runOnce

Process the "RunOnce" registry key instead of the "Run" key
 
.PARAMETER verify

Verify the executable exists when creating a new entry and error if it does not.

.PARAMETER icon

Optional icon file and offset to be used when creating a shortcut, e.g. "c:\windows\system32\shell32.dll,10"

.PARAMETER description

Optional description/comment to be used when creating a shortcut

.EXAMPLE

& "C:\Scripts\Autorun.ps1" -list -registry -allusers -run evernote

Show all autorun entries in the registry for all users that contain the string "evernote"

.EXAMPLE

& "C:\Scripts\Autorun.ps1" -name "Edit Hosts File" -run "notepad.exe c:\windows\system32\drivers\etc\hosts" -description "Look at the hosts files"

Add a shortcut for the user called "Edit Hosts File" which will run the given command line

.EXAMPLE

& "C:\Scripts\Autorun.ps1" -name "Edit Hosts File" -remove

Remove the shortcut for the user called "Edit Hosts File" 

.NOTES

Will work when start menus are redirected.

#>

[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='High')]

Param
(
    [switch]$allusers ,
    [switch]$registry ,
    [switch]$wow6432node ,
    [string]$run ,
    [switch]$remove ,
    [switch]$list ,
    [switch]$runOnce ,
    [string]$name ,
    [string]$icon ,
    [string]$description ,
    [switch]$verify
)

if( $runOnce )
{
    $registry = $True
}

if( $wow6432node )
{
    $registry = $True
    $allusers = $True
}

[int]$count = 0

$items = New-Object System.Collections.ArrayList

if( $registry )
{
    $regKey = '\Microsoft\Windows\CurrentVersion\Run'
    if( $runOnce )
    {
        $regKey += 'Once'
    }
    if( $wow6432node )
    {
        $regKey = '\Software\Wow6432Node' + $regkey
    }
    else
    {
        $regKey = '\Software' + $regkey
    }

    if( $allusers )
    {
        $regKey = 'HKLM:' + $regKey
    }
    else
    {
        $regKey = 'HKCU:' + $regKey
    }

    if( $remove -or $list )
    {
        if( $list -or ! [string]::IsNullOrEmpty( $run ) -or ! [string]::IsNullOrEmpty( $name ) ) 
        {
            Get-Item -Path $regKey | Select-Object -ExpandProperty property | ForEach-Object `
            { 
                [string]$value = (Get-ItemProperty -Path $regKey -Name $_).$_
                if( $value -match $run -and $_ -match $name )
                {
                    if( $list )
                    {
                        [void]$items.Add( [pscustomobject][ordered]@{
                            'Key' = $regKey
                            'Value Name' = $_
                            'Value Data' = $value })
                    }
                    elseif( $pscmdlet.ShouldProcess( "Value `"$_`"" , "Remove `"$value`"" ) )
                    {
                        Remove-ItemProperty -Path $regKey -Name $_
                        if( $? )
                        {
                            $count++
                        }
                    } 
                }
            }
        }
        else
        {
            Write-Error "Must specify -name or -run when removing an entry"
            return
        }
    }
    else ## adding
    {
        if( [string]::IsNullOrEmpty( $name ) )
        {
            Write-Error "Must specify the name of the registry value via the -name option"
            return
        }
        ## See if it exists already
        [string]$existing = Get-ItemProperty -Path $regKey -Name $name -ErrorAction SilentlyContinue | Select -ExpandProperty $name
        [bool]$carryOn = $true
        if( ! [string]::IsNullOrEmpty( $existing ) )
        {
            if( $existing -eq $run -or ! $pscmdlet.ShouldProcess( "`"$name`" already exists as `"$existing`"" , "Overwrite" ) )
            {
                $carryOn = $false
            }
        }
        if( $carryOn )
        {
            if( ! $verify -or ( Get-Command -Name $run -CommandType Application ) )
            {
                Set-ItemProperty -Path $regKey -Name $name -Value $run
            }
            else
            {
                Throw "Failed to verify that `"$run`" exists"
            }
        }
    }
}
else ## file system not registry
{
    if( $allusers )
    {
        $folder = [environment]::GetFolderPath('CommonStartUp') 
    }
    else
    {
        $folder = [environment]::GetFolderPath('StartUp')
    }
    
    [hashtable]$windowStyle = @{
        1 = 'Normal'
        3 = 'Maximised'
        7 = 'Minimised' }

    $ShellObject = New-Object -ComObject Wscript.Shell

    if( $remove -or $list )
    {
        if( $list -or ! [string]::IsNullOrEmpty( $run ) -or ! [string]::IsNullOrEmpty( $name ) )
        {
            Get-ChildItem -Path $folder -Include '*.lnk' -Recurse | ForEach-Object `
            {
                $shortcut = $ShellObject.CreateShortcut($_.FullName)
                [string]$targetPath =  $shortcut.TargetPath 
                if( $targetPath -match $run -and $_.Name -match $name )
                {
                    if( $list )
                    {
                        [void]$items.Add( [pscustomobject][ordered]@{
                            'Shortcut' = $_.FullName
                            'Target' = "`"$targetPath`" $($shortcut.Arguments)"
                            'Description' = $shortcut.Description
                            'Window Style' = $WindowStyle[ $shortcut.WindowStyle ]
                            'Hotkey' = $shortcut.Hotkey })
                    }
                    elseif( $pscmdlet.ShouldProcess( "$($_.FullName)" , "Delete" ) )
                    {
                        Remove-Item -Path $_ -Force
                        if( $? )
                        {
                            $count++
                        }
                    }
                }
            }
        }
        else
        {
            Write-Error "Must specify -name or -run when removing an entry"
        }
    }
    else ## add
    {
        if( [string]::IsNullOrEmpty( $name ) )
        {
            Write-Error "Must specify the name of the shortcut via the -name option"
            return
        }
        [string]$fullPath = Join-Path -Path  $folder -ChildPath ($name + '.lnk')
        if( ! ( Test-Path $fullPath ) -or $pscmdlet.ShouldProcess( "`"$fullPath`" already exists" , "Overwrite" ) )
        {           
            ## Shortcut needs exe (TargetPath) separate to the arguments so we must split out
            ## If first character is quote then find matching one else we break on whitespace
            [string]$target = $run
            [string]$arguments = $null
            if( $run.StartsWith( "`"" ) )
            {
                $end = $run.Substring( 1 ).IndexOf("`"")
                $target = $run.Substring( 1 , $end )
                $arguments = $run.Substring($end + 2).TrimStart() ## skip passed " and index is zero based
            }
            elseif( $run.IndexOf( ' ' ) -ge 0 )
            {
                $target = $run.Substring( 0 , $run.IndexOf( ' ' ) )
                $arguments = $run.Substring( $run.IndexOf(' ' ) ).TrimStart()
            }
            $shortcut = $ShellObject.CreateShortcut( $fullPath )
            $shortcut.TargetPath = $target
            if( ! [string]::IsNullOrEmpty( $arguments ) )
            {
                $shortcut.Arguments = $arguments
            }
            if( ! [string]::IsNullOrEmpty( $icon ) )
            {
                $shortcut.IconLocation = $icon
            }
            if( ! [string]::IsNullOrEmpty( $description ) )
            {
                $shortcut.Description = $description
            }
            if( ! $verify -or ( Get-Command -Name $target -CommandType Application -ErrorAction Stop ) )
            {
                $shortcut.Save()
            }
        }
    }
}

if( $list )
{
    if( ! $items -or ! $items.Count )
    {
        Write-Warning ( "Found no{0} items" -f $(if( ! [string]::IsNullOrEmpty( $name ) -or ! [string]::IsNullOrEmpty( $run ) ) { ' matching' } ) )
    }
    else
    {
        $items
    }
}

if( $remove -and ! $count )
{
    Write-Warning "Removed no items"
}