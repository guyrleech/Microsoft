#requires -version 3

<#
.SYNOPSIS

Look at specified executables and show their executable type

.DESCRIPTION

Executables have a PE header which contains the machine type they run on so this script fetches this and translates to a human readable description, if the file is an executable and has a PE header

.PARAMETER files

Comma separated list of files to examine

.PARAMETER folders

Comma separated list of folders to examine

.PARAMETER recurse

.PARAMETER explain

Output more detail for the .NET PE kinds

Recurse folders

.PARAMETER quiet

Do not report errors such as not being able to open the file or it not being an executable

.EXAMPLE

& '.\Get file bitness.ps1' -file file1.exe,file2.exe 

Show the bitness of file1.exe and file2.exe in the current folder

.EXAMPLE

& '.\Get file bitness.ps1' -folder c:\temp -recurse -quiet

Show the bitness of all executable files found in the c:\temp folder and subfolders but do not report any errors such as a file not being an executable

.NOTES

Based on code from @shaylevy although at the time of writing the script is flawed as it doesn't read enough data from the file to read the PE header

https://www.powershellmagazine.com/2013/03/08/pstip-how-to-determine-if-a-file-is-32bit-or-64bit/

#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$true,ParameterSetName='files',HelpMessage='Comma separated list of files to process',ValueFromPipeline=$true)]
    [string[]]$files ,
    [Parameter(Mandatory=$true,ParameterSetName='folders',HelpMessage='Comma separated list of folders to process')]
    [string[]]$folders ,
    [Parameter(Mandatory=$false,ParameterSetName='folders')]
    [switch]$recurse ,
    [switch]$explain ,
    [switch]$quiet
)

Begin
{
    [int]$MACHINE_OFFSET = 4
    [int]$PE_POINTER_OFFSET = 60

    [hashtable]$machineTypes = @{
        0x0 = 'Native'
        0x1d3 = 'Matsushita AM33'
        0x8664 = 'x64'
        0x1c0 = 'ARM little endian'
        0xaa64 = 'ARM64 little endian'
        0x1c4 = 'ARM Thumb-2 little endian'
        0xebc = 'EFI byte code'
        0x14c = 'x86' ## 'Intel 386 or later processors and compatible processors'
        0x200 = 'Intel Itanium processor family'
        0x9041 = 'Mitsubishi M32R little endian'
        0x266 = 'MIPS16'
        0x366 = 'MIPS with FPU'
        0x466 = 'MIPS16 with FPU'
        0x1f0 = 'Power PC little endian'
        0x1f1 = 'Power PC with floating point support'
        0x166 = 'MIPS little endian'
        0x5032 = 'RISC-V 32-bit address space'
        0x5064 = 'RISC-V 64-bit address space'
        0x5128 = 'RISC-V 128-bit address space'
        0x1a2 = 'Hitachi SH3'
        0x1a3 = 'Hitachi SH3 DSP'
        0x1a6 = 'Hitachi SH4'
        0x1a8 = 'Hitachi SH5'
        0x1c2 = 'Thumb'
        0x169 = 'MIPS little-endian WCE v2'
    }

    [hashtable]$processorAchitectures = @{
        'None'  = 'None'
        'MSIL'  = 'AnyCPU'
        'X86'   = 'x86'
        'I386'   = 'x86'
        'IA64'  = 'Itanium'
        'Amd64' = 'x64'
        'Arm' = 'ARM'
    }

    [hashtable]$pekindsExplanations = @{
     'ILOnly' = 'MSIL processor neutral'
     'NotAPortableExecutableImage' = 'Not in portable executable (PE) file format'
     'PE32Plus'	= 'Requires a 64-bit platform'
     'Preferred32Bit' = 'Platform-agnostic but should be run on 32-bit platform'
     'Required32Bit' = 'Runs on a 32-bit platform or in the 32-bit WOW environment on a 64-bit platform'
     'Unmanaged32Bit'  = 'Contains pure unmanaged code'
    }

    If( $PSBoundParameters[ 'folders' ] )
    {
        $files = @( ForEach( $folder in $folders )
        {
            Get-ChildItem -Path $folder -File -Recurse:$recurse -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
        })
    }
}

Process
{
    ForEach( $file in $files )
    {
        $data = New-Object System.Byte[] 4096
        Try
        {
            $stream = New-Object System.IO.FileStream -ArgumentList $file,Open,Read
        }
        Catch
        {
            $stream = $null
            if( ! $quiet )
            {
                Write-Error -Exception $_
            }
        }

        Try
        {
            $runtimeAssembly = [System.Reflection.Assembly]::ReflectionOnlyLoadFrom( $file )
        }
        Catch
        {
            $runtimeAssembly = $null
        }
    
        Try
        {
            $assembly = [System.Reflection.AssemblyName]::GetAssemblyName( $file )
        }
        Catch
        {
            $assembly = $null
        }

        If( $stream )
        {
            [uint16]$machineUint = 0xffff
            [int]$read = $stream.Read( $data , 0 ,$data.Count )
            If( $read -gt $PE_POINTER_OFFSET )
            {
                If( $data[0] -eq 0x4d -and $data[1] -eq 0x5a ) ## MZ
                {
                    [int]$PE_HEADER_ADDR = [System.BitConverter]::ToInt32( $data, $PE_POINTER_OFFSET )
                    [int]$typeOffset = $PE_HEADER_ADDR + $MACHINE_OFFSET
                    If( $data[ $PE_HEADER_ADDR ] -eq 0x50 -and $data[ $PE_HEADER_ADDR + 1 ] -eq 0x45 ) ## PE
                    {
                        If( $read -gt $typeOffset + [System.Runtime.InteropServices.Marshal]::SizeOf( $machineUint ) )
                        {
                            [uint16]$machineUint = [System.BitConverter]::ToUInt16( $data, $typeOffset )
                            $versionInfo = Get-ItemProperty -Path $file -ErrorAction SilentlyContinue | Select-Object -ExpandProperty VersionInfo
                            If( $runtimeAssembly -and ( $module = ($runtimeAssembly.GetModules() | Select -First 1) ) )
                            {
                                $pekinds = New-Object -TypeName System.Reflection.PortableExecutableKinds
                                $imageFileMachine = New-Object -TypeName System.Reflection.ImageFileMachine
                                $module.GetPEKind( [ref]$pekinds , [ref]$imageFileMachine )
                            }
                            Else
                            {
                                $pekinds = $null
                                $imageFileMachine = $null
                            }

                            [pscustomobject][ordered]@{
                                'File' = $file
                                'Architecture' = $machineTypes[ [int]$machineUint ]
                                '.NET Architecture' = $(If( $assembly ) { $processorAchitectures[ $assembly.ProcessorArchitecture.ToString() ] } else { 'Not .NET' } )
                                '.NET PE Kind' = $( If( $pekinds ) { if( $explain ) { ($pekinds.ToString() -split ',\s?' | ForEach-Object { $pekindsExplanations[ $_ ] }) -join ',' } else { $pekinds.ToString() } }  else { 'Not .NET' } )
                                '.NET Platform' = $(If( $imageFileMachine ) { $processorAchitectures[ $imageFileMachine.ToString() ] } else { 'Not .NET' } )
                                'Company' = $versionInfo | Select-Object -ExpandProperty CompanyName
                                'File Version' = $versionInfo | Select-Object -ExpandProperty FileVersionRaw
                                'Product Name' = $versionInfo | Select-Object -ExpandProperty ProductName
                            }
                        }
                        Else
                        {
                            Write-Warning "Only read $($data.Count) bytes from `"$file`" so can't reader header at offset $typeOffset"
                        }
                    }
                    ElseIf( ! $quiet )
                    {
                        Write-Warning "`"$file`" does not have a PE header signature"
                    }
                }
                ElseIf( ! $quiet )
                {
                    Write-Warning "`"$file`" is not an executable"
                }
            }
            ElseIf( ! $quiet )
            {
                Write-Warning "Only read $read bytes from `"$file`", not enough to get header at $PE_POINTER_OFFSET"
            }
            $stream.Close()
            $stream = $null
        }
    }
}
