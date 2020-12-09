
<#
.SYNOPSIS

Create an exe that displays a message box

.PARAMETER exename

The name of the executable to produce

.PARAMETER title

The title of the message box

.PARAMETER body

The body text of the message box

.PARAMETER icon

The icon, if any, to show in the message box

.PARAMETER button

The button(s) to show in the message box. The button clicked will be the exit value of the script (see notes)

.PARAMETER platform

The platform to compile for

.PARAMETER overwrite

Overwrite the exe if it already exists otherwise the script will exit prematurely and not overwrite the existing file

.EXAMPLE

& '.\Make MessageBox exe.ps1' -exename powershell.exe -title "PowerShell" -body "PowerShell says `"No`"" -icon Error -button OK -overwrite

Create an exe called powershell.exe which when run displays the given title and body with an "OK" button and error icon.

.EXAMPLE

& '.\Make MessageBox exe.ps1' -exename yesno.exe -title "Question from Guy" -body "Continue ?" -icon Question -button YesNo

Create an exe called yesno.exe which when run displays the given title and body with "yes" and "no" buttons and question icon.
Will return 6 if "yes" is pressed or "7" if "no" is pressed

.NOTES

Based on code at http://stekodyne.com/?p=11

Return codes as per https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.dialogresult?view=netframework-4.5

Modification History:

  09/12/2020  GRL  Initial release
#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$true,HelpMessage="Name of exe to produce including folder")]
    [string]$exename  ,
    [Parameter(Mandatory=$true,HelpMessage="Message title")]
    [string]$title ,
    [Parameter(Mandatory=$true,HelpMessage="Message body")]
    [string]$body ,
    [ValidateSet('Asterisk', 'Error', 'Exclamation', 'Hand', 'Information' , 'None' , 'Question' , 'Stop' , 'Warning' )]
    [string]$icon = 'none' ,
    [ValidateSet('AbortRetryIgnore', 'OK', 'OKCancel', 'RetryCancel', 'YesNo' , 'YesNoCancel' )]
    [string]$button = 'OK' ,
    [ValidateSet('x86', 'Itanium', 'x64', 'anycpu', 'anycpu32bitpreferred')]
    [string]$platform = 'x64' ,
    [switch]$overwrite
)


# Compile the CSharp code
function Compile-CSharp (
  [string] $code,
  [string] $OutputEXE,
  [array]  $references,
  [string] $platform   
)
{
    $codeProvider = New-Object -TypeName Microsoft.CSharp.CSharpCodeProvider
    $compilerParameters = New-Object -TypeName System.CodeDom.Compiler.CompilerParameters
    $compilerParameters.CompilerOptions = "-platform:$platform"
    foreach ($reference in $references)
    {
      $null = $compilerParameters.ReferencedAssemblies.Add( $reference );
    }
    $compilerParameters.GenerateInMemory = $false
    $compilerParameters.GenerateExecutable = $true
    
    $compilerParameters.OutputAssembly =  $OutputEXE
    $compiledCode = $codeProvider.CompileAssemblyFromSource(
      $compilerParameters,
      $code
    )

    if ( $compiledCode.Errors.Count)
    {
        $codeLines = $code.Split("`n");
        foreach ($compilerError in $compiledCode.Errors)
        {
            write-host "Error: $($codeLines[$($compilerError.Line - 1)])"
            write-host $compilerError
        }
        Throw "Errors encountered while compiling code"
    }
}

#do not clobber file
if( ! $overwrite -and ( Test-Path -Path $exename -ErrorAction SilentlyContinue ) )
{
	Throw "$exename already exists - aborting"
}

Write-Verbose "Compiling code to `"$exename`" for $platform"

## escape any quotes in our strings
$body = $body -replace '\"' , '\"'
$title = $title -replace '\"' , '\"'

#the actual csharp code, 

$csharp = @"
    using System;
    using System.Windows.Forms;

    namespace RefreshNow 
    { 
        static class Program 
        { 
            static void Main( string[] args ) 
            {   
                System.Environment.Exit( (int)MessageBox.Show( "$body" , "$title" , MessageBoxButtons.$button , MessageBoxIcon.$icon ) ) ;
            } 
       } 
   }
"@

Compile-CSharp -code $csharp -references 'System.Windows.Forms.dll' -outputexe $exename -platform $platform
