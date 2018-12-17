<#
Create an exe that will crash so that memory dump capture can be tested

Based on code at http://stekodyne.com/?p=11

@guyrleech 2018
#>

[CmdletBinding()]

Param
(
    [Parameter(Mandatory=$true,HelpMessage="Architecture for exe: x86, Itanium, x64, anycpu, or anycpu32bitpreferred. The default is x86.")]
    [ValidateSet('x86', 'Itanium', 'x64', 'anycpu', 'anycpu32bitpreferred')]
    [string]$platform = "x86" ,
    [Parameter(Mandatory=$false,HelpMessage="Name of exe to produce including folder.")]
    [string]$exename = "$($pwd)\wcrasher.$($platform).exe" ,
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
    $codeProvider = New-Object Microsoft.CSharp.CSharpCodeProvider
    $compilerParameters = New-Object System.CodeDom.Compiler.CompilerParameters
    ## Need unsafe option as we are deliberately dereferencing a null pointer
    $compilerParameters.CompilerOptions = "-unsafe -platform:$platform"
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
        throw "Errors encountered while compiling code"
    }
}


#-----------------------------------------------------------------------------------------------------------

#the actual csharp code, 

$csharp = @'
    using System;
    using System.Windows.Forms;

    namespace RefreshNow 
    { 
        static class Program 
        { 
            static void Main() 
            {   
		        DialogResult result = MessageBox.Show( "Click OK to crash the application",
			        "Wcrasher", MessageBoxButtons.OKCancel );
                unsafe
		        {
			        int *badptr = (int *)0 ;
			        if( DialogResult.OK == result )
			        {
				        *badptr = 0xbad ;
			        }
		        }
            } 
       } 
   }
'@

#do not clobber file
if( ! $overwrite -and ( Test-Path -Path $exename -ErrorAction SilentlyContinue ) )
{
	Throw "$exename already exists - aborting"
}

Write-Verbose "Compiling code to `"$exename`" for $platform"

Compile-CSharp -code $csharp -references 'System.Windows.Forms.dll' -outputexe $exename -platform $platform
