#requires -version 3
<#
    Get the named extended property(s) from the file or all available properties

    With code from https://rkeithhill.wordpress.com/2005/12/10/msh-get-extended-properties-of-a-file/

    @guyrleech 17/12/2019
#>


<#
.SYNOPSIS

Get the named extended property(s) from the file or all available properties

.DESCRIPTION

Useful where a file has extended properties/metadata such as an image file or msi file

.PARAMETER filename

The name of the file to retrieve the properties of

.PARAMETER properties

Comma separated list of the properties to retrieve. If not specified, will return all available properties 

.EXAMPLE

Get-ExtendedProperties -fileName "C:\@guyrleech\googlechromestandaloneenterprise64.v79.msi" -properties Title,Subject,Categories,Tags,Comments,Name,Authors

Returns the requested properties for the specified file

.EXAMPLE

Get-ExtendedProperties -fileName "C:\@guyrleech\picture.jpg" 

Returns all extended properties for the specified file

#>

Function Get-ExtendedProperties
{
    [CmdletBinding()]

    Param
    (
        [Parameter(Mandatory=$true,HelpMessage='File name to retrieve properties of')]
        [ValidateScript({Test-Path -Path $_})]
        [string]$fileName ,
        [AllowNull()]
        [string[]]$properties
    )

    [hashtable]$propertiesToIndex = @{}
    ## need to use absolute paths
    $fileName = Resolve-Path -Path $fileName | Select-Object -ExpandProperty Path
    $shellApp = New-Object -Com shell.application
    $myFolder = $shellApp.Namespace( (Split-Path -Path $fileName -Parent) )
    $myFile = $myFolder.Items().Item( (Split-Path -Path $fileName -Leaf) )

    0..500 | ForEach-Object `
    {
        If( $key = $myFolder.GetDetailsOf( $null , $_ ) )
        {
            Try
            {
                $propertiesToIndex.Add( $key , $_ )
            }
            Catch
            {
            }
        }
    }

    Write-Verbose "Got $($propertiesToIndex.Count) unique property names"

    If( ! $PSBoundParameters[ 'properties' ] -or ! $properties -or ! $properties.Count )
    {
        ForEach( $property in $propertiesToIndex.GetEnumerator() )
        {
            $thisProperty = $myFolder.GetDetailsOf( $myFile , $property.Value )
            If( ! [string]::IsNullOrEmpty( $thisProperty ) )
            {
                [pscustomobject]@{ 
                    'Property' = $property.Name
                    'Value' = $thisProperty
                }
            }
        }
    }
    Else
    {
        ForEach( $property in $properties )
        {
            $index = $propertiesToIndex[ $property ]
            If( $index -ne $null )
            {
                $myFolder.GetDetailsOf( $myFile , $index -as [int] )
            }
            Else
            {
                Write-Warning "No index for property `"$property`""
            }
        }
    }
}