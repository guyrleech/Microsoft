 #requires -version 3
 ## adapted from https://stackoverflow.com/questions/8743122/how-do-i-find-the-msi-product-version-number-using-powershell
 ## and https://adamtheautomator.com/powershell-windows-installer-msi-properties/
 
<#
.SYNOPSIS

Retrieve Windows Installer properties of MSI files

.PARAMETER path

Path to the MSI file(s) to query

.PARAMETER properties

A comma separated list of MSI properties to retrieve. Will retrieve all properties that match when -regex is specified

.PARAMETER regex

Treat the properties specified via -properties as regular expressions

.PARAMETER summaryInformation

Return summary information stream properties rather than properties

.PARAMETER quiet

Do not output warnings or errors

.EXAMPLE

Get-MSIProperty -Path c:\temp\fred.msi,c:\temp\bloggs.msi

Retrieve the "ProductVersion" property from the two specified MSI files

.EXAMPLE

Get-MSIProperty -Path c:\temp\fred.msi,c:\temp\bloggs.msi -Properties Product -regex

Retrieve the all properties which match the regular expression "Product" from the two specified MSI files such as "ProductVersion" and "ProductName"

.EXAMPLE

Get-MSIProperty -Path c:\temp\fred.msi,c:\temp\bloggs.msi

Retrieve the "ProductVersion" property from the two specified MSI files

.EXAMPLE

Get-ChildItem -Path c:\temp\*.msi | Get-MSIProperty

Retrieve the "ProductVersion" property from the two specified MSI files

.NOTES

Modification History:

   @guyrleech 11/04/20  Initial release
   @guyrleech 01/10/20  Added Summary Information Stream retrieval
   @guyrleech 01/10/20  Added description for security summary information field
#>

 Function Get-MSIProperty
 {
    [CmdletBinding()]

    Param
    (
        [Parameter(Mandatory=$true,HelpMessage='Path to msi file',ValueFromPipeline=$true)]
        [string[]]$path ,
        [string[]]$properties = @( 'ProductVersion' ) ,
        [switch]$regex ,
        [switch]$summaryInformation ,
        [switch]$quiet
    )

    Begin
    {
        if( $summaryInformation -and $regex )
        {
            Throw "Cannot use -regex with -summaryInformation"
        }
        if( ! ( $windowsInstaller = New-Object -Com WindowsInstaller.Installer ) )
        {
            Throw "Failed to create Windows Installer object"
        }
        ## https://docs.microsoft.com/en-us/windows/win32/msi/summary-information-stream-property-set
        [hashtable]$summaryInformationStreamProperties = @{
            1 =  'Codepage'
            2 =  'Title'
            3 =  'Subject'
            4 =  'Author'
            5 =  'Keywords'
            6 =  'Comments'
            7 =  'Template'
            8 =  'Last Saved By'
            9 =  'Revision Number'
            11 = 'Last Printed'
            12 = 'Create Time/Date'
            13 = 'Last Save Time/Date'
            14 = 'Page Count'
            15 = 'Word Count'
            16 = 'Character Count'
            18 = 'Creating Application'
            19 = 'Security'
        }
        ## https://docs.microsoft.com/en-us/windows/win32/msi/security-summary
        [hashtable]$securityProperties = @{
            0	= 'No restriction'
            2	= 'Read-only recommended'
            4	= 'Read-only enforced'
        }
    }

    Process
    {
        ForEach( $file in $path )
        {
            Try
            {
                $database = $windowsInstaller.GetType().InvokeMember( 'OpenDatabase', 'InvokeMethod', $Null, $windowsInstaller, @( $file , 0 ) ) ## 0 = open read-only - https://docs.microsoft.com/en-us/windows/win32/msi/installer-opendatabase
            }
            Catch
            {
                $database = $null
                if( ! $quiet )
                {
                    Write-Error -Exception $_.Exception
                }
            }

            if( $database )
            {
                if( $summaryInformation )
                {
                    if( $summary = $database.SummaryInformation(4) )
                    {
                        ForEach( $summaryInformationStreamProperty in $summaryInformationStreamProperties.GetEnumerator() )
                        {
                            if( ( ( ! $properties -or ! $properties.Count -or ! $PSBoundParameters[ 'properties' ] ) -or $summaryInformationStreamProperty.Value -in $properties ) -and ( $null -ne ( $info = $summary.Property( $summaryInformationStreamProperty.Name ) ) ) )
                            {
                                if( $summaryInformationStreamProperty.Value -eq 'Security' )
                                {
                                    ## convert number to meaningful string
                                    if( $securityDescription = $securityProperties[ $info ] )
                                    {
                                        $info = $securityDescription
                                    }
                                }
                                [pscustomobject][ordered]@{
                                        'File' = $file
                                        'Summary Property' = $summaryInformationStreamProperty.Value
                                        'PID' = $summaryInformationStreamProperty.Name
                                        'Value' = $info 
                                    }
                            }
                            elseif( ! $quiet )
                            {
                                Write-Warning -Message "Failed to retrieve summary information stream property PID $($summaryInformationStreamProperty.Name) from $file"
                            }
                        }
                    }
                    else
                    {
                        Write-Warning -Message "Failed to retrieve summary information from $file"
                    }
                }
                else
                {
                    ForEach( $property in $properties )
                    {
                        [string]$query = 'SELECT * FROM Property'
                        if( ! $regex )
                        {
                            $query += " WHERE Property = '$property'"
                        }
                        ## else we will look at each record returned and see if it matches the property which is a regex. Can't do a "like" query it seems

                        if( $View = $database.GetType().InvokeMember( 'OpenView' , 'InvokeMethod', $Null, $database, $query ) )
                        {
                            $View.GetType().InvokeMember( 'Execute' , 'InvokeMethod', $Null, $View, $Null )
                            [int]$recordCount = 0

                            while( $record = $View.GetType().InvokeMember( 'Fetch' , 'InvokeMethod', $Null, $View, $Null ) )
                            {
                                $recordCount++
                                [string]$propertyName  = $record.GetType().InvokeMember( 'StringData', 'GetProperty', $Null, $record, 1 )
                                if( ! $regex -or $propertyName -match $property )
                                {
                                    [string]$propertyValue = $record.GetType().InvokeMember( 'StringData', 'GetProperty', $Null, $record, 2 )

                                    if( ! [string]::IsNullOrEmpty( $propertyValue ))
                                    {
                                        [pscustomobject][ordered]@{
                                            'File' = $file
                                            'Property' = $propertyName
                                            'Value' = $propertyValue.Trim()
                                        }
                                    }
                                    elseif( ! $quiet )
                                    {
                                        Write-Warning "No $property property found"
                                    }
                                }
                                else
                                {
                                    Write-Verbose -Message "Ignoring property `"$propertyName`""
                                }
                            }
                    
                            $View.GetType().InvokeMember( "Close", "InvokeMethod", $Null, $View, $Null )

                            if( ! $quiet -and ! $recordCount )
                            {
                                Write-Warning "No $property record found"
                            }
                            $View.GetType().InvokeMember( 'Close' , 'InvokeMethod' ,$null,$View,$null)
                            $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject( $view )
                            $view = $null
                        }
                        elseif( ! $quiet )
                        {
                            Write-Warning "Failed to get $property view"
                        }
                    }
                }
                $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject( $database )
                $database = $null
            }
            elseif( ! $quiet )
            {
                Write-Warning "Failed to open MSI database"
            }
        }
    }

    End
    {
       $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject( $windowsInstaller )
       $windowsInstaller = $null
    }
}
