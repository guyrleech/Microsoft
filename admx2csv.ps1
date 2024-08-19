<#
.SYNOPSIS
    Convert GPO ADMX files to csv

.PARAMETER policyDir
    The directory containing the ADMX files. Default is %windir%\policyDefinitions
.PARAMETER language
    The language to use for the ADMX files. Default is 'en-US'
.PARAMETER outputfilename
    The name of the output csv file
.PARAMETER fileRegex
    A regex to filter the ADMX file names
.PARAMETER overwrite
    Overwrite the output file if it already exists
.PARAMETER delimiter    
    The delimiter to use in the output file. Default is ','. Some countires like the Netherlaands use a semicolon as delimiter!

.EXAMPLE
    .\admxcsv.ps1 -outputfilename "C:\temp\local.admx.csv" -overwrite
    Parse all ADMX files in %windir%\policyDefinitions and write the settings to to C:\temp\admx.csv, overwriting it if it already exists

.EXAMPLE
    .\admxcsv.ps1 -outputfilename "C:\temp\domain.admx.csv" -policyDir "\\domain.com\sysvol\domain.com\policies\PolicyDefinitions"
    Parse all ADMX files in the domain store and write the settings to to C:\temp\admx.csv, overwriting it if it already exists

.EXAMPLE
    .\admxcsv.ps1 -outputfilename "C:\temp\domain.admx.csv" -policyDir "\\domain.com\sysvol\domain.com\policies\PolicyDefinitions" -fileRegex ^terminalserver
    Parse all ADMX files in the domain store that start with "terminalserver" and write the settings to to C:\temp\admx.csv, overwriting it if it already exists

.NOTES
    Original code from @chentiangemalc https://chentiangemalc.wordpress.com/2014/10/02/powershell-script-to-extract-info-from-admx/
#>

[CmdletBinding()]

Param
(
    [string]$policyDir = "$env:windir\policyDefinitions" ,
    [string]$language = 'en-US' ,
    [Parameter(Mandatory=$true)]
    [string]$outputfilename ,
    [string]$fileRegex ,
    [switch]$overwrite ,
    [string]$delimiter = ','
)

if( (Test-Path -Path $outputfilename ) -and -not $overwrite )
{
    Throw "Output file `"$outputfilename`" already exists but -overwrite not specified"
}

$table = New-Object -typename System.Data.DataTable

[void]$table.Columns.Add("ADMX")
[void]$table.Columns.Add("Parent Category")
[void]$table.Columns.Add("Name")
[void]$table.Columns.Add("Display Name")
[void]$table.Columns.Add("Class")
[void]$table.Columns.Add("Explain Text")
[void]$table.Columns.Add("Supported On")
[void]$table.Columns.Add("Key")
[void]$table.Columns.Add("Value Name")

$admxFiles = @( Get-ChildItem $policyDir -filter *.admx | Where-Object -Property Name -Match $fileRegex )

Write-Verbose -Message "Got $($admxFiles.Count) admx files"

[int]$counter = 0

ForEach ($file in $admxFiles)
{
    $counter++
    Write-Verbose -Message "$counter / $($admxFiles.Count) : $($file.Name)"

    [xml]$data = Get-Content "$policyDir\$($file.Name)"
    [xml]$lang = Get-Content "$policyDir\$language\$($file.Name.Replace(".admx",".adml"))"

    $policyText = $lang.policyDefinitionResources.resources.stringTable.ChildNodes

    if( -Not $data.PolicyDefinitions.PSObject.properties[ 'policies' ] )
    {
        Write-Warning -Message "$($file.Name): No policies found"
        continue
    }

    $data.PolicyDefinitions.policies.ChildNodes | ForEach-Object {

        $policy = $_

        if ($policy -ne $null)
        {
            if ($policy.Name -ne "#comment")
            {
                ##Write-Verbose "`tProcessing policy $($policy.Name)"
                $displayName = ($policyText | Where-Object { $_.PSObject.Properties[ 'id' ] -and $_.id -eq $policy.displayName.Substring(9).TrimEnd(')') }) | Select-Object -ExpandProperty '#text' -ErrorAction SilentlyContinue
                $explainText = ($policyText | Where-Object { $_.PSObject.Properties[ 'id' ] -and $_.psobject.properties[ 'explainText' ] -and $_.id -eq $policy.explainText.Substring(9).TrimEnd(')') }) | Select-Object -ExpandProperty '#text' -ErrorAction SilentlyContinue
              
                if ($policy.SupportedOn.ref.Contains(":"))
                {        
                    $source=$policy.SupportedOn.ref.Split(":")[0]
                    $valueName=$policy.SupportedOn.ref.Split(":")[1]
                    $admlFile="$policyDir\$language\$source.adml"
                    [xml]$adml=Get-Content $admlFile -ErrorAction SilentlyContinue
                    if( $null -ne $adml )
                    {
                        $resourceText= $adml.policyDefinitionResources.resources.stringTable.ChildNodes
                        $supportedOn=($resourceText | Where-Object { $_.PSObject.Properties[ 'id' ] -and $_.id -eq $valueName }) | Select-Object -ExpandProperty '#text' -ErrorAction SilentlyContinue
                    }
                    elseif( Test-Path -Path $admlFile )
                    {
                        Write-Warning "$($file.name): policy $($policy.Name): Bad XML found in ADML file $admlFile"
                    }
                    else
                    {
                        Write-Warning "$($file.name): policy $($policy.Name): ADML file $admlFile not found"
                    }
                }
                else
                {
                    if( $data.policyDefinitions.PSObject.properties[ 'supportedon' ] -and ( $supportedOnID = $data.policyDefinitions.supportedOn.definitions.ChildNodes | Where-Object Name -eq $policy.supportedOn.ref  | Select-Object -ExpandProperty DisplayName -ErrorAction SilentlyContinue ))
                    {
                        $supportedOn = ($policyText | Where-Object { $_.PSObject.Properties[ 'id' ] -and $_.id -eq $supportedOnID.Substring(9).TrimEnd(')') }) | Select-Object -ExpandProperty '#text' -ErrorAction SilentlyContinue
                    }
                    else
                    {
                        $supportedOn = $null
                    }
                }

                if ($policy.parentCategory.ref.Contains(":"))
                {
                    $source=$policy.parentCategory.ref.Split(":")[0]
                    $valueName=$policy.parentCategory.ref.Split(":")[1]
                    $admlFile="$policyDir\$language\$source.adml"
                    [xml]$adml=Get-Content $admlFile -ErrorAction SilentlyContinue
                    if( $null -ne $adml )
                    {
                        $resourceText= $adml.policyDefinitionResources.resources.stringTable.ChildNodes
                        $parentCategory=($resourceText | Where-Object { $_.PSObject.Properties[ 'id' ] -and $_.id -eq $valueName }) | Select-Object -ExpandProperty '#text' -ErrorAction SilentlyContinue
                    }
                    elseif( Test-Path -Path $admlFile )
                    {
                        Write-Warning "$($file.name): policy $($policy.Name): Bad XML found in ADML file $admlFile"
                    }
                    else
                    {
                        Write-Warning "$($file.name): policy $($policy.Name): ADML file $admlFile not found"
                    }
                } 
                else
                {
                    $parentCategory = $null
                     if( -Not [string]::IsNullOrEmpty( ($parentCategoryID = ($data.policyDefinitions.categories.ChildNodes | Where-Object { $_.Name -eq $policy.parentCategory.ref }) | Select-Object -ExpandProperty DisplayName -ErrorAction SilentlyContinue )))
                     {
                        $parentCategory =  ($policyText | Where-Object { $_.PSObject.Properties[ 'id' ] -and $_.id -eq $parentCategoryID.Substring(9).TrimEnd(')') }) | Select-Object -ExpandProperty '#text' -ErrorAction SilentlyContinue
                    }
                }

                if( -Not $policy.PSObject.properties[ 'valueName' ] )
                {
                    ##TODO what if more than 1 ?
                    $value = $null
                    try
                    {
                        ## not XML but it works!
                        if( $policy.PSObject.properties[ 'elements' ] -and $policy.elements.ChildNodes.Count -gt 0 )
                        {
                            $value = ($policy.elements.ChildNodes | Select-Object -ExpandProperty valueName -ErrorAction SilentlyContinue) -join ' ; '
                        }
                        else
                        {
                            Write-Warning "$($file.name): policy $($policy.Name) has no elements"
                        }
                        Add-Member -InputObject $policy -MemberType NoteProperty -Name valueName -Value $value
                    }
                    catch
                    {
                        Write-Warning "$($file.name): policy $($policy.Name): $_"
                    }
                }

                [void]$table.Rows.Add(
                    $file.Name,
                    $parentCategory,
                    $policy.Name,
                    $displayName,
                    $policy.class,
                    $explainText,
                    $supportedOn,
                    $policy.key,
                    $policy.valueName)
            }
        }

    }
}

Write-Verbose -Message "Got $($table.Rows.Count) rows"

$table | Export-Csv $outputfilename -NoTypeInformation -Delimiter $delimiter -Force
