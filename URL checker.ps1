<#
.SYNOPSIS
    Decodes URLs from the clipboard and displays domain/registrar info using a WPF GUI or can be called from the command line wh a domain name

.DESCRIPTION
    This script provides a WPF GUI for decoding URLs (such as those found in email links) from the clipboard. It can also display domain and registrar information for a given URL or domain name. The GUI includes buttons to decode the clipboard URL, show domain info, and open a "Buy me a coffee" link.

.PARAMETER url
    The URL or domain name to look up. If not specified, the script uses the clipboard contents.

.PARAMETER allProperties
    If specified, returns all properties from the RDAP lookup for the domain.

.PARAMETER ignoredVCards
    An array of vCard property names to ignore when displaying domain info. Default is 'version'.

.PARAMETER joiner
    String used to join array values in vCard data. Default is ', '.

.PARAMETER buttonBackgroundColour
    Background colour for the main window buttons. Default is 'GreenYellow'.

.PARAMETER fontName
    Font family for the output text box. Default is 'Courier New'.

.PARAMETER noResolving
    If specified, does not resolve the domain to an IP address to get country information.
    This is useful if you only want registrar and domain info without making DNS queries.

.PARAMETER fontSize
    Font size for the output text box. Default is 16.

.PARAMETER readOnly
    Whether the output text box is read-only. Accepts 'True' or 'False'. Default is 'False'.

.EXAMPLE
    & '.\URL checker.ps1' -domain example.com

    Displays domain information for the sepcified domain

.EXAMPLE
    & '.\URL checker.ps1'

    Displays a GUI allowing a URL or domain name tobe pasted into the clipboard and decoded, showing domain information if available upon button clicking

.NOTES
    Modification History:

    2025/06/18  @guyrleech  Script born
    2025/07/15  @guyrleech  Added domain info button to get registrar and domain info
    2025/07/16  @guyrleech  Added declutter option to remove crap after ?
#>

[CmdletBinding()]

Param
(
    [Alias('domain','fqdn')]
    [string]$url ,
    [switch]$allProperties ,
    [switch]$noResolving ,
    [string[]]$ignoredVCards = @( 'version' ) ,
    [string]$joiner = ', ' ,
    [string]$buttonBackgroundColour = 'GreenYellow' ,
    [string]$fontName = 'Courier New' ,
    [int]$fontSize = 16 ,
    [ValidateSet('True', 'False')]
    [string]$readOnly = 'False'
)

Function
Get-RegistrarAndDomainInfo
{
    Param
    (
        [Parameter(Mandatory=$true)]
        [Alias('domain','fqdn')]
        [string]$url ,
        [switch]$allProperties ,
        [switch]$noResolving ,
        [string[]]$ignoredVCards = @( 'version' ) ,
        [string]$joiner = ', ' 
    )
    
    ## street addresses in vcard data can be arrays including arrays so flatten and caller joins

    Function
    Expand-StringArray
    {
        Param( [string[]]$string )

        ForEach( $element in $string )
        {
            if( $element -is [array] )
            {
                Expand-StringArray -string $element
            }
            elseif( -Not [string]::IsNullOrEmpty( $element ) )
            {
                $element
            }
        }

        }
    function
    Get-IPCountry
    {
        param
        (
            [Parameter(Mandatory = $true)]
            [ValidatePattern('^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$')]
            [string]$IPAddress
        )

        $uri = "http://ip-api.com/json/$IPAddress"

        try 
        {
            $response = $null
            $response = Invoke-RestMethod -Uri $uri -Method Get -ErrorAction SilentlyContinue
            if ($response.status -ieq "success")
            {
                $response.country ## output
            }
            else
            {
                Write-Error "API request failed: $($response.message)"
            }
        }
        catch [System.Net.WebException]
        {
            $statusCode = $_.Exception.Response.StatusCode.Value__
            if ($statusCode -eq 404)
            {
                Write-Error "IP address $IPAddress not found in the GeoIP database."
            }
            else
            {
                Write-Error "HTTP error $statusCode occurred: $($_.Exception.Message)"
            }
        }
        catch {
            Write-Error "Non-HTTP error occurred: $($_.Exception.Message)"
        }
    }
    ## from https://client.rdap.org/
    ## only use domain one currently

    [string]$registryURLsJSON = @'
    {
        "https://data.iana.org/rdap/dns.json": "domain", 
        "https://data.iana.org/rdap/asn.json": "autnum",
        "https://data.iana.org/rdap/ipv4.json": "ip4",
        "https://data.iana.org/rdap/ipv6.json": "ip6",
        "https://data.iana.org/rdap/object-tags.json": "entity"
    }
'@

    $registryURLs = $registryURLsJSON | ConvertFrom-Json

    [hashtable]$registryData = @{}

    ForEach( $entry in $registryURLs.PSObject.Properties )
    {
        if( $entry.value -ieq 'domain' )
        {
            $response = $null
            $response = Invoke-RestMethod -UseBasicParsing -Uri $entry.Name 
            if( $null -ne $response )
            {
                $registryData.Add( $entry.Value , $response )
            }
            else
            {
                Write-Warning -Message "Bad response for $($entry.Name)"
            }
        }
        ## else TODO see what we can use these other URIs for
    }
    
    $FQDN = $url -replace '^https?://' -replace '/.*$'
    $tld = $FQDN -replace '^.*\.'

    Write-Verbose -Message "tld for $url is $tld FQDN $FQDN"
    
    $registryEntry = $registryData[ 'domain' ].services | Where-Object { $_[0] -ieq $tld }

    if( [string]::IsNullOrEmpty( $registryEntry ) )
    {
        Write-Error "No registry for tld $tld found out of the $($registry[ 'domain' ].Count) tlds searched"
        return $null
    }

    $registryURL = $registryEntry[ -1 ]

    if( [string]::IsNullOrEmpty( $registryURL ) )
    {
        Throw "No registry URL for tld $($registryEntry[0])"
    }

    [string]$domainName = $FQDN
    [string]$lastParentName = $null
    do
    {

        ## https://rdap.nominet.uk/uk/domain/itv.co.uk?jscard=1
        [string]$queryURL = "$registryURL/domain/$domainName`?jscard=1" -replace '([^:])//+' , '$1/' ## registry URL may end in / so avoid //

        Write-Verbose -Message "registry is $registryURL query $queryURL"

        ## may need to loop since if is a separately managed subdomain it could come back as not found so we need to look at its parent and so on eg anything.slack.com

        $lookupResponse = $null
        try
        {
            $lookupResponse = Invoke-RestMethod -UseBasicParsing -Uri $queryURL
            break
        }
        catch
        {
            $exception = $_
            if( $exception.Exception.Response.StatusCode.value__ -eq 404 )
            {
                $parentDomain = $domainName -replace '^[^\.]*\.'
                Write-Verbose "Lookup failed for $domainName so trying parent $parentDomain : $exception "
                if( $parentDomain -eq $lastParentName )
                {
                    Write-Verbose "No parent domain found for $domainName as have already tried its parent $lastParentName"
                    break
                }
                $domainName = $parentDomain
                $lastParentName = $parentDomain
            }
            else
            {
                Write-Error "Query $queryURL errored : $exception" ## TODO could we get 429 errors and need to implement retries after back off ?
                break
            }
        }
    }
    while( -Not [string]::IsNullOrEmpty( $domainName ) )

    if( $null -eq $lookupResponse )
    {
        Write-Error "No data found for $FQDN from query to $queryURL"
        return $null
    }

    [string[]]$countries = $null

    if( -Not $noResolving )
    {
        ## resolve the domain to an IP address and get the country
        $resolution = $null
        $resolution = @( Resolve-DnsName -Name $FQDN -Type A -ErrorAction SilentlyContinue | Where-Object type -ieq 'A' )
        if( $null -ne $resolution -and $resolution.Count -gt 0)
        {
            $countries = @( ForEach( $address in $resolution )
            {
                $country = $null
                $country = Get-IPCountry -IPAddress $address.IPaddress
                if( -Not [string]::IsNullOrEmpty( $country ))
                {
                    $country
                }
            } ) 
            $countries = $countries | Sort-Object -Unique
        }
        else
        {
                $countries = 'Name resolution failed'
        }
    }

    if( $allProperties )
    {
        $lookupResponse
    }
    else
    {
        $now = Get-Date
        $output = $lookupResponse | Select-Object -Property @{ Name = 'FQDN' ; Expression = { $FQDN }} ,
            @{ Name = 'Domain'       ; Expression = { $lookupResponse.ldhName }} ,
            @{ Name = 'Registrar'    ; Expression = { $registryURL }},
            @{ Name = 'Registered'   ; Expression = { $script:registered = ( $_.events | Where-Object eventAction -ieq  "registration" | Select-Object -ExpandProperty eventDate ) -as [datetime] ; $script:registered }} ,
            @{ Name = 'Days Old'     ; Expression = { [math]::Round( ( $now - $Script:registered ).TotalDays , 1 ) }} ,
            @{ Name = 'Expires'      ; Expression = { ( $_.events | Where-Object eventAction -ieq  "expiration"   | Select-Object -ExpandProperty eventDate ) -as [datetime] }} ,
            @{ Name = 'Last Changed' ; Expression = { ( $_.events | Where-Object eventAction -ieq  "last changed" | Select-Object -ExpandProperty eventDate ) -as [datetime] }} ,
            @{ Name = 'Countries'    ; Expression = { $countries -join $joiner }},
            @{ Name = 'Name Servers' ; Expression = { ( $_.nameservers|Select-Object -ExpandProperty ldhname ) -replace '\.+$' -join ' ' }}
        ForEach( $vcard in ($lookupResponse.entities | Select-Object -ExpandProperty vcardArray -ErrorAction SilentlyContinue ))
        {
            if( $vcard -is [array] ) ## there will be a single string entry called "vcard" which we ignore
            {
                 <#
                     "vcardArray":  [
                                "vcard",
                                [
                                    [
                                        "version",
                                        {

                                        },
                                        "text",
                                        "4.0"
                                    ],
                                    [
                                        "fn",
                                        {

                                        },
                                        "text",
                                        "CSC Corporate Domains, Inc"
                                    ],
                                    [
                                        "tel",
                                        {
                                            "type":  "voice"
                                        },
                                        "text",
                                        "020-7565-4090"
                                    ],
                                    [
                                        "url",
                                        {

                                        },
                                        "uri",
                                        "https://www.cscdbs.com/"
                                    ]
                                ]
                            ],
                    #>
                ForEach( $element in $vcard )
                {
                    $title = $element[0]
                    $value = $element[ -1 ]
                    if( -Not [string]::IsNullOrEmpty( $title ) -and $title -notin $ignoredVCards -and -Not [string]::IsNullOrEmpty( $value ))
                    {
                        if( $value -is [array] )
                        {
                            $value = ( Expand-StringArray -string $value ) -join $joiner
                        }
                        Add-Member -InputObject $output -MemberType NoteProperty -Name $title -Value $value -Force ## overwrite if already exists - TODO maybe have both
                    }
                }
            }
        }
        $output
    }
}

if( [string]::IsNullOrEmpty( $url ))
{
    # Load required WPF assemblies
    Add-Type -AssemblyName PresentationFramework
    $mainwindowXAML = @"
    <Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="URL Decoder" Height="300" Width="700" Topmost="False">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>

            <Button Name="btnDecode" Content="_Decode URL" Grid.Row="0" Grid.Column="0" Margin="10" Width="120" Background="$buttonBackgroundColour"
                HorizontalAlignment="Left"/>
            <Button Name="btnClean" Content="De_clutter URL" Grid.Row="0" Grid.Column="1" Margin="10" Width="120" Background="$buttonBackgroundColour"
                HorizontalAlignment="Left"/>
            <Button Name="btnCoffee" Content="☕ Buy me a coffee" Grid.Row="0" Grid.Column="2" Margin="10" Width="140" Background="Gold"
                HorizontalAlignment="Center"/>
            <Button Name="btnClear" Content="_Clear" Grid.Row="0" Grid.Column="3" Margin="10" Width="80" Background="$buttonBackgroundColour"
                HorizontalAlignment="Center"/>
            <Button Name="btnDomainInfo" Content="Domain _Info" Grid.Column="4" Grid.Row="0" Margin="10" Width="120" Background="$buttonBackgroundColour"
                HorizontalAlignment="Right"/>
            <TextBox Name="OutputTextBox"
                TextWrapping="Wrap"
                AcceptsReturn="True"
                IsReadOnly="$readOnly"
                Grid.Row="1"
                Margin="10"
                VerticalScrollBarVisibility="Auto"
                ScrollViewer.HorizontalScrollBarVisibility="Disabled"
                FontFamily="$fontName"
                FontSize="$fontSize"
                MaxWidth = "500" 
                Grid.ColumnSpan="5"
                HorizontalAlignment="Stretch"
                VerticalAlignment="Stretch"/>
            <Label Content="Double click URL above to open" Grid.Row="3" Margin="10" Grid.ColumnSpan="4" HorizontalAlignment="Stretch"/>
            <Label Content="" Name="labelStatus" Grid.Row="2" Margin="10" Grid.ColumnSpan="4" HorizontalAlignment="Stretch"/>
        </Grid>
    </Window>
"@

    $domainInfoXAML = @"
    <Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Domain Info" Height="500" Width="700" WindowStartupLocation="CenterScreen" ResizeMode="CanResize">
        <Grid Margin="10">
            <Grid.RowDefinitions>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <TextBox Name="InfoTextBox"
                    Grid.Row="0"
                    FontFamily="Consolas"
                    FontSize="14"
                    TextWrapping="NoWrap"
                    VerticalScrollBarVisibility="Auto"
                    HorizontalScrollBarVisibility="Auto"
                    IsReadOnly="True"
                    AcceptsReturn="True"
                    Margin="0,0,0,0"
                    Text=""/>
            <Button Name="btnClose"
                    Grid.Row="1"
                    Content="_Close"
                    HorizontalAlignment="Right"
                    Width="80"
                    Height="30"
                    IsCancel="True"
                    Margin="0,10,0,0"/>
        </Grid>
    </Window>
"@

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

    Add-Type -AssemblyName PresentationCore , PresentationFramework , System.Windows.Forms , System.Web

    if( -Not ( $mainwindow = New-GUI -inputXAML $mainwindowXAML ) )
    {
        Throw 'Failed to create WPF from XAML'
    }

    Function
    Get-CleanURL
    {
        [CmdletBinding()]
        Param
        (
            [switch]$declutter
        )
        [string]$source = 'Text box'
        [string]$clipboard = $WPFOutputTextBox.Text
        if( [string]::IsNullOrEmpty( $clipboard ))
        {
            $clipboard = Get-Clipboard
            $source = 'Clipboard'
        }
        if( [string]::IsNullOrEmpty( $clipboard ))
        {
            Write-Verbose "Clipboard is empty"
            $WPFlabelStatus.Foreground = [System.Windows.Media.Brushes]::Red
            $wpfLabelStatus.Content = "ERROR: Clipboard is empty"
            return
        }
        if( $clipboard -notmatch '^https?://' )
        {
            Write-Verbose "$source does not contain a URL: $clipboard"
            $WPFlabelStatus.Foreground = [System.Windows.Media.Brushes]::Red
            $wpfLabelStatus.Content = "ERROR: $source does not contain a URL"
            return
        }
        [string]$cleanText = [System.Web.HttpUtility]::UrlDecode( $clipboard )
        if( $declutter )
        {
            Write-Verbose "Decluttering URL"
            $cleanText = $clipboard -replace '\?.*$'
            Write-Verbose "$source text is $clipboard`nCleaned is $cleanText"
        }
        else
        {
            Write-Verbose "Cleaning URL"
            ## remove safelinks and google tracking
            $cleanText = $clipboard -replace '^https://\w+\.safelinks\.protection\.outlook.com/\?url=' -replace '^https://www.google.com/url\?q='
        }
        Write-Verbose "$source text is $clipboard`nCleaned is $cleanText"
        if( $cleanText -eq $clipboard )
        {
            $WPFlabelStatus.Foreground = [System.Windows.Media.Brushes]::Orange
            $wpfOutputTextBox.Text = $clipboard
            $wpfLabelStatus.Content = "URL not changed" 
        }
        else
        {
            Write-Verbose "URL cleaned"
            $WPFlabelStatus.Foreground = [System.Windows.Media.Brushes]::Green
            $WPFlabelStatus.Content = ''
            $wpfOutputTextBox.Text = $cleanText -replace '&data=.+$'
        }
        $WPFOutputTextBox.ScrollToHome()
    }

    Function Get-DomainInfo
    {
        [CmdletBinding()]
        Param
        (
            [string]$Url
        )
        if( $Url -notmatch '^https?://' -and $url -notmatch '^[\w-]+(\.[\w-]+)+$' )
        {
            Write-Verbose "URL does not match expected format: $Url"
            $WPFlabelStatus.Foreground = [System.Windows.Media.Brushes]::Red
            $wpfLabelStatus.Content = "ERROR: URL does not match expected format"
            return
        }
        $wpfLabelStatus.Content = ''
        $domain = $Url -replace '^https?://([^/]+).*$' , '$1'
        Write-Verbose "Domain extracted: $domain"
        $domainInfo = $null
        $error.Clear()
        $domainInfo = Get-RegistrarAndDomainInfo -Url $domain -allProperties:$allProperties -ignoredVCards $ignoredVCards -joiner $joiner -noResolving:$noResolving

        if( $null -eq $domainInfo )
        {
            [void][Windows.MessageBox]::Show( "$domain : $($error[0])" , 'Domain Lookup Failure' , 'Ok' ,'Error' )
        }
        else
        {
            $infoText = $domainInfo | Out-String
            $domainInfoPopup = New-GUI -inputXAML $domainInfoXAML
            if( -Not $domainInfoPopup )
            {
                Write-Error "Failed to create domain info popup"
            }
            else
            {
                $wpfBtnClose.Add_Click({ $domainInfoPopup.Close() })
                $wpfInfoTextBox.Text = $infoText ## .Replace("`r`n", "`n").Replace('"','&quot;')
                $null = $domainInfoPopup.ShowDialog()
            }
        }
    }

    $wpfbtnDecode.Add_Click({
                $_.Handled = $true
                Write-Verbose "Decode clicked"
                Get-CleanURL
            })

    $WPFbtnDomainInfo.Add_Click({
                $_.Handled = $true
                Write-Verbose "Domain info clicked"
                [string]$text = $wpfOutputTextBox.Text
                if( [string]::IsNullOrEmpty( $text ))
                {
                    $text = Get-Clipboard
                    $wpfOutputTextBox.Text = $text
                }
                Get-DomainInfo -Url $text
            })

    $wpfBtnCoffee.Add_Click({
        $_.Handled = $true
        [string]$beggingURL = 'https://www.buymeacoffee.com/guyrleech'
        Start-Process -FilePath $beggingURL -Verb Open
    })
    
    $wpfoutputTextBox.Add_MouseDoubleClick({
        $_.Handled = $true
        Write-Verbose "Output TextBox clicked"
        if( $wpfOutputTextBox.Text -match '^https?://' )
        {
            Start-Process -FilePath $wpfOutputTextBox.Text -verb Open
        }
    })

    $wpfbtnClean.Add_Click({
        $_.Handled = $true
        Write-Verbose "Clean URL clicked"
        Get-cleanURL -declutter
    })
    $wpfBtnClear.Add_Click({
        $_.Handled = $true
        $wpfOutputTextBox.Text = ""
        $wpfLabelStatus.Content = ""
        $wpfLabelStatus.Foreground = [System.Windows.Media.Brushes]::Gray
    })
    $returned = $mainwindow.ShowDialog()
}
else ## url parameter passed so just get domain info and output
{
    Get-RegistrarAndDomainInfo -Url $url -allProperties:$allProperties -ignoredVCards $ignoredVCards -joiner $joiner -noResolving:$noResolving 
}
