<#
    Find all IIS servers in the domain, optionally filtered by OU, Group and/or name, or contained in a text file, check IIS certificates and email results of soon to expire ones or write to file or grid view

    @guyrleech 2018

    Modification History:

    05/01/19  GRL  Added parameter to save discovered servers to text file
                   Added -overwrite switch
#>

<#
.SYNOPSIS

Find all IIS servers in the domain, optionally filtered by OU, Group and/or name, check IIS certificates and email results of soon to expire ones or write to file or grid view

.PARAMETER servers

Comma separated list of server names to check

.PARAMETER serversFile

A text file containing one server per line. Lines not starting with an alphanumeric character are ignored as are characters after a space to the end of the line

.PARAMETER outputFile

Write the names of servers where IIS is discovered, along with subject and expiry date of any IIS certificates found, to the named file. Fails if file exists unless -overwrite is specified
The file can then be given to -serversFile for input so that only previously known IIS servers are queried.

.PARAMETER serverLike

Servers matching this regular expression will be checked

.PARAMETER OUs

A comma separated list of Organisational Units whose members will be checked, if they match the -serverLike argument if specified

.PARAMETER groups

A comma separated list of Active Directory groups  whose members will be checked, if they match the -serverLike argument if specified

.PARAMETER csv

Name of a non-existent CSV file which will be created with the list of soon to expire certificates

.PARAMETER overwrite

If the files specified via -csv or -outputFile already exist then they will fail to be written unless -overwrite is specified

.PARAMETER gridView

A sortable/filterable grid view will be shown with all soon to expire certificates

.PARAMETER emailIfNone

Will still send an email even if no soon to expire certificates are found. Use to know the script has run even if their are no soon to expire certificates found

.PARAMETER mailServer

The SMTP mail server to use to send the email message

.PARAMETER proxyMailServer

The name of a computer which is allowed to relay SMTP messages. Use this option if the computer running the script is prohibited from sending email via the SMTP email server

.PARAMETER from

The email address from which the email message will appear to be sent from

.PARAMETER subject

The subject for the email. A default is provided if one is not specified

.PARAMETER recipients

A comma separated list of recipients to send the email message to

.PARAMETER expiryDays

The number of days within which a cerificate will be reported if it expires

.PARAMETER remoteTimeout

The time in seconds that a remote command to interrogate IIS is allowed to run for

.PARAMETER domain

The Active Directory domain in which the computers reside

.PARAMETER OS

The operating system running on the computers

.EXAMPLE

& '.\Find and check IIS server certs.ps1' -serverLike 'IIS' -gridview

Find all servers in the domain with 'IIS' in their name and show any IIS certificates expiring within the next 60 days in a grid view

.EXAMPLE

& '.\Find and check IIS server certs.ps1' -OUs 'constoso.com/Sites/Dewsbury/Computers/IIS' -expirydays 90 -recipients bob@uncle.com -mailServer smtpserver

Find all servers in the specified OU, find any IIS certificates expiring within the next 90 days and email the results to bob@uncle.com via the smtp email server called 'smtpserver'

.EXAMPLE

& '.\Find and check IIS server certs.ps1' -Groups 'IIS Servers','Web Servers' -expirydays 45 -csv \\fileserver\share\reports\iis.expiring.certs.csv

Find all servers in the two specified AD groups, find any IIS certificates expiring within the next 45 days and write to the csv file specified in the UNC

.EXAMPLE

& '.\Find and check IIS server certs.ps1' -servers webserver01,webserver02,webserver03 -expirydays 45 -csv \\fileserver\share\reports\iis.expiring.certs.csv

Find any IIS certificates expiring within the next 45 days on the three specified servers and write to the csv file specified in the UNC

.EXAMPLE

& '.\Find and check IIS server certs.ps1' -serversFile \\fileserver\share\data\iis-servers.txt -expirydays 42 -csv \\fileserver\share\reports\iis.expiring.certs.csv

Find any IIS certificates expiring within the next 45 days on the servers contained withint the file "iis-servers.txt" and write to the csv file specified in the UNC

.NOTES

Requires the 'Guys.Common.Functions.psm1' module, available at github.com/guyrleech/Citrix

#>

[CmdletBinding()]

Param
(
    [string[]]$servers ,
    [string]$serversFile ,
    [string]$serverLike ,
    [string[]]$OUs ,
    [string[]]$groups ,
    [string]$csv ,
    [string]$outputFile ,
    [switch]$overwrite ,
    [switch]$gridView,
    [switch]$emailIfNone ,
    [string]$mailserver ,
    [string]$proxyMailServer = $env:Computername ,
    [string]$from = "$env:Computername@$env:userdnsdomain" ,
    [string]$subject  ,
    [string[]]$recipients ,
    [int]$expiryDays = 60 ,
    [int]$remoteTimeout = 30 ,
    [string]$domain ,
    [string]$OS = 'Windows*Server*' ,
    [string]$guysModule = 'Guys.Common.Functions.psm1'
)

Function Get-GroupMembers( $group , [string]$serverLike )
{
    $group.psbase.invoke('members') | ForEach-Object `
    {
        $adspath = $_.GetType().InvokeMember( 'ADSPath' ,  'GetProperty',  $null,  $_, $null)
        $class = $_.GetType().InvokeMember( 'Class' ,  'GetProperty',  $null,  $_, $null)

        if( $class -eq 'group' )
        {
            Get-GroupMembers -group ([ADSI]"$adspath,$class") -serverLike $serverLike
        }
        elseif( $class -eq 'user' )
        {
            $account = ([ADSI]"$adspath,$class")
            if( $account -and $account.UserFlags -and $account.UserFlags.Value -band 0x3000 -and $account.Name -match $serverLike ) # SERVER_TRUST_ACCOUNT or WORKSTATION_TRUST_ACCOUNT
            {
                $account.name -replace '\$$' , ''
            }
        }
    }
}

Import-Module (Join-Path ( & { Split-Path -Path $myInvocation.ScriptName -Parent } ) $guysModule ) -ErrorAction Stop

if( $PSBoundParameters[ 'serversFile' ] )
{
    $servers += @( Get-Content -Path $serversFile -ErrorAction Stop | ForEach-Object `
    {
        if( $_ -match '^[-a-z0-9]' )
        {
            ($_ -split '\s')[0]
        }
    })
}

## scheduled tasks don't handle arrays properly
if( $OUs -and $OUs.Count -eq 1 -and $OUs[0].IndexOf( ',' ) -ge 0 )
{
    $OUs = $OUs[0] -split ','
}

if( $groups -and $groups.Count -eq 1 -and $groups[0].IndexOf( ',' ) -ge 0 )
{
    $groups = $groups[0] -split ','
}

if( ! $PSBoundParameters[ 'domain' ] )
{
    $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name
}

if( $OUs -and $OUs.Count )
{
    ForEach( $OU in $OUs )
    {  
        ## see if canonical name (e.g. copied from AD Users & computers) and convert to distinguished
        if( $OU.IndexOf('=') -lt 0 )
        {
            try
            {
                ## http://www.itadmintools.com/2011/09/translate-active-directory-name-formats.html
                $NameTranslate = New-Object -ComObject NameTranslate 
                [System.__ComObject].InvokeMember( 'Init' , 'InvokeMethod', $null , $NameTranslate , @( 3 , $null ) ) ## ADS_NAME_INITTYPE_GC
                [System.__ComObject].InvokeMember( 'Set' , 'InvokeMethod', $null , $NameTranslate , (2 ,$OU)) ## CANONICALNAME
                $OU = [System.__ComObject].InvokeMember('Get' , 'InvokeMethod' , $null , $NameTranslate , 1) ## DISTINGUISHEDNAME
            }
            catch
            {
                Write-Warning "Failed to translate OU `"$OU`" from canonical name to distinguished`n$_"
                $OU = $null
            }
        }
        if( ! [string]::IsNullOrEmpty( $OU ) )
        {
            $searcher = [adsisearcher]([adsi]"LDAP://$OU")
            $searcher.Filter = "(&(objectCategory=Computer)(operatingsystem=$OS))"
            $searcher.PropertiesToLoad.AddRange( @( 'name' , 'operatingsystem' , 'adspath' ) )
            [System.Collections.ArrayList]$found = $searcher.FindAll()
            if( ! $found -or ! $found.Count )
            {
                Write-Warning "ADSI search returned no computers running a `"$OS`" operating system"
            }
            $servers += @( ($found | Select -ExpandProperty 'properties').name | Where-Object { $_ -match $serverLike } )
        }
    }
}

if( $groups -and $groups.Count )
{
    $servers += @( ForEach( $group in $groups )
    {
        $thisGroup = [ADSI]"WinNT://$domain/$group,group"
        try
        {
            if( $thisGroup -and $thisGroup.PSObject.properties[ 'Path' ] )
            {
                ## recursively get group members
                Get-GroupMembers -group $thisGroup -serverLike $serverLike
            }
            else
            {
                Write-Warning "Failed to find group `"$group`""
            }
        }
        catch
        {
            Write-Warning "Failed to find group `"$group`""
        }
    })
}

## only do global search for computers by name if not tried to find any via other means
if( ( ! $servers -or ! $servers.Count ) -and ! $OUs  -and ! $groups -and $PSBoundParameters[ 'serverLike' ] )
{
    $searcher = [adsisearcher]"(&(objectCategory=Computer)(operatingsystem=$OS))"
    $searcher.PropertiesToLoad.AddRange( @( 'name' , 'operatingsystem' , 'adspath' ) )
    [System.Collections.ArrayList]$found = $searcher.FindAll()
    if( ! $found -or ! $found.Count )
    {
        Write-Warning 'ADSI search returned no computers running a Windows Server operating system'
    }
    $servers += ($found | Select -ExpandProperty 'properties').name | Where-Object { $_ -match $serverLike }
}

if( ! $servers -or ! $servers.Count )
{
    Throw 'No computers to operate on. Use a combination of -servers, -OUs, -groups and -serverLike to specify'
}

[int]$counter = 0
[scriptblock]$remoteWork = `
{
    Get-Service -Name 'W3SVC' -ErrorAction SilentlyContinue
    Import-Module -Name 'WebAdministration' -ErrorAction SilentlyContinue
    $bindings = Get-ChildItem IIS:SSLBindings
    @( ForEach( $binding in $bindings )
    {
        Get-ChildItem cert:localMachine/My | Where-Object { $_.Thumbprint -eq $binding.thumbprint } | ForEach-Object `
        {
            Add-Member -InputObject $_ -MemberType NoteProperty -Name 'Binding' -Value $binding.PSChildName
            $_
        }
    } )
}

[datetime]$expiryDate = (Get-Date).AddDays( $expiryDays )
[hashtable]$remoteParams = @{
    'jobTimeout' = $remoteTimeout
    'work' = $remoteWork }
if( $VerbosePreference -eq 'Continue' )
{
    $remoteParams.Add( 'Verbose' , $true )
}
else
{
    $remoteParams.Add( 'Quiet' , $true )
}

## Remove any duplicates
$servers = $servers | Sort -Unique

Write-Verbose "Got $($servers.Count) computers to check: $($servers -join ',')"

## Keep a list of servers so we can write it to a ttext file if requested
$IISservers = New-Object -TypeName 'System.Collections.ArrayList'

[array]$results = @( ForEach( $server in $servers )
{
    $counter++
    Write-Verbose "$counter / $($servers.Count) : $server"
    ## Get W3SVC service and any certs but do via runspace so can time out quickly
    $webService,[array]$IIScertificates = Get-RemoteInfo -computer $server @remoteParams
    if( $webService )
    {
        Write-Verbose "$server has IIS"
        [string]$outputLine = "$server # "
        ## Now check for any expiring certificates within the window
        if( $IIScertificates -and $IIScertificates.Count )
        {
            $outputLine += "$($IIScertificates.Count) certs "
            ForEach( $cert in $IIScertificates )
            {
                $outputLine += " : `"$($cert.subject)`" $(Get-Date $cert.NotAfter -Format G)"
                if( $cert.NotAfter -le $expiryDate )
                {
                    Write-Warning "Cert from $server expires on $($cert.NotAfter)"
                    $cert | Select @{n='Computer';e={$_.PSComputerName}} , Binding , Subject, Thumbprint , NotAfter
                }
            }
        }
        else
        {
            $outputLine += 'no IIS certs found'
        }
        [void]$IISservers.Add( $outputLine )
    }
} )

[hashtable]$clobber = @{ 'NoClobber' = (! $overwrite) }

if( ! $results -or ! $results.Count )
{
    Write-Output "Found no certificates on $($servers.Count) computers expiring on or before $(Get-Date -Date $expiryDate -Format G)"
}
elseif( ! [string]::IsNullOrEmpty( $csv ) )
{
    $results | Export-Csv -Path $csv -NoTypeInformation @clobber
}

if( $IISservers -and $IISservers.Count -and ! [string]::IsNullOrEmpty( $outputFile ) )
{
    $IISservers | Out-File -FilePath $outputFile @clobber
}

if( $recipients -and $recipients.Count -and ! [string]::IsNullOrEmpty( $mailserver ) `
    -and (( $results -and $results.Count ) -or $emailIfNone ) )
{
    if( $recipients -and $recipients.Count -eq 1 -and $recipients[0].IndexOf(',') -ge 0 )
    {
        $recipients = $recipients[0] -split ','
    }
    
    if( [string]::IsNullOrEmpty( $subject ) )
    {
        $subject = "$($results.Count) IIS certificates found on $($servers.Count) computers expiring on or before $(Get-Date -Date $expiryDate -Format G)"
    }

    [string]$style = "<style>BODY{font-family: Arial; font-size: 10pt;}"
    $style += "TABLE{border: 1px solid black; border-collapse: collapse;}"
    $style += "TH{border: 1px solid black; background: #dddddd; padding: 5px;}"
    $style += "TD{border: 1px solid black; padding: 5px;}"
    $style += "</style>"

    [string]$htmlBody = $results | Sort -Property 'NotAfter' | ConvertTo-Html -Fragment -PreContent "<h2>IIS certificates expiring on or before $(Get-Date -Date $expiryDate -Format G)<h2>" | Out-String
    $htmlBody = ConvertTo-Html -PostContent $htmlBody -Head $style

    Invoke-Command -ComputerName $proxyMailServer -ScriptBlock { Send-MailMessage -Subject $using:subject -BodyAsHtml -Body $using:htmlBody -From $using:from -To $using:recipients -SmtpServer $using:mailserver }
}

if( $results -and $results.Count -and $gridView )
{
    $selected = @( $results  | Out-GridView -Title "$($results.Count) certificates across $($servers.Count) computers expiring on or before $(Get-Date -Date $expiryDate -Format G)" -PassThru )
    if( $selected -and $selected.Count )
    {
        $selected | clip.exe
    }
}
