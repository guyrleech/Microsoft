<#
    Get unique email addresses from Outlook folders

    Guy Leech, 2018
#>

<#
.SYNOPSIS

Interrogate Outlook data to find SMTP email addresses and output to csv or on-screen together with other contextual information such as the date and subject

.DESCRIPTION

CC: recipients and required and optional meeting attendees will also be captured.

Only the first occurence of an email address will be output so the results will only contain unique email addresses.

.PARAMETER includedFolders

A comma separated list of folder name regular expressions to interrogate. Folders not in the list will not be interrogated.

.PARAMETER excludedFolders

A comma separated list of folder name regular expressions not to interrogate. Folders not in the list will be interrogated.

.PARAMETER excludedAddresses

A comma separated list of email address regular expressions which will be ignored.

.PARAMETER includedAddresses

A comma separated list of email address regular expressions which will be included and those not matching will be excluded.

.PARAMETER csv

The name, and path, to a file which will have the results written to it. If not specified then output will be to an on screen grid view

.EXAMPLE

& '.\Outlook Miner.ps1' -includedFolders Inbox,'Sent Items' -excludedAddresses support,sales,reply,donotreply,trello.com -csv .\outlook.addresses.csv

Interrogate the users Innox and Sent Items folders for SMTP email addresses and output them to the named csv file

.NOTES

It works with mail folders and calendars. It does not work with contact items since these can be exported via Outlook anyway.

#>

[CmdletBinding()]

Param
(
    [string[]]$includedFolders ,
    [string[]]$excludedFolders ,
    [string[]]$excludedAddresses = @( 'DoNotReply' , 'NoReply' ) ,
    [string[]]$includedAddresses ,
    [string]$csv ,
    [int]$progressEvery = 20 ,
    [int]$maxItems 
)

Function Is-Excluded( [string]$item , [string[]]$exclusions )
{
    [bool]$excluded = $false

    if( [string]::IsNullOrEmpty( $item ) )
    {
        $excluded = $true 
    }
    else
    {
        ForEach( $exclusion in $exclusions )
        {
            if( $item -match $exclusion )
            {
                $excluded = $true
                break
            }
        }
    }
    $excluded
}

Function Is-Included( [string]$item , [string[]]$inclusions )
{
    [bool]$included = $false

    if( ! $inclusions -or ! $inclusions.Count )
    {
        $included = $true ## no explicit inclusions so everything is included
    }
    else
    {
        ForEach( $inclusion in $inclusions )
        {
            if( $item -match $inclusion )
            {
                $included = $true
                break
            }
        }
    }
    $included
}

Function Process-Recipients( $recipients , $item , [hashtable]$results , [string[]]$excludedAddresses , [string[]]$includedAddresses , [hashtable]$addressesExcluded )
{
    $recipients | ForEach-Object `
    {
        $recipient = $_
        try
        {
            ## will either be a recipient object or just a string depending on whether a mail or calendar item
            [string]$address = $null
            if( Get-Member -InputObject $recipient -Name Address -ErrorAction SilentlyContinue )
            {
                $address = $recipient.Address
            }
            else
            {
                $address = $recipient.Trim()
            }
            if( ! [string]::IsNullOrEmpty( $address ) -and $address.IndexOf( '@' ) -gt 0 )
            {
                if( ! ( Is-Excluded $address $excludedAddresses ) `
                    -and ( Is-Included $address $includedAddresses ) )
                {
                    ## We don't have complete info when sent to SMTP address
                    $results.Add( $address , ( New-Object -TypeName PSCustomObject -Property @{ 'Email address' = $address ; 'Date' = $item.SentOn ;'Subject' = $item.Subject ; 'Name' = $null ; 'Domain' = ($address -split '@')[-1] ; 'Folder' = $folder.Name} ) )
                    Write-Verbose "`t$address"
                }
                else
                {
                    $addressesExcluded.Add( $address , $folder.Name )
                }
            }
        }
        catch [System.ArgumentException]
        {
            if( $_.Exception.Message -notmatch '^Item has already been added\. Key in dictionary' ) ## duplicate hash table entry
            {
                throw $_
            }
        }
        catch
        {
            throw $_
        }
    }
}

if( ! $progressEvery )
{
    $ProgressPreference = 'SilentlyContinue'
}

$null = Add-type -assembly "Microsoft.Office.Interop.Outlook"

$outlook = New-Object -ComObject outlook.application

$namespace = $outlook.GetNameSpace("MAPI")

$olClass = "Microsoft.Office.Interop.Outlook.OlObjectClass" -as [type] 

if( ! $outlook -or ! $namespace )
{
    Write-Error "Failed to create Outlook objects"
    return 1
}

[int]$folders = 0
[long]$messages = 0

[hashtable]$addresses = @{}
[hashtable]$addressesExcluded = @{}

[string[]]$ownEmailAddresses = @( $namespace.Accounts|select -ExpandProperty CurrentUser|Select -ExpandProperty Address )

$namespace.Folders| ForEach-Object { $_.Folders } | ForEach-Object `
{
    $folder = $_ 

    Write-Verbose "Got $($folder.Items.Count) items in folder `"$($folder.Name)`""

    [bool]$skipFolder = $false
    ForEach( $excludedFolder in $excludedFolders )
    {
        if( $excludedFolder -match $folder.Name )
        {
            $skipFolder = $true
            break
        }
    }

    if( ! $skipFolder -and $includedFolders -and $includedFolders.Count )
    {
        $skipFolder = $true

        ForEach( $includedFolder in $includedFolders )
        {
            if( $includedFolder -match $folder.Name )
            {
                $skipFolder = $false
                break
            }
        }
    }

    if( ! $skipFolder )
    {
        $folders++
        [int]$counter = 0

        $folder.Items | ForEach-Object `
        {
            if( $counter % $progressEvery -eq 0 )
            {
                Write-Progress -Activity "Reading items in folder `"$($folder.Name)`"" -PercentComplete ( ($counter / $folder.items.Count ) * 100 ) -Status "Found $($addresses.Count) unique addresses in total"
            }

            try
            {
                $messages++
                if( ! $maxItems -or $counter -lt $maxItems ) 
                {
                    if( $_.Class -eq $olClass::olMail )
                    {
                        $email = $_
                        ## see if we are the initiator in which case we add the recipients instead, e.g. Sent Items
                        if( $ownEmailAddresses -contains $email.SenderEmailAddress -and $email.Recipients.Count )
                        {
                            Process-Recipients $email.Recipients $email $addresses $excludedAddresses  $includedAddresses  $addressesExcluded
                        }
                        elseif( $email.SenderEmailType -ne 'EX' -and ! [string]::IsNullOrEmpty( $email.SenderEmailAddress ) )
                        {
                            if( ! ( Is-Excluded $email.SenderEmailAddress $excludedAddresses ) `
                                -and ( Is-Included $email.SenderEmailAddress $includedAddresses ) )
                            {
                                $addresses.Add( $email.SenderEmailAddress , ( New-Object -TypeName PSCustomObject -Property @{ 'Email address' = $email.SenderEmailAddress ; 'Date' = $email.SentOn ; 'Subject' = $email.Subject ; 'Name' = $email.SenderName ; 'Domain' = ($email.SenderEmailAddress -split '@')[-1] ; 'Folder' = $folder.Name} ) )
                                Write-Verbose "`t$($email.SenderEmailAddress)"
                                Process-Recipients $email.Recipients $email $addresses $excludedAddresses  $includedAddresses  $addressesExcluded
                            }
                            else
                            {
                                $addressesExcluded.Add( $email.SenderEmailAddress , $folder.Name )
                            }
                        }
                    }
                    elseif( $_.Class -eq $olClass::olAppointment )
                    {
                        $appointment = $_
                        if( $appointment.Organizer -and $appointment.Organizer.Indexof( '@' ) -gt 0 `
                            -and ! ( Is-Excluded $appointment.Organizer $excludedAddresses ) `
                            -and ( Is-Included $appointment.Organizer $includedAddresses ) )
                        {
                            $addresses.Add( $appointment.Organizer , ( New-Object -TypeName PSCustomObject -Property @{ 'Email address' = $appointment.Organizer ; 'Date' = $appointment.Start ; 'Subject' = $appointment.Subject ; 'Name' = $email.SenderName ; 'Domain' = ($appointment.Organizer -split '@')[-1] ; 'Folder' = $folder.Name} ) )
                            Write-Verbose "`t$($appointment.Organizer)"
                        }
                        ## can't add a returned hashtable to main hashtable as a single exception would cause all addresses in the returned data to fail to add
                        Process-Recipients ( $appointment.RequiredAttendees -split ';' ) $appointment $addresses $excludedAddresses  $includedAddresses  $addressesExcluded
                        Process-Recipients ( $appointment.OptionalAttendees -split ';' ) $appointment $addresses $excludedAddresses  $includedAddresses  $addressesExcluded
                    }
                }
            }
            catch [System.ArgumentException]
            {
                if( $_.Exception.Message -notmatch '^Item has already been added\. Key in dictionary' ) ## duplicate hash table entry
                {
                    throw $_
                }
            }
            $counter++
        }
    }
    else
    {
        Write-Verbose "Skipping `"$($folder.Name)`""
    }
}

Write-Progress -PercentComplete 100 -Activity "Finished"

[string]$message = "Got $($addresses.Count) unique email addresses from $messages items in $folders folders. $($addressesExcluded.Count) excluded addresses"

Write-Verbose $message

if( $addresses -and $addresses.Count )
{
    if( [string]::IsNullOrEmpty( $csv ) )
    {
        $selected = $addresses.GetEnumerator() | Select -ExpandProperty Value | Out-GridView -PassThru -Title $message
        if( $selected )
        {
            $selected | clip.exe
        }
    }
    else
    {
        $addresses.GetEnumerator() | Select -ExpandProperty Value | Export-Csv -NoTypeInformation -NoClobber -Path $csv
    }
}
