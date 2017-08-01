<#
.SYNOPSIS
Checks active devices against the Spiceworks database based on an SQLite query

.DESCRIPTION
Useful for finding device info for SpiceWorks, which is also a limitation as SpiceWorks must be up-to-date and accurate for it to be effective.
Note that this script uses getSpiceData, another internal script.
#>


Import-Module ImportExcel

#Email parameters
$subject="Devices Status Report"
$body="This is a test of the CheckActiveDevices script."
$from="noreply@mdlogistics.com"
$server="email.mdlogistics.com"

#Array containing the email addresses you wish to send the output to
$emailAddresses = @('sgyll@emaildomain.com', 'otheruser@emaildomain.com')

#Path where the file will be saved
$outputFile = ('C:\Temp\Device Status Report.xlsx')

#Initiates the arrays that will be populated below, then used to export to seperate sheets in the Excel file
$final1301 = @()
$final700 = @()
$final2150 = @()
$finalReno = @()
$finalUnknown = @()

#You can edit this sqlite query depending on what you want to find
$query = "SELECT name, type, description, manufacturer, model, ip_address, location FROM devices WHERE location LIKE '%1301%' AND name LIKE '%AP%'"

#Gets data from SpiceWorks based on $query above
$spiceImport = getSpiceData -query $query

#Prints the query to the console
Write-Host 'Getting device data using the following query:'
Write-Host -ForegroundColor Yellow "$query"

#For each device in the list pulled from SpiceWorks
foreach ($device in $spiceImport)
{
    $name = $device.name

    $testnet = ''
    #Tries to ping the device, then stores results in variable $testnet
    try
    {
        $testnet = Test-Connection $device.ip_address -Count 2 -ErrorAction Stop   
    }
    #If unable to ping the computer, it displays that it is not connected.
    catch
    {
        Write-Host "$name is not connected" -ForegroundColor Red
        $device | Add-Member -MemberType NoteProperty -Name Status -Value OFFLINE    
    }

    #If the ping was successful, then...
    if($testnet)
    {
        Write-Host "$name is online" -ForegroundColor Green
        $device | Add-Member -MemberType NoteProperty -Name Status -Value ONLINE
    }

    #checks the location member of $device to match it to the corresponding location, then adds it to the appropriate array
    #If a location cannot be determined, the $device is processed as default, thereby adding it to $finalUnknown
    switch -Wildcard ($device.location)
    {
        '*Location1*'
        {
            $finalL1 += $device
        }

        '*Location2*'
        {
            $finalL2 += $device
        }

        '*Location3*'
        {
            $finalL3 += $device
        }

        '*Location4*'
        {
            $finalL4 += $device
        }

        default
        {
            $finalUnknown += $device
        }
    } #end switch($devicelocation)
} #end foreach ($device in $deviceImport)

#if the object is not empty, output to excel with conditional formatting, else export the blank sheet anyway ###Could probably be cleaned up a lot
if($final1301.Count -ne 0)
{
    $final1301 | Export-Excel $outputFile -WorkSheetname '1301 Perry Road' -AutoSize -BoldTopRow -AutoFilter -ConditionalText $(
        New-ConditionalText OFFLINE darkred -BackgroundColor pink
        New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
        )
}
else
{
    $final1301 | Export-Excel $outputFile -WorkSheetname '1301 Perry Road'
}

if($final700.Count -ne 0)
{
    $final700 | Export-Excel $outputFile -WorkSheetname '700 Perry Road' -AutoSize -BoldTopRow -AutoFilter -ConditionalText $(
        New-ConditionalText OFFLINE darkred -BackgroundColor pink
        New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
        )
}
else
{
    $final700 | Export-Excel $outputFile -WorkSheetname '700 Perry Road'
}

if($final2150.Count -ne 0)
{
    $final2150 | Export-Excel $outputFile -WorkSheetname '2150 Stanley Road' -AutoSize -BoldTopRow -AutoFilter -ConditionalText $(
        New-ConditionalText OFFLINE darkred -BackgroundColor pink
        New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
        )
}
else
{
    $final2150 | Export-Excel $outputFile -WorkSheetname '2150 Stanley Road'
}

if($finalReno.Count -ne 0)
{
    $finalReno | Export-Excel $outputFile -WorkSheetname 'Reno' -AutoSize -BoldTopRow -AutoFilter -ConditionalText $(
        New-ConditionalText OFFLINE darkred -BackgroundColor pink
        New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
        )
}
else
{
    $finalReno | Export-Excel $outputFile -WorkSheetname 'Reno'
}

if($finalUnknown.Count -ne 0)
{
    $finalUnknown | Export-Excel $outputFile -WorkSheetname 'Unknown' -AutoSize -BoldTopRow -AutoFilter -ConditionalText $(
        New-ConditionalText OFFLINE darkred -BackgroundColor pink
        New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
        )
}
else
{
    $finalUnknown | Export-Excel $outputFile -WorkSheetname 'Unknown'
}

#send email to each email address in $emailAddresses

foreach ($email in $emailAddresses)
{
    Send-MailMessage -To $email -From $from -Body $body -SmtpServer $server -Subject $subject -Attachments $outputFile
}

