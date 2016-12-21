Import-Module ImportExcel

$todaysDate = Get-Date -Format d

$subject="Devices Status Report"
$body="This is a test of the CheckActiveDevices script."
$from="noreply@mdlogistics.com"
$server="email.mdlogistics.com"

$outputFile = ('C:\Temp\Device Status Report.xlsx')
#$deviceImport = Import-CSV -Path '\\MDLFPS01\PowerShell\Device List\DevicesToCheck.csv'

#Cisco only
$spiceImport = getSpiceData -query "SELECT name, type, description, manufacturer, model, ip_address, location FROM devices WHERE manufacturer LIKE 'Cisco'"

$final1301 = @()
$final700 = @()
$final2150 = @()
$finalReno = @()
$finalUnknown = @()

$deviceList = foreach ($device in $spiceImport)
{
$deviceIP = $device.ip_address
$deviceName = $device.name
$deviceType = $device.type
$deviceDescription = $device.description
$deviceManufacturer = $device.manufacturer
$deviceModel = $device.model
$deviceLocation = $device.location
$status = ''

  $testnet = ''
    #Tries to ping the device, then stores results in variable $testnet
    try
    {
        $testnet = Test-Connection $deviceIP -Count 2 -ErrorAction Stop
        
    }
    #If unable to ping the computer, it displays that it is not connected.
    catch
    {
        Write-Host "$deviceName is not connected" -ForegroundColor Red
        $device | Add-Member -MemberType NoteProperty -Name Status -Value OFFLINE
        
    }

    if($testnet)
    {
        Write-Host "$deviceName is online" -ForegroundColor Green
        $device | Add-Member -MemberType NoteProperty -Name Status -Value OFFLINE
    }

    switch -Wildcard ($device.location)
    {
        '*1301*'
        {
            $final1301 += $device
         <#   $device | Export-Excel $outputFile -WorkSheetname '1301 Perry Road' -AutoSize -ConditionalText $(
                New-ConditionalText OFFLINE darkred -BackgroundColor pink
                New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
            )
            <#
            $1301obj = New-Object -TypeName PSobject
            $1301obj | Add-Member -MemberType NoteProperty -Name DeviceIP -Value $deviceIP
            $1301obj | Add-Member -MemberType NoteProperty -Name DeviceName -Value $deviceName
            $1301obj | Add-Member -MemberType NoteProperty -Name DeviceType -Value $deviceType
            $1301obj | Add-Member -MemberType NoteProperty -Name DeviceDescription -Value $deviceDescription
            $1301obj | Add-Member -MemberType NoteProperty -Name DeviceManufacturer -Value $deviceManufacturer
            $1301obj | Add-Member -MemberType NoteProperty -Name DeviceModel -Value $deviceModel
            $1301obj | Add-Member -MemberType NoteProperty -Name DeviceLocation -Value '1301 Perry Road'
            $1301obj | Add-Member -MemberType NoteProperty -Name Status -Value $status
            <#$1301obj | Export-Excel $outputFile -WorkSheetname '1301 Perry Road' -AutoSize -ConditionalText $(
                New-ConditionalText OFFLINE darkred -BackgroundColor pink
                New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
            )#>
        }

        '*700*'
        {
            $final700 += $device
        <#$device | Export-Excel $outputFile -WorkSheetname '700 Perry Road' -AutoSize -ConditionalText $(
                New-ConditionalText OFFLINE darkred -BackgroundColor pink
                New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
            )
        <#
            $700obj = New-Object -TypeName PSobject
            $700obj | Add-Member -MemberType NoteProperty -Name DeviceIP -Value $deviceIP
            $700obj | Add-Member -MemberType NoteProperty -Name DeviceName -Value $deviceName
            $700obj | Add-Member -MemberType NoteProperty -Name DeviceType -Value $deviceType
            $700obj | Add-Member -MemberType NoteProperty -Name DeviceDescription -Value $deviceDescription
            $700obj | Add-Member -MemberType NoteProperty -Name DeviceManufacturer -Value $deviceManufacturer
            $700obj | Add-Member -MemberType NoteProperty -Name DeviceModel -Value $deviceModel
            $700obj | Add-Member -MemberType NoteProperty -Name DeviceLocation -Value '700 Perry Road'
            $700obj | Add-Member -MemberType NoteProperty -Name Status -Value $status
            $700obj | Export-Excel $outputFile -WorkSheetname '700 Perry Road' -AutoSize -ConditionalText $(
                New-ConditionalText OFFLINE darkred -BackgroundColor pink
                New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
            )
            #>
        }

        '*2150*'
        {
            $final2150 += $device
        <#$device | Export-Excel $outputFile -WorkSheetname '2150 Stanley Road' -AutoSize -ConditionalText $(
                New-ConditionalText OFFLINE darkred -BackgroundColor pink
                New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
            )
        <#
            $2150obj = New-Object -TypeName PSobject
            $2150obj | Add-Member -MemberType NoteProperty -Name DeviceIP -Value $deviceIP
            $2150obj | Add-Member -MemberType NoteProperty -Name DeviceName -Value $deviceName
            $2150obj | Add-Member -MemberType NoteProperty -Name DeviceType -Value $deviceType
            $2150obj | Add-Member -MemberType NoteProperty -Name DeviceDescription -Value $deviceDescription
            $2150obj | Add-Member -MemberType NoteProperty -Name DeviceManufacturer -Value $deviceManufacturer
            $2150obj | Add-Member -MemberType NoteProperty -Name DeviceModel -Value $deviceModel
            $2150obj | Add-Member -MemberType NoteProperty -Name DeviceLocation -Value '2150 Stanley Road'
            $2150obj | Add-Member -MemberType NoteProperty -Name Status -Value $status
            $2150obj | Export-Excel $outputFile -WorkSheetname '2150 Stanley Road' -AutoSize -ConditionalText $(
                New-ConditionalText OFFLINE darkred -BackgroundColor pink
                New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
            )
            #>
        }

        '*Reno*'
        {
        $finalReno += $device
        <#$device | Export-Excel $outputFile -WorkSheetname 'Reno' -AutoSize -ConditionalText $(
                New-ConditionalText OFFLINE darkred -BackgroundColor pink
                New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
            )
        <#
            $renoobj = New-Object -TypeName PSobject
            $renoobj | Add-Member -MemberType NoteProperty -Name DeviceIP -Value $deviceIP
            $renoobj | Add-Member -MemberType NoteProperty -Name DeviceName -Value $deviceName
            $renoobj | Add-Member -MemberType NoteProperty -Name DeviceType -Value $deviceType
            $renoobj | Add-Member -MemberType NoteProperty -Name DeviceDescription -Value $deviceDescription
            $renoobj | Add-Member -MemberType NoteProperty -Name DeviceManufacturer -Value $deviceManufacturer
            $renoobj | Add-Member -MemberType NoteProperty -Name DeviceModel -Value $deviceModel
            $renoobj | Add-Member -MemberType NoteProperty -Name DeviceLocation -Value 'Reno'
            $renoobj | Add-Member -MemberType NoteProperty -Name Status -Value $status
            $renoobj | Export-Excel $outputFile -WorkSheetname 'Reno' -AutoSize -ConditionalText $(
                New-ConditionalText OFFLINE darkred -BackgroundColor pink
                New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
            )
            #>
        }

        default
        {
        $finalUnknown += $device
        <#$device | Export-Excel $outputFile -WorkSheetname 'Unknown' -AutoSize -ConditionalText $(
                New-ConditionalText OFFLINE darkred -BackgroundColor pink
                New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
            )
        <#
            $defobj = New-Object -TypeName PSobject
            $defobj | Add-Member -MemberType NoteProperty -Name DeviceIP -Value $deviceIP
            $defobj | Add-Member -MemberType NoteProperty -Name DeviceName -Value $deviceName
            $defobj | Add-Member -MemberType NoteProperty -Name DeviceType -Value $deviceType
            $defobj | Add-Member -MemberType NoteProperty -Name DeviceDescription -Value $deviceDescription
            $defobj | Add-Member -MemberType NoteProperty -Name DeviceManufacturer -Value $deviceManufacturer
            $defobj | Add-Member -MemberType NoteProperty -Name DeviceModel -Value $deviceModel
            $defobj | Add-Member -MemberType NoteProperty -Name DeviceLocation -Value $deviceLocation
            $defobj | Add-Member -MemberType NoteProperty -Name Status -Value $status
            $defobj | Export-Excel $outputFile -WorkSheetname 'Undetermined' -AutoSize  -ConditionalText $(
                New-ConditionalText OFFLINE darkred -BackgroundColor pink
                New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
            )
            #>
        }

    } #end switch($devicelocation)
}#end foreach ($device in $deviceImport)

$final1301 | Export-Excel $outputFile -WorkSheetname '1301 Perry Road' -AutoSize  -ConditionalText $(
                New-ConditionalText OFFLINE darkred -BackgroundColor pink
                New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
            )

$final700 | Export-Excel $outputFile -WorkSheetname '700 Perry Road' -AutoSize  -ConditionalText $(
                New-ConditionalText OFFLINE darkred -BackgroundColor pink
                New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
            )

$final2150 | Export-Excel $outputFile -WorkSheetname '2150 Stanley Road' -AutoSize  -ConditionalText $(
                New-ConditionalText OFFLINE darkred -BackgroundColor pink
                New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
          )

$finalReno | Export-Excel $outputFile -WorkSheetname 'Reno' -AutoSize  -ConditionalText $(
                New-ConditionalText OFFLINE darkred -BackgroundColor pink
                New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
            )

$finalUnknown | Export-Excel $outputFile -WorkSheetname 'Unknown' -AutoSize  -ConditionalText $(
                New-ConditionalText OFFLINE darkred -BackgroundColor pink
                New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
            )
#$finalOutput | Export-Excel $outputFile -WorkSheetname '1301 Perry Road' -AutoSize

#send email
#Send-MailMessage -To 'sgyll@mdlogistics.com' -From $from -Body $body -SmtpServer $server -Subject $subject -Attachments $outputFile
<#
#send email
Send-MailMessage -To 'rdalton@mdlogistics.com' -From $from -Body $body -SmtpServer $server -Subject $subject -Attachments '\\MDLFPS01\PowerShell\Device List\Offline Device Report.xlsx'

#send email
Send-MailMessage -To 'eratcliff@mdlogistics.com' -From $from -Body $body -SmtpServer $server -Subject $subject -Attachments '\\MDLFPS01\PowerShell\Device List\Offline Device Report.xlsx'
#>

