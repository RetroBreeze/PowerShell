Import-Module ImportExcel

$todaysDate = Get-Date -Format d

$subject="Devices Status Report"
$body="This is a test of the CheckActiveDevices script."
$from="noreply@mdlogistics.com"
$server="email.mdlogistics.com"

$outputFile = ('\\MDLFPS01\PowerShell\Device List\Device Status Report ' + $todaysDate + '.xlsx')
$deviceImport = Import-CSV -Path '\\MDLFPS01\PowerShell\Device List\DevicesToCheck.csv'

$finalOutput = @{}

$deviceList = foreach ($device in $deviceImport)
{
$deviceIP = $device.IP
$deviceName = $device.Name
$deviceType = $device.Type
$deviceNotes = $device.Notes
$deviceLocation = $device.Location
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
        $status = 'OFFLINE'
    }

    if($testnet)
    {
        Write-Host "$deviceName is online" -ForegroundColor Green
        $status = 'ONLINE'
    }

    switch -Wildcard ($deviceLocation)
    {
        '*1301*'
        {
            $1301obj = New-Object -TypeName PSobject
            $1301obj | Add-Member -MemberType NoteProperty -Name DeviceIP -Value $deviceIP
            $1301obj | Add-Member -MemberType NoteProperty -Name DeviceName -Value $deviceName
            $1301obj | Add-Member -MemberType NoteProperty -Name DeviceType -Value $deviceType
            $1301obj | Add-Member -MemberType NoteProperty -Name DeviceNotes -Value $deviceNotes
            $1301obj | Add-Member -MemberType NoteProperty -Name DeviceLocation -Value '1301 Perry Road'
            $1301obj | Add-Member -MemberType NoteProperty -Name Status -Value $status
            $1301obj | Export-Excel $outputFile -WorkSheetname $1301obj.DeviceLocation -AutoSize -BoldTopRow -FreezeTopRow -AutoFilter -NoClobber -ConditionalText $(
                New-ConditionalText OFFLINE darkred -BackgroundColor pink
                New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
            )
        }

        '*700*'
        {
            $700obj = New-Object -TypeName PSobject
            $700obj | Add-Member -MemberType NoteProperty -Name DeviceIP -Value $deviceIP
            $700obj | Add-Member -MemberType NoteProperty -Name DeviceName -Value $deviceName
            $700obj | Add-Member -MemberType NoteProperty -Name DeviceType -Value $deviceType
            $700obj | Add-Member -MemberType NoteProperty -Name DeviceNotes -Value $deviceNotes
            $700obj | Add-Member -MemberType NoteProperty -Name DeviceLocation -Value '700 Perry Road'
            $700obj | Add-Member -MemberType NoteProperty -Name Status -Value $status
            $700obj | Export-Excel $outputFile -WorkSheetname $700obj.DeviceLocation -AutoSize -BoldTopRow -FreezeTopRow -AutoFilter -NoClobber -ConditionalText $(
                New-ConditionalText OFFLINE darkred -BackgroundColor pink
                New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
            )
        }

        '*2150*'
        {
            $2150obj = New-Object -TypeName PSobject
            $2150obj | Add-Member -MemberType NoteProperty -Name DeviceIP -Value $deviceIP
            $2150obj | Add-Member -MemberType NoteProperty -Name DeviceName -Value $deviceName
            $2150obj | Add-Member -MemberType NoteProperty -Name DeviceType -Value $deviceType
            $2150obj | Add-Member -MemberType NoteProperty -Name DeviceNotes -Value $deviceNotes
            $2150obj | Add-Member -MemberType NoteProperty -Name DeviceLocation -Value '2150 Stanley Road'
            $2150obj | Add-Member -MemberType NoteProperty -Name Status -Value $status
            $2150obj | Export-Excel $outputFile -WorkSheetname $2150Obj.DeviceLocation -AutoSize -BoldTopRow -FreezeTopRow -AutoFilter -NoClobber -ConditionalText $(
                New-ConditionalText OFFLINE darkred -BackgroundColor pink
                New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
            )
        }

        '*Reno*'
        {
            $renoobj = New-Object -TypeName PSobject
            $renoobj | Add-Member -MemberType NoteProperty -Name DeviceIP -Value $deviceIP
            $renoobj | Add-Member -MemberType NoteProperty -Name DeviceName -Value $deviceName
            $renoobj | Add-Member -MemberType NoteProperty -Name DeviceType -Value $deviceType
            $renoobj | Add-Member -MemberType NoteProperty -Name DeviceNotes -Value $deviceNotes
            $renoobj | Add-Member -MemberType NoteProperty -Name DeviceLocation -Value 'Reno'
            $renoobj | Add-Member -MemberType NoteProperty -Name Status -Value $status
            $renoobj | Export-Excel $outputFile -WorkSheetname Reno -AutoSize -BoldTopRow -FreezeTopRow -AutoFilter -NoClobber -ConditionalText $(
                New-ConditionalText OFFLINE darkred -BackgroundColor pink
                New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
            )
        }

        default
        {
            $defobj = New-Object -TypeName PSobject
            $defobj | Add-Member -MemberType NoteProperty -Name DeviceIP -Value $deviceIP
            $defobj | Add-Member -MemberType NoteProperty -Name DeviceName -Value $deviceName
            $defobj | Add-Member -MemberType NoteProperty -Name DeviceType -Value $deviceType
            $defobj | Add-Member -MemberType NoteProperty -Name DeviceNotes -Value $deviceNotes
            $defobj | Add-Member -MemberType NoteProperty -Name DeviceLocation -Value $deviceLocation
            $defobj | Add-Member -MemberType NoteProperty -Name Status -Value $status
            $defobj | Export-Excel $outputFile -WorkSheetname Undetermined -AutoSize -BoldTopRow -FreezeTopRow -AutoFilter -NoClobber -ConditionalText $(
                New-ConditionalText OFFLINE darkred -BackgroundColor pink
                New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
            )
        }

    } #end switch($devicelocation)
}#end foreach ($device in $deviceImport)

send email
Send-MailMessage -To 'sgyll@mdlogistics.com' -From $from -Body $body -SmtpServer $server -Subject $subject -Attachments $outputFile
<#
#send email
Send-MailMessage -To 'rdalton@mdlogistics.com' -From $from -Body $body -SmtpServer $server -Subject $subject -Attachments '\\MDLFPS01\PowerShell\Device List\Offline Device Report.xlsx'

#send email
Send-MailMessage -To 'eratcliff@mdlogistics.com' -From $from -Body $body -SmtpServer $server -Subject $subject -Attachments '\\MDLFPS01\PowerShell\Device List\Offline Device Report.xlsx'
#>

