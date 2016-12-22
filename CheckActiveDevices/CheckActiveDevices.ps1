Import-Module ImportExcel

$subject="Devices Status Report"
$body="This is a test of the CheckActiveDevices script."
$from="noreply@mdlogistics.com"
$server="email.mdlogistics.com"

$outputFile = ('C:\Temp\Device Status Report.xlsx')

#You can edit this sqlite query depending on what you wan
$query = "SELECT name, type, description, manufacturer, model, ip_address, location FROM devices WHERE manufacturer LIKE 'Cisco'"

Write-Host 'Pulling device data using the following query:'
Write-Host -ForegroundColor Yellow "$query"
$spiceImport = getSpiceData -query $query

$final1301 = @()
$final700 = @()
$final2150 = @()
$finalReno = @()
$finalUnknown = @()



$deviceList = foreach ($device in $spiceImport)
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

    if($testnet)
    {
        Write-Host "$name is online" -ForegroundColor Green
        $device | Add-Member -MemberType NoteProperty -Name Status -Value ONLINE
    }

    switch -Wildcard ($device.location)
    {
        '*1301*'
        {
            $final1301 += $device
        }

        '*700*'
        {
            $final700 += $device
        }

        '*2150*'
        {
            $final2150 += $device
        }

        '*Reno*'
        {
        $finalReno += $device
        }

        default
        {
        $finalUnknown += $device
        }

    } #end switch($devicelocation)
}#end foreach ($device in $deviceImport)

if($final1301.Count -ne 0){
$final1301 | Export-Excel $outputFile -WorkSheetname '1301 Perry Road' -AutoSize  -ConditionalText $(
                New-ConditionalText OFFLINE darkred -BackgroundColor pink
                New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
            )
            }

if($final700.Count -ne 0){
$final700 | Export-Excel $outputFile -WorkSheetname '700 Perry Road' -AutoSize  -ConditionalText $(
                New-ConditionalText OFFLINE darkred -BackgroundColor pink
                New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
            )
            }

if($final2150.Count -ne 0){
$final2150 | Export-Excel $outputFile -WorkSheetname '2150 Stanley Road' -AutoSize  -ConditionalText $(
                New-ConditionalText OFFLINE darkred -BackgroundColor pink
                New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
          )
          }

if($finalReno.Count -ne 0){
$finalReno | Export-Excel $outputFile -WorkSheetname 'Reno' -AutoSize  -ConditionalText $(
                New-ConditionalText OFFLINE darkred -BackgroundColor pink
                New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
            )
            }

if($finalUnknown.Count -ne 0){
$finalUnknown | Export-Excel $outputFile -WorkSheetname 'Unknown' -AutoSize  -ConditionalText $(
                New-ConditionalText OFFLINE darkred -BackgroundColor pink
                New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
            )
            }

#send email
#Send-MailMessage -To 'sgyll@mdlogistics.com' -From $from -Body $body -SmtpServer $server -Subject $subject -Attachments $outputFile
<#
#send email
Send-MailMessage -To 'rdalton@mdlogistics.com' -From $from -Body $body -SmtpServer $server -Subject $subject -Attachments '\\MDLFPS01\PowerShell\Device List\Offline Device Report.xlsx'

#send email
Send-MailMessage -To 'eratcliff@mdlogistics.com' -From $from -Body $body -SmtpServer $server -Subject $subject -Attachments '\\MDLFPS01\PowerShell\Device List\Offline Device Report.xlsx'
#>

