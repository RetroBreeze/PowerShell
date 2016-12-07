Import-Module ImportExcel

$outputFile = 'C:\Users\sgyll\Desktop\dlist.xlsx'
$deviceImport = Import-CSV -Path 'C:\Temp\deviceList.csv'

$finalOutput = @{}

$deviceList = foreach ($device in $deviceImport)
{

$deviceIP = $device.IP
$deviceName = $device.Name
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

      <#$obj = New-Object -TypeName PSobject
        $obj | Add-Member -MemberType NoteProperty -Name DeviceIP -Value $deviceIP
        $obj | Add-Member -MemberType NoteProperty -Name DeviceName -Value $deviceName
        $obj | Add-Member -MemberType NoteProperty -Name DeviceLocation -Value $deviceLocation
        $obj | Add-Member -MemberType NoteProperty -Name Status -Value $status#>

        switch($deviceLocation){

        '1301 Perry Road'{
        $1301obj = New-Object -TypeName PSobject
        $1301obj | Add-Member -MemberType NoteProperty -Name DeviceIP -Value $deviceIP
        $1301obj | Add-Member -MemberType NoteProperty -Name DeviceName -Value $deviceName
        $1301obj | Add-Member -MemberType NoteProperty -Name DeviceLocation -Value $deviceLocation
        $1301obj | Add-Member -MemberType NoteProperty -Name Status -Value $status
        $1301obj
        }
        '700 Perry Road'{
        $700obj = New-Object -TypeName PSobject
        $700obj | Add-Member -MemberType NoteProperty -Name DeviceIP -Value $deviceIP
        $700obj | Add-Member -MemberType NoteProperty -Name DeviceName -Value $deviceName
        $700obj | Add-Member -MemberType NoteProperty -Name DeviceLocation -Value $deviceLocation
        $700obj | Add-Member -MemberType NoteProperty -Name Status -Value $status
        $700obj
        }
        '2150 Stanley Road'{
        $2150obj = New-Object -TypeName PSobject
        $2150obj | Add-Member -MemberType NoteProperty -Name DeviceIP -Value $deviceIP
        $2150obj | Add-Member -MemberType NoteProperty -Name DeviceName -Value $deviceName
        $2150obj | Add-Member -MemberType NoteProperty -Name DeviceLocation -Value $deviceLocation
        $2150obj | Add-Member -MemberType NoteProperty -Name Status -Value $status
        $2150obj
        }
        } #end switch($devicelocation)

}#end foreach ($device in $deviceImport)

ExportExcel -Show -AutoSize $outputFile $devicelist
