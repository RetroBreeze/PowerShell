$HOST.UI.RawUI.BackgroundColor = "Black"
$HOST.UI.RawUI.ForegroundColor = "Green"

$header = 'Hostname','OSName','OSVersion','OSManufacturer','OSConfig','Buildtype',`
'RegisteredOwner','RegisteredOrganization','ProductID','InstallDate','StartTime','Manufacturer',`
'Model','Type','Processor','BIOSVersion','WindowsFolder','SystemFolder','StartDevice','Culture',`
'UICulture','TimeZone','PhysicalMemory','AvailablePhysicalMemory','MaxVirtualMemory',`
'AvailableVirtualMemory','UsedVirtualMemory','PagingFile','Domain','LogonServer','Hotfix',`
'NetworkAdapter'

$disk = Get-WmiObject Win32_LogicalDisk -ComputerName $env:computername -Filter "DeviceID='C:'" |
Select-Object Size,FreeSpace

$result = systeminfo.exe /FO CSV | Select-Object -Skip 1 | ConvertFrom-Csv -Header $header

$UserName, $ComputerName = $env:username, $env:computername
$Date = Get-Date -format D
$DateStr = $Date.ToString()

Write-Host "Good morning! Welcome to Basic Auto-Start." -ForegroundColor Yellow
sleep -m 100
"`n"  
Write-Host "Today's date is $DateStr." -ForegroundColor Yellow 
sleep -m 80
Write-Host "`n`nUsername: $UserName
Computer Name: $ComputerName"
OS: $result.OSName $result.OSVersion "
Computer Info: "$result.Model $result.Type"
RAM: "$result.PhysicalMemory "| Available RAM: "$result.AvailablePhysicalMemory"" 
Write-Host "C Size: "([math]::Round($disk.Size / 1GB))"GB" "| Free: " ([math]::Round($disk.FreeSpace / 1GB))"GB "
"`n"
sleep -m 60

Write-Host "Press any key to continue..."

$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
Write-Host

Write-Host "Opening daily software..." -ForegroundColor Yellow
"`n`n`n"

Invoke-Item "C:\Program Files\Microsoft Office\Office14\OUTLOOK.exe"
Write-Host "Outlook..."
Invoke-Item "C:\Program Files\Microsoft Office\Office14\ONENOTE.exe"
Write-Host "OneNote..."
Invoke-Item "C:\Windows\System32\WindowsPowerShell\v1.0\powershell_ise.exe"
Write-Host "Powershell ISE..."
Invoke-Item "C:\Program Files (x86)\Mitel\Unified Communicator Advanced 7.0\UCA.exe"
Write-Host "MiCollab..."
Invoke-Item "C:\Windows\SystemApps\Microsoft.MicrosoftEdge_8wekyb3d8bbwe\MicrosoftEdge.exe"
Write-Host "Edge..."
Invoke-Item "C:\Users\sgyll\Desktop\admin.bat"
Write-Host "Admin Console..."

sleep 5

Write-Host "See ya!"
sleep -m 20
#cls
