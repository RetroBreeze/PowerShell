$domaincontrollers =@('MDLAD01.MDL.Local', 'MDLAD02.MDL.Local')

$days=((Get-Date).AddDays(-7))

foreach($dc in $domaincontrollers){

$eventlog = Get-WinEvent -ComputerName MDLAD01.MDL.Local -FilterHashtable @{Logname='Security'; StartTime=$days}| select TimeCreated,ID,Message | Export-Csv -Path C:\Temp\passwords.csv

}