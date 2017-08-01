<#
.SYNOPSIS
Shows the incorrect password attempts made on each domain controller in the last 7 days.
#>

$domaincontrollers =@('AD01.Domain.Local', 'AD02.Domain.Local')

$days=((Get-Date).AddDays(-7))

foreach($dc in $domaincontrollers){

$eventlog = Get-WinEvent -ComputerName $dc -FilterHashtable @{Logname='Security'; StartTime=$days}| select TimeCreated,ID,Message

}