#Import-Module ImportExcel
$path = 'C:/Temp/BuildNumbers.xlsx'

$computers = Get-ADComputer -Filter 'Name -like "mdl*"' -Properties * | Select-Object name,OperatingSystemVersion | Where-Object {($_.name -match '[MDL]\d{3}$' -and $_.name -notmatch 'MDL154' -and $_.name -notmatch 'MDL245')}
$computers | Export-Excel $path -Show -AutoSize -AutoFilter
