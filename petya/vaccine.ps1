$files =@("\\mdlfps01\IT\petyaVaccine\perfc","\\mdlfps01\IT\petyaVaccine\perfc.dat", "\\mdlfps01\IT\petyaVaccine\perfc.dll")
$computers=@()

$OU = Get-ADOrganizationalUnit -Filter {Name -like '*Workstations*'} | Select-Object -ExpandProperty DistinguishedName

for($i=0; $i -lt $OU.Length; $i++){
$computers += Get-ADComputer -SearchBase $OU[$i] -Filter '*' -SearchScope Subtree | Select-Object name
}

foreach($target in $computers)
{
    Copy-Item -Path \\mdlfps01\IT\petyaVaccine\perfc -Destination "\\$target\c$\Windows\perfc";
    Copy-Item -Path \\mdlfps01\IT\petyaVaccine\perfc.dat -Destination "\\$target\c$\Windows\perfc.dat";
    Copy-Item -Path \\mdlfps01\IT\petyaVaccine\perfc.dll -Destination "\\$target\c$\Windows\perfc.dll";
}