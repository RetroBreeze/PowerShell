$files =@("\\mdlfps01\IT\petyaVaccine\perfc","\\mdlfps01\IT\petyaVaccine\perfc.dat", "\\mdlfps01\IT\petyaVaccine\perfc.dll")
$computers=@()

$failed = @()

$OU = Get-ADOrganizationalUnit -Filter {Name -like '*Workstations*'} | Select-Object -ExpandProperty DistinguishedName

for($i=0; $i -lt $OU.Length; $i++){
$computers += Get-ADComputer -SearchBase $OU[$i] -Filter '*' -SearchScope Subtree | Select-Object name
}

foreach($target in $computers)
{
    $x = Copy-Item -Path \\mdlfps01\IT\petyaVaccine\perfc -Destination "\\$target\c$\Windows\perfc" -PassThru -force -ErrorAction silentlyContinue
    if ($x) { Write-Host "$target : perfc file copied" -ForegroundColor Green }
    else { Write-Host "$target : perfc file not copied" -ForegroundColor Red; $failed += $target} 

    $y = Copy-Item -Path \\mdlfps01\IT\petyaVaccine\perfc.dat -Destination "\\$target\c$\Windows\perfc.dat" -PassThru -force -ErrorAction silentlyContinue
    if ($y) {  Write-Host "$target : perfc.dat copied" -ForegroundColor Green }
    else {Write-Host "$target : perfc.dat not copied" -ForegroundColor Red;  $failed += $target} 

    $z = Copy-Item -Path \\mdlfps01\IT\petyaVaccine\perfc.dll -Destination "\\$target\c$\Windows\perfc.dll" -PassThru -force -ErrorAction silentlyContinue
    if ($z) {  Write-Host "$target : perfc.dll copied" -ForegroundColor Green }
    else {Write-Host "$target : perfc.dll not copied" -ForegroundColor Red;  $failed += $target} 
}

Write-Host "Consider checking the following:`n$failed" -ForegroundColor Red
Read-Host "Done"