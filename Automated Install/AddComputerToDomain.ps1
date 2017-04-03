$Type = Get-WmiObject Win32_Computersystem | Select-object Model

$name = $env:computername

$domainList = @{
#1301 Perry Road
"1301/Office/Desktops" = "OU=Desktops,OU=Office-Workstations,OU=1301,OU=Shared Pharma,DC=MDL,DC=Local";
"1301/Ops/Desktops" = "OU=Desktops,OU=Ops-Workstations,OU=1301,OU=Shared Pharma,DC=MDL,DC=Local";
"1301/Receiving/Desktops" = "OU=Desktops,OU=RecOffice-Workstations,OU=1301,OU=Shared Pharma,DC=MDL,DC=Local";
"1301/Office/Laptops" = "OU=Laptops,OU=Office-Workstations,OU=1301,OU=Shared Pharma,DC=MDL,DC=Local";
"1301/Ops/Laptops" = "OU=Laptops,OU=Ops-Workstations,OU=1301,OU=Shared Pharma,DC=MDL,DC=Local";
"1301/Receiving/Laptops" = "OU=Laptops,OU=RecOffice-Workstations,OU=1301,OU=Shared Pharma,DC=MDL,DC=Local";

#2150 Stanley Road
"2150/Office/Desktops" = "OU=Desktops,OU=Office-Workstations,OU=Stanley,OU=Shared Pharma,DC=MDL,DC=Local";
"2150/Ops/Desktops" = "OU=Desktops,OU=Ops-Workstations,OU=Stanley,OU=Shared Pharma,DC=MDL,DC=Local";
"2150/Receiving/Desktops" = "OU=Desktops,OU=RecOffice-Workstations,OU=Stanley,OU=Shared Pharma,DC=MDL,DC=Local";
"2150/Retail Ops/Desktops" = "OU=Desktops,OU=Retail-Ops-Workstations,OU=Stanley,OU=Shared Pharma,DC=MDL,DC=Local";
"2150/Office/Laptops" = "OU=Laptops,OU=Office-Workstations,OU=Stanley,OU=Shared Pharma,DC=MDL,DC=Local";
"2150/Ops/Laptops" = "OU=Laptops,OU=Ops-Workstations,OU=Stanley,OU=Shared Pharma,DC=MDL,DC=Local";
"2150/Receiving/Laptops" = "OU=Laptops,OU=RecOffice-Workstations,OU=Stanley,OU=Shared Pharma,DC=MDL,DC=Local";
"2150/Retail Ops/Laptops" = "OU=Laptops,OU=Retail-Ops-Workstations,OU=Stanley,OU=Shared Pharma,DC=MDL,DC=Local";

#700 Perry Road
"700/Office/Desktops" = "OU=Desktops,OU=Office-Workstations,OU=700,OU=Retail,DC=MDL,DC=Local";
"700/IT/Desktops" = "OU=Desktops,OU=IT-Workstations,OU=700,OU=Retail,DC=MDL,DC=Local";
"700/Ops/Desktops" = "OU=Desktops,OU=Ops-Workstations,OU=700,OU=Retail,DC=MDL,DC=Local";
"700/Office/Laptops" = "OU=Laptops,OU=Office-Workstations,OU=700,OU=Retail,DC=MDL,DC=Local";
"700/IT/Laptops" = "OU=Laptops,OU=IT-Workstations,OU=700,OU=Retail,DC=MDL,DC=Local";
"700/Ops/Laptops" = "OU=Laptops,OU=Ops-Workstations,OU=700,OU=Retail,DC=MDL,DC=Local";
}

$selectedOU = $domainList.GetEnumerator() | sort -Property name | Out-GridView -Title "Select an OU"  -PassThru
#$selectedOU = $domainList | Out-GridView -Title "Select an OU"  -PassThru
$cred = Get-Credential
Add-Computer -DomainName MDL.Local -Credential $cred -OUPath $selectedOU -Restart

