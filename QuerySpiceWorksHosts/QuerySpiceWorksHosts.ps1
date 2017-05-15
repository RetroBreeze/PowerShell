$computers = Get-ADComputer -Filter * -SearchBase "OU=Stanley,OU=Shared Pharma,DC=MDL,DC=Local"

$myList

$output = @()

foreach($comp in $computers.name)
{
    #$ip = [System.Net.Dns]::GetHostAddresses("$comp")
    $query = "SELECT name, ip_address, location FROM devices WHERE name LIKE '$comp'"

    #Gets data from SpiceWorks based on $query above
    $spiceImport = getSpiceData -query $query

    if($spiceImport.Columns = ''){
    Write-Host "No data found for $comp"
    }
    Else
    {
    $output += $spiceImport
    }
}
$output
