Import-Module ActiveDirectory

$outputFile = ('C:\Temp\SEKO Group Report.xlsx')

$Groups = Get-ADGroup -LDAPFilter "(name=*SEKO*)"

Foreach($group in $groups)
{
    $entry = Get-ADGroupMember $group | Select-Object name, distinguishedName
    if($entry.count -ne 0)
    {
     $entry | Export-Excel $outputFile -WorksheetName $Group.Name -AutoSize -BoldTopRow -AutoFilter
    }
}