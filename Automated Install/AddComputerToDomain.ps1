<#
.SYNOPSIS
This script will add the local computer to the domain selected in out-gridview.

.DESCRIPTION
Domains are hard-coded as strings with friendly names for ease of use. This script is to be used as part of the AutoInstall bundle. Domains have been dummified intentionally for github)
#>

#Hard-coded array containing list of domains available. First strings are friendly names, second string is the correctly syntaxed OU path.
$domainList = @{
"dummyLocation/OU1/Desktops" = "OU=Desktops,OU=OU1,OU=dummyLocation,DC=Domain,DC=Local";
"dummyLocation/OU2/Desktops" = "OU=Desktops,OU=OU2,OU=dummyLocation,DC=Domain,DC=Local";
"dummyLocation/OU3/Desktops" = "OU=Desktops,OU=OU3,OU=dummyLocation,DC=Domain,DC=Local";
"dummyLocation/OU1/Laptops" = "OU=Laptops,OU=OU1,OU=dummyLocation,DC=Domain,DC=Local";
"dummyLocation/OU2/Laptops" = "OU=Laptops,OU=OU2,OU=dummyLocation,DC=Domain,DC=Local";
"dummyLocation/OU3/Laptops" = "OU=Laptops,OU=OU3,OU=dummyLocation,DC=Domain,DC=Local";

}

#Create a gridview with passthru containing friendly names
$selectedOU = $domainList.GetEnumerator() | sort -Property name | Out-GridView -Title "Select an OU"  -PassThru

#Get domain credentials
$cred = Get-Credential

#Add the computer to the domain
Add-Computer -DomainName MDL.Local -Credential $cred -OUPath $selectedOU -Restart

