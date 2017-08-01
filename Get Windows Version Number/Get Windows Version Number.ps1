#Get the windows version number for every computer on the network (we name our computers MDLXXX)
$computers = Get-ADComputer -Filter 'Name -like "mdl*"' -Properties * | Select-Object name,OperatingSystemVersion | Where-Object {($_.name -match '[MDL]\d{3}$' -and $_.name -notmatch 'MDL154' -and $_.name -notmatch 'MDL245') | Sort-Object name}
