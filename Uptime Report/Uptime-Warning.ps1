. 'C:\Users\sgyll\Documents\GitLab\Shem-PowerShell\Uptime Report\Get-UptimeDays.ps1'

#email Variables
$subject="MDL IT Support - Computer Restart Notice"
$from="noreply@mdlogistics.com"
$server="email.mdlogistics.com"


#$computers = Get-ADComputer -Filter 'Name -like "mdl*"' | Where-Object {$_.name -match '[MDL]\d{3}$' -and $_.name -notmatch 'MDL154' -and $_.name -notmatch 'MDL245' }
$computers = Get-ADComputer mdl365

$Workstations = foreach($computer in $computers)
{
    $test = Test-Connection $computer.name -Count 1 -ErrorAction SilentlyContinue
        if($test)
            {
                Get-UptimeDays -ComputerName $computer.Name
            
            }

}

foreach($Workstation in $Workstations)
{
    if($Workstation.Uptime -gt 0)
    {
        #$Workstation
        #Restart-Computer -ComputerName $Workstation.ComputerName -Force
        $body="Please restart your computer. `nIt has been at least 14 days since your last restart. `nYour computer ($($Workstation.ComputerName)) must be restarted to install updates. If you don't restart, your computer will be restarted Saturday at Midnight."
        
        $email = sgyll@mdlogistics.com #$($Workstation.user).TrimStart('M','D','L','\') + '@mdlogistics.com'
        
        Send-MailMessage -To $email -From $from -Body $body -SmtpServer $server -Subject $subject -BodyAsHtml
        
    }
}