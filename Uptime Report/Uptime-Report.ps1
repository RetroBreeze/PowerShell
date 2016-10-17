. '\\mdlfps01\Powershell\Get-Uptime.ps1'

#email Variables
$subject="Computer Uptime Report"
$body="See attached"
$from="noreply@mdlogistics.com"
$server="email.mdlogistics.com"

$computers = Get-ADComputer -Filter 'Name -like "mdl260"' | Where-Object {$_.name -match '[MDL]\d{3}$' -and $_.name -notmatch 'MDL154' -and $_.name -notmatch 'MDL245' }

#$computers =@('mdl260','mdl365')

$Workstations = foreach($computer in $computers)
{
    $test = Test-Connection $computer.name -Count 1 -ErrorAction SilentlyContinue
        if($test)
            {
                Get-Uptime -ComputerName $computer.Name
            
            }

}

if(Test-Path "$env:TEMP\uptime.xlsx")
{
    Remove-Item "$env:TEMP\uptime.xlsx" -Force
}

$Spreadsheet = foreach($workstation in $Workstations)
{
<#
    $props = @{
        'Computer Name'=$Workstation.ComputerName;
        'User'=$Workstation.User;
        'Last Boot Up Time'=$Workstation.LastBootUpTime;
        'Uptime'=$Workstation.Uptime;
        'OU'=(Get-ADComputer $Workstation.ComputerName).DistinguishedName
    }
#>

    $propsobj = New-Object -TypeName psobject
    $propsObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $props.'Computer Name'
    $propsObj | Add-Member -MemberType NoteProperty -Name User -Value $props.User
    $propsObj | Add-Member -MemberType NoteProperty -Name LastBootUpTime -Value $props.'Last Boot Up Time'
    $propsObj | Add-Member -MemberType NoteProperty -Name Uptime -Value $props.Uptime
    $propsObj | Add-Member -MemberType NoteProperty -Name OU -Value $props.OU
    $propsobj
}


$Spreadsheet | Sort-Object 'Last Boot Up Time' | Export-Excel -FreezeTopRow -AutoFilter -AutoSize -Path C:\Temp\uptime.xlsx
<#
#send email
Send-MailMessage -To 'eratcliff@mdlogistics.com' -From $from -Body $body -SmtpServer $server -Subject $subject -Attachments $env:TEMP\uptime.xlsx

#send email
Send-MailMessage -To 'rdalton@mdlogistics.com' -From $from -Body $body -SmtpServer $server -Subject $subject -Attachments $env:TEMP\uptime.xlsx

#>
#send email
#Send-MailMessage -To 'sgyll@mdlogistics.com' -From $from -Body $body -SmtpServer $server -Subject $subject -Attachments $env:TEMP\uptime.xlsx
