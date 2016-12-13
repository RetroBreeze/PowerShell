Import-Module ImportExcel

$subject="Windows Update Log"
$body="This is a test of the Windows Update Log script. Please see the attached Excel file, which contains a list of updates installed on three computers on the network."
$from="noreply@mdlogistics.com"
$server="email.mdlogistics.com"

$path = 'C:\Temp\WindowsUpdateLog.xlsx'
$computers = Get-ADComputer -Filter 'Name -like "mdl*"' -Properties * | Select-Object name | Where-Object {($_.name -match '[MDL]\d{3}$' -and $_.name -notmatch 'MDL154' -and $_.name -notmatch 'MDL245')}
#$computers = @('MDL260')
$temp = New-Object System.Collections.ArrayList
 
foreach ($comp in $computers)
{

if($comp -ne $ENV:COMPUTERNAME){
$out = Invoke-Command -ComputerName $comp.name -ErrorAction SilentlyContinue -ScriptBlock{

$wu = new-object -com “Microsoft.Update.Searcher”

$totalupdates = $wu.GetTotalHistoryCount()
$all = $wu.QueryHistory(0,$totalupdates)

Foreach ($update in $all)
        {
    $string = $update.title
    $dateString = $update.Date
    
    $Regex = “KB\d*”
    $KB = $string | Select-String -Pattern $regex | Select-Object { $_.Matches }
     $output = New-Object -TypeName PSobject
     $output | add-member NoteProperty “Date” -value $dateString
     $output | add-member NoteProperty “HotFixID” -value $KB.‘ $_.Matches ‘.Value
     $output | add-member NoteProperty “Title” -value $string
     $output 
        }
    }
$temp.Clear()
$temp += $out
$temp | Where {$_.Date -ne '12/30/1899 12:00:00 AM'} | Sort-Object Date -Descending | Export-Excel $path -WorkSheetname $comp -AutoSize

}
else
{

$doodle = @()

$wu = new-object -com “Microsoft.Update.Searcher”

$totalupdates = $wu.GetTotalHistoryCount()
$all = $wu.QueryHistory(0,$totalupdates)

Foreach ($update in $all)
        {
    $string = $update.title
    $dateString = $update.Date
    
    $Regex = “KB\d*”
    $KB = $string | Select-String -Pattern $regex | Select-Object { $_.Matches }
     $output = New-Object -TypeName PSobject
     $output | add-member NoteProperty “Date” -value $dateString
     $output | add-member NoteProperty “HotFixID” -value $KB.‘ $_.Matches ‘.Value
     $output | add-member NoteProperty “Title” -value $string
     $doodle +=$output 
        }
$temp.Clear()
$temp += $doodle
$temp | Where {$_.Date -ne '12/30/1899 12:00:00 AM'} | Sort-Object Date -Descending | Export-Excel $path -WorkSheetname $comp.name -AutoSize
    }
}

<#
#send email
Send-MailMessage -To 'sgyll@mdlogistics.com' -From $from -Body $body -SmtpServer $server -Subject $subject -Attachments C:\Temp\WindowsUpdateLog.xlsx

#send email
Send-MailMessage -To 'rdalton@mdlogistics.com' -From $from -Body $body -SmtpServer $server -Subject $subject -Attachments C:\Temp\WindowsUpdateLog.xlsx

#send email
Send-MailMessage -To 'eratcliff@mdlogistics.com' -From $from -Body $body -SmtpServer $server -Subject $subject -Attachments C:\Temp\WindowsUpdateLog.xlsx
#>
