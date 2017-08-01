<#
.SYNOPSIS
Creates an excel file for each computer listed in the $computers array, which contains the update history for that device.
#>

Import-Module ImportExcel

#Email settings
$subject="Windows Update Log"
$body="This is a test of the Windows Update Log script. Please see the attached Excel file, which contains a list of updates installed on three computers on the network."
$from=""
$server=""
$path = 'C:\Temp\WindowsUpdateLog.xlsx'

#$computers = Get-ADComputer -Filter 'Name -like "mdl279"' -Properties * | Select-Object name | Where-Object {($_.name -match '[MDL]\d{3}$' -and $_.name -notmatch 'MDL154' -and $_.name -notmatch 'MDL245')}
$computers = @('#List of computers here')
$temp = New-Object System.Collections.ArrayList
 
foreach ($comp in $computers)
{

if($comp.name -ne $ENV:COMPUTERNAME){

#Create an WSUS update searcher object, gets the update history, then parses it out with a regular expression
$out = Invoke-Command -ComputerName $comp -ErrorAction SilentlyContinue -ScriptBlock{

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

#Clear the old $temp variable and sort it before outputting to the excel sheet
$temp.Clear()
$temp += $out
$temp | Where {$_.Date -ne '12/30/1899 12:00:00 AM'} | Sort-Object Date -Descending | Export-Excel $path -WorkSheetname $comp -AutoSize

}
}
else
{

$outputArray = @()

#Create a new Update Searcher object
$wu = new-object -com “Microsoft.Update.Searcher”
$totalupdates = $wu.GetTotalHistoryCount()
$all = $wu.QueryHistory(0,$totalupdates)

#Format the data into an object, then add to the $outputArray array
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
     $outputArray +=$output 
        }

#Clear the $temp variable then add $outputArray before sorting and exporting the array to Excel 
$temp.Clear()
$temp += $outputArray
$temp | Where {$_.Date -ne '12/30/1899 12:00:00 AM'} | Sort-Object Date -Descending | Export-Excel $path -WorkSheetname $comp.name -AutoSize
    }


<#
#send email
Send-MailMessage -To 'sgyll@mdlogistics.com' -From $from -Body $body -SmtpServer $server -Subject $subject -Attachments C:\Temp\WindowsUpdateLog.xlsx

#send email
Send-MailMessage -To 'rdalton@mdlogistics.com' -From $from -Body $body -SmtpServer $server -Subject $subject -Attachments C:\Temp\WindowsUpdateLog.xlsx

#send email
Send-MailMessage -To 'eratcliff@mdlogistics.com' -From $from -Body $body -SmtpServer $server -Subject $subject -Attachments C:\Temp\WindowsUpdateLog.xlsx
#>
