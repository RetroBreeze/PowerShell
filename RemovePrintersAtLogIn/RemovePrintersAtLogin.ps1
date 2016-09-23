"`n"
$pArray = @()
$printers = Get-Printer
foreach ($p in $printers)
    {
        if($p.ComputerName -eq "mdlfps01")
            {
                $name = $p.Name
                $date = Get-Date
                Write-Host "Removing $name from $ENV:COMPUTERNAME" -ForegroundColor Yellow
                #Remove-Printer $p

                $obj = New-Object -TypeName psobject
                $obj | Add-Member -MemberType NoteProperty -Name Date -Value $date.ToShortDateString()
                $obj | Add-Member -MemberType NoteProperty -Name PrinterName -Value $name
                $obj | Add-Member -MemberType NoteProperty -Name Hostname -Value $ENV:COMPUTERNAME
                $pArray += $obj
            }
    }

if($pArray.Count -ne 0)
    {
        $pArray | Export-Csv -Path "\\mdlfps01\IT\Printer Removal Log\Printer Removal Log.csv" -Append -NoTypeInformation -Force
        $pArray = ''
        "`nLog saved to \\mdlfps01\IT\Printer Removal Log\Printer Removal Log.csv!"
    }
    else
    {
        "`nNo printers were found on mdlfps01."
    }