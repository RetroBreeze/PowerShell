$printers = import-csv -Path '\\mdlphawmsdev\c$\Scripts\Stan_Printers.csv'

$zebraPrinters = $printers | Where {$_.Type -eq 'Zebra'}
$laserPrinters = $printers | Where {$_.Type -eq 'Laser'}

$zebraDriver = 'C:\Windows\inf\ntprint.inf'
$zebraModel = 'Generic / Text Only'


$laserModel = 'HP Universal Printing PCL 6'

foreach ($zebraPrinter in $zebraPrinters)
{
    $name = $zebraPrinter.PrinterName
    $model = $zebraPrinter.Driver
    $type = $zebraPrinter.Type
    $ip = $zebraPrinter.IP

    Write-Output "Creating printer port IP $ip"
    
    Invoke-Command -ScriptBlock{
    & cscript C:\Windows\System32\Printing_Admin_Scripts\en-US\prnport.vbs -a -r "IP_$ip" -h $ip -o raw
    }

    Write-Output "Installing printer $name"

    Invoke-Command -ScriptBlock{
    & RunDLL32.exe printui.dll,PrintUIEntry /if /b "$name" /r "IP_$ip" /f "$zebraDriver" /m "$zebraModel"
    }
    Start-Sleep -s 10
    "$name installed!"
}

foreach ($laserPrinter in $laserPrinters)
{
    $name = $laserPrinter.PrinterName
    $model = $laserPrinter.Driver
    $type = $laserPrinter.Type
    $ip = $laserPrinter.IP

    Write-Output "Creating printer port IP $ip"
    
    Invoke-Command -ScriptBlock{
    & cscript C:\Windows\System32\Printing_Admin_Scripts\en-US\prnport.vbs -a -r "IP_$ip" -h $ip -o raw
    }

    Write-Output "Installing printer $name"

    Invoke-Command -ScriptBlock{
    & RunDLL32.exe printui.dll,PrintUIEntry /if /b $name /r "IP_$ip" /m $laserModel
    }
    Start-Sleep -s 30
    "$name installed!"
}