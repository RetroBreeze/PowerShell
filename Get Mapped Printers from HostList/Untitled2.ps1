$printerPath = "\\mdlfps01\1301_HP_Color_Printer"

Invoke-Command -ComputerName "mdl365" -ScriptBlock{
#(New-Object -ComObject WScript.Network).RemovePrinterConnection("1301 HP Color Printer on mdlfps01", $True, $True)

$net = New-Object -ComObject "Wscript.Network"
$net.EnumPrinterConnections()

#(New-Object -ComObject WScript.Network).AddWindowsPrinterConnection("\\mdlfps01\1301_HP_Color_Printer")
}

Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Connections' | 
ForEach-Object -Process {
    $PSItem |
    Get-ItemProperty -Name Printer |
    Select-Object -ExpandProperty Printer
}

$net = New-Object -ComObject "Wscript.Network"
$net.RemovePrinterConnection("\\mdlfps01\1301_HP_Black_White")
$net.AddWindowsPrinterConnection("\\mdlfps01\1301_HP_Black_White")
$net.AddWindowsPrinterConnection("\\mdlfps01\1301_HP_Color_Printer")
$net.EnumPrinterConnections()

$Mapped = Get-Printer | Select-Object Name,ComputerName,Type,DriverName | Where-Object Type -EQ "Connection"

$Mapped