$computers = gethostlist

foreach ($computer in $computers ) {
$testnet = ''
    #Tries to ping the computer, then stores results in variable $testnet
    try
    {
        $testnet = Test-Connection $computer -Count 2 -ErrorAction Stop
    }
    #If unable to ping the computer, it displays that it is not connected.
    catch
    { Write-Host  "$computer not connected`n" -ForegroundColor Red}

    #If ping was successful, the script moves on, if not it loops back to the next computer
    if($testnet)
    {
    Write-Output "
    Printers on $($computer)
    ------------------------------ "
    gwmi -Class Win32_printer -ComputerName $computer | select Name,Portname
    get-printer -computername $computer  | select Name,Portname
    "`n"
    }

 }

gwmi -Class Win32_printer -ComputerName "mdl/brumple" | select name,type,port
Get-WmiObject -Class Win32_Printer -Computer "mdl354" | Select name, type, port | ft -auto

$doodle = Get-Printer -ComputerName "mdl354" | ? published
$doodle