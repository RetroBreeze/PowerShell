$computers = Get-ADComputer -Filter 'OperatingSystem -like "*xp*"' -Properties name, operatingsystem, operatingsystemversion | select name, operatingsystem, operatingsystemversion
$outputObj =@()
$offline =@()

foreach ($comp in $computers){

$testnet = ''

    #Tries to ping the computer, then stores results in variable $testnet
    try
    {
        $testnet = Test-Connection $comp.name -Count 2 -ErrorAction Stop
    }
    #If unable to ping the computer, it displays that it is not connected.
    catch
    {
    $offline += $comp
    }

    #If ping was successful, the script moves on, if not it loops back to the next computer
    if($testnet)
    {
    $comp
    $outputObj += $comp
    }
    }

    $outputObj
    $offline