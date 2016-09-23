<# Mapped Printer Script created by sphinxpup on Reddit.com/r/PowerShell; Heavily modified and expanded by Shem Gyll #>
$cn = gethostlist
$output = ''
#makes sure RR is not disabled
#Set-Service -Name RemoteRegistry -StartUpType Automatic
#Start-Service -Name RemoteRegistry

foreach($comp in $cn){

$testnet = ''
    #Tries to ping the computer, then stores results in variable $testnet
    try
    {
        $testnet = Test-Connection $comp -Count 2 -ErrorAction Stop
    }
    #If unable to ping the computer, it displays that it is not connected.
    catch
    {
        Write-Host  "$comp not connected`n" -ForegroundColor Red
    }

    #If ping was successful, the script moves on, if not it loops back to the next computer
    if($testnet)
    {
        $s = New-PSSession -ComputerName $comp
        Invoke-Command -Session $s -ScriptBlock{
            #gets list of SIDS to recurse through to get the user
            $sids = Get-ChildItem Registry::HKEY_USERS -Exclude ".Default","*Classes*" | Select-Object Name -ExpandProperty Name
            #gets the users mapped printers by recursing through sids to determine who has mapped printers
    
            foreach($sid in $sids)
                {
                  #Get-ChildItem Registry::$sid\Printers
                  if (Test-Path -Path Registry::$sid\Printers\Connections)  
                    {
                        $goodsid = $sid

                        #trims string
                        $unsid = $sid.Substring(11)
                        #gets username from sid
                        $objSID = New-Object System.Security.Principal.SecurityIdentifier($unsid)
                        $objUser = $objSID.Translate( [System.Security.Principal.NTAccount])

                        #Might want to UN for some reason
                        $un = $objUser.Value

                        #This is a workaround to a bug that stops $_.PSComputerName from being used pretty much anywhere in the output.
                        [string[]]$Computername = $env:COMPUTERNAME

                        #Stores results into $output
                        $output = Get-ChildItem Registry::$goodsid\Printers\Connections

                        #Expresses $output as a table with three columns: Printer Name, Hostname and Username.
                        $output | Format-Table @{Expression={$_.PSChildName};Label="Printer Name";width=40}, @{Expression={$Computername};Label="Hostname";width=16}, @{Expression={$un.ToUpper()};Label="Username";width = 20}
                        
                }
            }
        }
    }
}

$path = Read-Host "Enter path for CSV export"
Write-Host "Please wait..."

$result = foreach($comp in $cn){

$testnet = ''
    #Tries to ping the computer, then stores results in variable $testnet
    try
    {
        $testnet = Test-Connection $comp -Count 2 -ErrorAction Stop
    }
    #If unable to ping the computer, it displays that it is not connected.
    catch
    {
        
    }

    #If ping was successful, the script moves on, if not it loops back to the next computer
    if($testnet)
    {
        $s = New-PSSession -ComputerName $comp
        Invoke-Command -Session $s -ScriptBlock{
            #gets list of SIDS to recurse through to get the user
            $sids = Get-ChildItem Registry::HKEY_USERS -Exclude ".Default","*Classes*" | Select-Object Name -ExpandProperty Name
            #gets the users mapped printers by recursing through sids to determine who has mapped printers
    
            foreach($sid in $sids)
                {
                  #Get-ChildItem Registry::$sid\Printers
                  if (Test-Path -Path Registry::$sid\Printers\Connections)  
                    {
                        $goodsid = $sid

                        #trims string
                        $unsid = $sid.Substring(11)
                        #gets username from sid
                        $objSID = New-Object System.Security.Principal.SecurityIdentifier($unsid)
                        $objUser = $objSID.Translate( [System.Security.Principal.NTAccount])

                        #Might want to UN for some reason
                        $un = $objUser.Value

                        #This is a workaround to a bug that stops $_.PSComputerName from being used pretty much anywhere in the output.
                        [string[]]$Computername = $env:COMPUTERNAME

                        #Stores results into $output
                        $output = Get-ChildItem Registry::$goodsid\Printers\Connections
                        $output | Select-Object PSChildName, PSComputerName
                }
            }
        }
    }
}
#>
$result | Export-Csv -Path $path 

Write-Host "Done!"

