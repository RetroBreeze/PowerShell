Invoke-Command -ComputerName "mdl365" -ScriptBlock{

printui.exe /dn /n "\\mdlfps01\1301 HP Black White"

}

<#
    $printServ = "mdlfps01"

    $sids = Get-ChildItem Registry::HKEY_USERS -Exclude ".Default","*Classes*" | Select-Object Name -ExpandProperty Name
            #gets the users mapped printers by recursing through sids to determine who has mapped printers
    
            foreach($sid in $sids)
                {
                    if (Test-Path -Path "Registry::$sid\Printers\Connections")  
                        {
                            $goodsid = $sid
                            $regEntry = Get-ChildItem "Registry::$goodsid\Printers\Connections"
                            foreach($key in $regEntry)
                            {
                             if($key.Name -Match "$printServ")
                        {
                            Write-Host "Removing $key"
                            $key | Remove-Item
                        }
                   }
              }

                     if(Test-Path -Path "REGISTRY::HKEY_USERS\$goodsid\SOFTWARE\Microsoft\Windows NT\Current Version\Devices") 
                        {
                            
                            $regEntry = Get-ChildItem "REGISTRY::$goodsid\SOFTWARE\Microsoft\Windows NT\Current Version\Devices"
                            foreach($key in $regEntry)
                            {
                                if($key.Name -Match "$printServ")
                                {
                                    Write-Host "Removing $key"
                                    $key | Remove-Item
                                }
                            }
                        }

                    if(Test-Path -Path "REGISTRY::HKEY_USERS\$goodsid\SOFTWARE\Microsoft\Windows NT\Current Version\PrinterPorts") 
                    {
                            
                        $regEntry = Get-ChildItem "REGISTRY::$goodsid\SOFTWARE\Microsoft\Windows NT\Current Version\PrinterPorts"
                        foreach($key in $regEntry)
                        {
                            if($key.Name -Match "$printServ")
                        {
                            Write-Host "Removing $key"
                            $key | Remove-Item
                        }
                        }
                    }

                    if(Test-Path -Path "REGISTRY::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\Client Side Rendering Print Provider\$goodsid\Printers\Connections") 
                    {
                            
                        $regEntry = Get-ChildItem "REGISTRY::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\Client Side Rendering Print Provider\$goodsid\Printers\Connections"
                        foreach($key in $regEntry)
                        {
                            if($key.Name -Match "$printServ")
                        {
                            Write-Host "Removing $key"
                            $key | Remove-Item
                        }
                        }
                    }
                    
                }

                if(Test-Path -Path "REGISTRY::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\Client Side Rendering Print Provider\Servers") 
                        {
                        
                            $regEntry = Get-ChildItem "REGISTRY::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\Client Side Rendering Print Provider\Servers"
                            foreach($key in $regEntry)
                            {
                            if($key.Name -Match "$printServ")
                            {
                                Write-Host "Removing $key"
                                $key | Remove-Item
                            }
                        }
                        }

                if(Test-Path -Path "REGISTRY::HKEY_CURRENT_USER\Printers\Devices") 
                        {
                        
                            $regEntry = Get-ChildItem "REGISTRY::HKEY_CURRENT_USER\Printers\Devices"
                            foreach($key in $regEntry)
                            {
                            if($key.Name -Match "$printServ")
                            {
                                Write-Host "Removing $key"
                                $key | Remove-Item
                            }
                        }
                        }


                if(Test-Path -Path "REGISTRY::HKEY_CURRENT_USER\Printers\Settings") 
                {
                    $keys =@()
                    $regEntry = Get-ChildItem "REGISTRY::HKEY_CURRENT_USER\Printers\Settings"
                    foreach($key in $regEntry)
                    {
                    if($key.Name -Match "$printServ")
                    {
                        Write-Host "Removing $key"
                        $key | Remove-Item
                    }
                }
                }

                if(Test-Path -Path "REGISTRY::HKEY_CURRENT_USER\Software\Microsoft\Windows NT\Current Version\Devices") 
                {
                    $keys =@()
                    $regEntry = Get-ChildItem "REGISTRY::HKEY_CURRENT_USER\Software\Microsoft\Windows NT\Current Version\Devices"
                    foreach($key in $regEntry)
                    {
                    if($key.Name -Match "$printServ")
                    {
                        Write-Host "Removing $key"
                        $key | Remove-Item
                    }
                    }
                }

                if(Test-Path -Path "HKEY_CURRENT_USER\Software\Microsoft\Windows NT\Current Version\PrinterPorts") 
                {
                    $regEntry = Get-ChildItem "HKEY_CURRENT_USER\Software\Microsoft\Windows NT\Current Version\PrinterPorts"
                    foreach($key in $regEntry)
                    {
                        if($key.Name -Match "$printServ")
                    {
                        Write-Host "Removing $key"
                        $key | Remove-Item
                    }
                    }
                }

}#>