## Completely automated post-OS-deployment of software and Windows 10 fixes. Based on the New-PC Guided setup.
Import-Module NetSecurity

$Type = Get-WmiObject Win32_Computersystem | Select-object Model
$strPath = ""

$objDomain = New-Object System.DirectoryServices.DirectoryEntry  
$strFilter = "(&(objectCategory=computer)(name=" + $env:computername + "))"

$objSearcher = New-Object System.DirectoryServices.DirectorySearcher  
$objSearcher.SearchRoot = $objDomain  
$objSearcher.Filter = $strFilter

$strPath = $objSearcher.FindOne().Path

function debloat
{
Write-host "Fixing privacy settings..."
. \\mdlfps01\Powershell\Debloat-Windows-10\scripts\fix-privacy-settings.ps1

Write-host "Blocking Telemetry"
. \\mdlfps01\Powershell\Debloat-Windows-10\scripts\block-telemetry.ps1
Write-Host "Snooper no snooping!"

Write-Host "Removing One Drive"
. \\mdlfps01\Powershell\Debloat-Windows-10\scripts\remove-onedrive.ps1
Write-Host "She gone."

Write-Host "Removing bloatware"
. \\mdlfps01\Powershell\Debloat-Windows-10\scripts\remove-default-apps.ps1
Write-Host "It'll probably reappear anyway"

Write-Host "Windows 10 fixes complete!"
}


function misc
{
    clear-host
    Write-Host "Debloat complete!"
    Write-Host "Disabling Windows Firewall..."
    Get-NetFirewallProfile | Set-NetFirewallProfile -Enabled False
    Write-Host "Done!"

    Write-Host "Setting power plan to High Performance..."
    Try {
        $HighPerf = powercfg -l | %{if($_.contains("High performance")) {$_.split()[3]}}
        $CurrPlan = $(powercfg -getactivescheme).split()[3]
        if ($CurrPlan -ne $HighPerf) {powercfg -setactive $HighPerf}
    } Catch {
        Write-Warning -Message "Unable to set power plan to high performance"
    }
    Write-Host "Done!"

    Write-Host "Performing Group Policy update..."
    gpupdate /force
    Write-Host "Done!"

    #Add to WSUS Group
    Invoke-Command -ComputerName mdlfinance -ScriptBlock
    {
    
        if($strPath -match 'Pharma')
        {
            if($strPath -match '1301')
            {
                if($strPath -match 'RecOffice')
                {
                    $targetGroup = '1301 Receiving'
                }

                if($strPath -match 'Office')
                {
                    $targetGroup = '1301 Corporate'
                }

                if($strPath -match 'Ops')
                {
                    $targetGroup = '1301 Ops'
                }
            }
        
            if($strPath -match 'Stanley')
            {
                if($strPath -match 'Office')
                {
                    $targetGroup = '***STANOFFICE***'
                }

                if($strPath -match 'Ops')
                {
                    $targetGroup = '***STANOPS***'
                }

                if($strPath -match 'Retail')
                {
                    $targetGroup = '***STANRETAIL***'
                }

                if($strPath -match 'RecOffice')
                {
                    $targetGroup = '***STANRECOFFICE***'
                }

                if($strPath -match 'SEKO-Office')
                {
                    $targetGroup = '***STANSEKO***'
                }
            }
        }

        if($strPath -match 'Retail')
        {
    
            if($strPath -match 'Office')
                {
                    $targetGroup = '700 Office'
                }

            if($strPath -match 'Ops')
                {
                    $targetGroup = '***700OPS***'
                }
                 
            if($strPath -match 'IT')
                {
                    $targetGroup = 'IT Department'
                }
        }



    Get-WsusComputer -NameIncludes $env:computername | Add-WsusComputer -TargetGroupName $targetGroup
    }

    Write-Host "Performing windows update..."
    wuauclt /detectnow /reportnow /downloadnow
    wuauclt /detectnow /reportnow /downloadnow
    Write-Host "Done!"

    Write-Host "Setting default file associations..."
    dism.exe /online /Import-DefaultAppAssociations:$PSScriptRoot\AppAssociations.xml
    Write-Host "Done!"    

    if($Type -Match "EliteBook")
    {
        laptop
    }
    else
    {
        install
    }
}


function install
{

Write-host "Installing software... You will need to click 'Close' when Ninite is complete."

$strPath = ""

$objDomain = New-Object System.DirectoryServices.DirectoryEntry  
$strFilter = "(&(objectCategory=computer)(name=" + $env:computername + "))"

$objSearcher = New-Object System.DirectoryServices.DirectorySearcher  
$objSearcher.SearchRoot = $objDomain  
$objSearcher.Filter = $strFilter

$strPath = $objSearcher.FindOne().Path

if($strPath -match 'Pharma'){
write-host "pharma"

. \\mdlfps01\Powershell\NewPC-Software.ps1 -pharma 
#complete
}

if($strPath -match 'Retail'){
write-host "retail"
. \\mdlfps01\Powershell\NewPC-Software.ps1 -retail 
#complete
}
}

function complete
{
cls
Write-Host "Installation is complete!"

Write-Host "DON'T FORGET TO ACTIVATE THE MICROSOFT OFFICE LICENSE!" -ForegroundColor Red
Write-Host "Press any key to exit..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
exit
}

function laptop
{
    Write-Host "Installing wireless profiles..."
    Start-Process netsh -ArgumentList 'wlan add profile filename="\\mdlfps01\Powershell\AutomatedInstall\WirelessProfiles\TrainingRoom.xml"' -wait
    Start-Process netsh -ArgumentList 'wlan add profile filename="\\mdlfps01\Powershell\AutomatedInstall\WirelessProfiles\700Office.xml"' -wait
    Start-Process netsh -ArgumentList 'wlan add profile filename="\\mdlfps01\Powershell\AutomatedInstall\StanGuest.xml"' -Wait
    Start-Process netsh -ArgumentList 'wlan add profile filename="\\mdlfps01\Powershell\AutomatedInstall\WirelessProfiles\PharmaConf.xml"' -wait
    Write-Host "Done!"

    Write-Host "Installing VPN Client and fixes..."
    . \\mdlfps01\Powershell\AutomatedInstall\VPNclient-install.ps1

    install
}

debloat -Wait
misc -Wait