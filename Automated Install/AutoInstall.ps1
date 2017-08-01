<#
.SYNOPSIS
Post-OS-installation automatic software deployment and system configuration

.DESCRIPTION
With Windows 10 Pro, we are unable to distribute a "golden image" of Windows 10 that contains all our software and configurations.
This script is a good "between" solution. After the OS has been installed, kick off AddToDomain.bat to add the PC to domain, then
this script via the AutoInstall.bat file to automatically handle software installation and system configuration.
#>


Import-Module NetSecurity

# Finds which domain and OU the computer is currently in
$Type = Get-WmiObject Win32_Computersystem | Select-object Model
$strPath = ""
$objDomain = New-Object System.DirectoryServices.DirectoryEntry  
$strFilter = "(&(objectCategory=computer)(name=" + $env:computername + "))"
$objSearcher = New-Object System.DirectoryServices.DirectorySearcher  
$objSearcher.SearchRoot = $objDomain  
$objSearcher.Filter = $strFilter
$strPath = $objSearcher.FindOne().Path

#Calls on W4RH4WK's Debloat-Windows-10 scripts (https://github.com/W4RH4WK/Debloat-Windows-10) to remove bloatware and other Windows 10 Quirks 
function debloat
{
Write-host "Fixing privacy settings..."
. \\networkshare\Powershell\Debloat-Windows-10\scripts\fix-privacy-settings.ps1
Write-Host "Done!"

Write-host "Blocking Telemetry"
. \\networkshare\Powershell\Debloat-Windows-10\scripts\block-telemetry.ps1
Write-Host "Done!"

Write-Host "Removing One Drive"
. \\networkshare\Powershell\Debloat-Windows-10\scripts\remove-onedrive.ps1
Write-Host "Done!"

Write-Host "Removing bloatware"
. \\networkshare\Powershell\Debloat-Windows-10\scripts\remove-default-apps.ps1
Write-Host "Done! It'll probably reappear anyway"

Write-Host "Windows 10 fixes complete!"
}

#Performs miscellaneous system configurations
function misc
{
    clear-host
    #disables Windows firewall
    Write-Host "Disabling Windows Firewall..."
    Get-NetFirewallProfile | Set-NetFirewallProfile -Enabled False
    Write-Host "Done!"

    #Sets power plan to High Performance
    Write-Host "Setting power plan to High Performance..."
    Try {
        $HighPerf = powercfg -l | %{if($_.contains("High performance")) {$_.split()[3]}}
        $CurrPlan = $(powercfg -getactivescheme).split()[3]
        if ($CurrPlan -ne $HighPerf) {powercfg -setactive $HighPerf}
    } Catch {
        Write-Warning -Message "Unable to set power plan to high performance"
    }
    Write-Host "Done!"

    #Performs Group Policy Update
    Write-Host "Performing Group Policy update..."
    gpupdate /force
    Write-Host "Done!"


    #Add computer to WSUS Group (quick and dirty way to do this, needs updating for sure) [actual OU names obfuscated]
    Invoke-Command -ComputerName updateserver -ScriptBlock
    {
        #Match the name of $strPath (the computer's OU path) to the corresponding WSUS group and set $targetGroup to that string
        if($strPath -match 'Pharma')
        {
            if($strPath -match 'Location1')
            {
                if($strPath -match 'Department1')
                {
                    $targetGroup = 'Department 1 Group'
                }

                if($strPath -match 'Department2')
                {
                    $targetGroup = 'Department 2 Group'
                }

                if($strPath -match 'Department3')
                {
                    $targetGroup = 'Department 3 Group'
                }
            }
        

        if($strPath -match 'Retail')
        {
            if($strPath -match 'Location2')
            {
    
                if($strPath -match 'Department1')
                    {
                        $targetGroup = 'Department 1 Group'
                    }

                if($strPath -match 'Department2')
                    {
                        $targetGroup = 'Department 2 Group'
                    }
                 
                if($strPath -match 'Department3')
                    {
                        $targetGroup = 'Department 3 Group'
                    }

            }
        }

    #Get the computer in WSUS and add it to the group that matches $targetGroup
    Get-WsusComputer -NameIncludes $env:computername | Add-WsusComputer -TargetGroupName $targetGroup
    }

    #Perform windows update (a colleague is certain you need to run this twice to make it work, so why not.)
    Write-Host "Performing windows update..."
    wuauclt /detectnow /reportnow /downloadnow
    wuauclt /detectnow /reportnow /downloadnow
    Write-Host "Done!"

    #Uses AppAssociations.xml to set some important file associations (mostly setting the PDF reader to Foxit, instead of Edge, which has caused us a lot
    # of problems in the past with other software.
    Write-Host "Setting default file associations..."
    dism.exe /online /Import-DefaultAppAssociations:$PSScriptRoot\AppAssociations.xml
    Write-Host "Done!"    

    #If the model is a laptop (we use HP, so we look for the "EliteBook" string), then skip to the laptop function, else do the install function
    if($Type -Match "EliteBook")
    {
        laptop
    }
    else
    {
        install
    }
}

#Installs our software packages
function install
{

Write-host "Installing software... You will need to click 'Close' when Ninite is complete."

#If the computer is in Pharma, install the pharma package using the -pharma argument on NewPC-Software.ps1 (another script on our network)
if($strPath -match 'Pharma'){
write-host "pharma"

. \\networkpath\Powershell\NewPC-Software.ps1 -pharma 
#complete
}

#If the computer is in Retail, install the retail package using the -retail argument on NewPC-Software.ps1 (another script on our network)
if($strPath -match 'Retail'){
write-host "retail"
. \\networkpath\Powershell\NewPC-Software.ps1 -retail 
Write-Host "Installation is complete!"
}
}

#Clears host, writes reminder about office license, then waits for user input to exit.
function complete
{
cls
Write-Host "DON'T FORGET TO ACTIVATE THE MICROSOFT OFFICE LICENSE!" -ForegroundColor Red
Write-Host "Press any key to exit..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
exit
}

#Installs wifi profiles from network folder using netsh
function laptop
{
    Write-Host "Installing wireless profiles..."
    Start-Process netsh -ArgumentList 'wlan add profile filename="\\Networkpath\WirelessProfiles\Profile1.xml"' -wait
    Start-Process netsh -ArgumentList 'wlan add profile filename="\\Networkpath\WirelessProfiles\Profile2.xml"' -wait
    Start-Process netsh -ArgumentList 'wlan add profile filename="\\Networkpath\WirelessProfiles\Profile2.xml"' -Wait
    Start-Process netsh -ArgumentList 'wlan add profile filename="\\Networkpath\WirelessProfiles\Profile4.xml"' -wait
    Start-Process netsh -ArgumentList 'wlan add profile filename="\\Networkpath\WirelessProfiles\Profile5.xml"' -wait

    Write-Host "Done!"
    
    #Installs Cisco VPN Client and the required fixes
    Write-Host "Installing VPN Client and fixes..."
    . \\mdlfps01\Powershell\AutomatedInstall\VPNclient-install.ps1

    install
}

#script actually starts here by kicking off debloat then moving on to misc.
debloat -Wait
misc -Wait