<#
.SYNOPSIS
Automatic installation of Cisco Systems VPN client and the fixes required to make it work with Windows 10.

.DESCRIPTION
The Cisco Systems VPN client does not work properly in Windows 10. A few fixes are required to get it to work.
Kick off this script and the client will be installed along with the required fixes.
#>

# Step 1: Install WinFix #
Write-Host "Installing WinFix..."
Set-Location "\\network-path-to-installation-folder\01 winfix and dneupdates"
Start-Process "winfix.exe" -ArgumentList "/S /noreboot" -Wait
Write-Host "Done!"

# Step 2: Install SonicWall #
Write-Host "Installing SonicWall..."
Set-Location "\\network-path-to-installation-folder\02 for windows10\sonic64"
Start-Process "GVCInstall64.msi" -ArgumentList "/q" -Wait
Write-Host "Done!"

# Step 3: Install VPN Client #
Write-Host "Installing the Cisco VPN Client..."
Set-Location "\\network-path-to-installation-folder\vpnclient-winx64-msi-5.0.07.0290-k9"
Start-Process "vpnclient_setup.msi" -ArgumentList "/q" -Wait
Write-Host "Done!"

# Step 4: Fix VPN Client registry entry; The string before "Cisco Systems VPN Adapter for 64-bit Windows" in the registry causes errors,
# therefore we change the string to the correct value.
Write-Host "Fixing registry DisplayName..."
Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\CVirtA" -Name DisplayName -Value “Cisco Systems VPN Adapter for 64-bit Windows”
Write-Host "Done!"

Write-Host "VPN Client installation is complete!"