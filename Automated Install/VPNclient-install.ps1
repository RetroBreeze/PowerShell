##     VPN Client installer and fixer     ##

# Step 1: Install WinFix #
Write-Host "Installing WinFix..."
Set-Location "\\mdlfps01\IT\Software\Installations\VPN Clients\vpn_packagev2\01 winfix and dneupdates"
Start-Process "winfix.exe" -ArgumentList "/S /noreboot" -Wait
Write-Host "Done!"

# Step 2: Install SonicWall #
Write-Host "Installing SonicWall..."
Set-Location "\\mdlfps01\IT\Software\Installations\VPN Clients\vpn_packagev2\02 for windows10\sonic64"
Start-Process "GVCInstall64.msi" -ArgumentList "/q" -Wait
Write-Host "Done!"

# Step 3: Install VPN Client #
Write-Host "Installing the Cisco VPN Client..."
Set-Location "\\mdlfps01\IT\Software\Installations\VPN Clients\vpn_packagev2\03 vpn software\vpnclient-winx64-msi-5.0.07.0290-k9"
Start-Process "vpnclient_setup.msi" -ArgumentList "/q" -Wait
Write-Host "Done!"

# Step 4: Fix VPN Client registry entry #
Write-Host "Fixing registry DisplayName..."
Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\CVirtA" -Name DisplayName -Value “Cisco Systems VPN Adapter for 64-bit Windows”
Write-Host "Done!"

Write-Host "VPN Client installation is complete!"