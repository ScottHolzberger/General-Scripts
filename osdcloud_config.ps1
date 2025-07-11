Write-Host -ForegroundColor Green "Starting OSDCloud ZTI"
Start-Sleep -Seconds 5

Start-OSDCloud -OSName 'Windows 11 23H2 x64' -OSActivation Retail -OSLanguage en-us -OSEdition Pro

#Restart from WinPE

Write-Host -ForegroundColor Green “Restarting in 20 seconds!”

Start-Sleep -Seconds 20

wpeutil reboot
