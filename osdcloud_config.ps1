Write-Host -ForegroundColor Green "Starting OSDCloud ZTI"
Start-Sleep -Seconds 5

Start-OSDCloud 

#Restart from WinPE

Write-Host -ForegroundColor Green “Restarting in 20 seconds!”

Start-Sleep -Seconds 20

wpeutil reboot
