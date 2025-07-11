Set-ExecutionPolicy RemoteSigned -Force
 
Install-Module OSD -Force
 
Import-Module OSD -Force
 
New-OSDCloudTemplate
 
New-OSDCloudWorkspace -WorkspacePath C:\OSDCloud
 
New-OSDCloudUSB
 
Edit-OSDCloudwinPE -workspacepath C:\OSDCloud -CloudDriver * -WebPSScript https://raw.githubusercontent.com/ScottHolzberger/General-Scripts/6a6f04abd01dbed901dd1ca1d669f4f8482809ea/osdcloud_config.ps1 -Verbose
 
New-OSDCloudISO
 
Update-OSDCloudUSB
