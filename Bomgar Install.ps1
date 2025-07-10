
#Start-Transcript -Path "C:\temp\BomgarInstall.txt"


$app = Get-Package -Name "BeyondTrust Remote Support Jump Client zahezone.beyondtrustcloud.com"

#Check is application is installed
if ($app)
{
    Write-Host "Application is installed"
}
else
{
    #Download BeyondTrust Installer
    If ([Environment]::Is64BitOperatingSystem){
        $URL = "https://zahezone.beyondtrustcloud.com/files/bomgar-scc-win64.msi"
        $Path= "C:\Temp\bomgar-scc-win64.msi"
        }
    Else {
        $URL = "https://zahezone.beyondtrustcloud.com/files/bomgar-scc-win32.msi"
        $Path= "C:\Temp\bomgar-scc-win32.msi"
        }
    
    If (Test-Path -Path "C:\Temp"){
        #Write-Host "Folder Exists"
        }
    Else {
        MD C:\Temp
        }
    
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -URI $URL -OutFile $Path
    
    #Install Bomgar
    Start-Process -FilePath "$env:systemroot\system32\msiexec.exe" -ArgumentList "/i C:\Temp\bomgar-scc-win64.msi KEY_INFO='w0eec30e1y8h15jihhg7yfe86g7z8hyhj7hiif7c40hc90' installdir=C:\ZZBomgar /qn"
    
Stop-Transcript
    
}