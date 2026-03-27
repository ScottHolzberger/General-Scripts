[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$providerName = "NuGet"
if ($null -eq (Get-PackageProvider -Name $providerName -ListAvailable -ErrorAction Ignore)) {
    Write-Host "Installing $providerName package provider..."
    Install-PackageProvider -Name $providerName -Force -ForceBootstrap -Verbose
} else {
    Write-Host "$providerName package provider is already installed."
}

$moduleName = "PSWindowsUpdate"
if (Get-Module -ListAvailable -Name $moduleName) {
    Write-Host "The $moduleName module is installed and available."
} else {
    Write-Host "The $moduleName module is not installed. Installing now..."
    Install-Module -Name PSWindowsUpdate -Force -Confirm:$False
    
}

#Install-Module -Name PSWindowsUpdate -Force -Confirm:$False
Import-Module PSWindowsUpdate
#Get-WindowsUpdate -AcceptAll -Install -AutoReboot
Get-WindowsUpdate -MicrosoftUpdate -AcceptAll -Install -AutoReboot -Verbose