<#
.DESCRIPTION
    Below Powershell script will Check if the "Launch as Seperate Process Registry Value is Enabled or Disabled
#>

$key = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
$value = (Get-ItemProperty -Path $key -Name SeparateProcess).SeparateProcess

if($value -eq 0){
        Write-host "Change Key"
        #Exit 1
}
Else{
        Write-host "Do Not Change Key"
        #Exit 0
}