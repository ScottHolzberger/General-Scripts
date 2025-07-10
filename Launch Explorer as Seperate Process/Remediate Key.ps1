$key = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
Set-ItemProperty -Path $key -Name SeparateProcess -Value 1


$key = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
$value = (Get-ItemProperty -Path $key -Name SeparateProcess).SeparateProcess

if($value -eq 1){
        Write-host "Key Successfully Changes"
        Exit 0
}
Else{
        Write-host "Key Not changed"
        Exit 1
}