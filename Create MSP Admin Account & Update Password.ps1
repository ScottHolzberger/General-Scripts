
$checkForUser = (Get-LocalUser).Name -Contains "msp.admin"

if ($checkForUser -eq $true){

Add-Type -AssemblyName System.Web
$Pass = [System.Web.Security.Membership]::GeneratePassword(15,6)
$Password = ConvertTo-SecureString $Pass -AsPlainText -Force

$useraccount = get-localuser -Name "msp.admin"
$useraccount | Set-LocalUser -Password $password

Write-Host "User Exists - Reset Password"

}

ElseIf ($checkForUser -eq $false){

# Generate Password
Add-Type -AssemblyName System.Web
$Pass = [System.Web.Security.Membership]::GeneratePassword(15,6)
$Password = ConvertTo-SecureString $Pass -AsPlainText -Force
$params = @{
    Name        = 'msp.admin'
    Password    = $Password
    FullName    = 'ZaheZone Local Admin'
    Description = 'ZaheZone Override local admin account'
}
New-LocalUser @params

Add-LocalGroupMember -Group "Administrators" -Member "msp.admin"

Write-Host "User Does Not Exist - Account Created"

}





