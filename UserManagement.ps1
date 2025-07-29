### Import ###
$users = Import-Csv -Path "C:\Temp\SSD_Users.csv"
$ACCPassword = (Read-Host -AsSecureString 'AccountPassword')

foreach ($user in $users) {
        New-ADUser  -Name $user.displayName `
                    -Path "OU=Users,OU=SilverStone,DC=ssd,DC=local" `
                    -SamAccountName $user.SAM `
                    -givenName $user.givenName `
                    -surname $user.surname `
                    -userPrincipalName $user.userPrincipalName  `
                    -displayName $user.displayName `
                    -EmailAddress  $user.mail `
                    -Title $user.jobTitle `
                    -OfficePhone $user.telephoneNumber `
                    -mobilePhone $user.mobilePhone `
                    -AccountPassword $ACCPassword `
                    -Enabled $true

    }

    ### Export ###
    get-mguser | Select userPrincipalName, displayName, givenName, surname, mobilePhone, mail, jobTitle, businessPhones | Export-Csv -Path "C:\Temp\users.csv" -NoTypeInformation