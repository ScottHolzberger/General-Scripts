### Import ###

$groups = Import-Csv -Path "C:\Temp\security_groups.csv"

foreach ($group in $groups) {
        New-ADGroup -Name $group.DisplayName `
                    -GroupScope "Global" `
                    -GroupCategory "Security" `
                    -Description $group.Description `
                    -Path "OU=Groups,OU=SilverStone,DC=ssd,DC=local" `
                    -SamAccountName $group.OnPremisesSamAccountName
    }


    ### Export ###
    Get-MgGroup -Filter "securityEnabled eq true" | Select-Object DisplayName, Description, OnPremisesSamAccountName |  Export-Csv -Path "C:\Temp\security_groups.csv" -NoTypeInformation