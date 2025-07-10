
#### **IT-Glue version**


########################## Office 365 ############################
$ApplicationId = '8e3d4e1e-49c5-4597-856b-28887a130208'
$ApplicationSecret = 'l5wT+lafIBEUdpbjlZ7/Hl6bjrPjkDAWDOy60wTkKF4=' | Convertto-SecureString -AsPlainText -Force
$TenantID = 'caa9ec21-bcab-4544-af18-bfc520f7d18b'
$RefreshToken = '0.AWcAIeypyqu8REWvGL_FIPfRix5OPY7FSZdFhWsoiHoTAghnAB0.AgABAAEAAADnfolhJpSnRYB1SVj-Hgd8AgDs_wUA9P-mX8GjydgdrdWVzONV-UvIEFUAf2Rf_wGVYgwTaI6IaV1DIfxeZLCO38WUqyycOzVOLfdv8juDwDzc0pRcYeklGz-jwIag4EMCW8iJSuIulYqB9LnWHUpmzWFn5sk6ZXKbsaK-bj8knnNzmqnh9ydB1JSKC8tyUfjBhmyhlJUfkISu3fmyFouS8AHR6VmqMgwWcBzuHsdt7AZ7o5PJuyxWLxEZ8arhSSwUivbDDQBlR8zRYcRLoDdAj3gjM7L594hBmYI_D3cFCMIHNYoj4aNbwC-3SwfByRnEjjRuKaE2go-MYxrP7Ns3KNC4i18lD6AkBu2NQ2AeoiZFf3dLDdsRPoltkXkOZMncJQqLuUd84LhuzwQ9WsN8OFTzKvnwbiE-EYhwzdq3Smndtx2gFhkcpAb20lpeU4zez7t374bljUi1LFmImenZZ6M6Z_14OuKG4BxgHM6p5Tznr3IQ0H0-rZGx_JJwFCz9odzx7PJ-x9YT_qJxDTpHbCqF1T0HV_mt91KmnES5LIWDer8n7duIofKZKME3WeMFX_FbDOLwLFewP0BAICYNpFE7y0ubsDyb53JdgqkSfxvQrVY9EiF29BKAvwGaRdTHbpD-t4Ag2NXUMhfj3_OKHITCS4HEp2yyqUr59yrUQFYBNuiJIuC5zIRVv6zPivWtSxZfe0mLVEHff71Yku5BktjfSbf8xCJ3RyAYcgmk_bswstSRMAOUXL3gtGpnIp6DtVpPY21EGdC5H0VuTkW_a4AiPlJuCtpsur4HDxshNUQBtZ7IaGKtj9XcjFYv72iMc1fJU2-CuVch_eMNJQhnTx-SifUEQDJj-R2dICGuj7EfadAvlV6TwamVaFzapWEn_ka9LrIadT6DJIMfA7p2B-RvWvcCP1jdprMo8zqf-IYkfptOU8xpf8wIZOqMKeepscxhYBj5MjWeRZ9x1_lHbephtvVSpbRYVoNMTiBXDfgJbwhH2GHQ0QJDM8SzclUNKdssCnjnM_dX5-sTgIAZncFq-hoxs62Cd2SalV8SxnRdVE5y3FmirWYI741UhvTU8E1DJgPTPPFI8cyLtJUB8QOjyFywhcG0DOfxcoKP26ZTf7NmYc4PGRXJNsCHlHGDU0zdV7u8tqosny5qIE4STDXSSAHoFXADvxEg'
$upn = 'scott@zahezone.com.au'
########################## IT-Glue ############################
$APIKEy = "ITG.5546a634c361bdd7163661d6fcced4ef.mI4rXC2zKwtj61fkejumkWQpp3vH6gSTkkzizzwha80-BkOrK3NjQy5DH7dG_TYf"
$APIEndpoint = "https://api.au.itglue.com"
$FlexAssetName = "Office365 Reports - AutoDoc v1"
$Description = "Office365 Reporting."
#some layout options, change if you want colours to be different or do not like the whitespace.
$TableHeader = "<table class=`"table table-bordered table-hover`" style=`"width:80%`">"
$TableStyling = "<th>", "<th style=`"background-color:#4CAF50`">"
###########################
#Grabbing ITGlue Module and installing.
If (Get-Module -ListAvailable -Name "ITGlueAPI") {
    Import-module ITGlueAPI
}
Else {
    Install-Module ITGlueAPI -Force
    Import-Module ITGlueAPI
}

#Settings IT-Glue logon information
Add-ITGlueBaseURI -base_uri $APIEndpoint
Add-ITGlueAPIKey $APIKEy
write-host "Checking if Flexible Asset exists in IT-Glue." -foregroundColor green
$FilterID = (Get-ITGlueFlexibleAssetTypes -filter_name $FlexAssetName).data
if (!$FilterID) {
    write-host "Does not exist, creating new." -foregroundColor green
    $NewFlexAssetData =
    @{
        type          = 'flexible-asset-types'
        attributes    = @{
            name        = $FlexAssetName
            icon        = 'sitemap'
            description = $description
        }
        relationships = @{
            "flexible-asset-fields" = @{
                data = @(
                    @{
                        type       = "flexible_asset_fields"
                        attributes = @{
                            order           = 1
                            name            = "Teams Device Reports"
                            kind            = "Textbox"
                            required        = $true
                            "show-in-list"  = $true
                            "use-for-title" = $true
                        }
                    },
                    @{
                        type       = "flexible_asset_fields"
                        attributes = @{
                            order          = 2
                            name           = "Teams User Reports"
                            kind           = "Textbox"
                            required       = $false
                            "show-in-list" = $false
                        }
                    },
                    @{
                        type       = "flexible_asset_fields"
                        attributes = @{
                            order          = 3
                            name           = "Email Reports"
                            kind           = "Textbox"
                            required       = $false
                            "show-in-list" = $false
                        }
                    },
                    @{
                        type       = "flexible_asset_fields"
                        attributes = @{
                            order          = 4
                            name           = "Mailbox Usage Reports"
                            kind           = "Textbox"
                            required       = $false
                            "show-in-list" = $false
                        }
                    },
                    @{
                        type       = "flexible_asset_fields"
                        attributes = @{
                            order          = 5
                            name           = "O365 Activations Reports"
                            kind           = "Textbox"
                            required       = $false
                            "show-in-list" = $false
                        }
                    },
                    @{
                        type       = "flexible_asset_fields"
                        attributes = @{
                            order          = 6
                            name           = "OneDrive Activity Reports"
                            kind           = "Textbox"
                            required       = $false
                            "show-in-list" = $false
                        }
                    },
                    @{
                        type       = "flexible_asset_fields"
                        attributes = @{
                            order          = 7
                            name           = "OneDrive Usage Reports"
                            kind           = "Textbox"
                            required       = $false
                            "show-in-list" = $false
                        }
                    },
                    @{
                        type       = "flexible_asset_fields"
                        attributes = @{
                            order          = 8
                            name           = "Sharepoint Usage Reports"
                            kind           = "Textbox"
                            required       = $false
                            "show-in-list" = $false
                        }
                    },
                    @{
                        type       = "flexible_asset_fields"
                        attributes = @{
                            order          = 9
                            name           = "TenantID"
                            kind           = "Text"
                            required       = $false
                            "show-in-list" = $false
                        }
                    }
                )
            }
        }
    }
    New-ITGlueFlexibleAssetTypes -Data $NewFlexAssetData
    $FilterID = (Get-ITGlueFlexibleAssetTypes -filter_name $FlexAssetName).data
}
$AllITGlueContacts = @()
#Grab all IT-Glue contacts to match the domain name.
write-host "Getting IT-Glue contact list" -foregroundColor green
$i = 0
do {
    $AllITGlueContacts += (Get-ITGlueContacts -page_size 1000 -page_number $i).data.attributes
    $i++
    Write-Host "Retrieved $($AllITGlueContacts.count) Contacts" -ForegroundColor Yellow
}while ($AllITGlueContacts.count % 1000 -eq 0 -and $AllITGlueContacts.count -ne 0)


write-host "Start documentation process." -foregroundColor green
$credential = New-Object System.Management.Automation.PSCredential($ApplicationId, $ApplicationSecret)
write-host "Generating access tokens" -ForegroundColor Green
$aadGraphToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes 'https://graph.windows.net/.default' -ServicePrincipal -Tenant $tenantID

$graphToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes 'https://graph.microsoft.com/.default' -ServicePrincipal -Tenant $tenantID
write-host "Connecting to MSOLService" -ForegroundColor Green
Connect-MsolService -AdGraphAccessToken $aadGraphToken.AccessToken -MsGraphAccessToken $graphToken.AccessToken
write-host "Grabbing client list" -ForegroundColor Green
$customers = Get-MsolPartnerContract -All
write-host "Connecting to clients" -ForegroundColor Green

foreach ($customer in $customers) {
    write-host "Generating token for $($Customer.name)" -ForegroundColor Green
    $graphToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes 'https://graph.microsoft.com/.default' -ServicePrincipal -Tenant $customer.TenantID
    $Header = @{
        Authorization = "Bearer $($graphToken.AccessToken)"
    }
    write-host "Gathering Reports for $($Customer.name)" -ForegroundColor Green
    #Gathers which devices currently use Teams, and the details for these devices.
    $TeamsDeviceReportsURI = "https://graph.microsoft.com/v1.0/reports/getTeamsDeviceUsageUserDetail(period='D7')"
    $TeamsDeviceReports = (Invoke-RestMethod -Uri $TeamsDeviceReportsURI -Headers $Header -Method Get -ContentType "application/json") -replace "ï»¿", "" | ConvertFrom-Csv | ConvertTo-Html -fragment -PreContent "<h1>Teams device report</h1>" | Out-String
    #Gathers which Users currently use Teams, and the details for these Users.
    $TeamsUserReportsURI = "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserDetail(period='D7')"
    $TeamsUserReports = (Invoke-RestMethod -Uri $TeamsUserReportsURI -Headers $Header -Method Get -ContentType "application/json") -replace "ï»¿", "" | ConvertFrom-Csv | ConvertTo-Html -fragment -PreContent "<h1>Teams user report</h1>" | Out-String
    #Gathers which users currently use email and the details for these Users
    $EmailReportsURI = "https://graph.microsoft.com/v1.0/reports/getEmailActivityUserDetail(period='D7')"
    $EmailReports = (Invoke-RestMethod -Uri $EmailReportsURI -Headers $Header -Method Get -ContentType "application/json") -replace "ï»¿", "" | ConvertFrom-Csv | ConvertTo-Html -fragment -PreContent "<h1>Email users Report</h1>" | Out-String
    #Gathers the storage used for each e-mail user.
    $MailboxUsageReportsURI = "https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='D7')"
    $MailboxUsage = (Invoke-RestMethod -Uri $MailboxUsageReportsURI -Headers $Header -Method Get -ContentType "application/json") -replace "ï»¿", "" | ConvertFrom-Csv | ConvertTo-Html -fragment -PreContent "<h1>Email storage report</h1>" | Out-String
    #Gathers the activations for each user of office.
    $O365ActivationsReportsURI = "https://graph.microsoft.com/v1.0/reports/getOffice365ActivationsUserDetail"
    $O365ActivationsReports = (Invoke-RestMethod -Uri $O365ActivationsReportsURI -Headers $Header -Method Get -ContentType "application/json") -replace "ï»¿", "" | ConvertFrom-Csv | ConvertTo-Html -fragment -PreContent "<h1>O365 Activation report</h1>" | Out-String
    #Gathers the Onedrive activity for each user.
    $OneDriveActivityURI = "https://graph.microsoft.com/v1.0/reports/getOneDriveActivityUserDetail(period='D7')"
    $OneDriveActivityReports = (Invoke-RestMethod -Uri $OneDriveActivityURI -Headers $Header -Method Get -ContentType "application/json") -replace "ï»¿", "" | ConvertFrom-Csv | ConvertTo-Html -fragment -PreContent "<h1>Onedrive Activity report</h1>" | Out-String
    #Gathers the Onedrive usage for each user.
    $OneDriveUsageURI = "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='D7')"
    $OneDriveUsageReports = (Invoke-RestMethod -Uri $OneDriveUsageURI -Headers $Header -Method Get -ContentType "application/json") -replace "ï»¿", "" | ConvertFrom-Csv | ConvertTo-Html -fragment -PreContent "<h1>OneDrive usage report</h1>" | Out-String
    #Gathers the Sharepoint usage for each user.
    $SharepointUsageReportsURI = "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='D7')"
    $SharepointUsageReports = (Invoke-RestMethod -Uri $SharepointUsageReportsURI -Headers $Header -Method Get -ContentType "application/json") -replace "ï»¿", "" | ConvertFrom-Csv | ConvertTo-Html -fragment -PreContent "<h1>Sharepoint usage report</h1>" | Out-String

    $FlexAssetBody =
    @{
        type       = 'flexible-assets'
        attributes = @{
            traits = @{
                'teams-device-reports'      = ($TableHeader + $TeamsDeviceReports) -replace $TableStyling
                'teams-user-reports'        = ($TableHeader + $TeamsUserReports ) -replace $TableStyling
                'email-reports'             = ($TableHeader + $EmailReports) -replace $TableStyling
                'mailbox-usage-reports'     = ($TableHeader + $MailboxUsage) -replace $TableStyling
                'o365-activations-reports'  = ($TableHeader + $O365ActivationsReports) -replace $TableStyling
                'onedrive-activity-reports' = ($TableHeader + $OneDriveActivityReports) -replace $TableStyling
                'onedrive-usage-reports'    = ($TableHeader + $OneDriveUsageReports) -replace $TableStyling
                'sharepoint-usage-reports'  = ($TableHeader + $SharepointUsageReports) -replace $TableStyling
                'tenantid'                  = $customer.TenantId
            }
        }
    }

    Write-Host "          Finding $($customer.name) in IT-Glue" -ForegroundColor Green
    $orgID = @()
    $customerdomains = Get-MsolDomain -TenantId $customer.tenantid
    foreach ($customerDomain in $customerdomains) {
        $orgID += ($AllITGlueContacts | Where-Object { $_.'contact-emails'.value -match $customerDomain.name }).'organization-id' | Select-Object -Unique
    }
    write-host "             Uploading Reports $($customer.name) into IT-Glue"  -ForegroundColor Green
    foreach ($org in $orgID) {
        $ExistingFlexAsset = (Get-ITGlueFlexibleAssets -filter_flexible_asset_type_id $($filterID.ID) -filter_organization_id $org).data | Where-Object { $_.attributes.traits.'tenantid' -eq $customer.TenantId }
        #If the Asset does not exist, we edit the body to be in the form of a new asset, if not, we just upload.
        if (!$ExistingFlexAsset) {
            $FlexAssetBody.attributes.add('organization-id', $org)
            $FlexAssetBody.attributes.add('flexible-asset-type-id', $($filterID.ID))
            write-host "                      Creating Reports $($customer.name) into IT-Glue organisation $org" -ForegroundColor Green
            New-ITGlueFlexibleAssets -data $FlexAssetBody
            start-sleep 2
        }
        else {
            write-host "                      Updating Reports $($customer.name) into IT-Glue organisation $org"  -ForegroundColor Green
            $ExistingFlexAsset = $ExistingFlexAsset | select-object -last 1
            Set-ITGlueFlexibleAssets -id $ExistingFlexAsset.id -data $FlexAssetBody
            start-sleep 2
        }

    }



}
