
#### **IT-Glue version**


########################## Office 365 ############################
$ApplicationId = 'ab2ca24f-cbdb-434f-a72f-9da0ac0cbde1'
$ApplicationSecret = 'POll/FuFvIN/sQVZ3tA+MEURvwCkBYt7T8ORzrXPVKI=' | Convertto-SecureString -AsPlainText -Force
$TenantID = 'caa9ec21-bcab-4544-af18-bfc520f7d18b'
$RefreshToken = '0.AWcAIeypyqu8REWvGL_FIPfRi0-iLKvby09Dpy-doKwMveFnAB0.AgABAAEAAADnfolhJpSnRYB1SVj-Hgd8AgDs_wUA9P-DZlSv3OrH4EA8OQgewusO8PXBsMdadNWBFxLG5ub5rTvrt6K9H9-lvq2Iu5w3ZHlM2KDT9w3evHoMXvqpiQNWfF_8EHPzLdsudP-xiuHLlR5Sfcs1wa3O_AtkukGWkcOnw7XT-FCdVAiddJ7LxjVboKrmP92szgSBAYJ-weNW_i8mQty9hRo0QAgyLUU8hKm1uLI3cnVszvGoG25KO_nvh9M93xKV4fY4xIzXOAi2QxsCYQMDoF-oPndVAVzM3j9fEPAOmT_u_nld81jba2iTKWJgjFi3FlufnJjK45NLpm_PGXea3CdSayjrgod2BDvw_g_hHp0XEt5mStpiOjERLuGJR9AwNalrHoLaOv2iGvkdx3sjmIdknQzoKYziEmp_LUX8oHRWH7_rw9GE41e8CZcNMUuJTuYdQ3eNa45RydvnCiQNSlfsfbBN5jhzgAjnfLdSu9O-ewvKUd29S2X75ruf84uZqjGWykNm7o7Lv2sRlzP2W024psc3SZt5gUkMnt6RdWxXK2aQZ4hypqmOJH67Tj6aoW0mdOBZOmTaIlXHa5THPbIg-WHZf9KwWrAo7ss5sT02-rwZt9KbFTke_rh6d-q64zmDZkNuYgkgSsASq4JmQmzwcAtcM7zHPlXvV3j03BNIguwJU8pTQDkfJGExiS0zpMBsEU4knCGtlKr2SeG-HMvVUdhvU6azrM0yBVBBWN2bGVZ8BKTJUaIpkxkEtxYgar8PRgZRlPTdiZVJKPagHB8_4d0Se2G_uQqi0FupapwTOrSbdUG7gPg4rf0Ktmiu4JpmIixHQdEhr2dg1WydPa6RrfXP5rpic4_Ak_yYuW7K9zqUz5nCGjoXpGt2hDkirdfBdWY8s3IvxZkq_YsBl4fqbcrN3SAoJI_xcmPxkQ5ofGSDxIwGdek1byKA-33a60sRXxJIhy8xEYWxGe1y46qcxv_W-36DaT8s301UyMzWwnDKVb7p0YJGEUInMPZfG0tQbjfJFKDXfd1xFe3mGGJgoHkqP8tCqecjRdv9kPBiEyvuIck-oHPhqQkOvWZWItFDmoqHK10ULL40-e-SNrf0WoVK811XUQIN'
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
