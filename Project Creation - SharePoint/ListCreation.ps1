# ==========================
# CONFIG
# ==========================
$SiteUrl = "https://zahe.sharepoint.com/sites/ZZ-Control"

# ==========================
# CONNECT
# ==========================
Connect-PnPOnline -Url $SiteUrl -ClientId "01e1b71f-cbcb-48df-a076-871aa4ba10d9" -Tenant "zahe.onmicrosoft.com" -CertificatePath "$PSScriptRoot\ZaheZone-PnP-Projects.pfx" -CertificatePassword (ConvertTo-SecureString -AsPlainText "UseA-LongRandomPasswordHere" -Force)

# ==========================
# FUNCTION: Create List If Not Exists
# ==========================
function Ensure-List {
    param (
        [string]$Title,
        [string]$Template = "GenericList"
    )

    $list = Get-PnPList | Where-Object { $_.Title -eq $Title }

    if (-not $list) {
        Write-Host "Creating list: $Title"
        New-PnPList -Title $Title -Template $Template -OnQuickLaunch:$false
    } else {
        Write-Host "List exists: $Title"
    }
}

# ==========================
# FUNCTION: Add Field If Not Exists
# ==========================
function Ensure-Field {
    param (
        [string]$List,
        [string]$DisplayName,
        [string]$InternalName,
        [string]$Type,
        [hashtable]$Options = @{}
    )

    $fields = Get-PnPField -List $List
    if ($fields.InternalName -contains $InternalName) {
        Write-Host "Field exists: $List → $DisplayName"
        return
    }

    Write-Host "Creating field: $List → $DisplayName"

    switch ($Type) {
        "Text" {
            Add-PnPField -List $List -DisplayName $DisplayName -InternalName $InternalName -Type Text
        }
        "Note" {
            Add-PnPField -List $List -DisplayName $DisplayName -InternalName $InternalName -Type Note
        }
        "Choice" {
            Add-PnPField -List $List -DisplayName $DisplayName -InternalName $InternalName -Type Choice -Choices $Options.Choices
        }
        "Lookup" {
            Add-PnPField -List $List -DisplayName $DisplayName -InternalName $InternalName -Type Lookup -LookupList $Options.LookupList -LookupField $Options.LookupField
        }
    }
}

# ==========================
# CREATE LISTS
# ==========================

Ensure-List -Title "Client Registry"
Ensure-List -Title "Service Catalogue"
Ensure-List -Title "Client Service Mapping"
Ensure-List -Title "Client Contacts"
Ensure-List -Title "Client Infrastructure"
Ensure-List -Title "Client Software"

# ==========================
# CLIENT REGISTRY
# ==========================

Ensure-Field -List "Client Registry" -DisplayName "ClientID" -InternalName "ClientID" -Type "Text"
Ensure-Field -List "Client Registry" -DisplayName "ClientName" -InternalName "ClientName" -Type "Text"
Ensure-Field -List "Client Registry" -DisplayName "PrimaryDomain" -InternalName "PrimaryDomain" -Type "Text"
Ensure-Field -List "Client Registry" -DisplayName "Status" -InternalName "Status" -Type "Choice" -Options @{Choices = @("Active","Onboarding","Offboarding")}
Ensure-Field -List "Client Registry" -DisplayName "SupportTier" -InternalName "SupportTier" -Type "Choice" -Options @{Choices = @("Standard","Premium","Enterprise")}
Ensure-Field -List "Client Registry" -DisplayName "Notes" -InternalName "Notes" -Type "Note"

# ==========================
# SERVICE CATALOGUE
# ==========================

Ensure-Field -List "Service Catalogue" -DisplayName "ServiceID" -InternalName "ServiceID" -Type "Text"
Ensure-Field -List "Service Catalogue" -DisplayName "ServiceName" -InternalName "ServiceName" -Type "Text"
Ensure-Field -List "Service Catalogue" -DisplayName "Category" -InternalName "Category" -Type "Choice" -Options @{Choices = @("PSA","RMM","Backup","Security","Other")}
Ensure-Field -List "Service Catalogue" -DisplayName "Vendor" -InternalName "Vendor" -Type "Text"

# ==========================
# CLIENT SERVICE MAPPING
# ==========================

$ClientRegistryId = (Get-PnPList -Identity "Client Registry").Id
$ServiceCatalogueId = (Get-PnPList -Identity "Service Catalogue").Id

Ensure-Field -List "Client Service Mapping" -DisplayName "MappingID" -InternalName "MappingID" -Type "Text"

Ensure-Field -List "Client Service Mapping" -DisplayName "Client" -InternalName "Client" -Type "Lookup" -Options @{
    LookupList = $ClientRegistryId
    LookupField = "ClientName"
}

Ensure-Field -List "Client Service Mapping" -DisplayName "Service" -InternalName "Service" -Type "Lookup" -Options @{
    LookupList = $ServiceCatalogueId
    LookupField = "ServiceName"
}

Ensure-Field -List "Client Service Mapping" -DisplayName "ExternalID" -InternalName "ExternalID" -Type "Text"
Ensure-Field -List "Client Service Mapping" -DisplayName "Metadata" -InternalName "Metadata" -Type "Note"

# ==========================
# CLIENT CONTACTS
# ==========================

Ensure-Field -List "Client Contacts" -DisplayName "ContactID" -InternalName "ContactID" -Type "Text"

Ensure-Field -List "Client Contacts" -DisplayName "Client" -InternalName "Client" -Type "Lookup" -Options @{
    LookupList = $ClientRegistryId
    LookupField = "ClientName"
}

Ensure-Field -List "Client Contacts" -DisplayName "Name" -InternalName "Name" -Type "Text"
Ensure-Field -List "Client Contacts" -DisplayName "Email" -InternalName "Email" -Type "Text"
Ensure-Field -List "Client Contacts" -DisplayName "Role" -InternalName "Role" -Type "Choice" -Options @{Choices = @("Technical","Billing","Decision Maker","Other")}
Ensure-Field -List "Client Contacts" -DisplayName "AuthorityLevel" -InternalName "AuthorityLevel" -Type "Choice" -Options @{Choices = @("Low","Medium","High")}

# ==========================
# CLIENT INFRASTRUCTURE
# ==========================

Ensure-Field -List "Client Infrastructure" -DisplayName "InfrastructureID" -InternalName "InfrastructureID" -Type "Text"

Ensure-Field -List "Client Infrastructure" -DisplayName "Client" -InternalName "Client" -Type "Lookup" -Options @{
    LookupList = $ClientRegistryId
    LookupField = "ClientName"
}

Ensure-Field -List "Client Infrastructure" -DisplayName "Type" -InternalName "Type" -Type "Choice" -Options @{Choices = @("Firewall","Switch","AP","Server","Network")}
Ensure-Field -List "Client Infrastructure" -DisplayName "Name" -InternalName "Name" -Type "Text"
Ensure-Field -List "Client Infrastructure" -DisplayName "Details" -InternalName "Details" -Type "Note"

# ==========================
# CLIENT SOFTWARE
# ==========================

Ensure-Field -List "Client Software" -DisplayName "SoftwareID" -InternalName "SoftwareID" -Type "Text"

Ensure-Field -List "Client Software" -DisplayName "Client" -InternalName "Client" -Type "Lookup" -Options @{
    LookupList = $ClientRegistryId
    LookupField = "ClientName"
}

Ensure-Field -List "Client Software" -DisplayName "Name" -InternalName "Name" -Type "Text"
Ensure-Field -List "Client Software" -DisplayName "Vendor" -InternalName "Vendor" -Type "Text"
Ensure-Field -List "Client Software" -DisplayName "SupportContact" -InternalName "SupportContact" -Type "Text"
Ensure-Field -List "Client Software" -DisplayName "Notes" -InternalName "Notes" -Type "Note"

Write-Host "✅ All lists and fields created successfully."