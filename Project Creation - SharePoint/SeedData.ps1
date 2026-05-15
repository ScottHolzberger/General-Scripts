# ==========================
# CONNECT
# ==========================
$SiteUrl = "https://zahe.sharepoint.com/sites/ZZ-Control"

#Connect-PnPOnline -Url $SiteUrl -Interactive

# ==========================
# HELPER: Ensure item exists
# ==========================
function Ensure-Item {
    param (
        [string]$List,
        [string]$KeyField,
        [string]$KeyValue,
        [hashtable]$Values
    )

    $existing = Get-PnPListItem -List $List -PageSize 1000 | Where-Object {
        $_.FieldValues[$KeyField] -eq $KeyValue
    }

    if ($existing) {
        Write-Host "Item exists: $List → $KeyValue"
        return $existing
    }

    Write-Host "Creating item: $List → $KeyValue"
    return Add-PnPListItem -List $List -Values $Values
}

# ==========================
# 1. SERVICE CATALOGUE
# ==========================

$halo = Ensure-Item -List "Service Catalogue" -KeyField "ServiceID" -KeyValue "SVC-HALO" -Values @{
    "Title"       = "Halo PSA"
    "ServiceID"   = "SVC-HALO"
    "ServiceName" = "Halo PSA"
    "Category"    = "PSA"
    "Vendor"      = "Halo"
}

$ninja = Ensure-Item -List "Service Catalogue" -KeyField "ServiceID" -KeyValue "SVC-NINJA" -Values @{
    "Title"       = "NinjaOne"
    "ServiceID"   = "SVC-NINJA"
    "ServiceName" = "NinjaOne"
    "Category"    = "RMM"
    "Vendor"      = "NinjaOne"
}

# ==========================
# 2. CLIENT REGISTRY
# ==========================

$client = Ensure-Item -List "Client Registry" -KeyField "ClientID" -KeyValue "CLIENT-001" -Values @{
    "Title"          = "Demo Client"
    "ClientID"       = "CLIENT-001"
    "ClientName"     = "Demo Client Pty Ltd"
    "PrimaryDomain"  = "demo.local"
    "Status"         = "Active"
    "SupportTier"    = "Standard"
}

# ==========================
# 3. CLIENT SERVICE MAPPING
# ==========================

# Get IDs for lookup fields
$clientItem = Get-PnPListItem -List "Client Registry" | Where-Object { $_.FieldValues["ClientID"] -eq "CLIENT-001" }

$haloItem = Get-PnPListItem -List "Service Catalogue" | Where-Object { $_.FieldValues["ServiceID"] -eq "SVC-HALO" }
$ninjaItem = Get-PnPListItem -List "Service Catalogue" | Where-Object { $_.FieldValues["ServiceID"] -eq "SVC-NINJA" }

Ensure-Item -List "Client Service Mapping" -KeyField "MappingID" -KeyValue "MAP-001" -Values @{
    "Title"      = "Demo-Halo"
    "MappingID"  = "MAP-001"
    "Client"     = $clientItem.Id
    "Service"    = $haloItem.Id
    "ExternalID" = "HALO-CLIENT-123"
}

Ensure-Item -List "Client Service Mapping" -KeyField "MappingID" -KeyValue "MAP-002" -Values @{
    "Title"      = "Demo-Ninja"
    "MappingID"  = "MAP-002"
    "Client"     = $clientItem.Id
    "Service"    = $ninjaItem.Id
    "ExternalID" = "NINJA-ORG-456"
}

Write-Host "✅ Seed data creation complete."