# ==========================================================
# Add Indexes – ZaheZone Client Tenancy Register
# ==========================================================

# AUTH
[string]$SiteUrl  = "https://zahe.sharepoint.com/sites/ZaheZoneClientTenancyRegister"
[string]$Tenant   = "zahe.onmicrosoft.com"
[string]$ClientId = "01e1b71f-cbcb-48df-a076-871aa4ba10d9"
[string]$PfxPath  = "$PSScriptRoot\ZaheZone-PnP-Projects.pfx"
[string]$PfxPasswordPlain = "UseA-LongRandomPasswordHere"

Connect-PnPOnline `
    -Url $SiteUrl `
    -Tenant $Tenant `
    -ClientId $ClientId `
    -CertificatePath $PfxPath `
    -CertificatePassword (ConvertTo-SecureString $PfxPasswordPlain -AsPlainText -Force)

Write-Host "✅ Connected"

# -------------------------------
# Helper: Ensure Index
# -------------------------------
function Ensure-Index {
    param(
        [string]$List,
        [string]$FieldInternalName
    )

    $field = Get-PnPField -List $List -Identity $FieldInternalName -ErrorAction SilentlyContinue

    if (-not $field) {
        Write-Host "⚠️ Field not found: $FieldInternalName (skipping)"
        return
    }

    if ($field.Indexed -eq $true) {
        Write-Host "Index exists: $FieldInternalName"
        return
    }

    Set-PnPField `
        -List $List `
        -Identity $FieldInternalName `
        -Values @{ Indexed = $true }

    Write-Host "✅ Indexed: $FieldInternalName"
}

# ==========================================================
# RECOMMENDATIONS REGISTER (MOST IMPORTANT)
# ==========================================================

# Filters + reporting
Ensure-Index "Recommendations Register" "Severity"
Ensure-Index "Recommendations Register" "RecommendationStatus"
Ensure-Index "Recommendations Register" "ServiceTier"
Ensure-Index "Recommendations Register" "AutoApplyEligible"

# Lookup (VERY important for performance)
Ensure-Index "Recommendations Register" "Client_x0020_Tenancy"
Ensure-Index "Recommendations Register" "Inventory_x0020_Reference"

# ==========================================================
# TENANT INVENTORY
# ==========================================================

Ensure-Index "Tenant Inventory" "Client_x0020_Tenancy"
Ensure-Index "Tenant Inventory" "SecurityDefaultsEnabled"
Ensure-Index "Tenant Inventory" "SSPREnabled"
Ensure-Index "Tenant Inventory" "DKIMStatus"

# ==========================================================
# CLIENT TENANCIES
# ==========================================================

Ensure-Index "Client Tenancies" "O365Status"
Ensure-Index "Client Tenancies" "LicensingTier"
Ensure-Index "Client Tenancies" "PrimaryDomain"

Write-Host "✅ All indexes applied successfully."