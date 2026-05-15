# ==========================================================
# Fix Missing Inventory Schema
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

# Ensure InventoryDate exists
$field = Get-PnPField -List "Tenant Inventory" -Identity "InventoryDate" -ErrorAction SilentlyContinue

if (-not $field) {
    Add-PnPField `
        -List "Tenant Inventory" `
        -DisplayName "Inventory Date" `
        -InternalName "InventoryDate" `
        -Type DateTime `
        -ErrorAction Stop

    Write-Host "✅ Created field: Inventory Date"
}
else {
    Write-Host "Field exists: Inventory Date"
}
